#!/usr/bin/env node
// Synchronise dist/ vers le bucket Cellar via le SDK AWS S3 v3.
// Variables d'environnement requises (charger via .env.local + `source .env.local` ou export) :
//   CELLAR_KEY_ID
//   CELLAR_SECRET
//   CELLAR_BUCKET     (ex: sheet-navigator)
//   CELLAR_ENDPOINT   (ex: https://cellar-c2.services.clever-cloud.com)
//
// Chaque objet est pousse avec ACL public-read (Cellar n'expose pas un bucket entier
// en public par defaut, il faut le marquer objet par objet). Le HTML recoit un
// Cache-Control court (5 min), les autres assets sont consideres immutables (1 an).

import { readdir, readFile, stat } from "node:fs/promises";
import { join, relative, extname, sep } from "node:path";
import { fileURLToPath } from "node:url";
import {
  S3Client,
  PutObjectCommand,
  ListObjectsV2Command,
  DeleteObjectsCommand,
} from "@aws-sdk/client-s3";

const REQUIRED = ["CELLAR_KEY_ID", "CELLAR_SECRET", "CELLAR_BUCKET", "CELLAR_ENDPOINT"];
for (const k of REQUIRED) {
  if (!process.env[k]) {
    console.error(`Variable d'environnement manquante : ${k}`);
    process.exit(1);
  }
}

const bucket = process.env.CELLAR_BUCKET;
const endpoint = process.env.CELLAR_ENDPOINT;

const client = new S3Client({
  endpoint,
  region: "us-east-1",
  credentials: {
    accessKeyId: process.env.CELLAR_KEY_ID,
    secretAccessKey: process.env.CELLAR_SECRET,
  },
  forcePathStyle: false,
});

const CONTENT_TYPES = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".mjs": "application/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".xml": "application/xml; charset=utf-8",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".gif": "image/gif",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
  ".map": "application/json; charset=utf-8",
  ".txt": "text/plain; charset=utf-8",
};

function contentTypeFor(path) {
  return CONTENT_TYPES[extname(path).toLowerCase()] ?? "application/octet-stream";
}

function cacheControlFor(path) {
  // HTML : court (5 min) pour pouvoir iterer sans bloquer le cache navigateur.
  // Reste : hashes dans les noms de fichier, donc immutable un an.
  return extname(path).toLowerCase() === ".html"
    ? "public, max-age=300"
    : "public, max-age=31536000, immutable";
}

async function* walk(dir) {
  for (const entry of await readdir(dir, { withFileTypes: true })) {
    const full = join(dir, entry.name);
    if (entry.isDirectory()) {
      yield* walk(full);
    } else if (entry.isFile()) {
      yield full;
    }
  }
}

async function uploadOne(localPath, key) {
  const body = await readFile(localPath);
  await client.send(
    new PutObjectCommand({
      Bucket: bucket,
      Key: key,
      Body: body,
      ACL: "public-read",
      ContentType: contentTypeFor(localPath),
      CacheControl: cacheControlFor(localPath),
    })
  );
  console.log(`  -> s3://${bucket}/${key}`);
}

async function listAllKeys() {
  const keys = [];
  let token;
  do {
    const out = await client.send(
      new ListObjectsV2Command({ Bucket: bucket, ContinuationToken: token })
    );
    for (const obj of out.Contents ?? []) {
      if (obj.Key) keys.push(obj.Key);
    }
    token = out.IsTruncated ? out.NextContinuationToken : undefined;
  } while (token);
  return keys;
}

async function pruneExcept(uploadedSet) {
  const all = await listAllKeys();
  // On garde manifest.xml a la racine (publie separement) et on supprime les
  // autres orphelins.
  const stale = all.filter((k) => !uploadedSet.has(k) && k !== "manifest.xml");
  if (stale.length === 0) return;
  console.log(`Suppression de ${stale.length} objet(s) obsolete(s) :`);
  for (let i = 0; i < stale.length; i += 1000) {
    const batch = stale.slice(i, i + 1000);
    await client.send(
      new DeleteObjectsCommand({
        Bucket: bucket,
        Delete: { Objects: batch.map((Key) => ({ Key })) },
      })
    );
    for (const k of batch) console.log(`  - ${k}`);
  }
}

async function main() {
  const __filename = fileURLToPath(import.meta.url);
  const root = join(__filename, "..", "..");
  const dist = join(root, "dist");

  try {
    const s = await stat(dist);
    if (!s.isDirectory()) throw new Error();
  } catch {
    console.error(`Le dossier dist/ est absent. Lance d'abord 'npm run build'.`);
    process.exit(1);
  }

  console.log(`Deploiement vers s3://${bucket} (${endpoint})`);
  const uploaded = new Set();

  for await (const file of walk(dist)) {
    const key = relative(dist, file).split(sep).join("/");
    await uploadOne(file, key);
    uploaded.add(key);
  }

  // Publie le manifest prod a la racine du bucket, pour permettre aux
  // collegues de le recuperer directement via curl / navigateur.
  const prodManifest = join(root, "manifests", "manifest.prod.xml");
  await uploadOne(prodManifest, "manifest.xml");
  uploaded.add("manifest.xml");

  await pruneExcept(uploaded);

  console.log("");
  console.log("Deploiement termine.");
  console.log(`URL du taskpane : ${endpoint.replace("https://", `https://${bucket}.`)}/taskpane.html`);
  console.log(`URL du manifest : ${endpoint.replace("https://", `https://${bucket}.`)}/manifest.xml`);
}

main().catch((err) => {
  console.error("Echec du deploiement :", err);
  process.exit(1);
});
