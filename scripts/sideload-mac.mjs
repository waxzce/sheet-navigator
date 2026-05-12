#!/usr/bin/env node
// Installe / desinstalle le manifest dans le dossier wef d'Excel sur Mac.
// Usage :
//   node scripts/sideload-mac.mjs           # installe manifest.local.xml
//   node scripts/sideload-mac.mjs --prod    # installe manifest.prod.xml
//   node scripts/sideload-mac.mjs --clean   # supprime le manifest installe

import { existsSync, mkdirSync, copyFileSync, readdirSync, unlinkSync } from "node:fs";
import { join } from "node:path";
import { homedir } from "node:os";

const args = process.argv.slice(2);
const isProd = args.includes("--prod");
const isClean = args.includes("--clean");

const wefDir = join(homedir(), "Library/Containers/com.microsoft.Excel/Data/Documents/wef");
const installedName = "sheet-navigator.manifest.xml";

if (isClean) {
  if (existsSync(wefDir)) {
    for (const f of readdirSync(wefDir)) {
      if (f.startsWith("sheet-navigator")) {
        unlinkSync(join(wefDir, f));
      }
    }
  }
  console.log("Manifests Sheet Navigator supprimes du dossier wef.");
  console.log(`  -> ${wefDir}`);
  console.log("Relance Excel pour que la modification soit prise en compte.");
  process.exit(0);
}

const source = isProd ? "manifests/manifest.prod.xml" : "manifests/manifest.local.xml";
if (!existsSync(source)) {
  console.error(`Manifest introuvable : ${source}`);
  process.exit(1);
}

mkdirSync(wefDir, { recursive: true });
copyFileSync(source, join(wefDir, installedName));

console.log(`Manifest ${isProd ? "PROD" : "LOCAL"} installe :`);
console.log(`  source : ${source}`);
console.log(`  cible  : ${join(wefDir, installedName)}`);
console.log("");
console.log("Etapes suivantes :");
console.log("  1. Quitte completement Excel (Cmd+Q)");
console.log("  2. Rouvre un classeur");
console.log("  3. Insertion > Complements > Mes complements > Developer Add-ins > Sheet Navigator");
