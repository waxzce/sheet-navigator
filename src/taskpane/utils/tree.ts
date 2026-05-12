import type { SheetInfo } from "../hooks/useWorksheets";

// Decoupe un nom de feuille en segments selon le delimiteur "/". On strippe
// les espaces et on filtre les segments vides (donc "A//B" -> ["A", "B"],
// "/A" -> ["A"], "A/" -> ["A"]). Si tout est vide ou que le nom est tout
// blanc, on retombe sur le nom brut comme segment unique.
export function splitPath(name: string): string[] {
  const parts = name
    .split("/")
    .map((s) => s.trim())
    .filter((s) => s.length > 0);
  return parts.length === 0 ? [name] : parts;
}

export type SheetTreeNode = {
  kind: "sheet";
  name: string; // dernier segment ("France" pour "Region/France")
  path: string; // chemin complet
  sheet: SheetInfo;
  depth: number;
};

export type FolderTreeNode = {
  kind: "folder";
  name: string; // dernier segment du chemin
  path: string; // chemin complet
  sheet: SheetInfo | null; // non-null si une feuille existe au chemin exact
  children: TreeNode[];
  depth: number;
};

export type TreeNode = SheetTreeNode | FolderTreeNode;

type Builder = {
  name: string;
  path: string;
  sheet: SheetInfo | null;
  children: Map<string, Builder>;
};

// Construit un arbre a partir des feuilles. `comparator` est applique a chaque
// niveau pour ordonner les enfants. Pour les dossiers purs (sans feuille
// associee), la "position" utilisee pour le tri est celle de la premiere
// feuille descendante (en ordre du workbook).
export function buildTree(
  sheets: SheetInfo[],
  comparator: (a: SheetInfo, b: SheetInfo) => number
): TreeNode[] {
  const roots = new Map<string, Builder>();

  for (const sheet of sheets) {
    const segments = splitPath(sheet.name);
    let level = roots;
    let path = "";

    for (let i = 0; i < segments.length; i++) {
      const seg = segments[i];
      path = path === "" ? seg : `${path}/${seg}`;

      let node = level.get(seg);
      if (!node) {
        node = { name: seg, path, sheet: null, children: new Map() };
        level.set(seg, node);
      }

      if (i === segments.length - 1) {
        node.sheet = sheet;
      }

      level = node.children;
    }
  }

  // Pour le tri, on a besoin d'un SheetInfo "representatif" pour les dossiers
  // sans feuille propre : on prend la premiere descendante.
  function firstSheet(b: Builder): SheetInfo | null {
    if (b.sheet) return b.sheet;
    for (const child of b.children.values()) {
      const found = firstSheet(child);
      if (found) return found;
    }
    return null;
  }

  function materialize(b: Builder, depth: number): TreeNode {
    if (b.children.size === 0) {
      // Pas d'enfants : c'est forcement une feuille pure (sinon comment cette
      // entree serait-elle dans l'arbre ?).
      if (!b.sheet) {
        // Impossible en pratique, mais on renvoie un folder vide plutot
        // que de crasher.
        return { kind: "folder", name: b.name, path: b.path, sheet: null, children: [], depth };
      }
      return { kind: "sheet", name: b.name, path: b.path, sheet: b.sheet, depth };
    }

    // Dossier (potentiellement avec feuille au chemin exact)
    const childBuilders = Array.from(b.children.values()).sort((x, y) => {
      const sx = firstSheet(x);
      const sy = firstSheet(y);
      if (!sx || !sy) return 0;
      return comparator(sx, sy);
    });

    const children = childBuilders.map((c) => materialize(c, depth + 1));
    return { kind: "folder", name: b.name, path: b.path, sheet: b.sheet, children, depth };
  }

  const rootBuilders = Array.from(roots.values()).sort((x, y) => {
    const sx = firstSheet(x);
    const sy = firstSheet(y);
    if (!sx || !sy) return 0;
    return comparator(sx, sy);
  });

  return rootBuilders.map((b) => materialize(b, 0));
}

// Donne le chemin du dossier parent ("Region/France" -> "Region",
// "France" -> null).
export function parentFolderPath(fullPath: string): string | null {
  const idx = fullPath.lastIndexOf("/");
  if (idx <= 0) return null;
  return fullPath.substring(0, idx);
}

// Dernier segment du chemin ("Region/France" -> "France").
export function leafName(fullPath: string): string {
  const idx = fullPath.lastIndexOf("/");
  if (idx < 0) return fullPath;
  return fullPath.substring(idx + 1);
}

// Compose un nouveau chemin "<folder>/<leaf>" en gerant le cas folder=null.
export function joinPath(folder: string | null, leaf: string): string {
  return folder ? `${folder}/${leaf}` : leaf;
}
