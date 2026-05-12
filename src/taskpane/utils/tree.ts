import type { SheetInfo } from "../hooks/useWorksheets";

// Separateur de hierarchie. Excel interdit "/" dans les noms de feuilles
// (cf. caracteres reserves : \ / ? * [ ]), donc on utilise ">" qui passe.
// Les espaces autour du separateur sont tolerees au parsing, ce qui permet
// "Region>France" comme "Region > France". La forme canonique au join
// utilise des espaces (" > ") pour la lisibilite.
export const SEPARATOR = ">";
export const SEPARATOR_DISPLAY = " > ";
const SEPARATOR_REGEX = /\s*>\s*/;

// Decoupe un nom de feuille en segments. Si tous les segments sont non vides
// apres strip, on retourne la liste ; sinon on traite le nom comme un leaf
// litteral (par exemple ">100" ou "A > > B" ne sont pas une hierarchie
// valide et restent une feuille a un seul segment).
export function splitPath(name: string): string[] {
  if (!name.includes(SEPARATOR)) return [name];
  const rawParts = name.split(SEPARATOR_REGEX);
  if (rawParts.some((p) => p.length === 0)) return [name];
  return rawParts;
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
      path = path === "" ? seg : `${path}${SEPARATOR_DISPLAY}${seg}`;

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

// Donne le chemin du dossier parent ("Region > France" -> "Region",
// "France" -> null). Renvoie null si splitPath ne reconnait pas le nom
// comme une hierarchie (cas des noms ambigus type ">100").
export function parentFolderPath(fullPath: string): string | null {
  const segments = splitPath(fullPath);
  if (segments.length < 2) return null;
  return segments.slice(0, -1).join(SEPARATOR_DISPLAY);
}

// Dernier segment du chemin ("Region > France" -> "France"). Si le nom
// n'est pas une hierarchie reconnue, renvoie le nom complet.
export function leafName(fullPath: string): string {
  const segments = splitPath(fullPath);
  return segments[segments.length - 1];
}

// Compose un nouveau chemin "<folder> > <leaf>" en gerant le cas folder=null.
export function joinPath(folder: string | null, leaf: string): string {
  return folder ? `${folder}${SEPARATOR_DISPLAY}${leaf}` : leaf;
}
