import * as React from "react";
import type { SheetInfo } from "../hooks/useWorksheets";
import { SheetRow } from "./SheetRow";
import { FolderRow } from "./FolderRow";
import {
  buildTree,
  parentFolderPath,
  leafName,
  type TreeNode,
  type FolderTreeNode,
} from "../utils/tree";

type Props = {
  // Liste deja filtree/triee a passer en mode plat (utilisee aussi pour les
  // resultats de recherche en breadcrumb).
  sheets: SheetInfo[];
  // Si vrai, on rend une liste plate avec breadcrumb (mode recherche).
  // Sinon, on construit un arbre avec dossiers via ">".
  flatMode: boolean;
  // Comparateur de tri applique a chaque niveau de l'arbre quand flatMode=false.
  compareSheets: (a: SheetInfo, b: SheetInfo) => number;
  collapsedFolders: Set<string>;
  onToggleFolder: (path: string) => void;
  onActivate: (id: string) => void;
  onToggleHidden: (id: string) => void;
  onPin: (id: string) => void;
  onReorder: (id: string, targetPosition: number) => void;
  onMove: (id: string, targetFolder: string | null, targetPosition?: number) => void;
};

type DragInfo = {
  id: string;
  sourceFolder: string | null;
};

// Position cible pour un drop. Quand `kind` vaut "folder", on droppe DANS le
// dossier `path` (le sheet sera renomme dossier/<leaf>). Quand `kind` vaut
// "sheet", on droppe juste avant la feuille cible (reorder + eventuellement
// rename si le dossier change).
type DropTarget =
  | { kind: "sheet"; id: string }
  | { kind: "folder"; path: string };

function dropTargetKey(t: DropTarget): string {
  return t.kind === "sheet" ? `sheet:${t.id}` : `folder:${t.path}`;
}

export function SheetList({
  sheets,
  flatMode,
  compareSheets,
  collapsedFolders,
  onToggleFolder,
  onActivate,
  onToggleHidden,
  onPin,
  onReorder,
  onMove,
}: Props): JSX.Element {
  const [dragging, setDragging] = React.useState<DragInfo | null>(null);
  const [dragOverKey, setDragOverKey] = React.useState<string | null>(null);

  if (sheets.length === 0) {
    return (
      <div style={{ padding: 16, color: "#605e5c", textAlign: "center", fontStyle: "italic" }}>
        Aucune feuille a afficher.
      </div>
    );
  }

  const tree: TreeNode[] = flatMode
    ? []
    : buildTree(sheets, compareSheets);

  function handleDropOnSheet(targetId: string) {
    if (!dragging || dragging.id === targetId) {
      setDragging(null);
      setDragOverKey(null);
      return;
    }
    const target = sheets.find((s) => s.id === targetId);
    if (!target) {
      setDragging(null);
      setDragOverKey(null);
      return;
    }
    const targetFolder = parentFolderPath(target.name);
    if (targetFolder === dragging.sourceFolder) {
      // Meme dossier : simple reorder.
      onReorder(dragging.id, target.position);
    } else {
      // Dossier different : rename + reorder.
      onMove(dragging.id, targetFolder, target.position);
    }
    setDragging(null);
    setDragOverKey(null);
  }

  function handleDropOnFolder(targetFolderPath: string) {
    if (!dragging) {
      setDragOverKey(null);
      return;
    }
    if (targetFolderPath === dragging.sourceFolder) {
      // Drop sur le dossier source : rien a faire.
      setDragging(null);
      setDragOverKey(null);
      return;
    }
    onMove(dragging.id, targetFolderPath);
    setDragging(null);
    setDragOverKey(null);
  }

  function renderFlat(items: SheetInfo[]) {
    return items.map((s) => {
      const breadcrumb = parentFolderPath(s.name);
      const leaf = leafName(s.name);
      return (
        <SheetRow
          key={s.id}
          sheet={s}
          depth={0}
          displayName={leaf}
          breadcrumb={breadcrumb}
          onActivate={() => onActivate(s.id)}
          onToggleHidden={() => onToggleHidden(s.id)}
          onPin={() => onPin(s.id)}
          onDragStart={() => setDragging({ id: s.id, sourceFolder: breadcrumb })}
          onDragOver={(e) => {
            e.preventDefault();
            const key = dropTargetKey({ kind: "sheet", id: s.id });
            if (dragOverKey !== key) setDragOverKey(key);
          }}
          onDrop={() => handleDropOnSheet(s.id)}
          draggingOver={dragOverKey === `sheet:${s.id}` && dragging?.id !== s.id}
        />
      );
    });
  }

  function renderTree(nodes: TreeNode[]): React.ReactNode[] {
    const out: React.ReactNode[] = [];
    for (const node of nodes) {
      if (node.kind === "sheet") {
        const s = node.sheet;
        const sourceFolder = parentFolderPath(s.name);
        out.push(
          <SheetRow
            key={`s:${s.id}`}
            sheet={s}
            depth={node.depth}
            displayName={node.name}
            onActivate={() => onActivate(s.id)}
            onToggleHidden={() => onToggleHidden(s.id)}
            onPin={() => onPin(s.id)}
            onDragStart={() => setDragging({ id: s.id, sourceFolder })}
            onDragOver={(e) => {
              e.preventDefault();
              const key = dropTargetKey({ kind: "sheet", id: s.id });
              if (dragOverKey !== key) setDragOverKey(key);
            }}
            onDrop={() => handleDropOnSheet(s.id)}
            draggingOver={dragOverKey === `sheet:${s.id}` && dragging?.id !== s.id}
          />
        );
      } else {
        const folder: FolderTreeNode = node;
        const collapsed = collapsedFolders.has(folder.path);
        out.push(
          <FolderRow
            key={`f:${folder.path}`}
            node={folder}
            collapsed={collapsed}
            childrenCount={folder.children.length}
            onToggle={() => onToggleFolder(folder.path)}
            onActivate={() => {
              if (folder.sheet) onActivate(folder.sheet.id);
            }}
            onDragOver={(e) => {
              e.preventDefault();
              const key = dropTargetKey({ kind: "folder", path: folder.path });
              if (dragOverKey !== key) setDragOverKey(key);
            }}
            onDrop={() => handleDropOnFolder(folder.path)}
            draggingOver={dragOverKey === `folder:${folder.path}`}
          />
        );
        if (!collapsed) {
          out.push(...renderTree(folder.children));
        }
      }
    }
    return out;
  }

  return (
    <div style={{ flex: "1 1 auto", overflowY: "auto", overflowX: "hidden" }}>
      {flatMode ? renderFlat(sheets) : renderTree(tree)}
    </div>
  );
}
