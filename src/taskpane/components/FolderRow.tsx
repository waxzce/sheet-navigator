import * as React from "react";
import { ChevronRight16Regular, ChevronDown16Regular } from "@fluentui/react-icons";
import type { FolderTreeNode } from "../utils/tree";

type Props = {
  node: FolderTreeNode;
  collapsed: boolean;
  onToggle: () => void;
  onActivate: () => void; // Active la feuille au chemin exact (si elle existe)
  onDragOver: (e: React.DragEvent) => void;
  onDrop: () => void;
  draggingOver: boolean;
  childrenCount: number;
};

// Ligne d'entete d'un dossier. Le triangle gere expand/collapse. Le nom est
// cliquable pour activer la feuille situee au chemin exact si elle existe
// (sinon le clic toggle le dossier aussi).
export function FolderRow({
  node,
  collapsed,
  onToggle,
  onActivate,
  onDragOver,
  onDrop,
  draggingOver,
  childrenCount,
}: Props): JSX.Element {
  const hasSheetHere = node.sheet !== null;
  const tabColor = node.sheet?.tabColor ?? null;
  const isActive = node.sheet?.isActive ?? false;

  return (
    <div
      onDragOver={onDragOver}
      onDrop={onDrop}
      style={{
        display: "flex",
        alignItems: "center",
        gap: 4,
        height: 28,
        paddingLeft: 8 + node.depth * 14,
        paddingRight: 8,
        cursor: "default",
        userSelect: "none",
        color: "#605e5c",
        fontSize: 12.5,
        fontWeight: 600,
        backgroundColor: isActive
          ? "var(--colorBrandBackground2, #ebf3fc)"
          : "transparent",
        borderTop: draggingOver ? "2px solid var(--colorBrandStroke1, #0f6cbd)" : "2px solid transparent",
      }}
    >
      <button
        type="button"
        onClick={(e) => {
          e.stopPropagation();
          onToggle();
        }}
        aria-label={collapsed ? "Deplier" : "Replier"}
        style={{
          background: "transparent",
          border: "none",
          padding: 0,
          margin: 0,
          width: 18,
          height: 18,
          display: "grid",
          placeItems: "center",
          cursor: "pointer",
          color: "#605e5c",
        }}
      >
        {collapsed ? <ChevronRight16Regular /> : <ChevronDown16Regular />}
      </button>
      {hasSheetHere && tabColor && (
        <span
          aria-hidden
          style={{
            flex: "0 0 auto",
            width: 6,
            height: 6,
            borderRadius: "50%",
            background: tabColor,
          }}
        />
      )}
      <span
        onClick={hasSheetHere ? onActivate : onToggle}
        title={
          hasSheetHere
            ? `Activer la feuille "${node.path}"`
            : `${childrenCount} feuille${childrenCount > 1 ? "s" : ""}`
        }
        style={{
          flex: "1 1 auto",
          whiteSpace: "nowrap",
          overflow: "hidden",
          textOverflow: "ellipsis",
          cursor: "pointer",
          color: hasSheetHere ? "var(--colorNeutralForeground1, #242424)" : "#605e5c",
          fontWeight: hasSheetHere ? 600 : 600,
        }}
      >
        {node.name}
      </span>
      <span
        aria-hidden
        style={{
          flex: "0 0 auto",
          fontSize: 11,
          color: "#a19f9d",
          fontWeight: 400,
          marginLeft: 4,
        }}
      >
        {childrenCount}
      </span>
    </div>
  );
}
