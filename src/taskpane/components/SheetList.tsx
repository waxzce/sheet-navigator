import * as React from "react";
import type { SheetInfo } from "../hooks/useWorksheets";
import { SheetRow } from "./SheetRow";

type Props = {
  sheets: SheetInfo[];
  onActivate: (id: string) => void;
  onToggleHidden: (id: string) => void;
  onPin: (id: string) => void;
  onReorder: (id: string, targetPosition: number) => void;
};

export function SheetList({ sheets, onActivate, onToggleHidden, onPin, onReorder }: Props): JSX.Element {
  const [draggingId, setDraggingId] = React.useState<string | null>(null);
  const [dragOverId, setDragOverId] = React.useState<string | null>(null);

  const handleDrop = (targetId: string) => {
    if (!draggingId || draggingId === targetId) {
      setDraggingId(null);
      setDragOverId(null);
      return;
    }
    const target = sheets.find((s) => s.id === targetId);
    if (target) {
      onReorder(draggingId, target.position);
    }
    setDraggingId(null);
    setDragOverId(null);
  };

  if (sheets.length === 0) {
    return (
      <div style={{ padding: 16, color: "#605e5c", textAlign: "center", fontStyle: "italic" }}>
        Aucune feuille a afficher.
      </div>
    );
  }

  return (
    <div style={{ flex: "1 1 auto", overflowY: "auto", overflowX: "hidden" }}>
      {sheets.map((s) => (
        <SheetRow
          key={s.id}
          sheet={s}
          onActivate={() => onActivate(s.id)}
          onToggleHidden={() => onToggleHidden(s.id)}
          onPin={() => onPin(s.id)}
          onDragStart={() => setDraggingId(s.id)}
          onDragOver={(e) => {
            e.preventDefault();
            if (dragOverId !== s.id) setDragOverId(s.id);
          }}
          onDrop={() => handleDrop(s.id)}
          draggingOver={dragOverId === s.id && draggingId !== s.id}
        />
      ))}
    </div>
  );
}
