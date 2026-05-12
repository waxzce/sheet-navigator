import * as React from "react";
import { Button, Tooltip } from "@fluentui/react-components";
import {
  Eye20Regular,
  EyeOff20Regular,
  Pin20Regular,
  PinOff20Regular,
} from "@fluentui/react-icons";
import type { SheetInfo } from "../hooks/useWorksheets";

type Props = {
  sheet: SheetInfo;
  onActivate: () => void;
  onToggleHidden: () => void;
  onPin: () => void;
  onDragStart: () => void;
  onDragOver: (e: React.DragEvent) => void;
  onDrop: () => void;
  draggingOver: boolean;
};

export function SheetRow({
  sheet,
  onActivate,
  onToggleHidden,
  onPin,
  onDragStart,
  onDragOver,
  onDrop,
  draggingOver,
}: Props): JSX.Element {
  const [hover, setHover] = React.useState(false);
  const isHidden = sheet.visibility !== Excel.SheetVisibility.visible;

  return (
    <div
      draggable
      onDragStart={onDragStart}
      onDragOver={onDragOver}
      onDrop={onDrop}
      onMouseEnter={() => setHover(true)}
      onMouseLeave={() => setHover(false)}
      onClick={onActivate}
      role="button"
      tabIndex={0}
      onKeyDown={(e) => {
        if (e.key === "Enter" || e.key === " ") {
          e.preventDefault();
          onActivate();
        }
      }}
      style={{
        display: "flex",
        alignItems: "center",
        gap: 8,
        height: 34,
        padding: "0 8px",
        cursor: "pointer",
        backgroundColor: sheet.isActive
          ? "var(--colorBrandBackground2, #ebf3fc)"
          : hover
          ? "var(--colorNeutralBackground1Hover, #f5f5f5)"
          : "transparent",
        borderLeft: sheet.isActive
          ? "3px solid var(--colorBrandBackground, #0f6cbd)"
          : "3px solid transparent",
        borderTop: draggingOver ? "2px solid var(--colorBrandStroke1, #0f6cbd)" : "2px solid transparent",
        opacity: isHidden ? 0.55 : 1,
        userSelect: "none",
      }}
    >
      <span
        aria-hidden
        style={{
          flex: "0 0 auto",
          width: 8,
          height: 8,
          borderRadius: "50%",
          background: sheet.tabColor ?? "#c8c6c4",
          border: sheet.tabColor ? "none" : "1px solid #a19f9d",
        }}
      />
      <span
        style={{
          flex: "1 1 auto",
          whiteSpace: "nowrap",
          overflow: "hidden",
          textOverflow: "ellipsis",
          fontWeight: sheet.isActive ? 600 : 400,
        }}
        title={sheet.name}
      >
        {sheet.name}
      </span>
      {(hover || sheet.isPinned) && (
        <Tooltip content={sheet.isPinned ? "Detacher" : "Epingler"} relationship="label">
          <Button
            size="small"
            appearance="subtle"
            icon={sheet.isPinned ? <Pin20Regular /> : <PinOff20Regular />}
            onClick={(e) => {
              e.stopPropagation();
              onPin();
            }}
          />
        </Tooltip>
      )}
      {hover && (
        <Tooltip content={isHidden ? "Afficher" : "Masquer"} relationship="label">
          <Button
            size="small"
            appearance="subtle"
            icon={isHidden ? <EyeOff20Regular /> : <Eye20Regular />}
            onClick={(e) => {
              e.stopPropagation();
              onToggleHidden();
            }}
          />
        </Tooltip>
      )}
    </div>
  );
}
