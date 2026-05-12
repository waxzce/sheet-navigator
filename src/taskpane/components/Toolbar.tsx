import { ToggleButton, Button, Tooltip } from "@fluentui/react-components";
import {
  Eye20Regular,
  EyeOff20Regular,
  ArrowSortDownLines20Regular,
  TextSortAscending20Regular,
  ArrowClockwise20Regular,
} from "@fluentui/react-icons";

export type SortMode = "workbook" | "alpha";

type Props = {
  showHidden: boolean;
  onToggleShowHidden: () => void;
  sortMode: SortMode;
  onToggleSort: () => void;
  onRefresh: () => void;
};

export function Toolbar({
  showHidden,
  onToggleShowHidden,
  sortMode,
  onToggleSort,
  onRefresh,
}: Props): JSX.Element {
  return (
    <div style={{ display: "flex", gap: 4, alignItems: "center" }}>
      <Tooltip content={showHidden ? "Masquer les feuilles cachees" : "Afficher les feuilles cachees"} relationship="label">
        <ToggleButton
          size="small"
          appearance="subtle"
          checked={showHidden}
          icon={showHidden ? <Eye20Regular /> : <EyeOff20Regular />}
          onClick={onToggleShowHidden}
        />
      </Tooltip>
      <Tooltip
        content={sortMode === "workbook" ? "Trier par nom" : "Trier par ordre du classeur"}
        relationship="label"
      >
        <Button
          size="small"
          appearance="subtle"
          icon={sortMode === "workbook" ? <ArrowSortDownLines20Regular /> : <TextSortAscending20Regular />}
          onClick={onToggleSort}
        />
      </Tooltip>
      <Tooltip content="Rafraichir" relationship="label">
        <Button size="small" appearance="subtle" icon={<ArrowClockwise20Regular />} onClick={onRefresh} />
      </Tooltip>
    </div>
  );
}
