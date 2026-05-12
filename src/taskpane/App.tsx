import { useMemo, useState } from "react";
import { useWorksheets, type SheetInfo } from "./hooks/useWorksheets";
import { SearchBar } from "./components/SearchBar";
import { Toolbar, type SortMode } from "./components/Toolbar";
import { SheetList } from "./components/SheetList";

// Normalisation insensible aux accents et a la casse pour la recherche
function normalize(s: string): string {
  return s.normalize("NFD").replace(/\p{Diacritic}/gu, "").toLowerCase();
}

function sortSheets(sheets: SheetInfo[], mode: SortMode): SheetInfo[] {
  const pinned = sheets.filter((s) => s.isPinned);
  const others = sheets.filter((s) => !s.isPinned);
  const compare = (a: SheetInfo, b: SheetInfo) =>
    mode === "alpha" ? a.name.localeCompare(b.name, undefined, { sensitivity: "base" }) : a.position - b.position;
  return [...pinned.sort(compare), ...others.sort(compare)];
}

export function App(): JSX.Element {
  const { sheets, loading, error, activate, reorder, toggleHidden, pin, refresh } = useWorksheets();
  const [query, setQuery] = useState("");
  const [showHidden, setShowHidden] = useState(false);
  const [sortMode, setSortMode] = useState<SortMode>("workbook");

  const displayed = useMemo(() => {
    let result = sheets;
    if (!showHidden) {
      result = result.filter((s) => s.visibility === Excel.SheetVisibility.visible);
    }
    if (query.trim()) {
      const q = normalize(query);
      result = result.filter((s) => normalize(s.name).includes(q));
    }
    return sortSheets(result, sortMode);
  }, [sheets, query, showHidden, sortMode]);

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100%" }}>
      <div
        style={{
          padding: 8,
          borderBottom: "1px solid var(--colorNeutralStroke2, #e1dfdd)",
          display: "flex",
          flexDirection: "column",
          gap: 8,
        }}
      >
        <SearchBar value={query} onChange={setQuery} />
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <span style={{ fontSize: 12, color: "#605e5c" }}>
            {displayed.length} feuille{displayed.length > 1 ? "s" : ""}
          </span>
          <Toolbar
            showHidden={showHidden}
            onToggleShowHidden={() => setShowHidden((v) => !v)}
            sortMode={sortMode}
            onToggleSort={() => setSortMode((m) => (m === "workbook" ? "alpha" : "workbook"))}
            onRefresh={() => void refresh()}
          />
        </div>
      </div>
      {error && (
        <div style={{ padding: 8, background: "#fde7e9", color: "#a4262c", fontSize: 12 }}>{error}</div>
      )}
      {loading ? (
        <div style={{ padding: 16, textAlign: "center", color: "#605e5c" }}>Chargement...</div>
      ) : (
        <SheetList
          sheets={displayed}
          onActivate={(id) => void activate(id)}
          onToggleHidden={(id) => void toggleHidden(id)}
          onPin={(id) => void pin(id)}
          onReorder={(id, position) => void reorder(id, position)}
        />
      )}
    </div>
  );
}
