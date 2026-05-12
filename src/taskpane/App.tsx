import { useMemo, useState } from "react";
import { useWorksheets, type SheetInfo } from "./hooks/useWorksheets";
import { SearchBar } from "./components/SearchBar";
import { Toolbar, type SortMode } from "./components/Toolbar";
import { SheetList } from "./components/SheetList";

// Normalisation insensible aux accents et a la casse pour la recherche.
function normalize(s: string): string {
  return s.normalize("NFD").replace(/\p{Diacritic}/gu, "").toLowerCase();
}

function makeComparator(mode: SortMode): (a: SheetInfo, b: SheetInfo) => number {
  return mode === "alpha"
    ? (a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: "base" })
    : (a, b) => a.position - b.position;
}

export function App(): JSX.Element {
  const {
    sheets,
    loading,
    error,
    collapsedFolders,
    activate,
    reorder,
    toggleHidden,
    pin,
    toggleFolder,
    moveSheet,
    refresh,
  } = useWorksheets();
  const [query, setQuery] = useState("");
  const [showHidden, setShowHidden] = useState(false);
  const [sortMode, setSortMode] = useState<SortMode>("workbook");

  const compareSheets = useMemo(() => makeComparator(sortMode), [sortMode]);
  const isSearching = query.trim().length > 0;

  // Liste filtree (visibilite + recherche) avant pin/tree.
  const filtered = useMemo(() => {
    let result = sheets;
    if (!showHidden) {
      result = result.filter((s) => s.visibility === Excel.SheetVisibility.visible);
    }
    if (isSearching) {
      const q = normalize(query);
      result = result.filter((s) => normalize(s.name).includes(q));
    }
    return result;
  }, [sheets, query, showHidden, isSearching]);

  // Pins en haut (plat, hors arbre) + le reste.
  const pinned = useMemo(
    () => filtered.filter((s) => s.isPinned).sort(compareSheets),
    [filtered, compareSheets]
  );
  const nonPinned = useMemo(
    () => filtered.filter((s) => !s.isPinned).sort(compareSheets),
    [filtered, compareSheets]
  );

  // En recherche : tout en plat avec breadcrumb. Sinon arbre dans la
  // deuxieme zone, pins en plat au-dessus.
  const flatMode = isSearching;

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
            {filtered.length} feuille{filtered.length > 1 ? "s" : ""}
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
        <div style={{ flex: "1 1 auto", display: "flex", flexDirection: "column", minHeight: 0 }}>
          {pinned.length > 0 && !isSearching && (
            <>
              <div
                style={{
                  padding: "6px 12px 2px",
                  fontSize: 11,
                  textTransform: "uppercase",
                  letterSpacing: "0.05em",
                  color: "#a19f9d",
                  fontWeight: 600,
                }}
              >
                Epingles
              </div>
              <SheetList
                sheets={pinned}
                flatMode
                compareSheets={compareSheets}
                collapsedFolders={collapsedFolders}
                onToggleFolder={toggleFolder}
                onActivate={(id) => void activate(id)}
                onToggleHidden={(id) => void toggleHidden(id)}
                onPin={(id) => void pin(id)}
                onReorder={(id, position) => void reorder(id, position)}
                onMove={(id, folder, position) => void moveSheet(id, folder, position)}
              />
              <div
                style={{
                  padding: "6px 12px 2px",
                  fontSize: 11,
                  textTransform: "uppercase",
                  letterSpacing: "0.05em",
                  color: "#a19f9d",
                  fontWeight: 600,
                  borderTop: "1px solid var(--colorNeutralStroke2, #e1dfdd)",
                  marginTop: 4,
                }}
              >
                Toutes les feuilles
              </div>
            </>
          )}
          <SheetList
            sheets={isSearching ? filtered.sort(compareSheets) : nonPinned}
            flatMode={flatMode}
            compareSheets={compareSheets}
            collapsedFolders={collapsedFolders}
            onToggleFolder={toggleFolder}
            onActivate={(id) => void activate(id)}
            onToggleHidden={(id) => void toggleHidden(id)}
            onPin={(id) => void pin(id)}
            onReorder={(id, position) => void reorder(id, position)}
            onMove={(id, folder, position) => void moveSheet(id, folder, position)}
          />
        </div>
      )}
    </div>
  );
}
