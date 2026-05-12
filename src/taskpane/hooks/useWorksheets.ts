import { useCallback, useEffect, useRef, useState } from "react";

// L'API office-js renvoie `visibility` comme une union enum + literals
// ("Visible" | "Hidden" | "VeryHidden"), donc on conserve cette largeur ici.
export type SheetVisibilityValue = Excel.SheetVisibility | "Visible" | "Hidden" | "VeryHidden";

export type SheetInfo = {
  id: string;
  name: string;
  position: number;
  visibility: SheetVisibilityValue;
  tabColor: string | null;
  isActive: boolean;
  isPinned: boolean;
};

export type UseWorksheetsResult = {
  sheets: SheetInfo[];
  loading: boolean;
  error: string | null;
  activate: (id: string) => Promise<void>;
  rename: (id: string, newName: string) => Promise<void>;
  reorder: (id: string, targetPosition: number) => Promise<void>;
  toggleHidden: (id: string) => Promise<void>;
  pin: (id: string) => Promise<void>;
  refresh: () => Promise<void>;
};

const SETTINGS_KEY = "sheetNavigator.pinnedIds";

function readPinned(): string[] {
  const raw = Office.context.document.settings.get(SETTINGS_KEY);
  if (!raw) {
    return [];
  }
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed.filter((x) => typeof x === "string") : [];
  } catch {
    return [];
  }
}

async function writePinned(pinned: string[]): Promise<void> {
  Office.context.document.settings.set(SETTINGS_KEY, JSON.stringify(pinned));
  await new Promise<void>((resolve, reject) => {
    Office.context.document.settings.saveAsync((res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(res.error);
      }
    });
  });
}

export function useWorksheets(): UseWorksheetsResult {
  const [sheets, setSheets] = useState<SheetInfo[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const pinnedRef = useRef<Set<string>>(new Set());

  const fetchSheets = useCallback(async () => {
    setError(null);
    try {
      await Excel.run(async (context) => {
        const wb = context.workbook;
        const sheetsCol = wb.worksheets;
        // Sur une collection, load() recoit la liste des proprietes a charger
        // pour chaque element de items[] (pas de prefixe "items/").
        sheetsCol.load("id,name,position,visibility,tabColor");
        const active = wb.worksheets.getActiveWorksheet();
        active.load("id");
        await context.sync();

        const next: SheetInfo[] = sheetsCol.items.map((ws) => ({
          id: ws.id,
          name: ws.name,
          position: ws.position,
          visibility: ws.visibility,
          tabColor: ws.tabColor && ws.tabColor.length > 0 ? ws.tabColor : null,
          isActive: ws.id === active.id,
          isPinned: pinnedRef.current.has(ws.id),
        }));
        setSheets(next);
      });
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setLoading(false);
    }
  }, []);

  // Chargement initial + abonnement aux evenements du classeur
  useEffect(() => {
    let disposed = false;
    const handlers: Array<{ remove: () => void }> = [];

    (async () => {
      pinnedRef.current = new Set(readPinned());

      try {
        await Excel.run(async (context) => {
          const wb = context.workbook;

          const onAdded = wb.worksheets.onAdded.add(async () => {
            if (!disposed) await fetchSheets();
          });
          const onDeleted = wb.worksheets.onDeleted.add(async () => {
            if (!disposed) await fetchSheets();
          });
          const onActivated = wb.worksheets.onActivated.add(async () => {
            if (!disposed) await fetchSheets();
          });
          const onNameChanged = wb.worksheets.onNameChanged.add(async () => {
            if (!disposed) await fetchSheets();
          });

          await context.sync();
          handlers.push(onAdded, onDeleted, onActivated, onNameChanged);
        });
      } catch (e) {
        if (!disposed) {
          setError(e instanceof Error ? e.message : String(e));
        }
      }

      if (!disposed) {
        await fetchSheets();
      }
    })();

    return () => {
      disposed = true;
      // Desabonnement propre : chaque handler retourne un objet avec remove()
      // dont l'invocation doit elle-meme etre encadree par Excel.run.
      Excel.run(async (context) => {
        for (const h of handlers) {
          try {
            h.remove();
          } catch {
            /* ignore */
          }
        }
        await context.sync();
      }).catch(() => {
        /* ignore at unmount */
      });
    };
  }, [fetchSheets]);

  const activate = useCallback(async (id: string) => {
    await Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getItem(id);
      ws.activate();
      ws.getCell(0, 0).select();
      await context.sync();
    });
  }, []);

  const rename = useCallback(async (id: string, newName: string) => {
    await Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getItem(id);
      ws.name = newName;
      await context.sync();
    });
  }, []);

  const reorder = useCallback(async (id: string, targetPosition: number) => {
    await Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getItem(id);
      ws.position = targetPosition;
      await context.sync();
    });
  }, []);

  const toggleHidden = useCallback(async (id: string) => {
    await Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getItem(id);
      ws.load("visibility");
      await context.sync();
      ws.visibility =
        ws.visibility === Excel.SheetVisibility.visible
          ? Excel.SheetVisibility.hidden
          : Excel.SheetVisibility.visible;
      await context.sync();
    });
  }, []);

  const pin = useCallback(async (id: string) => {
    const next = new Set(pinnedRef.current);
    if (next.has(id)) {
      next.delete(id);
    } else {
      next.add(id);
    }
    pinnedRef.current = next;
    setSheets((prev) => prev.map((s) => (s.id === id ? { ...s, isPinned: next.has(id) } : s)));
    await writePinned(Array.from(next));
  }, []);

  return { sheets, loading, error, activate, rename, reorder, toggleHidden, pin, refresh: fetchSheets };
}
