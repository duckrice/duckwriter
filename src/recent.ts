// The documents opened or saved lately, newest first, remembered on this computer.

const KEY = "duckwriter.recent";
const MAX = 12;

export interface RecentFile {
  path: string;
  /** When it was last opened or saved (milliseconds). */
  time: number;
}

export function recentFiles(): RecentFile[] {
  try {
    const list = JSON.parse(localStorage.getItem(KEY) ?? "[]");
    return Array.isArray(list) ? list.filter((f) => typeof f?.path === "string") : [];
  } catch {
    return [];
  }
}

function store(list: RecentFile[]) {
  try {
    localStorage.setItem(KEY, JSON.stringify(list.slice(0, MAX)));
  } catch {
    // Not remembered this time.
  }
}

const same = (a: string, b: string) => a.toLowerCase() === b.toLowerCase();

export function addRecent(path: string) {
  store([{ path, time: Date.now() }, ...recentFiles().filter((f) => !same(f.path, path))]);
}

export function removeRecent(path: string) {
  store(recentFiles().filter((f) => !same(f.path, path)));
}
