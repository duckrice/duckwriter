// Where the others are: a coloured cursor with their name, and their selection.

import { Plugin, PluginKey } from "prosemirror-state";
import { Decoration, DecorationSet } from "prosemirror-view";

export interface Peer {
  id: string;
  name: string;
  color: string;
  anchor: number;
  head: number;
}

interface PresenceState {
  peers: Map<string, Peer>;
  decorations: DecorationSet;
}

export const presenceKey = new PluginKey<PresenceState>("presence");

/** A colour for each person, the same on every computer. */
export const PEER_COLORS = ["#d9480f", "#1971c2", "#2b8a3e", "#9c36b5", "#e8590c", "#0c8599", "#c2255c", "#5f3dc4"];

export function colorFor(name: string): string {
  let hash = 7;
  for (const ch of name) hash = (hash * 31 + ch.charCodeAt(0)) >>> 0;
  return PEER_COLORS[hash % PEER_COLORS.length];
}

function decorate(doc: import("prosemirror-model").Node, peers: Map<string, Peer>): DecorationSet {
  const decos: Decoration[] = [];
  const size = doc.content.size;
  for (const p of peers.values()) {
    const head = Math.max(0, Math.min(p.head, size));
    const anchor = Math.max(0, Math.min(p.anchor, size));
    const from = Math.min(head, anchor);
    const to = Math.max(head, anchor);
    if (to > from) decos.push(Decoration.inline(from, to, { class: "peer-selection", style: `background-color:${p.color}33` }));
    decos.push(
      Decoration.widget(
        head,
        () => {
          const caret = document.createElement("span");
          caret.className = "peer-caret";
          caret.style.borderColor = p.color;
          const label = document.createElement("span");
          label.className = "peer-label";
          label.style.background = p.color;
          label.textContent = p.name || "Someone";
          caret.append(label);
          return caret;
        },
        { key: `peer-${p.id}-${p.name}-${p.color}`, side: 1, ignoreSelection: true },
      ),
    );
  }
  return DecorationSet.create(doc, decos);
}

export function presencePlugin() {
  return new Plugin<PresenceState>({
    key: presenceKey,
    state: {
      init: () => ({ peers: new Map(), decorations: DecorationSet.empty }),
      apply(tr, old, _o, state) {
        const meta = tr.getMeta(presenceKey) as { set?: Peer; remove?: string; clear?: true } | undefined;
        if (!meta && !tr.docChanged) return old;
        const peers = new Map(old.peers);
        // Their cursors move with the text.
        if (tr.docChanged) for (const [id, p] of peers) peers.set(id, { ...p, anchor: tr.mapping.map(p.anchor), head: tr.mapping.map(p.head) });
        if (meta?.set) peers.set(meta.set.id, meta.set);
        if (meta?.remove) peers.delete(meta.remove);
        if (meta?.clear) peers.clear();
        return { peers, decorations: decorate(state.doc, peers) };
      },
    },
    props: { decorations: (state) => presenceKey.getState(state)?.decorations },
  });
}
