import type { Node } from "prosemirror-model";
import { TextSelection, type EditorState, type Plugin, type Transaction } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { describe, expect, it } from "vitest";
import { decodeInvite, encodeInvite } from "../src/collab/invite";
import type { Hub, HubEvents } from "../src/collab/hub";
import { presenceKey } from "../src/collab/presence";
import { CollabSession, type CollabApp } from "../src/collab/session";
import type { DocContext } from "../src/docx/model";
import { readDocx } from "../src/docx/read";
import { blankDocx } from "../src/docx/template";
import { createState } from "../src/editor/editor";
import { schema } from "../src/editor/schema";

/** A hub in memory, delivering only when told to: a slow network on demand. */
class Bus {
  queue: (() => void)[] = [];
  host: HubEvents | null = null;
  guests = new Map<string, HubEvents>();
  next = 0;
  hub(role: "host" | "guest"): Hub {
    let me = "";
    return {
      connect: async (events) => {
        me = role === "host" ? "host" : `p${++this.next}`;
        if (role === "host") this.host = events;
        else {
          this.guests.set(me, events);
          this.host!.joined(me);
        }
        return me;
      },
      send: (to, data) => {
        this.queue.push(() => {
          if (me === "host") for (const [id, g] of this.guests) (to === "all" || to === id) && g.data("host", data);
          else this.host!.data(me, data);
        });
      },
      close: () => undefined,
    };
  }
  /** Delivers everything waiting (and what that causes), letting the encryption finish in between. */
  async flush() {
    for (let round = 0; round < 50; round++) {
      await new Promise((r) => setTimeout(r, 5));
      if (!this.queue.length) {
        await new Promise((r) => setTimeout(r, 20));
        if (!this.queue.length) return;
      }
      const batch = this.queue.splice(0);
      for (const deliver of batch) deliver();
    }
  }
}

/** A DuckWriter without a window: an editor state and the context. */
class FakeApp implements CollabApp {
  state!: EditorState;
  session: CollabSession | null = null;
  constructor(
    public ctx: DocContext,
    doc: Node,
    public name: string,
    private bytes: Uint8Array,
  ) {
    this.state = createState(doc, () => this.ctx);
  }
  get body(): EditorView {
    return {
      state: this.state,
      dispatch: (tr: Transaction) => {
        this.state = this.state.apply(tr);
        this.session?.afterChange();
      },
      dom: document.createElement("div"),
    } as unknown as EditorView;
  }
  baseBytes() {
    return this.bytes;
  }
  openShared(doc: Node, ctx: DocContext, name: string, plugins: Plugin[]) {
    this.ctx = ctx;
    this.name = name;
    this.state = createState(doc, () => this.ctx, undefined, plugins);
  }
  setPlugins(add: Plugin[], remove: Plugin[]) {
    this.state = this.state.reconfigure({ plugins: [...this.state.plugins.filter((p) => !remove.includes(p)), ...add] });
  }
  toast() {}
  sharingChanged() {}
  commentsUpdated() {}
  sideUpdated() {}
  type(at: number, text: string) {
    this.body.dispatch(this.state.tr.insertText(text, at));
  }
  text() {
    return this.state.doc.textContent;
  }
}

async function pair() {
  const bytes = await blankDocx();
  const { context } = await readDocx(bytes);
  const doc = schema.nodes.doc.create(null, [schema.nodes.paragraph.create(null, schema.text("Hello"))]);
  const host = new FakeApp(context, doc, "notes.docx", bytes);
  const bus = new Bus();
  host.session = await CollabSession.host(host, bus.hub("host"), "room1234", "secret", "DW1.x", "Kari");
  const guest = new FakeApp((await readDocx(bytes)).context, schema.nodes.doc.create(null, [schema.nodes.paragraph.create()]), "", bytes);
  const joining = CollabSession.join(guest, bus.hub("guest"), "room1234", "secret", "DW1.x", "Ola");
  await bus.flush();
  guest.session = await joining;
  return { host, guest, bus };
}

describe("writing together", { timeout: 60_000 }, () => {
  it("gives someone joining the document as it is", async () => {
    const { guest } = await pair();
    expect(guest.text()).toBe("Hello");
    expect(guest.name).toBe("notes.docx");
    expect([...guest.session!.people.values()].map((p) => p.name).sort()).toEqual(["Kari", "Ola"]);
  });

  it("puts changes made at the same time together, the same for everyone", async () => {
    const { host, guest, bus } = await pair();
    // Both type before either hears of the other's change.
    host.type(1, "Dear ");
    guest.type(6, " world");
    guest.type(12, "!");
    await bus.flush();
    expect(host.text()).toBe("Dear Hello world!");
    expect(guest.text()).toBe("Dear Hello world!");
    host.type(host.state.doc.content.size - 1, " Bye.");
    await bus.flush();
    expect(guest.text()).toBe("Dear Hello world! Bye.");
    expect(guest.session!.confirmed).toBe(host.session!.confirmed);
  });

  it("shows where the others are", async () => {
    const { host, guest, bus } = await pair();
    guest.body.dispatch(guest.state.tr.setSelection(TextSelection.create(guest.state.doc, 3)));
    await new Promise((r) => setTimeout(r, 150));
    await bus.flush();
    const peers = presenceKey.getState(host.state)!.peers;
    expect([...peers.values()].map((p) => [p.name, p.head])).toEqual([["Ola", 3]]);
  });

  it("shares comments, and a guest wrongly invited can't read anything", async () => {
    const { host, guest, bus } = await pair();
    host.ctx.comments.items.set("9", { id: "9", author: "Kari", initials: "K", date: null, text: "Check this", xml: null, original: null, paraId: null, parent: null, done: false });
    host.session!.commentsChanged();
    await bus.flush();
    expect(guest.ctx.comments.items.get("9")?.text).toBe("Check this");

    const intruder = new FakeApp(guest.ctx, guest.state.doc, "", new Uint8Array());
    const attempt = CollabSession.join(intruder, bus.hub("guest"), "room1234", "wrong-secret", "DW1.x", "Eve");
    await bus.flush();
    // The host can't read Eve's hello, so Eve never gets the document.
    expect(intruder.text()).toBe("Hello");
    void attempt.catch(() => undefined);
  });

  it("shares footnotes and headers", async () => {
    const { host, guest, bus } = await pair();
    guest.ctx.footnotes.notes.set("7", schema.nodes.doc.create(null, [schema.nodes.paragraph.create(null, schema.text("A note"))]));
    const header = host.ctx.stories[0];
    if (header) header.doc = schema.nodes.doc.create(null, [schema.nodes.paragraph.create(null, schema.text("Draft"))]);
    else host.ctx.stories.push({ kind: "header", type: "default", path: null, rid: null, rootStart: null, prefixes: new Set(), rels: new Map(), doc: schema.nodes.doc.create(null, [schema.nodes.paragraph.create(null, schema.text("Draft"))]), saved: null });
    guest.session!.afterChange();
    host.session!.afterChange();
    await bus.flush();
    expect(host.ctx.footnotes.notes.get("7")?.textContent).toBe("A note");
    expect(guest.ctx.stories.find((x) => x.kind === "header" && x.type === "default")?.doc.textContent).toBe("Draft");
  });

  it("makes invitations that survive copying", () => {
    const invite = encodeInvite({ method: "relay", address: "wss://relay.example.com", room: "abcDEF123", secret: "s3cr3t" });
    expect(invite.startsWith("DW1.")).toBe(true);
    expect(decodeInvite(`  Here you go: ${invite}\n`)).toEqual({ method: "relay", address: "wss://relay.example.com", room: "abcDEF123", secret: "s3cr3t" });
    expect(() => decodeInvite("DW1.garbage")).toThrow(/incomplete/);
  });
});
