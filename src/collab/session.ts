// Writing together. The person sharing (the host) is the authority: everyone
// sends their changes to the host, which puts them in order and sends them
// back to all (ProseMirror's collab protocol). Whoever's change arrives second
// has it rebased over the first, so nothing typed is lost and everyone ends up
// with the same text. Also passed around: where everyone's cursor is,
// comments, and pictures added.
//
// What is shared: the main text change by change; comments, pictures,
// sources, footnotes, headers and footers whole (the last one saved wins if
// two people edit the same footnote or header at once).

import { Node, type Schema } from "prosemirror-model";
import { collab, getVersion, receiveTransaction, sendableSteps } from "prosemirror-collab";
import type { Plugin, Transaction } from "prosemirror-state";
import { Step } from "prosemirror-transform";
import type { EditorView } from "prosemirror-view";
import type { CommentData, DocContext } from "../docx/model";
import type { Source } from "../docx/bibliography";
import { readDocx } from "../docx/read";
import { schema } from "../editor/schema";
import { citationsKey } from "../editor/citations";
import { fromBase64, open, roomKey, seal, toBase64 } from "./crypto";
import type { Hub } from "./hub";
import { colorFor, presenceKey, presencePlugin } from "./presence";

export interface CollabApp {
  readonly body: EditorView;
  readonly ctx: DocContext;
  readonly name: string;
  /** The document's file as last opened or saved, which a guest starts from. */
  baseBytes(): Uint8Array;
  /** Shows the shared document (a guest, on joining). */
  openShared(doc: Node, ctx: DocContext, name: string, plugins: Plugin[]): void;
  /** Adds or takes away editor plugins (the host, starting or stopping). */
  setPlugins(add: Plugin[], remove: Plugin[]): void;
  toast(text: string, kind?: "info" | "error"): void;
  /** Who is here changed, or the session started or ended. */
  sharingChanged(): void;
  commentsUpdated(): void;
  /** Footnotes, headers or footers came from someone else. */
  sideUpdated(): void;
}

export interface Person {
  id: string;
  name: string;
  color: string;
}

type Message =
  | { t: "hello"; name: string }
  | { t: "welcome"; you: string; version: number; doc: unknown; base: string; name: string; comments: CommentData[]; sources: Source[]; style: string; page: DocContext["page"]; media: { key: string; ext: string; data: string }[]; people: Person[]; notes: [string, unknown][]; stories: { kind: "header" | "footer"; type: "default" | "first" | "even"; doc: unknown }[]; titlePage: boolean }
  | { t: "steps"; version: number; steps: unknown[]; clientID: string }
  | { t: "confirmed"; steps: unknown[]; clientIDs: string[] }
  | { t: "presence"; id: string; name: string; color: string; anchor: number; head: number }
  | { t: "people"; people: Person[] }
  | { t: "comments"; items: CommentData[] }
  | { t: "sources"; sources: Source[]; style: string }
  | { t: "media"; key: string; ext: string; data: string }
  | { t: "notes"; notes: [string, unknown][] }
  | { t: "stories"; stories: { kind: "header" | "footer"; type: "default" | "first" | "even"; doc: unknown }[]; titlePage: boolean }
  | { t: "resync" }
  | { t: "end" };

export class CollabSession {
  readonly people = new Map<string, Person>();
  me = "";
  private key!: CryptoKey;
  private plugins: Plugin[] = [];
  /** The host's count of changes made, in order (the authority's version). */
  private version = 0;
  private sentMedia = new Set<string>();
  private mediaUrls = new Map<string, string>();
  private presenceTimer: number | undefined;
  /** The footnotes and headers as last sent or received, to notice changes. */
  private seenNotes = new Map<string, Node>();
  private seenStories = new Map<string, Node>();
  private seenTitlePage = false;
  private lastPresence = "";
  private ended = false;
  private joined: ((ok: boolean, error?: Error) => void) | null = null;

  private constructor(
    private app: CollabApp,
    private hub: Hub,
    readonly role: "host" | "guest",
    readonly invite: string,
    readonly myName: string,
  ) {}

  /** Starts sharing the open document. */
  static async host(app: CollabApp, hub: Hub, room: string, secret: string, invite: string, name: string): Promise<CollabSession> {
    const s = new CollabSession(app, hub, "host", invite, name);
    s.key = await roomKey(secret, room);
    s.me = await hub.connect(s.events());
    s.people.set(s.me, { id: s.me, name, color: colorFor(name) });
    s.plugins = [collab({ version: 0, clientID: s.me }), presencePlugin()];
    app.setPlugins(s.plugins, []);
    s.noteSides();
    app.sharingChanged();
    return s;
  }

  /** Joins a document someone shares; resolves once it is open. */
  static async join(app: CollabApp, hub: Hub, room: string, secret: string, invite: string, name: string): Promise<CollabSession> {
    const s = new CollabSession(app, hub, "guest", invite, name);
    s.key = await roomKey(secret, room);
    const welcomed = new Promise<void>((resolve, reject) => {
      s.joined = (ok, error) => (ok ? resolve() : reject(error));
      setTimeout(() => reject(new Error("The person sharing didn't answer. Is the invitation still valid?")), 60_000);
    });
    s.me = await hub.connect(s.events());
    await s.send("host", { t: "hello", name });
    try {
      await welcomed;
    } catch (error) {
      hub.close();
      throw error;
    }
    return s;
  }

  private events() {
    return {
      data: (from: string, data: string) => void this.receive(from, data),
      joined: () => undefined,
      left: (peer: string) => this.left(peer),
      closed: (reason: string) => this.end(reason),
    };
  }

  private async send(to: string, message: Message) {
    if (this.ended) return;
    this.hub.send(to, await seal(this.key, JSON.stringify(message)));
  }

  private async receive(from: string, data: string) {
    let message: Message;
    try {
      message = JSON.parse(await open(this.key, data));
    } catch {
      // Not sealed with this room's key: someone with an old or wrong invitation.
      return;
    }
    if (this.role === "host") await this.hostReceives(from, message);
    else this.guestReceives(message);
  }

  // ---- the host ------------------------------------------------------------------------

  private async hostReceives(from: string, m: Message) {
    const view = this.app.body;
    if (m.t === "hello") {
      this.people.set(from, { id: from, name: m.name, color: colorFor(m.name) });
      await this.welcome(from);
      void this.send("all", { t: "people", people: [...this.people.values()] });
      this.app.toast(`${m.name || "Someone"} joined`);
      this.app.sharingChanged();
    } else if (m.t === "steps") {
      this.accept(m.version, m.steps, from);
    } else if (m.t === "presence") {
      this.showPeer(m);
      void this.send("all", m);
    } else if (m.t === "comments") {
      this.mergeComments(m.items);
      void this.send("all", m);
    } else if (m.t === "media") {
      this.addMedia(m.key, m.ext, m.data);
      void this.send("all", m);
    } else if (m.t === "notes" || m.t === "stories") {
      this.takeSide(m);
      void this.send("all", m);
    } else if (m.t === "sources") {
      this.setSources(m.sources, m.style);
      void this.send("all", m);
    } else if (m.t === "resync") await this.welcome(from);
    void view;
  }

  /** Sends a guest the document as it is now. */
  private async welcome(to: string) {
    const { ctx, body } = this.app;
    const media = [...ctx.media].map(([key, e]) => ({ key, ext: e.ext, data: toBase64(e.bytes) }));
    for (const m of media) this.sentMedia.add(m.key);
    await this.send(to, {
      t: "welcome",
      you: to,
      version: this.version,
      doc: confirmedDoc(body).toJSON(),
      base: toBase64(this.app.baseBytes()),
      name: this.app.name,
      comments: [...ctx.comments.items.values()],
      sources: ctx.bibliography.sources,
      style: ctx.bibliography.style,
      page: ctx.page,
      media,
      people: [...this.people.values()],
      ...this.sideParts(),
    });
  }

  /** The footnotes and headers, whole. */
  private sideParts() {
    const { ctx } = this.app;
    return {
      notes: [...ctx.footnotes.notes].map(([id, doc]) => [id, doc.toJSON()] as [string, unknown]),
      stories: ctx.stories.map((s) => ({ kind: s.kind, type: s.type, doc: s.doc.toJSON() })),
      titlePage: ctx.titlePage,
    };
  }

  /** Remembers the footnotes and headers as they are, so only later changes are sent. */
  private noteSides() {
    const { ctx } = this.app;
    this.seenNotes = new Map(ctx.footnotes.notes);
    this.seenStories = new Map(ctx.stories.map((s) => [`${s.kind}:${s.type}`, s.doc]));
    this.seenTitlePage = ctx.titlePage;
  }

  /** Sends footnotes and headers that changed here. */
  private sendSides() {
    const { ctx } = this.app;
    const notes = [...ctx.footnotes.notes].filter(([id, doc]) => this.seenNotes.get(id) !== doc);
    if (notes.length) void this.send(this.role === "host" ? "all" : "host", { t: "notes", notes: notes.map(([id, doc]) => [id, doc.toJSON()]) });
    const changed = ctx.stories.some((s) => this.seenStories.get(`${s.kind}:${s.type}`) !== s.doc) || ctx.titlePage !== this.seenTitlePage;
    if (changed) {
      const { stories, titlePage } = this.sideParts();
      void this.send(this.role === "host" ? "all" : "host", { t: "stories", stories, titlePage });
    }
    if (notes.length || changed) this.noteSides();
  }

  /** Footnotes or headers from someone else. */
  private takeSide(m: Extract<Message, { t: "notes" }> | Extract<Message, { t: "stories" }>) {
    const { ctx } = this.app;
    if (m.t === "notes") {
      for (const [id, json] of m.notes) ctx.footnotes.notes.set(id, Node.fromJSON(schema, this.localImages(json)));
    } else {
      for (const s of m.stories) {
        const doc = Node.fromJSON(schema, this.localImages(s.doc));
        const story = ctx.stories.find((x) => x.kind === s.kind && x.type === s.type);
        if (story) story.doc = doc;
        else ctx.stories.push({ kind: s.kind, type: s.type, path: null, rid: null, rootStart: null, prefixes: new Set(), rels: new Map(), doc, saved: null });
      }
      ctx.titlePage = m.titlePage;
    }
    this.noteSides();
    this.app.sideUpdated();
  }

  /** Changes made from version `version` on; out-of-date ones are dropped (they come back rebased). */
  private accept(version: number, stepsJson: unknown[], clientID: string) {
    if (version !== this.version || !stepsJson.length) return;
    const view = this.app.body;
    let steps: Step[];
    try {
      steps = stepsJson.map((j) => Step.fromJSON(schema, this.localImages(j)));
      view.dispatch(remote(receiveTransaction(view.state, steps, steps.map(() => clientID))));
    } catch (error) {
      console.warn("A change couldn't be applied; the guest is sent the document again", error);
      if (clientID !== this.me) void this.welcome(clientID);
      return;
    }
    this.version += steps.length;
    void this.send("all", { t: "confirmed", steps: stepsJson, clientIDs: steps.map(() => clientID) });
  }

  // ---- a guest -----------------------------------------------------------------------------

  private guestReceives(m: Message) {
    const view = this.app.body;
    if (m.t === "welcome") void this.welcomed(m);
    else if (m.t === "confirmed") {
      try {
        const steps = m.steps.map((j) => Step.fromJSON(schema, this.localImages(j)));
        view.dispatch(remote(receiveTransaction(view.state, steps, m.clientIDs, { mapSelectionBackward: true })));
      } catch (error) {
        console.warn("Out of step with the host; asking for the document again", error);
        void this.send("host", { t: "resync" });
        return;
      }
      this.afterChange();
    } else if (m.t === "presence") {
      if (m.id !== this.me) this.showPeer(m);
    } else if (m.t === "people") {
      this.people.clear();
      for (const p of m.people) this.people.set(p.id, p);
      // Someone who left: their cursor goes.
      const shown = presenceKey.getState(view.state)?.peers;
      for (const id of shown?.keys() ?? []) if (!this.people.has(id)) view.dispatch(view.state.tr.setMeta(presenceKey, { remove: id }));
      this.app.sharingChanged();
    } else if (m.t === "comments") this.mergeComments(m.items);
    else if (m.t === "media") this.addMedia(m.key, m.ext, m.data);
    else if (m.t === "notes" || m.t === "stories") this.takeSide(m);
    else if (m.t === "sources") this.setSources(m.sources, m.style);
    else if (m.t === "end") this.end("The person sharing stopped. Your copy stays open; save it to keep it.");
  }

  private async welcomed(m: Extract<Message, { t: "welcome" }>) {
    try {
      const { context } = await readDocx(fromBase64(m.base));
      context.page = m.page;
      context.comments.items = new Map(m.comments.map((c) => [c.id, c]));
      context.comments.version++;
      context.bibliography.sources = m.sources;
      context.bibliography.style = m.style as DocContext["bibliography"]["style"];
      for (const media of m.media) {
        const bytes = fromBase64(media.data);
        context.media.set(media.key, { bytes, ext: media.ext, rid: null });
        this.mediaUrls.set(media.key, URL.createObjectURL(new Blob([bytes as BlobPart])));
      }
      this.ctxForImages = context;
      const doc = Node.fromJSON(schema as Schema, this.localImages(m.doc));
      for (const [id, json] of m.notes) context.footnotes.notes.set(id, Node.fromJSON(schema, this.localImages(json)));
      for (const s of m.stories) {
        const storyDoc = Node.fromJSON(schema, this.localImages(s.doc));
        const story = context.stories.find((x) => x.kind === s.kind && x.type === s.type);
        if (story) story.doc = storyDoc;
        else context.stories.push({ kind: s.kind, type: s.type, path: null, rid: null, rootStart: null, prefixes: new Set(), rels: new Map(), doc: storyDoc, saved: null });
      }
      context.titlePage = m.titlePage;
      this.people.clear();
      for (const p of m.people) this.people.set(p.id, p);
      this.people.set(this.me, { id: this.me, name: this.myName, color: colorFor(this.myName) });
      this.plugins = [collab({ version: m.version, clientID: this.me }), presencePlugin()];
      this.app.openShared(doc, context, m.name, this.plugins);
      this.ctxForImages = null;
      this.noteSides();
      this.app.sharingChanged();
      this.joined?.(true);
      this.joined = null;
    } catch (error) {
      this.joined?.(false, error instanceof Error ? error : new Error(String(error)));
      this.joined = null;
    }
  }

  // ---- both ----------------------------------------------------------------------------------

  /** After every change here: send what is new (pictures first), and where the cursor is. */
  afterChange() {
    if (this.ended || !this.plugins.length) return;
    this.sendSides();
    const state = this.app.body.state;
    const sendable = sendableSteps(state);
    if (sendable) {
      const json = sendable.steps.map((s) => s.toJSON());
      this.sendNewMedia(json);
      if (this.role === "host") this.accept(sendable.version, json, this.me);
      else void this.send("host", { t: "steps", version: sendable.version, steps: json, clientID: this.me });
    }
    this.schedulePresence();
  }

  private schedulePresence() {
    if (this.presenceTimer) return;
    this.presenceTimer = window.setTimeout(() => {
      this.presenceTimer = undefined;
      const { anchor, head } = this.app.body.state.selection;
      const me = this.people.get(this.me);
      const message = { t: "presence" as const, id: this.me, name: this.myName, color: me?.color ?? colorFor(this.myName), anchor, head };
      const key = `${anchor}:${head}`;
      if (key === this.lastPresence) return;
      this.lastPresence = key;
      void this.send(this.role === "host" ? "all" : "host", message);
    }, 120);
  }

  private showPeer(p: { id: string; name: string; color: string; anchor: number; head: number }) {
    const view = this.app.body;
    if (!presenceKey.getState(view.state)) return;
    view.dispatch(view.state.tr.setMeta(presenceKey, { set: { id: p.id, name: p.name, color: p.color, anchor: p.anchor, head: p.head } }).setMeta("addToHistory", false));
  }

  /** Comments changed here: everyone gets them. */
  commentsChanged() {
    if (this.ended) return;
    const items = [...this.app.ctx.comments.items.values()];
    void this.send(this.role === "host" ? "all" : "host", { t: "comments", items });
  }

  /** The sources (for citations) changed here. */
  sourcesChanged() {
    if (this.ended) return;
    const bib = this.app.ctx.bibliography;
    void this.send(this.role === "host" ? "all" : "host", { t: "sources", sources: bib.sources, style: bib.style });
  }

  private mergeComments(items: CommentData[]) {
    const own = this.app.ctx.comments;
    for (const item of items) own.items.set(item.id, { ...own.items.get(item.id), ...item });
    own.version++;
    this.app.commentsUpdated();
  }

  private setSources(sources: Source[], style: string) {
    const bib = this.app.ctx.bibliography;
    bib.sources = sources;
    bib.style = style as typeof bib.style;
    bib.version++;
    this.app.body.dispatch(this.app.body.state.tr.setMeta(citationsKey, true));
  }

  /** Pictures added here go out before the change that uses them. */
  private sendNewMedia(steps: unknown[]) {
    const keys = new Set<string>();
    const walk = (v: unknown) => {
      if (Array.isArray(v)) v.forEach(walk);
      else if (v && typeof v === "object") {
        const o = v as { type?: string; attrs?: { media?: string } };
        if (o.type === "image" && o.attrs?.media) keys.add(o.attrs.media);
        Object.values(o).forEach(walk);
      }
    };
    walk(steps);
    for (const key of keys) {
      const entry = this.app.ctx.media.get(key);
      if (!entry || this.sentMedia.has(key)) continue;
      this.sentMedia.add(key);
      void this.send(this.role === "host" ? "all" : "host", { t: "media", key, ext: entry.ext, data: toBase64(entry.bytes) });
    }
  }

  private addMedia(key: string, ext: string, data: string) {
    if (this.app.ctx.media.has(key)) return;
    const bytes = fromBase64(data);
    this.app.ctx.media.set(key, { bytes, ext, rid: null });
    this.sentMedia.add(key);
    this.mediaUrls.set(key, URL.createObjectURL(new Blob([bytes as BlobPart])));
  }

  /** Used while a guest's document is being made, before the app has its context. */
  private ctxForImages: DocContext | null = null;

  /** Pictures in a change point at the sender's copies; here they show this computer's. */
  private localImages<T>(json: T): T {
    const ctx = this.ctxForImages ?? this.app.ctx;
    const walk = (v: unknown): unknown => {
      if (Array.isArray(v)) return v.map(walk);
      if (!v || typeof v !== "object") return v;
      const o = { ...(v as Record<string, unknown>) };
      for (const [k, value] of Object.entries(o)) o[k] = walk(value);
      if (o.type === "image" && o.attrs && typeof o.attrs === "object") {
        const a = o.attrs as { rid?: string; media?: string; src?: string };
        a.src = (a.rid && ctx.imageSources.get(a.rid)) || (a.media && this.mediaUrls.get(a.media)) || a.src || "";
      }
      return o;
    };
    return walk(json) as T;
  }

  private left(peer: string) {
    if (peer === "host" && this.role === "guest") return this.end("The person sharing has gone. Your copy stays open; save it to keep it.");
    const person = this.people.get(peer);
    this.people.delete(peer);
    const view = this.app.body;
    view.dispatch(view.state.tr.setMeta(presenceKey, { remove: peer }));
    if (this.role === "host") void this.send("all", { t: "people", people: [...this.people.values()] });
    if (person) this.app.toast(`${person.name || "Someone"} left`);
    this.app.sharingChanged();
  }

  /** Stops sharing (the host) or leaves (a guest); the document stays open. */
  stop() {
    if (this.ended) return;
    if (this.role === "host") void this.send("all", { t: "end" }).then(() => setTimeout(() => this.hub.close(), 300));
    else this.hub.close();
    this.finish();
  }

  private end(reason: string) {
    if (this.ended) return;
    this.app.toast(reason);
    this.hub.close();
    this.finish();
  }

  private finish() {
    this.ended = true;
    clearTimeout(this.presenceTimer);
    this.app.setPlugins([], this.plugins);
    this.plugins = [];
    this.app.sharingChanged();
  }

  get active(): boolean {
    return !this.ended;
  }

  /** The version the host has given out (for tests). */
  get confirmed(): number {
    return this.role === "host" ? this.version : getVersion(this.app.body.state);
  }
}

/** Marks a change from someone else: not tracked again here, and not one to undo. */
function remote(tr: Transaction): Transaction {
  return tr.setMeta("noTrack", true);
}

/**
 * The host's document as put in order so far. The host's own changes are
 * accepted as soon as they are made, so there are none waiting.
 */
function confirmedDoc(view: EditorView): Node {
  return view.state.doc;
}
