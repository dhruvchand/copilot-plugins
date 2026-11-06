// Teams chat state for the canvas: chat list with unread detection, display
// names, per-chat history, sending, and read state. All Microsoft 365 access
// goes through the WorkIQ transport in workiq.mjs.

import { createHash } from "node:crypto";
import { mkdirSync, readFileSync, renameSync, writeFileSync } from "node:fs";
import { mkdir, readdir, readFile, rename, unlink, writeFile } from "node:fs/promises";
import { dirname, join } from "node:path";
import { WorkIQError } from "./workiq.mjs";

const FULL_PAGE = 50;
const POLL_PAGE = 12;
const HISTORY_PAGE = 30;
const REFRESH_PAGE = 10;
const POLL_MS = 20_000;
const FULL_REFRESH_MS = 5 * 60_000;
const HISTORY_STALE_MS = 90_000;
const NAME_TTL_MS = 7 * 24 * 3600_000;
const NAME_RETRY_MS = 10 * 60_000;
const NAME_BATCH = 8;
const MAX_MESSAGE_CHARS = 28_000;
const MAX_DRAFT_CHARS = 30_000;
const DRAFT_TTL_MS = 14 * 24 * 3600_000;
const CHAT_SELECT = "id,topic,chatType,viewpoint,webUrl,createdDateTime,lastUpdatedDateTime";
const ONE_ON_ONE = /^19:([0-9a-f-]{36})_([0-9a-f-]{36})@unq\.gbl\.spaces$/i;
const REACTION_EMOJI = { like: "👍", heart: "❤️", laugh: "😆", surprised: "😮", sad: "😢", angry: "😡" };
const ENTITIES = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'", nbsp: " " };

const chatsPath = (top) =>
    `/me/chats?$expand=lastMessagePreview&$orderby=lastMessagePreview/createdDateTime desc&$top=${top}&$select=${CHAT_SELECT}`;

async function readJson(path) {
    try {
        return JSON.parse(await readFile(path, "utf8"));
    } catch {
        return null;
    }
}

async function writeJsonAtomic(path, value) {
    await mkdir(dirname(path), { recursive: true, mode: 0o700 });
    const temp = `${path}.${process.pid}.${Date.now()}.tmp`;
    await writeFile(temp, JSON.stringify(value), { mode: 0o600 });
    await rename(temp, path);
}

export function htmlToText(html) {
    if (!html) return "";
    return String(html)
        // Drop content that never renders in the canvas, so the agent sees no
        // text that is invisible to the user.
        .replace(/<!--[\s\S]*?(?:-->|$)/g, "")
        .replace(/<(script|style|noscript|template|head|title|iframe|frame|frameset|object|embed|svg|math|form|button|textarea|select|video|audio|canvas|xml|attachment)\b[\s\S]*?(?:<\/\1\s*>|$)/gi, "")
        // The canvas renders an emoji's alt text, falling back to its contents.
        .replace(/<emoji\b([^>]*)>([\s\S]*?)<\/emoji\s*>/gi, (_, attrs, inner) => /\balt="([^"]*)"/i.exec(attrs)?.[1] || inner)
        .replace(/<emoji\b([^>]*)>/gi, (_, attrs) => /\balt="([^"]*)"/i.exec(attrs)?.[1] ?? "")
        .replace(/<img\b[^>]*>/gi, (tag) =>
            /schema\.skype\.com\/Emoji/i.test(tag) ? (/\balt="([^"]*)"/i.exec(tag)?.[1] ?? "") : " [image] ",
        )
        .replace(/<(?:br|\/p|\/div|\/li|\/h\d|\/blockquote|\/pre)\b[^>]*>/gi, "\n")
        .replace(/<[^>]+>/g, "")
        .replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (match, entity) => {
            if (entity[0] === "#") {
                const code = entity[1].toLowerCase() === "x" ? parseInt(entity.slice(2), 16) : parseInt(entity.slice(1), 10);
                return Number.isFinite(code) && code > 0 && code < 0x110000 ? String.fromCodePoint(code) : match;
            }
            return ENTITIES[entity.toLowerCase()] ?? match;
        })
        .replace(/[ \t\u00a0]+/g, " ")
        .replace(/ *\n */g, "\n")
        .replace(/\n{3,}/g, "\n\n")
        .trim();
}

function escapeHtml(value) {
    return String(value)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
}

export function textToHtml(text) {
    const linked = escapeHtml(text).replace(/\bhttps?:\/\/[^\s<]+/g, (url) => {
        const href = url.replace(/(?:&#39;|&quot;|[.,;:!?)\]])+$/, "");
        return `<a href="${href}">${href}</a>${url.slice(href.length)}`;
    });
    return `<p>${linked.replace(/\r?\n/g, "<br>")}</p>`;
}

function safeJson(value) {
    if (typeof value !== "string") return value ?? null;
    try {
        return JSON.parse(value);
    } catch {
        return null;
    }
}

function safeUrl(value) {
    try {
        const url = new URL(value);
        return url.protocol === "https:" || url.protocol === "http:" ? url.href : null;
    } catch {
        return null;
    }
}

function relativeLink(link) {
    if (!link) return null;
    if (link.startsWith("/")) return link;
    try {
        const url = new URL(link);
        return url.pathname.replace(/^\/(?:v1\.0|beta)(?=\/)/, "") + url.search;
    } catch {
        return null;
    }
}

function identityName(identity) {
    return identity?.user?.displayName ?? identity?.application?.displayName ?? identity?.device?.displayName ?? null;
}

function formatDuration(iso) {
    const match = /^P(?:\d+D)?T?(?:(\d+)H)?(?:(\d+)M)?(?:(\d+(?:\.\d+)?)S)?$/.exec(iso ?? "");
    if (!match) return "";
    const [, h, m, s] = match;
    const parts = [];
    if (h) parts.push(`${h}h`);
    if (m) parts.push(`${m}m`);
    if (s && !h) parts.push(`${Math.round(Number(s))}s`);
    return parts.join(" ");
}

function describeEvent(detail) {
    if (!detail) return "Activity";
    const type = String(detail["@odata.type"] ?? "").replace(/^#microsoft\.graph\./, "");
    const who = identityName(detail.initiator) ?? "Someone";
    const names = (list) => (list ?? []).map((m) => m.displayName).filter(Boolean).join(", ") || "people";
    switch (type) {
        case "membersAddedEventMessageDetail":
            return `${who} added ${names(detail.members)}`;
        case "membersDeletedEventMessageDetail":
            return `${who} removed ${names(detail.members)}`;
        case "membersJoinedEventMessageDetail":
            return `${names(detail.members)} joined`;
        case "membersLeftEventMessageDetail":
            return `${names(detail.members)} left`;
        case "chatRenamedEventMessageDetail":
            return `${who} renamed the chat to “${detail.chatDisplayName ?? ""}”`;
        case "callStartedEventMessageDetail":
            return "Call started";
        case "callEndedEventMessageDetail": {
            const duration = formatDuration(detail.callDuration);
            return duration ? `Call ended · ${duration}` : "Call ended";
        }
        case "callRecordingEventMessageDetail":
            return "Call recording is ready";
        case "callTranscriptEventMessageDetail":
            return "Call transcript is ready";
        case "meetingPolicyUpdatedEventMessageDetail":
            return "Meeting settings were updated";
        case "messagePinnedEventMessageDetail":
            return `${who} pinned a message`;
        case "messageUnpinnedEventMessageDetail":
            return `${who} unpinned a message`;
        default: {
            const words = type.replace(/EventMessageDetail$/, "").replace(/([a-z])([A-Z])/g, "$1 $2").toLowerCase();
            return words ? words[0].toUpperCase() + words.slice(1) : "Activity";
        }
    }
}

function previewOf(message) {
    if (!message) return "";
    if (message.isDeleted || message.deletedDateTime) return "This message was deleted";
    if (message.messageType && message.messageType !== "message") return describeEvent(message.eventDetail);
    const html = message.body?.content ?? "";
    const text = (message.body?.contentType === "html" ? htmlToText(html) : String(html)).replace(/\s+/g, " ").trim();
    if (text && text !== "[image]") return text.length > 180 ? `${text.slice(0, 179)}…` : text;
    if (/<img\b/i.test(html)) return "Sent an image";
    if (/<attachment\b/i.test(html) || message.attachments?.length) return "Sent an attachment";
    return text;
}

function normAttachment(attachment) {
    if (!attachment) return null;
    const type = attachment.contentType ?? "";
    if (type === "messageReference" || type === "forwardedMessageReference") {
        const ref = safeJson(attachment.content);
        return {
            kind: "quote",
            id: attachment.id ?? null,
            forwarded: type === "forwardedMessageReference",
            sender: identityName(ref?.messageSender) ?? null,
            preview: htmlToText(ref?.messagePreview ?? "").slice(0, 400),
        };
    }
    if (type.startsWith("application/vnd.microsoft.card")) {
        return { kind: "card", id: attachment.id ?? null, name: attachment.name || "Card" };
    }
    return {
        kind: "file",
        id: attachment.id ?? null,
        name: attachment.name || (type.startsWith("image/") ? "Image" : "Attachment"),
        url: safeUrl(attachment.contentUrl),
    };
}

function summarizeReactions(list, meId) {
    const byType = new Map();
    for (const reaction of list ?? []) {
        const key = reaction.reactionType ?? "?";
        const emoji = REACTION_EMOJI[key] ?? (/\p{Extended_Pictographic}/u.test(key) ? key : reaction.displayName || key);
        const entry = byType.get(key) ?? { emoji, count: 0, mine: false };
        entry.count++;
        if (meId && reaction.user?.user?.id === meId) entry.mine = true;
        byType.set(key, entry);
    }
    return [...byType.values()];
}

function normMessage(raw, meId, partial = false) {
    const user = raw.from?.user ?? null;
    const app = raw.from?.application ?? null;
    const kind = raw.messageType ?? "message";
    const deleted = Boolean(raw.deletedDateTime || raw.isDeleted);
    const isHtml = raw.body?.contentType === "html";
    return {
        id: String(raw.id),
        at: raw.createdDateTime ?? null,
        editedAt: raw.lastEditedDateTime ?? null,
        kind,
        deleted,
        fromId: user?.id ?? app?.id ?? null,
        fromName: user?.displayName ?? app?.displayName ?? null,
        mine: Boolean(meId && user?.id === meId),
        html: deleted || !isHtml ? null : raw.body?.content ?? "",
        text: deleted || isHtml ? null : raw.body?.content ?? "",
        subject: raw.subject || null,
        importance: raw.importance && raw.importance !== "normal" ? raw.importance : null,
        attachments: deleted ? [] : (raw.attachments ?? []).map(normAttachment).filter(Boolean),
        reactions: summarizeReactions(raw.reactions, meId),
        mentionsMe: Boolean(meId && (raw.mentions ?? []).some((m) => m.mentioned?.user?.id === meId)),
        event: kind === "message" ? null : describeEvent(raw.eventDetail),
        partial,
    };
}

function toStatusError(error) {
    const throttled = error?.status === 429 || error?.code === "CircuitOpen" || error?.code === "cooldown";
    return {
        message: throttled
            ? "Microsoft Graph is rate-limiting these requests."
            : error?.message ?? String(error),
        code: error?.code || (error?.status ? String(error.status) : ""),
        throttled,
        retryAt: error?.retryAfterMs ? Date.now() + error.retryAfterMs : null,
    };
}

const time = (iso) => (iso ? Date.parse(iso) || 0 : 0);

function firstName(name) {
    const trimmed = String(name ?? "").trim();
    if (/,/.test(trimmed)) return trimmed.split(",")[1]?.trim().split(/\s+/)[0] || trimmed;
    return trimmed.split(/\s+/)[0] || trimmed;
}

export class TeamsStore {
    constructor({ wq, cacheDir, debug = () => {} }) {
        this.wq = wq;
        this.namesPath = cacheDir ? join(cacheDir, "names.json") : null;
        this.draftsDir = cacheDir ? join(cacheDir, "drafts") : null;
        this.debug = debug;
        this.me = null;
        this.chats = new Map();
        this.moreLink = null;
        this.loaded = false;
        this.fullAt = 0;
        this.status = { loading: false, error: null };
        this.histories = new Map();
        this.drafts = new Map();
        this.draftRevs = new Map();
        this.draftTouched = new Map();
        this.draftSaveTimer = null;
        this.pendingSends = new Map();
        this.reconcileDelays = [4000, 8000, 15_000, 30_000];
        this.names = { users: {}, chats: {} };
        this.nameMisses = new Map();
        this.listeners = new Set();
        this.active = new Map();
        this.watching = false;
        this.pollTimer = null;
        this.retryTimers = new Map();
        this.historyFailures = 0;
        this.inflight = new Map();
        this.chatsSignature = "";
        this.saveTimer = null;
        this.ready = Promise.all([this.loadNames(), this.loadDrafts()]).catch((error) =>
            this.debug(`cache load failed: ${error?.message ?? error}`),
        );
    }

    on(listener) {
        this.listeners.add(listener);
        return () => this.listeners.delete(listener);
    }

    emit(event) {
        for (const listener of this.listeners) {
            try {
                listener(event);
            } catch (error) {
                this.debug(`listener failed: ${error?.stack ?? error}`);
            }
        }
    }

    once(key, fn) {
        const existing = this.inflight.get(key);
        if (existing) return existing;
        const promise = Promise.resolve()
            .then(fn)
            .finally(() => this.inflight.delete(key));
        this.inflight.set(key, promise);
        return promise;
    }

    setStatus(patch) {
        this.status = { ...this.status, ...patch };
        this.emit({ type: "status" });
    }

    // ---- identity --------------------------------------------------------

    async ensureMe() {
        if (this.me) return this.me;
        return this.once("me", async () => {
            const [me, org] = await this.wq.fetchMany([
                "/me?$select=id,displayName,userPrincipalName",
                "/organization?$select=id",
            ]);
            if (!me.ok) throw me.error;
            this.me = {
                id: me.data.id,
                name: me.data.displayName ?? me.data.userPrincipalName ?? "Me",
                tenantId: org.ok ? org.data?.value?.[0]?.id ?? null : null,
            };
            return this.me;
        });
    }

    learnFrom(preview) {
        const user = preview?.from?.user;
        if (!user?.id) return;
        if (user.displayName) this.rememberUser(user.id, user.displayName);
        if (this.me && !this.me.tenantId && user.id === this.me.id && user.tenantId) this.me.tenantId = user.tenantId;
    }

    // ---- display names ---------------------------------------------------

    async loadNames() {
        if (!this.namesPath) return;
        const parsed = await readJson(this.namesPath);
        if (parsed && typeof parsed.users === "object" && typeof parsed.chats === "object") this.names = parsed;
    }

    saveNamesSoon() {
        if (!this.namesPath || this.saveTimer) return;
        this.saveTimer = setTimeout(async () => {
            this.saveTimer = null;
            try {
                const users = Object.entries(this.names.users).sort((a, b) => b[1].at - a[1].at).slice(0, 2000);
                const chats = Object.entries(this.names.chats).sort((a, b) => b[1].at - a[1].at).slice(0, 1000);
                await writeJsonAtomic(this.namesPath, { users: Object.fromEntries(users), chats: Object.fromEntries(chats) });
            } catch (error) {
                this.debug(`name cache write failed: ${error?.message ?? error}`);
            }
        }, 2000);
        this.saveTimer.unref?.();
    }

    rememberUser(id, name) {
        const key = String(id).toLowerCase();
        const current = this.names.users[key];
        if (current?.name === name && Date.now() - current.at < NAME_TTL_MS / 2) return;
        this.names.users[key] = { name, at: Date.now() };
        this.saveNamesSoon();
    }

    otherUserId(chatId) {
        const match = ONE_ON_ONE.exec(chatId);
        if (!match || !this.me) return null;
        const me = this.me.id.toLowerCase();
        const [a, b] = [match[1].toLowerCase(), match[2].toLowerCase()];
        if (a === me) return b;
        if (b === me) return a;
        return null;
    }

    titleFor(chat) {
        if (chat.topic) return chat.topic;
        if (chat.type === "oneOnOne") {
            const other = this.otherUserId(chat.id);
            const known = other && this.names.users[other]?.name;
            if (known) return known;
            // Bots and external users often aren't resolvable by id, but the
            // last message from them carries their display name.
            if (chat.last?.fromName && chat.last.fromId && chat.last.fromId !== this.me?.id) return chat.last.fromName;
        }
        const names = this.names.chats[chat.id]?.names;
        if (names?.length) {
            if (names.length === 1) return names[0];
            const shown = names.slice(0, 3).map(firstName).join(", ");
            return names.length > 3 ? `${shown} +${names.length - 3}` : shown;
        }
        if (chat.type === "oneOnOne") return "Chat";
        return chat.type === "meeting" ? "Meeting chat" : "Group chat";
    }

    needsNames(chat) {
        if (chat.topic) return false;
        if (chat.type === "oneOnOne") {
            const other = this.otherUserId(chat.id);
            if (other && this.names.users[other]?.name) return false;
        } else {
            const cached = this.names.chats[chat.id];
            if (cached && Date.now() - cached.at < NAME_TTL_MS) return false;
        }
        const missedAt = this.nameMisses.get(chat.id);
        return !(missedAt && Date.now() - missedAt < NAME_RETRY_MS);
    }

    async resolveNames() {
        return this.once("names", async () => {
            const meId = this.me?.id?.toLowerCase();
            if (!meId) return;
            const pending = [...this.chats.values()]
                .filter((chat) => !chat.hidden && this.needsNames(chat))
                .sort((a, b) => time(b.activityAt) - time(a.activityAt))
                .slice(0, 24);
            for (let i = 0; i < pending.length; i += NAME_BATCH) {
                const batch = pending.slice(i, i + NAME_BATCH);
                let results;
                try {
                    results = await this.wq.fetchMany(batch.map((chat) => `/chats/${chat.id}/members`));
                } catch (error) {
                    this.debug(`member lookup failed: ${error?.message ?? error}`);
                    return;
                }
                batch.forEach((chat, index) => {
                    const result = results[index];
                    if (!result.ok) {
                        this.nameMisses.set(chat.id, Date.now());
                        return;
                    }
                    const members = result.data?.value ?? [];
                    for (const member of members) {
                        if (member.userId && member.displayName) this.rememberUser(member.userId, member.displayName);
                    }
                    const others = members
                        .filter((member) => String(member.userId ?? "").toLowerCase() !== meId)
                        .map((member) => member.displayName)
                        .filter(Boolean);
                    this.names.chats[chat.id] = { names: others.slice(0, 12), count: members.length, at: Date.now() };
                });
                this.saveNamesSoon();
                this.emitChats();
            }
        });
    }

    // ---- chat list -------------------------------------------------------

    isUnread(chat) {
        if (chat.markedUnread) return true;
        const last = chat.last;
        if (!last || last.deleted || last.kind !== "message") return false;
        if (this.me && last.fromId === this.me.id) return false;
        if (!chat.lastReadAt) return true;
        return time(last.at) > time(chat.lastReadAt);
    }

    mergeChats(list) {
        const changed = [];
        const seeded = [];
        for (const raw of list) {
            if (!raw?.id) continue;
            const previous = this.chats.get(raw.id);
            const preview = raw.lastMessagePreview ?? null;
            const next = {
                id: raw.id,
                type: raw.chatType ?? "group",
                topic: raw.topic?.trim() || null,
                webUrl: safeUrl(raw.webUrl),
                hidden: Boolean(raw.viewpoint?.isHidden),
                lastReadAt: raw.viewpoint?.lastMessageReadDateTime ?? null,
                preview,
                last: preview
                    ? {
                          id: String(preview.id),
                          at: preview.createdDateTime ?? null,
                          fromId: preview.from?.user?.id ?? preview.from?.application?.id ?? null,
                          fromName: identityName(preview.from),
                          kind: preview.messageType ?? "message",
                          deleted: Boolean(preview.isDeleted),
                          text: previewOf(preview),
                      }
                    : null,
                activityAt: preview?.createdDateTime ?? raw.lastUpdatedDateTime ?? raw.createdDateTime ?? null,
                markedUnread: false,
            };
            if (
                previous?.markedUnread &&
                previous.last?.id === next.last?.id &&
                !(next.lastReadAt && next.last && time(next.lastReadAt) >= time(next.last.at))
            ) {
                next.markedUnread = true;
            }
            // History may already hold a newer message (e.g. one just sent) than
            // the chat-list preview, which can lag behind by a few seconds.
            if (previous?.last && next.last && time(previous.last.at) > time(next.last.at)) {
                next.last = previous.last;
                next.activityAt = previous.activityAt;
            }
            this.learnFrom(preview);
            this.chats.set(next.id, next);
            if (next.last && next.last.id !== previous?.last?.id) {
                changed.push(next.id);
                const history = this.histories.get(next.id);
                if (history && preview && !history.byId.has(String(preview.id))) {
                    this.addMessages(history, [preview], true);
                    // Only the newest preview is known; anything between it and the
                    // last loaded page is missing until the newest page reloads.
                    history.stale = true;
                    seeded.push(next.id);
                }
            }
        }
        // Push previews straight away so new messages appear even while the
        // history endpoint is throttled.
        for (const chatId of seeded) this.emit({ type: "messages", chatId });
        return changed;
    }

    view(chat) {
        const meId = this.me?.id;
        return {
            id: chat.id,
            type: chat.type,
            title: this.titleFor(chat),
            topic: chat.topic,
            webUrl: chat.webUrl,
            lastReadAt: chat.lastReadAt,
            unread: this.isUnread(chat),
            memberCount: this.names.chats[chat.id]?.count ?? (chat.type === "oneOnOne" ? 2 : null),
            activityAt: chat.activityAt,
            last: chat.last
                ? {
                      id: chat.last.id,
                      at: chat.last.at,
                      fromName: chat.last.fromName,
                      mine: Boolean(meId && chat.last.fromId === meId),
                      kind: chat.last.kind,
                      text: chat.last.text,
                  }
                : null,
        };
    }

    listChats() {
        return [...this.chats.values()]
            .filter((chat) => !chat.hidden || this.isUnread(chat))
            .sort((a, b) => time(b.activityAt) - time(a.activityAt))
            .map((chat) => this.view(chat));
    }

    emitChats(force = false) {
        const chats = this.listChats();
        const signature = JSON.stringify([chats.map((c) => [c.id, c.title, c.unread, c.last?.id, c.last?.text]), Boolean(this.moreLink)]);
        if (!force && signature === this.chatsSignature) return;
        this.chatsSignature = signature;
        this.emit({ type: "chats" });
    }

    snapshot() {
        return {
            me: this.me ? { id: this.me.id, name: this.me.name } : null,
            loaded: this.loaded,
            chats: this.listChats(),
            hasMore: Boolean(this.moreLink),
            status: this.status,
            pending: [...this.pendingSends.keys()],
        };
    }

    async refreshChats({ full = false } = {}) {
        const wantFull = full || !this.loaded;
        return this.once(wantFull ? "chats-full" : "chats-poll", async () => {
            if (!this.loaded) this.setStatus({ loading: true });
            let data;
            try {
                await this.ensureMe();
                await this.ready;
                data = await this.wq.fetchOne(chatsPath(wantFull ? FULL_PAGE : POLL_PAGE));
            } catch (error) {
                this.setStatus({ loading: false, error: toStatusError(error) });
                throw error;
            }
            const changed = this.mergeChats(data?.value ?? []);
            if (wantFull) {
                this.moreLink = relativeLink(data?.["@odata.nextLink"]);
                this.fullAt = Date.now();
            }
            this.loaded = true;
            this.setStatus({ loading: false, error: null });
            this.emitChats();
            this.resolveNames().catch((error) => this.debug(`names failed: ${error?.message ?? error}`));
            return changed;
        });
    }

    async loadMoreChats() {
        if (!this.moreLink) return false;
        return this.once("chats-more", async () => {
            const data = await this.wq.fetchOne(this.moreLink);
            this.mergeChats(data?.value ?? []);
            this.moreLink = relativeLink(data?.["@odata.nextLink"]);
            this.emitChats(true);
            this.resolveNames().catch((error) => this.debug(`names failed: ${error?.message ?? error}`));
            return true;
        });
    }

    async findChat(chatId) {
        const existing = this.chats.get(chatId);
        if (existing) return existing;
        await this.ensureMe();
        const data = await this.wq.fetchOne(`/me/chats/${chatId}?$expand=lastMessagePreview&$select=${CHAT_SELECT}`);
        this.mergeChats([data]);
        this.emitChats();
        this.resolveNames().catch(() => {});
        return this.chats.get(chatId);
    }

    prime() {
        this.refreshChats({ full: !this.loaded }).catch(() => {});
    }

    // ---- history ---------------------------------------------------------

    history(chatId) {
        let entry = this.histories.get(chatId);
        if (!entry) {
            entry = { byId: new Map(), olderLink: null, fetchedAt: 0, error: null };
            this.histories.set(chatId, entry);
            const preview = this.chats.get(chatId)?.preview;
            if (preview) this.addMessages(entry, [preview], true);
        }
        return entry;
    }

    addMessages(entry, list, partial = false) {
        const meId = this.me?.id;
        for (const raw of list) {
            if (!raw?.id) continue;
            const id = String(raw.id);
            const existing = entry.byId.get(id);
            if (partial && existing && !existing.partial) continue;
            entry.byId.set(id, normMessage(raw, meId, partial));
            this.learnFrom(raw);
        }
    }

    messagesView(chatId) {
        const entry = this.histories.get(chatId);
        if (!entry) return { chatId, messages: [], hasOlder: false, error: null, olderError: null, loadedAt: 0, complete: false };
        const messages = [...entry.byId.values()].sort((a, b) => time(a.at) - time(b.at) || a.id.localeCompare(b.id));
        const lastId = this.chats.get(chatId)?.last?.id;
        return {
            chatId,
            messages,
            hasOlder: Boolean(entry.olderLink),
            error: entry.error,
            olderError: entry.olderError ?? null,
            loadedAt: entry.fetchedAt,
            // True only when the newest page loaded successfully and nothing newer
            // has arrived since; while history is throttled the view may hold just
            // the chat-list preview, which can hide messages in between.
            complete: Boolean(entry.fetchedAt) && !entry.error && !entry.stale && (!lastId || entry.byId.get(lastId)?.partial === false),
        };
    }

    syncLastFromHistory(chatId) {
        const chat = this.chats.get(chatId);
        const entry = this.histories.get(chatId);
        if (!chat || !entry) return;
        let newest = null;
        for (const message of entry.byId.values()) {
            if (message.kind === "message" && (!newest || time(message.at) > time(newest.at))) newest = message;
        }
        if (!newest || (chat.last && time(chat.last.at) >= time(newest.at))) return;
        chat.last = {
            id: newest.id,
            at: newest.at,
            fromId: newest.fromId,
            fromName: newest.fromName,
            kind: newest.kind,
            deleted: newest.deleted,
            text: newest.deleted ? "This message was deleted" : (newest.html != null ? htmlToText(newest.html) : newest.text ?? "").replace(/\s+/g, " ").slice(0, 180),
        };
        chat.activityAt = newest.at;
        this.emitChats();
    }

    scheduleHistoryRetry(chatId, delayMs) {
        clearTimeout(this.retryTimers.get(chatId));
        const timer = setTimeout(() => {
            this.retryTimers.delete(chatId);
            if (this.watchedChats().has(chatId)) this.loadMessages(chatId, { force: true }).catch(() => {});
        }, Math.max(delayMs, 1000) + 500);
        timer.unref?.();
        this.retryTimers.set(chatId, timer);
    }

    async loadMessages(chatId, { force = false } = {}) {
        await this.findChat(chatId);
        const entry = this.history(chatId);
        if (!force && entry.fetchedAt && Date.now() - entry.fetchedAt < HISTORY_STALE_MS) return this.messagesView(chatId);
        return this.once(`history:${chatId}`, async () => {
            await this.ensureMe();
            const first = !entry.fetchedAt;
            try {
                const data = await this.wq.fetchOne(
                    `/chats/${chatId}/messages?$top=${first ? HISTORY_PAGE : REFRESH_PAGE}&$orderby=createdDateTime desc`,
                );
                const list = data?.value ?? [];
                const overlaps = list.some((m) => entry.byId.get(String(m.id))?.partial === false);
                if (!first && !overlaps && list.length >= REFRESH_PAGE) {
                    entry.byId.clear();
                    entry.olderLink = relativeLink(data?.["@odata.nextLink"]);
                }
                this.addMessages(entry, list);
                if (first) entry.olderLink = relativeLink(data?.["@odata.nextLink"]);
                entry.fetchedAt = Date.now();
                entry.error = null;
                entry.stale = false;
                this.historyFailures = 0;
                this.syncLastFromHistory(chatId);
            } catch (error) {
                if (!(error instanceof WorkIQError)) throw error;
                entry.error = toStatusError(error);
                if (entry.error.throttled || error.transient) {
                    // Back off exponentially on real throttling so retries don't keep
                    // WorkIQ's circuit breaker open; local cooldown hits don't count.
                    if (error.code !== "cooldown") this.historyFailures = (this.historyFailures ?? 0) + 1;
                    const failures = Math.max(this.historyFailures ?? 0, 1);
                    const delay = Math.min(5 * 60_000, Math.max(error.retryAfterMs || 0, 15_000 * 2 ** (failures - 1)));
                    entry.error.retryAt = Date.now() + delay;
                    this.scheduleHistoryRetry(chatId, delay);
                }
            }
            this.emit({ type: "messages", chatId });
            return this.messagesView(chatId);
        });
    }

    async loadOlder(chatId) {
        const entry = this.history(chatId);
        if (!entry.olderLink) return this.messagesView(chatId);
        return this.once(`older:${chatId}`, async () => {
            try {
                const data = await this.wq.fetchOne(entry.olderLink);
                this.addMessages(entry, data?.value ?? []);
                entry.olderLink = relativeLink(data?.["@odata.nextLink"]);
                entry.olderError = null;
            } catch (error) {
                if (!(error instanceof WorkIQError)) throw error;
                // Kept apart from the newest-page error, which gates auto mark-read.
                entry.olderError = toStatusError(error);
            }
            this.emit({ type: "messages", chatId });
            return this.messagesView(chatId);
        });
    }

    // ---- writes ----------------------------------------------------------

    tenantFor(chat) {
        if (this.me?.tenantId) return this.me.tenantId;
        const user = chat?.preview?.from?.user;
        if (user?.id && user.id === this.me?.id && user.tenantId) return user.tenantId;
        try {
            return chat?.webUrl ? new URL(chat.webUrl).searchParams.get("tenantId") : null;
        } catch {
            return null;
        }
    }

    inlineQuote(chatId, messageId) {
        const message = this.histories.get(chatId)?.byId.get(String(messageId));
        if (!message || message.deleted) return "";
        const text = (message.html != null ? htmlToText(message.html) : message.text ?? "").trim();
        if (!text) return "";
        return `<blockquote><b>${escapeHtml(message.fromName ?? "Message")}</b><br>${escapeHtml(text.slice(0, 300))}</blockquote>`;
    }

    async postMessage(chatId, html, quoteId) {
        const parent = `/chats/${chatId}/messages`;
        if (!quoteId) return this.wq.create(parent, { body: { contentType: "html", content: html } });
        try {
            return await this.wq.action(`${parent}/replyWithQuote`, {
                messageIds: [String(quoteId)],
                replyMessage: { body: { contentType: "html", content: html } },
            });
        } catch (error) {
            if (!(error instanceof WorkIQError) || error.ambiguous) throw error;
            // Fall back only when the quote endpoint itself is unavailable. Any other
            // rejection (e.g. a stale or inaccessible quoted message) fails closed so
            // a quoted reply never silently becomes a plain message.
            const unsupported = [405, 501].includes(error.status) || error.code === "path_denied";
            if (!unsupported) {
                if (!error.throttled && error.code !== "cooldown" && error.status >= 400 && error.status < 500) {
                    throw new WorkIQError(`Teams rejected the quoted reply (${error.message}). Remove the quote to send it as a normal message.`, {
                        status: error.status,
                        code: "quote_rejected",
                    });
                }
                throw error;
            }
            const quote = this.inlineQuote(chatId, quoteId);
            if (!quote) {
                throw new WorkIQError("Teams can't send quoted replies here, and the quoted message isn't loaded. Remove the quote, or reload the chat and try again.", {
                    code: "quote_unavailable",
                });
            }
            this.debug(`replyWithQuote unavailable (${error.status} ${error.message}); sending an inline quote`);
            return this.wq.create(parent, { body: { contentType: "html", content: quote + html } });
        }
    }

    async send(chatId, text, { quoteId = null, origin = null, draftRev = null } = {}) {
        const body = String(text ?? "").trim();
        if (!body) throw new WorkIQError("Type a message first.", { code: "empty" });
        if (body.length > MAX_MESSAGE_CHARS) throw new WorkIQError("That message is too long for Teams.", { code: "too_long" });
        if (this.pendingSends.has(chatId)) {
            throw new WorkIQError("Your previous message here is still waiting for Teams to confirm it. Check the conversation before sending again.", {
                code: "send_pending",
            });
        }
        await this.ensureMe();
        const chat = await this.findChat(chatId);
        const pending = {
            text: body,
            origin,
            draftRev: Number.isInteger(draftRev) ? draftRev : null,
            startedAt: Date.now(),
            prevLastId: chat?.last?.id ?? null,
            knownIds: new Set(this.histories.get(chatId)?.byId.keys() ?? []),
        };
        this.pendingSends.set(chatId, pending);
        let created;
        try {
            created = await this.postMessage(chatId, textToHtml(body), quoteId);
        } catch (error) {
            if (error instanceof WorkIQError && error.ambiguous) this.resolveUnconfirmed(chatId, pending, error);
            else this.pendingSends.delete(chatId);
            throw error;
        }
        this.pendingSends.delete(chatId);
        return this.applySent(chatId, created, pending);
    }

    applySent(chatId, created, pending) {
        const me = this.me;
        const entry = this.history(chatId);
        const chat = this.chats.get(chatId);
        if (created?.id) {
            this.addMessages(entry, [created]);
            if (chat) {
                const at = created.createdDateTime ?? new Date().toISOString();
                chat.last = { id: String(created.id), at, fromId: me.id, fromName: me.name, kind: "message", deleted: false, text: previewOf(created) || pending.text.slice(0, 180) };
                chat.preview = created;
                chat.activityAt = at;
                chat.lastReadAt = at;
                chat.markedUnread = false;
                chat.hidden = false;
            }
        } else {
            entry.fetchedAt = 0;
        }
        // Clear the draft only if it still represents what was sent: unchanged
        // since the send started, identical text, or already empty. A newer draft
        // written meanwhile (another panel, or Copilot) is left alone.
        const draft = this.draftState(chatId);
        if (draft.rev === pending.draftRev || draft.text.trim() === pending.text || !draft.text.trim()) {
            this.writeDraft(chatId, "", null, pending.origin, { sentText: pending.text });
        }
        this.emit({ type: "messages", chatId });
        this.emit({ type: "send_state", chatId, origin: pending.origin, state: "delivered", text: pending.text });
        this.emitChats(true);
        if (!created?.id) this.loadMessages(chatId, { force: true }).catch(() => {});
        return created?.id ? normMessage(created, me.id) : null;
    }

    // Ends a pending send exactly once; whichever outcome arrives first wins.
    settlePending(chatId, pending, state, { created = null, message = null } = {}) {
        if (this.pendingSends.get(chatId) !== pending) return false;
        this.pendingSends.delete(chatId);
        if (state === "delivered") this.applySent(chatId, created, pending);
        else this.emit({ type: "send_state", chatId, origin: pending.origin, state, message });
        return true;
    }

    // A send whose outcome is unknown keeps blocking new sends to that chat. The
    // original call's late result and chat reconciliation run side by side, so a
    // call that never settles can't leave the chat blocked forever.
    resolveUnconfirmed(chatId, pending, error) {
        this.emit({ type: "send_state", chatId, origin: pending.origin, state: "pending" });
        error.late?.then(
            (created) => this.settlePending(chatId, pending, "delivered", { created }),
            (lateError) => {
                if (lateError instanceof WorkIQError && !lateError.ambiguous) {
                    this.settlePending(chatId, pending, "failed", {
                        message: `Teams didn't accept your last message${lateError.message ? ` (${lateError.message})` : ""}. It wasn't sent.`,
                    });
                }
            },
        );
        this.reconcileSend(chatId, pending);
    }

    sentText(raw) {
        const content = raw?.body?.content;
        const text = raw?.body?.contentType === "html" ? htmlToText(content) : content;
        return String(text ?? "").replace(/\s+/g, " ").trim();
    }

    // Find the sent message in the chat-list preview or in reloaded history.
    findSent(chatId, pending) {
        if (!this.me) return null;
        const sent = pending.text.replace(/\s+/g, " ").trim();
        const matches = (raw) => {
            if (!raw?.id || raw.from?.user?.id !== this.me.id) return false;
            if (String(raw.id) === pending.prevLastId || pending.knownIds.has(String(raw.id))) return false;
            if (time(raw.createdDateTime) < pending.startedAt - 30_000) return false;
            const posted = this.sentText(raw);
            // An inline-quote fallback prefixes the quoted text.
            return posted === sent || posted.endsWith(sent);
        };
        const preview = this.chats.get(chatId)?.preview;
        if (matches(preview)) return preview;
        for (const message of this.histories.get(chatId)?.byId.values() ?? []) {
            if (message.partial || !message.mine) continue;
            const raw = {
                id: message.id,
                createdDateTime: message.at,
                from: { user: { id: message.fromId } },
                body: message.html != null ? { contentType: "html", content: message.html } : { contentType: "text", content: message.text },
            };
            if (matches(raw)) return { ...raw, from: { user: { id: this.me.id, displayName: this.me.name } } };
        }
        return null;
    }

    reconcileSend(chatId, pending, attempt = 0) {
        const delays = this.reconcileDelays;
        if (attempt >= delays.length) {
            this.settlePending(chatId, pending, "unknown", {
                message: "Teams never confirmed your last message. Check the conversation in Teams before sending it again.",
            });
            return;
        }
        const timer = setTimeout(async () => {
            if (this.pendingSends.get(chatId) !== pending) return;
            try {
                await this.refreshChats();
            } catch {
                // Try again on the next attempt.
            }
            // A newer incoming message can push the sent one out of the preview,
            // so also look at the newest history page when it isn't throttled.
            if (this.pendingSends.get(chatId) === pending && !this.findSent(chatId, pending) && !this.wq.cooling("messages")) {
                await this.loadMessages(chatId, { force: true }).catch(() => {});
            }
            if (this.pendingSends.get(chatId) !== pending) return;
            const found = this.findSent(chatId, pending);
            if (found) {
                this.settlePending(chatId, pending, "delivered", { created: found });
                return;
            }
            this.reconcileSend(chatId, pending, attempt + 1);
        }, delays[attempt]);
        timer.unref?.();
    }

    async markRead(chatId) {
        const chat = await this.findChat(chatId);
        if (!chat || !this.isUnread(chat)) return false;
        return this.once(`read:${chatId}`, async () => {
            const me = await this.ensureMe();
            const tenantId = this.tenantFor(chat);
            if (!tenantId) throw new WorkIQError("Couldn't determine your tenant to mark this chat read.", { code: "no_tenant" });
            await this.wq.action(`/chats/${chatId}/markChatReadForUser`, { user: { id: me.id, tenantId } });
            chat.lastReadAt = chat.last?.at && time(chat.last.at) > time(chat.lastReadAt) ? chat.last.at : chat.lastReadAt ?? new Date().toISOString();
            chat.markedUnread = false;
            this.emitChats(true);
            return true;
        });
    }

    async markUnread(chatId) {
        const chat = await this.findChat(chatId);
        if (!chat?.last) return false;
        return this.once(`unread:${chatId}`, async () => {
            const me = await this.ensureMe();
            const tenantId = this.tenantFor(chat);
            if (!tenantId) throw new WorkIQError("Couldn't determine your tenant to mark this chat unread.", { code: "no_tenant" });
            const before = new Date(time(chat.last.at) - 1000).toISOString();
            await this.wq.action(`/chats/${chatId}/markChatUnreadForUser`, {
                user: { id: me.id, tenantId },
                lastMessageReadDateTime: before,
            });
            chat.lastReadAt = before;
            chat.markedUnread = true;
            this.emitChats(true);
            return true;
        });
    }

    // ---- drafts ------------------------------------------------------------
    // Each chat's draft has a revision. Client edits apply only on top of the
    // revision they last saw, so stale or out-of-order writes can't overwrite
    // newer text or bring back a message that was already sent. Drafts persist
    // under the user's Copilot home, one atomic file per chat, so they survive
    // extension restarts and other sessions' copies of this extension can't
    // clobber each other's updates.

    draftRev(chatId) {
        return this.draftRevs.get(chatId) ?? 0;
    }

    draftState(chatId) {
        const draft = this.drafts.get(chatId);
        return { text: draft?.text ?? "", quoteId: draft?.quoteId ?? null, rev: this.draftRev(chatId) };
    }

    writeDraft(chatId, text, quoteId, origin = null, { sentText = null } = {}) {
        const rev = this.draftRev(chatId) + 1;
        this.draftRevs.set(chatId, rev);
        const value = String(text ?? "").slice(0, MAX_DRAFT_CHARS);
        if (value.trim() || quoteId) this.drafts.set(chatId, { text: value, quoteId: quoteId ? String(quoteId) : null, at: Date.now() });
        else this.drafts.delete(chatId);
        this.draftTouched.set(chatId, Date.now());
        this.saveDraftsSoon();
        this.emit({ type: "draft", chatId, origin, sentText });
        return rev;
    }

    updateDraft(chatId, text, quoteId, baseRev, origin) {
        if (baseRev !== this.draftRev(chatId)) return { ok: false, ...this.draftState(chatId) };
        return { ok: true, rev: this.writeDraft(chatId, text, quoteId, origin) };
    }

    draftsView() {
        const view = {};
        for (const chatId of new Set([...this.draftRevs.keys(), ...this.drafts.keys()])) view[chatId] = this.draftState(chatId);
        return view;
    }

    draftFile(chatId) {
        return join(this.draftsDir, `${createHash("sha256").update(chatId).digest("hex").slice(0, 40)}.json`);
    }

    draftRecord(chatId, at) {
        const draft = this.drafts.get(chatId);
        return draft ? { chatId, text: draft.text, quoteId: draft.quoteId, at } : { chatId, deleted: true, at };
    }

    async loadDrafts() {
        if (!this.draftsDir) return;
        let files = [];
        try {
            files = (await readdir(this.draftsDir)).filter((name) => name.endsWith(".json"));
        } catch {
            return;
        }
        const cutoff = Date.now() - DRAFT_TTL_MS;
        for (const name of files) {
            const path = join(this.draftsDir, name);
            const entry = await readJson(path);
            if (!entry || typeof entry.chatId !== "string" || !(entry.at > cutoff)) {
                await unlink(path).catch(() => {});
                continue;
            }
            if (entry.deleted || typeof entry.text !== "string" || this.drafts.has(entry.chatId)) continue;
            this.drafts.set(entry.chatId, { text: entry.text.slice(0, MAX_DRAFT_CHARS), quoteId: typeof entry.quoteId === "string" ? entry.quoteId : null, at: entry.at });
            this.draftRevs.set(entry.chatId, this.draftRev(entry.chatId) + 1);
        }
    }

    // Newest write wins per chat; deletions leave a tombstone so an older copy
    // held by another session can't bring the draft back.
    async saveDrafts() {
        if (!this.draftsDir || !this.draftTouched.size) return;
        const touched = [...this.draftTouched];
        this.draftTouched.clear();
        for (const [chatId, at] of touched) {
            try {
                const path = this.draftFile(chatId);
                const existing = await readJson(path);
                if (existing?.at > at) continue;
                await writeJsonAtomic(path, this.draftRecord(chatId, at));
            } catch (error) {
                // Retry with the next save unless a newer edit already supersedes it.
                if (!(this.draftTouched.get(chatId) > at)) this.draftTouched.set(chatId, at);
                this.debug(`draft save failed: ${error?.message ?? error}`);
            }
        }
    }

    saveDraftsSoon() {
        if (!this.draftsDir || this.draftSaveTimer) return;
        this.draftSaveTimer = setTimeout(() => {
            this.draftSaveTimer = null;
            // Serialize saves so an older snapshot can never land after a newer one.
            this.draftSaveChain = (this.draftSaveChain ?? Promise.resolve())
                .then(() => this.saveDrafts())
                .catch((error) => this.debug(`draft save failed: ${error?.message ?? error}`));
        }, 600);
        this.draftSaveTimer.unref?.();
    }

    flushDraftsSync() {
        if (!this.draftsDir || !this.draftTouched.size) return;
        clearTimeout(this.draftSaveTimer);
        this.draftSaveTimer = null;
        mkdirSync(this.draftsDir, { recursive: true, mode: 0o700 });
        for (const [chatId, at] of this.draftTouched) {
            const path = this.draftFile(chatId);
            let existing = null;
            try {
                existing = JSON.parse(readFileSync(path, "utf8"));
            } catch {
                existing = null;
            }
            if (existing?.at > at) continue;
            const temp = `${path}.${process.pid}.${Date.now()}.tmp`;
            writeFileSync(temp, JSON.stringify(this.draftRecord(chatId, at)), { mode: 0o600 });
            renameSync(temp, path);
        }
        this.draftTouched.clear();
    }

    // ---- polling while a canvas is connected ------------------------------

    setActive(clientId, chatId) {
        if (chatId) this.active.set(clientId, chatId);
        else this.active.delete(clientId);
    }

    dropClient(clientId) {
        this.active.delete(clientId);
    }

    watchedChats() {
        return new Set(this.active.values());
    }

    setWatching(on) {
        if (on === this.watching) return;
        this.watching = on;
        clearTimeout(this.pollTimer);
        this.pollTimer = null;
        if (on) this.schedulePoll(this.loaded ? 2000 : 0);
    }

    schedulePoll(delayMs) {
        clearTimeout(this.pollTimer);
        this.pollTimer = setTimeout(() => this.pollTick(), delayMs);
        this.pollTimer.unref?.();
    }

    async pollTick() {
        if (!this.watching) return;
        let delay = POLL_MS;
        try {
            const changed = new Set(await this.refreshChats({ full: !this.loaded || Date.now() - this.fullAt > FULL_REFRESH_MS }));
            for (const chatId of this.watchedChats()) {
                const entry = this.histories.get(chatId);
                if (!entry || this.wq.cooling("messages")) continue;
                if (entry.error?.retryAt && Date.now() < entry.error.retryAt) continue;
                if (changed.has(chatId) || !entry.fetchedAt || Date.now() - entry.fetchedAt > HISTORY_STALE_MS) {
                    this.loadMessages(chatId, { force: true }).catch(() => {});
                }
            }
        } catch (error) {
            const permanent = ["not_connected", "needs_auth", "denied"].includes(error?.code);
            delay = permanent ? 60_000 : Math.max(POLL_MS, error?.retryAfterMs ?? 0);
        }
        if (this.watching) this.schedulePoll(delay);
    }

    stop() {
        this.setWatching(false);
        for (const timer of this.retryTimers.values()) clearTimeout(timer);
        this.retryTimers.clear();
    }

    // ---- agent helpers ---------------------------------------------------

    matchChats(query) {
        const needle = String(query ?? "").trim().toLowerCase();
        if (!needle) return [];
        const chats = this.listChats();
        const exact = chats.filter((c) => c.title.toLowerCase() === needle);
        if (exact.length) return exact;
        const prefix = chats.filter((c) => c.title.toLowerCase().startsWith(needle));
        if (prefix.length) return prefix;
        return chats.filter((c) => c.title.toLowerCase().includes(needle));
    }

    transcript(chatId, limit = 20) {
        const view = this.messagesView(chatId);
        return view.messages.slice(-limit).map((m) => ({
            id: m.id,
            at: m.at,
            from: m.mine ? `${this.me?.name ?? "Me"} (me)` : m.fromName ?? (m.kind === "message" ? "Unknown" : "System"),
            text: m.deleted
                ? "(deleted)"
                : m.kind !== "message"
                  ? m.event
                  : [
                        ...m.attachments.filter((a) => a.kind === "quote").map((a) => `> ${a.sender ?? ""}: ${a.preview}`),
                        m.html != null ? htmlToText(m.html) : m.text ?? "",
                        ...m.attachments.filter((a) => a.kind !== "quote").map((a) => `[${a.kind}: ${a.name}]`),
                    ]
                        .filter(Boolean)
                        .join("\n"),
        }));
    }
}
