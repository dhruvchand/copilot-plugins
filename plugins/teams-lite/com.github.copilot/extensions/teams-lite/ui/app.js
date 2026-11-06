(() => {
    "use strict";

    const config = JSON.parse(document.getElementById("tl-config").textContent);
    const TOKEN = config.token;
    const INSTANCE = config.instanceId || "";
    const CLIENT_ID = crypto.randomUUID ? crypto.randomUUID() : `${Date.now().toString(36)}${Math.random().toString(36).slice(2)}`;
    const SVG_NS = "http://www.w3.org/2000/svg";
    const ICONS = {
        reply: "M10 9V5l-7 7 7 7v-4.1c5 0 8.5 1.6 11 5.1-1-5-4-10-11-11z",
        copy: "M16 1H4a2 2 0 0 0-2 2v14h2V3h12V1zm3 4H8a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h11a2 2 0 0 0 2-2V7a2 2 0 0 0-2-2zm0 16H8V7h11v14z",
        file: "M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8l-6-6zm4 18H6V4h7v5h5v11z",
        card: "M20 4H4a2 2 0 0 0-2 2v12a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2V6a2 2 0 0 0-2-2zm0 14H4v-6h16v6zm0-10H4V6h16v2z",
        markRead: "M21.99 8c0-.72-.37-1.35-.94-1.7L12 1 2.95 6.3C2.38 6.65 2 7.28 2 8v10c0 1.1.9 2 2 2h16c1.1 0 2-.9 2-2l-.01-10zM12 13 3.74 7.84 12 3l8.26 4.84L12 13z",
        markUnread: "M20 4H4a2 2 0 0 0-2 2v12a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2V6a2 2 0 0 0-2-2zm0 4-8 5-8-5V6l8 5 8-5v2z",
    };

    const $ = (selector) => document.querySelector(selector);
    const els = {
        app: $("#app"),
        list: $("#chat-list"),
        search: $("#search"),
        refresh: $("#refresh"),
        countUnread: $("#count-unread"),
        banner: $("#banner"),
        tabs: [...document.querySelectorAll(".tabs [data-filter]")],
        emptyPane: $("#empty-pane"),
        chatPane: $("#chat-pane"),
        back: $("#back"),
        chatAvatar: $("#chat-avatar"),
        chatTitle: $("#chat-title"),
        chatSubtitle: $("#chat-subtitle"),
        toggleRead: $("#toggle-read"),
        toggleReadIcon: $("#toggle-read path"),
        openTeams: $("#open-teams"),
        reloadChat: $("#reload-chat"),
        historyBanner: $("#history-banner"),
        messages: $("#messages"),
        composer: $("#composer"),
        input: $("#input"),
        send: $("#send"),
        draftBar: $("#draft-bar"),
        draftBarLabel: $("#draft-bar-label"),
        draftBarText: $("#draft-bar-text"),
        draftUse: $("#draft-use"),
        draftDismiss: $("#draft-dismiss"),
        quoteBar: $("#quote-bar"),
        quoteName: $("#quote-name"),
        quoteText: $("#quote-text"),
        quoteCancel: $("#quote-cancel"),
        sendError: $("#send-error"),
    };

    const state = {
        me: null,
        loaded: false,
        chats: [],
        hasMore: false,
        status: { loading: true, error: null },
        connected: true,
        filter: "unread",
        query: "",
        activeId: null,
        histories: new Map(),
        drafts: new Map(),
        dividerAt: new Map(),
        sending: new Set(),
        pendingChats: new Set(),
        sendErrors: new Map(),
        engaged: false,
        refreshing: false,
        loadingOlder: false,
        keepOffset: false,
        pinned: true,
    };

    // ---- helpers ---------------------------------------------------------

    function h(tag, props, ...children) {
        const el = document.createElement(tag);
        if (props) {
            for (const [key, value] of Object.entries(props)) {
                if (value == null || value === false) continue;
                if (key === "class") el.className = value;
                else if (key === "dataset") Object.assign(el.dataset, value);
                else if (key.startsWith("on")) el.addEventListener(key.slice(2), value);
                else el.setAttribute(key, value === true ? "" : String(value));
            }
        }
        for (const child of children.flat()) {
            if (child == null || child === false) continue;
            el.append(child instanceof Node ? child : document.createTextNode(String(child)));
        }
        return el;
    }

    function icon(name, size = 16) {
        const svg = document.createElementNS(SVG_NS, "svg");
        svg.setAttribute("viewBox", "0 0 24 24");
        svg.setAttribute("width", String(size));
        svg.setAttribute("height", String(size));
        svg.setAttribute("aria-hidden", "true");
        const path = document.createElementNS(SVG_NS, "path");
        path.setAttribute("fill", "currentColor");
        path.setAttribute("d", ICONS[name]);
        svg.append(path);
        return svg;
    }

    async function api(method, path, body) {
        const headers = { "x-tl-token": TOKEN };
        if (body) headers["content-type"] = "application/json";
        const response = await fetch(path, { method, headers, body: body ? JSON.stringify(body) : undefined, cache: "no-store" });
        let data = {};
        try {
            data = await response.json();
        } catch {
            // Non-JSON error bodies fall through to the generic message.
        }
        if (!response.ok) {
            const error = new Error(data.error || `Request failed (${response.status}).`);
            error.status = response.status;
            error.code = data.code;
            error.retryAfterMs = data.retryAfterMs || 0;
            error.data = data;
            throw error;
        }
        return data;
    }

    const chatById = (id) => (id ? state.chats.find((chat) => chat.id === id) ?? null : null);
    const time = (iso) => (iso ? Date.parse(iso) || 0 : 0);
    const isNarrow = () => els.app.classList.contains("narrow");

    const fmtTime = new Intl.DateTimeFormat(undefined, { hour: "numeric", minute: "2-digit" });
    const fmtWeekday = new Intl.DateTimeFormat(undefined, { weekday: "short" });
    const fmtShort = new Intl.DateTimeFormat(undefined, { month: "short", day: "numeric" });
    const fmtLong = new Intl.DateTimeFormat(undefined, { weekday: "long", month: "long", day: "numeric" });
    const fmtFull = new Intl.DateTimeFormat(undefined, { dateStyle: "medium", timeStyle: "short" });

    function daysAgo(iso) {
        const date = new Date(iso);
        const now = new Date();
        const start = (d) => new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
        return Math.round((start(now) - start(date)) / 86_400_000);
    }

    function listTime(iso) {
        if (!iso) return "";
        const days = daysAgo(iso);
        if (days <= 0) return fmtTime.format(new Date(iso));
        if (days === 1) return "Yesterday";
        if (days < 7) return fmtWeekday.format(new Date(iso));
        return fmtShort.format(new Date(iso));
    }

    function dayLabel(iso) {
        const days = daysAgo(iso);
        if (days <= 0) return "Today";
        if (days === 1) return "Yesterday";
        return fmtLong.format(new Date(iso));
    }

    function firstName(name) {
        const trimmed = String(name ?? "").trim();
        if (trimmed.includes(",")) return trimmed.split(",")[1]?.trim().split(/\s+/)[0] || trimmed;
        return trimmed.split(/\s+/)[0] || trimmed;
    }

    function initials(name) {
        const raw = String(name || "");
        const source = raw.replace(/\([^)]*\)|\[[^\]]*\]/g, " ").trim() || raw;
        const words = source.split(/[^\p{L}\p{N}]+/u).filter(Boolean);
        if (!words.length) return "?";
        const letters = words.length === 1 ? [...words[0]].slice(0, 2) : [[...words[0]][0], [...words[words.length - 1]][0]];
        return letters.join("").toUpperCase();
    }

    function hue(seed) {
        let hash = 0;
        for (const ch of String(seed)) hash = (hash * 31 + ch.codePointAt(0)) >>> 0;
        return hash % 360;
    }

    function avatar(name, seed, { small = false, group = false } = {}) {
        const el = h("div", { class: `avatar${small ? " sm" : ""}${group ? " group" : ""}`, "aria-hidden": "true" }, initials(name));
        el.style.background = `hsl(${hue(seed || name)} 42% 42%)`;
        return el;
    }

    // ---- safe rendering of Teams message HTML -----------------------------

    const ALLOWED = new Set([
        "P", "BR", "B", "STRONG", "I", "EM", "U", "S", "STRIKE", "DEL", "INS", "A", "UL", "OL", "LI", "CODE", "PRE",
        "BLOCKQUOTE", "SPAN", "DIV", "H1", "H2", "H3", "H4", "H5", "H6", "TABLE", "THEAD", "TBODY", "TFOOT", "TR",
        "TD", "TH", "HR", "SUB", "SUP", "SMALL", "MARK",
    ]);
    const DROPPED = new Set([
        "SCRIPT", "STYLE", "IFRAME", "FRAME", "FRAMESET", "OBJECT", "EMBED", "TEMPLATE", "NOSCRIPT", "SVG", "MATH", "LINK", "META",
        "TITLE", "HEAD", "FORM", "INPUT", "BUTTON", "SELECT", "TEXTAREA", "VIDEO", "AUDIO", "SOURCE", "CANVAS", "ATTACHMENT", "XML",
    ]);
    const parser = new DOMParser();

    function safeHref(raw) {
        try {
            const url = new URL(raw);
            return ["http:", "https:", "mailto:"].includes(url.protocol) ? url.href : null;
        } catch {
            return null;
        }
    }

    function hostedImagePath(src) {
        try {
            const url = new URL(src);
            if (url.hostname !== "graph.microsoft.com") return null;
            let path = decodeURIComponent(url.pathname).replace(/^\/(?:v1\.0|beta)(?=\/)/, "");
            // Chat-list previews omit the chat id in hosted image URLs.
            if (path.startsWith("/chats//") && state.activeId) path = `/chats/${state.activeId}/${path.slice("/chats//".length)}`;
            return /^\/chats\/[^/?#]+\/messages\/[^/?#]+\/hostedContents\/[^/?#]+\/\$value$/.test(path) ? path : null;
        } catch {
            return null;
        }
    }

    function imagePlaceholder() {
        return h("span", { class: "image-placeholder", title: "Open in Teams to view this image" }, "Image");
    }

    function imageNode(node) {
        const alt = node.getAttribute("alt") || "";
        if (/schema\.skype\.com\/Emoji/i.test(node.getAttribute("itemtype") || "")) return document.createTextNode(alt);
        const path = hostedImagePath(node.getAttribute("src") || "");
        if (!path) return imagePlaceholder();
        const img = h("img", { alt: alt || "Image", loading: "lazy", decoding: "async" });
        const width = Number.parseInt(node.getAttribute("width") || "", 10);
        const height = Number.parseInt(node.getAttribute("height") || "", 10);
        if (width > 0 && height > 0 && width < 5000 && height < 5000) {
            img.width = width;
            img.height = height;
        }
        img.addEventListener("error", () => img.replaceWith(imagePlaceholder()), { once: true });
        img.addEventListener("load", () => {
            if (state.pinned) scrollToBottom();
        }, { once: true });
        img.src = `/api/blob?p=${encodeURIComponent(path)}&t=${encodeURIComponent(TOKEN)}`;
        return img;
    }

    function walk(source, target) {
        for (const node of source.childNodes) {
            if (node.nodeType === Node.TEXT_NODE) {
                target.append(node.nodeValue);
                continue;
            }
            if (node.nodeType !== Node.ELEMENT_NODE) continue;
            const tag = node.tagName.toUpperCase();
            if (DROPPED.has(tag)) continue;
            if (tag === "AT") {
                const mention = h("span", { class: "mention" }, node.textContent);
                if (state.me?.name && node.textContent.trim() === state.me.name) mention.classList.add("me");
                target.append(mention);
                continue;
            }
            if (tag === "EMOJI") {
                target.append(node.getAttribute("alt") || node.textContent || "");
                continue;
            }
            if (tag === "IMG") {
                target.append(imageNode(node));
                continue;
            }
            if (tag === "CODEBLOCK") {
                const pre = h("pre");
                walk(node, pre);
                target.append(pre);
                continue;
            }
            if (!ALLOWED.has(tag)) {
                walk(node, target);
                continue;
            }
            const el = document.createElement(tag.toLowerCase());
            if (tag === "A") {
                const href = safeHref(node.getAttribute("href") || "");
                if (href) {
                    el.setAttribute("href", href);
                    el.setAttribute("rel", "noopener noreferrer");
                    el.dataset.external = "1";
                    el.title = href;
                }
            }
            if (tag === "TD" || tag === "TH") {
                const span = Number.parseInt(node.getAttribute("colspan") || "", 10);
                if (span > 1 && span < 50) el.colSpan = span;
            }
            walk(node, el);
            target.append(el);
        }
    }

    function renderHtml(target, html) {
        const doc = parser.parseFromString(`<!doctype html><body>${html}</body>`, "text/html");
        walk(doc.body, target);
    }

    function plainText(message) {
        if (!message) return "";
        if (message.html != null) {
            const doc = parser.parseFromString(`<!doctype html><body>${message.html}</body>`, "text/html");
            for (const img of doc.querySelectorAll("img")) img.replaceWith(img.getAttribute("alt") || "[image]");
            for (const br of doc.querySelectorAll("br")) br.replaceWith("\n");
            for (const block of doc.querySelectorAll("p, div, li")) block.append("\n");
            return (doc.body.textContent || "").replace(/\n{3,}/g, "\n\n").trim();
        }
        return String(message.text ?? "").trim();
    }

    // ---- chat list -------------------------------------------------------

    function visibleChats() {
        const query = state.query.trim().toLowerCase();
        return state.chats.filter((chat) => {
            if (state.filter === "unread" && !chat.unread && chat.id !== state.activeId) return false;
            if (!query) return true;
            return chat.title.toLowerCase().includes(query) || (chat.last?.text ?? "").toLowerCase().includes(query);
        });
    }

    function previewLine(chat) {
        const last = chat.last;
        if (!last) return "";
        let prefix = "";
        if (last.kind === "message") {
            if (last.mine) prefix = "You: ";
            else if (chat.type !== "oneOnOne" && last.fromName) prefix = `${firstName(last.fromName)}: `;
        }
        return `${prefix}${last.text || ""}`;
    }

    function chatRow(chat) {
        return h(
            "button",
            {
                type: "button",
                class: `chat-item${chat.unread ? " unread" : ""}`,
                "aria-current": chat.id === state.activeId ? "true" : null,
                dataset: { chatId: chat.id },
                title: chat.title,
                onclick: () => openChat(chat.id, { user: true }),
            },
            avatar(chat.title, chat.id, { group: chat.type !== "oneOnOne" }),
            h("div", { class: "chat-main" }, h("div", { class: "title" }, chat.title), h("div", { class: "preview" }, previewLine(chat) || "\u00a0")),
            h(
                "div",
                { class: "meta" },
                h("time", { datetime: chat.last?.at || chat.activityAt || null }, listTime(chat.last?.at || chat.activityAt)),
                chat.unread ? h("span", { class: "unread-dot", role: "img", "aria-label": "Unread" }) : null,
            ),
        );
    }

    function renderList() {
        const unread = state.chats.filter((chat) => chat.unread).length;
        els.countUnread.textContent = unread ? String(unread) : "";
        els.countUnread.hidden = !unread;
        document.title = unread ? `(${unread}) Teams` : "Teams";
        for (const tab of els.tabs) tab.setAttribute("aria-selected", String(tab.dataset.filter === state.filter));

        const focusedId = document.activeElement?.closest?.(".chat-item")?.dataset.chatId ?? null;
        const items = visibleChats();
        const fragment = document.createDocumentFragment();
        if (!state.loaded) {
            if (!state.status.error) for (let i = 0; i < 7; i++) fragment.append(h("li", { class: "skeleton", "aria-hidden": "true" }));
        } else if (!items.length) {
            const text = state.query ? "No chats match your search." : state.filter === "unread" ? "You're all caught up." : "No chats yet.";
            fragment.append(
                h(
                    "li",
                    { class: "empty" },
                    h("div", null, text),
                    state.filter === "unread" && !state.query ? h("button", { type: "button", class: "text-btn", onclick: () => setFilter("all") }, "Show all chats") : null,
                ),
            );
        }
        for (const chat of items) fragment.append(h("li", null, chatRow(chat)));
        if (state.loaded && state.hasMore && state.filter === "all" && !state.query) {
            fragment.append(h("li", { class: "more" }, h("button", { type: "button", class: "text-btn", onclick: loadMoreChats }, "Load more chats")));
        }
        els.list.replaceChildren(fragment);
        if (focusedId) els.list.querySelector(`.chat-item[data-chat-id="${CSS.escape(focusedId)}"]`)?.focus({ preventScroll: true });
    }

    function renderBanner() {
        const error = state.status.error;
        let text = "";
        let tone = "warn";
        let action = null;
        if (!state.connected) {
            text = "Reconnecting to the Teams canvas…";
        } else if (error) {
            tone = error.throttled ? "warn" : "error";
            const seconds = error.retryAt ? Math.max(0, Math.ceil((error.retryAt - Date.now()) / 1000)) : 0;
            text = error.throttled
                ? `Microsoft Graph is rate-limiting Teams requests.${seconds ? ` Retrying in ${seconds}s.` : ""}`
                : error.message;
            action = h("button", { type: "button", class: "text-btn", onclick: refreshChats }, "Retry");
        }
        els.banner.hidden = !text;
        if (!text) return;
        els.banner.className = `banner ${tone}`;
        els.banner.replaceChildren(h("span", null, text), action);
    }

    function setFilter(filter) {
        state.filter = filter === "all" ? "all" : "unread";
        renderList();
        saveView();
    }

    async function refreshChats() {
        if (state.refreshing) return;
        state.refreshing = true;
        els.refresh.classList.add("spinning");
        try {
            await api("POST", "/api/refresh");
        } catch (error) {
            state.status = { ...state.status, error: { message: error.message, throttled: error.status === 429, retryAt: error.retryAfterMs ? Date.now() + error.retryAfterMs : null } };
            renderBanner();
        } finally {
            state.refreshing = false;
            els.refresh.classList.remove("spinning");
        }
        if (state.activeId) loadHistory(state.activeId, true);
    }

    async function loadMoreChats(event) {
        const button = event?.currentTarget;
        if (button) {
            button.disabled = true;
            button.textContent = "Loading…";
        }
        try {
            await api("POST", "/api/more");
        } catch (error) {
            if (button) {
                button.disabled = false;
                button.textContent = "Load more chats";
            }
            state.status = { ...state.status, error: { message: error.message, throttled: error.status === 429 } };
            renderBanner();
        }
    }

    // ---- conversation ----------------------------------------------------

    function subtitle(chat) {
        if (chat.type === "oneOnOne") return "Chat";
        const kind = chat.type === "meeting" ? "Meeting chat" : "Group chat";
        return chat.memberCount ? `${kind} · ${chat.memberCount} people` : kind;
    }

    function renderHeader() {
        const chat = chatById(state.activeId);
        els.emptyPane.hidden = Boolean(state.activeId);
        els.chatPane.hidden = !state.activeId;
        els.app.classList.toggle("show-chat", Boolean(state.activeId));
        if (!state.activeId) return;
        const title = chat?.title ?? "Chat";
        els.chatAvatar.replaceChildren(avatar(title, state.activeId, { group: chat ? chat.type !== "oneOnOne" : false }));
        els.chatTitle.textContent = title;
        els.chatTitle.title = title;
        els.chatSubtitle.textContent = chat ? subtitle(chat) : "";
        els.openTeams.disabled = !chat?.webUrl;
        const unread = Boolean(chat?.unread);
        const label = unread ? "Mark as read" : "Mark as unread";
        els.toggleRead.disabled = !chat?.last;
        els.toggleRead.title = label;
        els.toggleRead.setAttribute("aria-label", label);
        els.toggleReadIcon.setAttribute("d", unread ? ICONS.markRead : ICONS.markUnread);
    }

    function fileChip(attachment, chat) {
        if (attachment.kind === "card") {
            return h(
                "button",
                { type: "button", class: "file-chip", title: "Cards open in Teams", disabled: chat?.webUrl ? null : true, onclick: () => openExternal(chat?.webUrl) },
                icon("card"),
                h("span", null, attachment.name || "Card"),
            );
        }
        return h(
            "button",
            { type: "button", class: "file-chip", title: attachment.url ? `Open ${attachment.name}` : attachment.name, disabled: attachment.url ? null : true, onclick: () => openExternal(attachment.url) },
            icon("file"),
            h("span", null, attachment.name || "Attachment"),
        );
    }

    function renderMessage(message, grouped, chat) {
        if (message.kind !== "message") {
            return h("div", { class: "event", title: message.at ? fmtFull.format(new Date(message.at)) : null }, message.event || "Activity");
        }
        const row = h("div", { class: `msg${message.mine ? " mine" : ""}${grouped ? " grouped" : ""}`, dataset: { id: message.id } });
        if (!message.mine) row.append(grouped ? h("div", { "aria-hidden": "true" }) : avatar(message.fromName, message.fromId, { small: true }));
        const column = h("div", { class: "msg-col" });
        if (!grouped) {
            column.append(
                h(
                    "div",
                    { class: "msg-head" },
                    message.mine ? null : h("span", { class: "name" }, message.fromName || "Unknown"),
                    message.at ? h("time", { datetime: message.at, title: fmtFull.format(new Date(message.at)) }, fmtTime.format(new Date(message.at))) : null,
                ),
            );
        }
        const bubble = h("div", { class: `bubble${message.mentionsMe ? " mention-me" : ""}` });
        if (grouped && message.at) bubble.title = fmtFull.format(new Date(message.at));
        if (message.importance) bubble.append(h("div", { class: "importance" }, message.importance === "urgent" ? "Urgent!" : "Important!"));
        if (message.subject) bubble.append(h("div", { class: "subject" }, message.subject));
        for (const attachment of message.attachments) {
            if (attachment.kind !== "quote") continue;
            const who = attachment.forwarded ? `Forwarded${attachment.sender ? ` from ${attachment.sender}` : ""}` : attachment.sender || "Message";
            bubble.append(h("div", { class: "quote" }, h("div", { class: "who" }, who), h("div", { class: "quote-text" }, attachment.preview)));
        }
        const content = h("div", { class: "content" });
        if (message.deleted) content.append(h("span", { class: "deleted" }, "This message was deleted."));
        else if (message.html != null) renderHtml(content, message.html);
        else content.textContent = message.text || "";
        bubble.append(content);
        const files = message.attachments.filter((attachment) => attachment.kind !== "quote");
        if (files.length) bubble.append(h("div", { class: "files" }, files.map((attachment) => fileChip(attachment, chat))));
        if (message.editedAt && !message.deleted) {
            bubble.append(h("span", { class: "edited", title: `Edited ${fmtFull.format(new Date(message.editedAt))}` }, "Edited"));
        }
        column.append(bubble);
        if (message.reactions?.length) {
            column.append(
                h(
                    "div",
                    { class: "reactions" },
                    message.reactions.map((reaction) => h("span", { class: `reaction${reaction.mine ? " mine" : ""}`, title: reaction.mine ? "You reacted" : null }, `${reaction.emoji} ${reaction.count}`)),
                ),
            );
        }
        row.append(column);
        if (!message.deleted) {
            row.append(
                h(
                    "div",
                    { class: "msg-actions" },
                    h("button", { type: "button", class: "icon-btn sm", title: "Reply", "aria-label": "Reply with quote", onclick: () => startQuote(message) }, icon("reply")),
                    h("button", { type: "button", class: "icon-btn sm", title: "Copy text", "aria-label": "Copy message text", onclick: () => copyText(plainText(message)) }, icon("copy")),
                ),
            );
        }
        return row;
    }

    function scrollToBottom() {
        els.messages.scrollTop = els.messages.scrollHeight;
    }

    function renderMessages() {
        const chatId = state.activeId;
        const chat = chatById(chatId);
        const view = chatId ? state.histories.get(chatId) : null;
        const box = els.messages;
        if (!chatId || !view) {
            box.replaceChildren();
            renderHistoryBanner(null);
            return;
        }
        const previousHeight = box.scrollHeight;
        const previousTop = box.scrollTop;
        const fragment = document.createDocumentFragment();
        if (view.hasOlder) {
            const label = state.loadingOlder ? "Loading…" : view.olderError ? "Couldn't load older messages — try again" : "Load older messages";
            fragment.append(
                h(
                    "div",
                    { class: "older" },
                    h("button", { type: "button", class: "text-btn", disabled: state.loadingOlder ? true : null, onclick: loadOlder }, label),
                ),
            );
        } else if (view.loadedAt && view.messages.length) {
            fragment.append(h("div", { class: "start" }, "Beginning of the conversation"));
        }
        const divider = state.dividerAt.get(chatId);
        let dividerShown = !divider;
        let lastDay = "";
        let lastSender = null;
        let lastTime = 0;
        for (const message of view.messages) {
            const at = time(message.at);
            const day = message.at ? new Date(message.at).toDateString() : "";
            if (day && day !== lastDay) {
                fragment.append(h("div", { class: "day-sep" }, dayLabel(message.at)));
                lastDay = day;
                lastSender = null;
            }
            if (!dividerShown && !message.mine && message.kind === "message" && at > time(divider)) {
                fragment.append(h("div", { class: "new-sep", role: "separator" }, "New"));
                dividerShown = true;
                lastSender = null;
            }
            const grouped = message.kind === "message" && lastSender === message.fromId && at - lastTime < 5 * 60_000;
            fragment.append(renderMessage(message, grouped, chat));
            lastSender = message.kind === "message" ? message.fromId : null;
            lastTime = at;
        }
        if (!view.messages.length) {
            fragment.append(
                h("div", { class: view.loading ? "loading" : "empty" }, view.loading ? "Loading messages…" : view.error ? "Messages aren't available right now." : "No messages yet."),
            );
        }
        box.replaceChildren(fragment);
        if (state.keepOffset) {
            box.scrollTop = previousTop + (box.scrollHeight - previousHeight);
            state.keepOffset = false;
        } else if (state.pinned) {
            scrollToBottom();
        } else {
            box.scrollTop = previousTop;
        }
        renderHistoryBanner(view);
    }

    function renderHistoryBanner(view) {
        const error = view?.error;
        els.historyBanner.hidden = !error;
        if (!error) return;
        const seconds = error.retryAt ? Math.max(0, Math.ceil((error.retryAt - Date.now()) / 1000)) : 0;
        const latestOnly = view.messages.length ? " Showing the latest message only." : "";
        const staysUnread = chatById(state.activeId)?.unread ? " This chat stays unread until history loads or you mark it read." : "";
        const text = error.throttled
            ? `Teams history is rate-limited right now.${latestOnly}${seconds ? ` Retrying in ${seconds}s.` : ""} You can still reply.${staysUnread}`
            : `Couldn't load messages: ${error.message}${staysUnread}`;
        els.historyBanner.className = `banner ${error.throttled ? "warn" : "error"}`;
        els.historyBanner.replaceChildren(
            h("span", null, text),
            h("button", { type: "button", class: "text-btn", onclick: () => loadHistory(state.activeId, true) }, "Retry"),
        );
    }

    function applyHistory(view) {
        if (!view?.chatId) return;
        state.histories.set(view.chatId, { ...view, loading: false });
        if (view.chatId === state.activeId) renderMessages();
    }

    async function loadHistory(chatId, force = false) {
        if (!chatId) return;
        const existing = state.histories.get(chatId);
        state.histories.set(chatId, { ...(existing ?? { chatId, messages: [], hasOlder: false, error: null, loadedAt: 0 }), loading: true });
        if (chatId === state.activeId && !existing) renderMessages();
        try {
            applyHistory(await api("GET", `/api/messages?chat=${encodeURIComponent(chatId)}${force ? "&force=1" : ""}`));
        } catch (error) {
            const entry = state.histories.get(chatId);
            entry.loading = false;
            entry.complete = false;
            entry.error = { message: error.message, throttled: error.status === 429, retryAt: error.retryAfterMs ? Date.now() + error.retryAfterMs : null };
            if (chatId === state.activeId) renderMessages();
        }
    }

    async function loadOlder() {
        const chatId = state.activeId;
        if (!chatId || state.loadingOlder) return;
        state.loadingOlder = true;
        state.pinned = false;
        renderMessages();
        try {
            const view = await api("POST", "/api/older", { chatId });
            state.keepOffset = true;
            state.loadingOlder = false;
            applyHistory(view);
        } catch (error) {
            state.loadingOlder = false;
            const entry = state.histories.get(chatId);
            if (entry) entry.olderError = { message: error.message, throttled: error.status === 429 };
            renderMessages();
        }
    }

    // ---- read state ------------------------------------------------------

    // Auto mark-read only when the user is actually looking at a complete,
    // up-to-date conversation: never for programmatic navigation, a hidden or
    // unfocused panel, or while history is unavailable and only a preview shows.
    function canAutoRead(chat) {
        if (!chat?.unread || !state.engaged || !state.pinned) return false;
        if (document.visibilityState !== "visible" || !document.hasFocus()) return false;
        return Boolean(state.histories.get(chat.id)?.complete);
    }

    let readTimer = null;
    function maybeMarkRead() {
        clearTimeout(readTimer);
        readTimer = setTimeout(() => {
            const chat = chatById(state.activeId);
            if (!canAutoRead(chat)) return;
            api("POST", "/api/read", { chatId: chat.id }).catch((error) => {
                if (error.status !== 429) console.warn("mark read failed", error);
            });
        }, 700);
    }

    function engage() {
        if (!state.activeId || state.engaged) return;
        state.engaged = true;
        maybeMarkRead();
    }

    async function toggleRead() {
        const chat = chatById(state.activeId);
        if (!chat) return;
        els.toggleRead.disabled = true;
        try {
            if (chat.unread) {
                await api("POST", "/api/read", { chatId: chat.id });
            } else {
                await api("POST", "/api/unread", { chatId: chat.id });
                closeChat();
            }
        } catch (error) {
            state.status = { ...state.status, error: { message: error.message, throttled: error.status === 429 } };
            renderBanner();
        } finally {
            renderHeader();
        }
    }

    // ---- navigation ------------------------------------------------------

    function saveView() {
        api("POST", "/api/view", { inst: INSTANCE, chatId: state.activeId, filter: state.filter }).catch(() => {});
    }

    function openChat(chatId, { user = false } = {}) {
        if (!chatId) return;
        if (state.activeId && state.activeId !== chatId) leaveChat(state.activeId);
        const switching = state.activeId !== chatId;
        state.activeId = chatId;
        state.engaged = user;
        const chat = chatById(chatId);
        if (switching) state.dividerAt.set(chatId, chat?.unread ? chat.lastReadAt || "1970-01-01T00:00:00.000Z" : null);
        state.pinned = true;
        renderList();
        renderHeader();
        restoreDraft();
        renderMessages();
        api("POST", "/api/active", { cid: CLIENT_ID, chatId }).catch(() => {});
        saveView();
        if (user && !isNarrow()) els.input.focus({ preventScroll: true });
        loadHistory(chatId).then(maybeMarkRead);
    }

    function closeChat() {
        if (state.activeId) leaveChat(state.activeId);
        state.activeId = null;
        state.engaged = false;
        renderHeader();
        renderList();
        renderMessages();
        api("POST", "/api/active", { cid: CLIENT_ID, chatId: null }).catch(() => {});
        saveView();
    }

    function openExternal(url) {
        if (!url) return;
        api("POST", "/api/open", { url }).catch((error) => console.warn("open failed", error));
    }

    async function copyText(text) {
        try {
            await navigator.clipboard.writeText(text);
        } catch {
            const area = h("textarea", { "aria-hidden": "true" });
            area.value = text;
            area.style.position = "fixed";
            area.style.opacity = "0";
            document.body.append(area);
            area.select();
            document.execCommand("copy");
            area.remove();
        }
    }

    // ---- drafts ----------------------------------------------------------
    // Each chat's draft tracks the server revision it's based on and the text
    // the server last confirmed ("synced"). Unsynced typing always wins over
    // incoming drafts; a competing draft is offered instead of overwriting it.

    function draftOf(chatId) {
        let draft = state.drafts.get(chatId);
        if (!draft) {
            draft = { text: "", quoteId: null, rev: 0, synced: { text: "", quoteId: null }, remote: null };
            state.drafts.set(chatId, draft);
        }
        return draft;
    }

    function captureEditor() {
        if (state.activeId) draftOf(state.activeId).text = els.input.value;
    }

    const isDirty = (draft) => draft.text !== draft.synced.text || (draft.quoteId ?? null) !== (draft.synced.quoteId ?? null);

    function adoptServerDraft(chatId, server, origin) {
        const draft = draftOf(chatId);
        if (!Number.isInteger(server?.rev) || server.rev < draft.rev) return;
        if (chatId === state.activeId) captureEditor();
        const incoming = { text: server.text ?? "", quoteId: server.quoteId ?? null };
        const known = server.rev === draft.rev && incoming.text === draft.synced.text && incoming.quoteId === (draft.synced.quoteId ?? null);
        if (known) return;
        draft.rev = server.rev;
        if (origin && origin === CLIENT_ID) {
            draft.synced = incoming;
            return;
        }
        // A draft cleared because its text was sent is authoritative for that
        // text: never push the already-sent message back as a draft.
        const sentText = typeof server.sentText === "string" ? server.sentText.trim() : null;
        if (sentText && draft.text.trim() === sentText) {
            draft.text = incoming.text;
            draft.quoteId = incoming.quoteId;
            draft.synced = incoming;
            draft.remote = null;
            if (chatId === state.activeId) restoreDraft();
            return;
        }
        const same = incoming.text === draft.text && incoming.quoteId === (draft.quoteId ?? null);
        if (same || !isDirty(draft)) {
            draft.text = incoming.text;
            draft.quoteId = incoming.quoteId;
            draft.synced = incoming;
            draft.remote = null;
            if (chatId === state.activeId) restoreDraft();
            return;
        }
        draft.synced = incoming;
        if (incoming.text.trim() && incoming.text !== draft.text) draft.remote = { ...incoming, origin };
        if (chatId === state.activeId) renderDraftBar();
        schedulePush(chatId);
    }

    const pushTimers = new Map();
    let pushChain = Promise.resolve();

    function schedulePush(chatId, delayMs = 700) {
        clearTimeout(pushTimers.get(chatId));
        pushTimers.set(chatId, setTimeout(() => pushDraft(chatId), delayMs));
    }

    function pushDraft(chatId) {
        clearTimeout(pushTimers.get(chatId));
        pushTimers.delete(chatId);
        pushChain = pushChain.then(async () => {
            const draft = state.drafts.get(chatId);
            if (!draft || !isDirty(draft) || state.sending.has(chatId)) return;
            const payload = { text: draft.text, quoteId: draft.quoteId ?? null };
            try {
                const result = await api("POST", "/api/draft", { chatId, ...payload, baseRev: draft.rev, cid: CLIENT_ID });
                draft.failures = 0;
                if (result.rev >= draft.rev) {
                    draft.rev = result.rev;
                    draft.synced = payload;
                }
            } catch (error) {
                if (error.status === 409 && error.data) {
                    adoptServerDraft(chatId, error.data, null);
                    return;
                }
                // Keep the draft dirty and retry with backoff (the extension may be
                // restarting); unload flushes still try to save it.
                draft.failures = (draft.failures ?? 0) + 1;
                if (draft.failures <= 6) schedulePush(chatId, Math.min(30_000, 1000 * 2 ** draft.failures));
            }
        });
        return pushChain;
    }

    // Final flush when the panel goes away; keepalive lets it outlive the page.
    function flushDraftOnUnload() {
        captureEditor();
        for (const [chatId, draft] of state.drafts) {
            if (!isDirty(draft) || state.sending.has(chatId)) continue;
            fetch("/api/draft", {
                method: "POST",
                keepalive: true,
                headers: { "x-tl-token": TOKEN, "content-type": "application/json" },
                body: JSON.stringify({ chatId, text: draft.text, quoteId: draft.quoteId ?? null, baseRev: draft.rev, cid: CLIENT_ID }),
            }).catch(() => {});
        }
    }

    function leaveChat(chatId) {
        if (chatId === state.activeId) captureEditor();
        const draft = state.drafts.get(chatId);
        if (draft && isDirty(draft)) pushDraft(chatId);
    }

    function restoreDraft() {
        const draft = state.activeId ? state.drafts.get(state.activeId) : null;
        els.input.value = draft?.text ?? "";
        autosize();
        renderQuoteBar();
        renderDraftBar();
        renderSendError();
        updateComposer();
    }

    function renderDraftBar() {
        const remote = state.activeId ? state.drafts.get(state.activeId)?.remote : null;
        els.draftBar.hidden = !remote;
        if (!remote) return;
        els.draftBarLabel.textContent = remote.origin === "copilot" ? "Copilot drafted a reply:" : "A newer draft arrived:";
        els.draftBarText.textContent = remote.text.replace(/\s+/g, " ").slice(0, 160);
    }

    function useRemoteDraft() {
        const chatId = state.activeId;
        const draft = chatId ? state.drafts.get(chatId) : null;
        if (!draft?.remote) return;
        draft.text = draft.remote.text;
        draft.quoteId = draft.remote.quoteId;
        draft.remote = null;
        restoreDraft();
        els.input.focus();
        els.input.setSelectionRange(els.input.value.length, els.input.value.length);
        schedulePush(chatId, 0);
    }

    function dismissRemoteDraft() {
        const draft = state.activeId ? state.drafts.get(state.activeId) : null;
        if (!draft) return;
        draft.remote = null;
        renderDraftBar();
    }

    // ---- composer ----------------------------------------------------------

    function setSendError(chatId, text) {
        if (!chatId) return;
        if (text) state.sendErrors.set(chatId, text);
        else state.sendErrors.delete(chatId);
        if (chatId === state.activeId) renderSendError();
    }

    function renderSendError() {
        const text = state.activeId ? state.sendErrors.get(state.activeId) ?? "" : "";
        els.sendError.textContent = text;
        els.sendError.hidden = !text;
    }

    function autosize() {
        els.input.style.height = "auto";
        els.input.style.height = `${Math.min(els.input.scrollHeight, 168)}px`;
    }

    function updateComposer() {
        const chatId = state.activeId;
        const busy = Boolean(chatId) && state.sending.has(chatId);
        const pending = Boolean(chatId) && state.pendingChats.has(chatId);
        els.send.disabled = !chatId || busy || pending || !els.input.value.trim();
        els.send.classList.toggle("busy", busy || pending);
        els.send.title = pending ? "Waiting for Teams to confirm your last message" : "Send (Enter)";
        els.input.readOnly = busy;
    }

    function renderQuoteBar() {
        const draft = state.activeId ? state.drafts.get(state.activeId) : null;
        if (!draft?.quoteId) {
            els.quoteBar.hidden = true;
            return;
        }
        const message = state.histories.get(state.activeId)?.messages.find((m) => m.id === draft.quoteId);
        els.quoteName.textContent = message?.fromName ?? "Message";
        els.quoteText.textContent = message ? plainText(message).replace(/\s+/g, " ").slice(0, 200) : "(not loaded)";
        els.quoteBar.hidden = false;
    }

    function startQuote(message) {
        const chatId = state.activeId;
        if (!chatId) return;
        captureEditor();
        draftOf(chatId).quoteId = message.id;
        renderQuoteBar();
        els.input.focus();
        schedulePush(chatId);
    }

    function cancelQuote() {
        const chatId = state.activeId;
        const draft = chatId ? state.drafts.get(chatId) : null;
        if (!draft?.quoteId) return;
        draft.quoteId = null;
        renderQuoteBar();
        schedulePush(chatId);
    }

    // Clear the composer only if it still holds exactly what was sent.
    function clearSentText(chatId, text) {
        const draft = draftOf(chatId);
        if (chatId === state.activeId) captureEditor();
        if (draft.text.trim() === String(text ?? "").trim()) {
            draft.text = "";
            draft.quoteId = null;
            draft.remote = null;
            draft.synced = { text: "", quoteId: null };
            if (chatId === state.activeId) restoreDraft();
        }
        setSendError(chatId, "");
    }

    async function send() {
        const chatId = state.activeId;
        if (!chatId || state.sending.has(chatId) || state.pendingChats.has(chatId)) return;
        captureEditor();
        const draft = draftOf(chatId);
        const text = draft.text.trim();
        if (!text) return;
        clearTimeout(pushTimers.get(chatId));
        pushTimers.delete(chatId);
        state.sending.add(chatId);
        setSendError(chatId, "");
        updateComposer();
        try {
            const { message } = await api("POST", "/api/send", { chatId, text, quoteId: draft.quoteId ?? null, draftRev: draft.rev, cid: CLIENT_ID });
            clearSentText(chatId, text);
            if (message) {
                const entry = state.histories.get(chatId);
                if (entry && !entry.messages.some((m) => m.id === message.id)) entry.messages = [...entry.messages, message];
            }
            state.pinned = true;
            if (state.activeId === chatId) renderMessages();
        } catch (error) {
            if (error.code === "unconfirmed") state.pendingChats.add(chatId);
            setSendError(chatId, error.message || "Couldn't send that message.");
        } finally {
            state.sending.delete(chatId);
            updateComposer();
            if (isDirty(draftOf(chatId))) schedulePush(chatId);
            if (state.activeId === chatId) els.input.focus();
        }
    }

    function applySendState(event) {
        const chatId = event.chatId;
        if (!chatId) return;
        if (event.state === "pending") {
            state.pendingChats.add(chatId);
        } else {
            state.pendingChats.delete(chatId);
            if (event.state === "delivered") {
                if (event.origin === CLIENT_ID) clearSentText(chatId, event.text);
            } else if (event.message) {
                setSendError(chatId, event.message);
            }
        }
        if (chatId === state.activeId) updateComposer();
    }

    // ---- live updates ----------------------------------------------------

    function applySnapshot(snapshot) {
        if (snapshot.me) state.me = snapshot.me;
        if (typeof snapshot.loaded === "boolean") state.loaded = snapshot.loaded;
        if (Array.isArray(snapshot.chats)) state.chats = snapshot.chats;
        if (typeof snapshot.hasMore === "boolean") state.hasMore = snapshot.hasMore;
        if (snapshot.status) state.status = snapshot.status;
        if (Array.isArray(snapshot.pending)) state.pendingChats = new Set(snapshot.pending);
        renderList();
        renderBanner();
        renderHeader();
        updateComposer();
    }

    function mergeDrafts(drafts) {
        for (const [chatId, server] of Object.entries(drafts ?? {})) adoptServerDraft(chatId, server, null);
    }

    function handleEvent(event) {
        switch (event.type) {
            case "hello":
                applySnapshot(event);
                mergeDrafts(event.drafts);
                if (state.activeId) {
                    api("POST", "/api/active", { cid: CLIENT_ID, chatId: state.activeId }).catch(() => {});
                    loadHistory(state.activeId);
                }
                break;
            case "chats":
                applySnapshot(event);
                break;
            case "status":
                state.status = event.status ?? state.status;
                if (typeof event.loaded === "boolean") state.loaded = event.loaded;
                renderBanner();
                renderList();
                break;
            case "messages":
                applyHistory(event);
                if (event.chatId === state.activeId) {
                    renderQuoteBar();
                    maybeMarkRead();
                }
                break;
            case "draft":
                adoptServerDraft(event.chatId, event, event.origin ?? null);
                break;
            case "send_state":
                applySendState(event);
                break;
            case "navigate":
                if (event.chatId) openChat(event.chatId);
                break;
        }
    }

    function connect() {
        const params = new URLSearchParams({ t: TOKEN, cid: CLIENT_ID, inst: INSTANCE });
        const source = new EventSource(`/events?${params}`);
        source.onopen = () => {
            state.connected = true;
            renderBanner();
        };
        source.onerror = () => {
            state.connected = false;
            renderBanner();
        };
        source.onmessage = (message) => {
            let event;
            try {
                event = JSON.parse(message.data);
            } catch {
                return;
            }
            handleEvent(event);
        };
    }

    // ---- wiring ----------------------------------------------------------

    function applyLayout() {
        els.app.classList.toggle("narrow", document.body.clientWidth < 680);
    }

    function bind() {
        for (const tab of els.tabs) tab.addEventListener("click", () => setFilter(tab.dataset.filter));
        els.search.addEventListener("input", () => {
            state.query = els.search.value;
            renderList();
        });
        els.refresh.addEventListener("click", refreshChats);
        els.back.addEventListener("click", closeChat);
        els.toggleRead.addEventListener("click", toggleRead);
        els.openTeams.addEventListener("click", () => openExternal(chatById(state.activeId)?.webUrl));
        els.reloadChat.addEventListener("click", () => loadHistory(state.activeId, true));
        els.quoteCancel.addEventListener("click", cancelQuote);
        els.draftUse.addEventListener("click", useRemoteDraft);
        els.draftDismiss.addEventListener("click", dismissRemoteDraft);
        els.composer.addEventListener("submit", (event) => {
            event.preventDefault();
            send();
        });
        els.input.addEventListener("input", () => {
            autosize();
            updateComposer();
            if (!state.activeId) return;
            captureEditor();
            if (!state.pendingChats.has(state.activeId)) setSendError(state.activeId, "");
            schedulePush(state.activeId);
        });
        els.input.addEventListener("keydown", (event) => {
            if (event.key === "Enter" && !event.shiftKey && !event.isComposing) {
                event.preventDefault();
                send();
            } else if (event.key === "Escape") {
                if (state.drafts.get(state.activeId)?.quoteId) cancelQuote();
                else if (isNarrow()) closeChat();
            }
        });
        els.messages.addEventListener("scroll", () => {
            const box = els.messages;
            state.pinned = box.scrollHeight - box.scrollTop - box.clientHeight < 60;
            if (state.pinned) maybeMarkRead();
        });
        // Any real interaction with the open conversation counts as reading it.
        for (const type of ["pointerdown", "keydown", "wheel"]) {
            els.chatPane.addEventListener(type, engage, { passive: true });
        }
        document.addEventListener("click", (event) => {
            const link = event.target.closest?.("a[href]");
            if (!link) return;
            event.preventDefault();
            if (link.dataset.external) openExternal(link.href);
        });
        document.addEventListener("keydown", (event) => {
            if (event.key === "/" && !event.metaKey && !event.ctrlKey && !["INPUT", "TEXTAREA"].includes(document.activeElement?.tagName)) {
                event.preventDefault();
                els.search.focus();
            }
        });
        document.addEventListener("visibilitychange", () => {
            if (document.visibilityState === "visible") maybeMarkRead();
            else flushDraftOnUnload();
        });
        window.addEventListener("focus", maybeMarkRead);
        window.addEventListener("pagehide", flushDraftOnUnload);
        new ResizeObserver(applyLayout).observe(document.body);
        setInterval(() => {
            if (state.status.error?.retryAt) renderBanner();
            const view = state.histories.get(state.activeId);
            if (view?.error?.retryAt) renderHistoryBanner(view);
        }, 1000);
    }

    async function init() {
        bind();
        applyLayout();
        renderList();
        renderHeader();
        try {
            const snapshot = await api("GET", `/api/state?inst=${encodeURIComponent(INSTANCE)}`);
            mergeDrafts(snapshot.drafts);
            if (snapshot.view?.filter) state.filter = snapshot.view.filter;
            applySnapshot(snapshot);
            if (snapshot.view?.chatId) openChat(snapshot.view.chatId);
        } catch (error) {
            state.status = { loading: false, error: { message: error.message } };
            renderBanner();
        }
        connect();
    }

    init();
})();
