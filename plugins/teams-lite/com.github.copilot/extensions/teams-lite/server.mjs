// Loopback HTTP server for the Teams canvas: serves the UI, a small JSON API,
// Server-Sent Events for live updates, and an image proxy for Teams-hosted
// images. Every request is bound to 127.0.0.1 and must carry the per-process
// capability token so other local pages can't read or send messages.

import { spawn } from "node:child_process";
import { randomBytes, timingSafeEqual } from "node:crypto";
import { readFile } from "node:fs/promises";
import { createServer } from "node:http";
import { join } from "node:path";

const MAX_BODY_BYTES = 128 * 1024;
const IMAGE_CACHE_LIMIT = 60;
const HOSTED_CONTENT = /^\/chats\/[^/?#]+\/messages\/[^/?#]+\/hostedContents\/[^/?#]+\/\$value$/;

class HttpError extends Error {
    constructor(status, message) {
        super(message);
        this.status = status;
    }
}

function sniffImage(bytes) {
    if (!bytes || bytes.length < 12) return null;
    if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) return "image/png";
    if (bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) return "image/jpeg";
    if (bytes.subarray(0, 4).toString("latin1") === "GIF8") return "image/gif";
    if (bytes.subarray(0, 4).toString("latin1") === "RIFF" && bytes.subarray(8, 12).toString("latin1") === "WEBP") return "image/webp";
    return null;
}

function openExternal(raw) {
    let url;
    try {
        url = new URL(String(raw));
    } catch {
        return false;
    }
    if (!["https:", "http:", "mailto:"].includes(url.protocol)) return false;
    const [command, args] =
        process.platform === "darwin"
            ? ["open", [url.href]]
            : process.platform === "win32"
              ? ["rundll32", ["url.dll,FileProtocolHandler", url.href]]
              : ["xdg-open", [url.href]];
    try {
        const child = spawn(command, args, { stdio: "ignore", detached: true });
        child.on("error", () => {});
        child.unref();
        return true;
    } catch {
        return false;
    }
}

function readJson(req) {
    return new Promise((resolve, reject) => {
        if (!String(req.headers["content-type"] ?? "").startsWith("application/json")) {
            reject(new HttpError(415, "Expected application/json."));
            return;
        }
        const chunks = [];
        let size = 0;
        req.on("data", (chunk) => {
            size += chunk.length;
            if (size > MAX_BODY_BYTES) {
                reject(new HttpError(413, "Request too large."));
                req.destroy();
                return;
            }
            chunks.push(chunk);
        });
        req.on("end", () => {
            try {
                const value = JSON.parse(Buffer.concat(chunks).toString("utf8") || "{}");
                resolve(value && typeof value === "object" && !Array.isArray(value) ? value : {});
            } catch {
                reject(new HttpError(400, "Invalid JSON."));
            }
        });
        req.on("error", reject);
    });
}

function sendJson(res, status, value) {
    const body = JSON.stringify(value);
    res.writeHead(status, {
        "Content-Type": "application/json; charset=utf-8",
        "Cache-Control": "no-store",
        "X-Content-Type-Options": "nosniff",
    });
    res.end(body);
}

function sendText(res, status, text) {
    res.writeHead(status, { "Content-Type": "text/plain; charset=utf-8", "Cache-Control": "no-store", "X-Content-Type-Options": "nosniff" });
    res.end(text);
}

const str = (value) => (typeof value === "string" && value.trim() ? value.trim() : null);
const clientId = (value) => {
    const id = str(value);
    return id && /^[A-Za-z0-9-]{8,64}$/.test(id) ? id : null;
};

function requireChatId(value) {
    const chatId = str(value);
    if (!chatId || chatId.length > 300 || /[\s/?#]/.test(chatId)) throw new HttpError(400, "A valid chatId is required.");
    return chatId;
}

export function createUiServer({ store, wq, uiDir, debug = () => {} }) {
    const token = randomBytes(24).toString("base64url");
    const tokenBytes = Buffer.from(token);
    let server = null;
    let port = 0;
    let starting = null;
    let closeTimer = null;
    const clients = new Map();
    const views = new Map();
    const images = new Map();
    const imageLoads = new Map();

    const tokenOk = (value) => {
        if (typeof value !== "string") return false;
        const bytes = Buffer.from(value);
        return bytes.length === tokenBytes.length && timingSafeEqual(bytes, tokenBytes);
    };

    function write(client, event) {
        try {
            client.res.write(`data: ${JSON.stringify(event)}\n\n`);
        } catch (error) {
            debug(`sse write failed: ${error?.message ?? error}`);
        }
    }

    function broadcast(event, filter = () => true) {
        for (const client of clients.values()) if (filter(client)) write(client, event);
    }

    store.on((event) => {
        if (event.type === "chats") broadcast({ type: "chats", ...store.snapshot() });
        else if (event.type === "status") broadcast({ type: "status", status: store.status, loaded: store.loaded });
        else if (event.type === "messages") broadcast({ type: "messages", ...store.messagesView(event.chatId) }, (c) => c.chatId === event.chatId);
        else if (event.type === "draft") {
            broadcast({ type: "draft", chatId: event.chatId, origin: event.origin ?? null, sentText: event.sentText ?? null, ...store.draftState(event.chatId) });
        } else if (event.type === "send_state") broadcast(event);
    });

    function updateWatching() {
        store.setWatching(clients.size > 0);
    }

    async function renderPage(instanceId) {
        const [template, styles, script] = await Promise.all(
            ["index.html", "styles.css", "app.js"].map((file) => readFile(join(uiDir, file), "utf8")),
        );
        if (/<\/script/i.test(script) || /<\/style/i.test(styles)) throw new Error("UI assets must not contain closing script/style tags.");
        const nonce = randomBytes(16).toString("base64url");
        const values = {
            NONCE: nonce,
            CONFIG: JSON.stringify({ token, instanceId }).replace(/</g, "\\u003c"),
            STYLES: styles,
            SCRIPT: script,
        };
        const body = template.replace(/__(NONCE|CONFIG|STYLES|SCRIPT)__/g, (_, key) => values[key]);
        const csp = [
            "default-src 'none'",
            `script-src 'nonce-${nonce}'`,
            "style-src 'unsafe-inline'",
            "img-src 'self' data:",
            "font-src 'self' data:",
            "connect-src 'self'",
            "base-uri 'none'",
            "form-action 'none'",
            "frame-ancestors 'none'",
        ].join("; ");
        return { body, csp };
    }

    function openEvents(req, res, url) {
        const cid = str(url.searchParams.get("cid"))?.slice(0, 64) ?? randomBytes(8).toString("hex");
        const instanceId = url.searchParams.get("inst") ?? "";
        res.writeHead(200, {
            "Content-Type": "text/event-stream; charset=utf-8",
            "Cache-Control": "no-store",
            Connection: "keep-alive",
            "X-Content-Type-Options": "nosniff",
        });
        res.write("retry: 3000\n\n");
        clients.get(cid)?.res.end();
        const client = {
            res,
            instanceId,
            chatId: null,
            heartbeat: setInterval(() => res.write(": ping\n\n"), 25_000),
        };
        clients.set(cid, client);
        updateWatching();
        store.ready.then(() => {
            if (clients.get(cid) === client) write(client, { type: "hello", ...store.snapshot(), drafts: store.draftsView() });
        });
        req.on("close", () => {
            clearInterval(client.heartbeat);
            if (clients.get(cid) === client) {
                clients.delete(cid);
                store.dropClient(cid);
            }
            updateWatching();
        });
    }

    async function serveImage(res, path) {
        if (!HOSTED_CONTENT.test(path)) throw new HttpError(400, "Unsupported image path.");
        let image = images.get(path);
        if (image) {
            images.delete(path);
        } else {
            let load = imageLoads.get(path);
            if (!load) {
                load = wq.blob(path).finally(() => imageLoads.delete(path));
                imageLoads.set(path, load);
            }
            const blob = await load;
            const type = sniffImage(blob.bytes);
            if (!type) throw new HttpError(415, "Not a supported image.");
            image = { bytes: blob.bytes, type };
        }
        images.set(path, image);
        if (images.size > IMAGE_CACHE_LIMIT) images.delete(images.keys().next().value);
        res.writeHead(200, {
            "Content-Type": image.type,
            "Content-Length": image.bytes.length,
            "Cache-Control": "private, max-age=3600",
            "Content-Security-Policy": "default-src 'none'; sandbox",
            "X-Content-Type-Options": "nosniff",
        });
        res.end(image.bytes);
    }

    async function handleApi(req, res, url) {
        const body = req.method === "POST" ? await readJson(req) : {};
        switch (`${req.method} ${url.pathname}`) {
            case "GET /api/state": {
                await store.ready;
                const view = views.get(url.searchParams.get("inst") ?? "") ?? null;
                return sendJson(res, 200, { ...store.snapshot(), drafts: store.draftsView(), view });
            }
            case "POST /api/view": {
                views.set(String(body.inst ?? ""), { chatId: str(body.chatId), filter: body.filter === "all" ? "all" : "unread" });
                return sendJson(res, 200, { ok: true });
            }
            case "POST /api/active": {
                const cid = String(body.cid ?? "");
                const chatId = str(body.chatId);
                const client = clients.get(cid);
                if (client) client.chatId = chatId;
                store.setActive(cid, chatId);
                return sendJson(res, 200, { ok: true });
            }
            case "POST /api/refresh":
                await store.refreshChats({ full: true });
                return sendJson(res, 200, { ok: true });
            case "POST /api/more":
                await store.loadMoreChats();
                return sendJson(res, 200, { ok: true });
            case "GET /api/messages": {
                const chatId = requireChatId(url.searchParams.get("chat"));
                return sendJson(res, 200, await store.loadMessages(chatId, { force: url.searchParams.get("force") === "1" }));
            }
            case "POST /api/older":
                return sendJson(res, 200, await store.loadOlder(requireChatId(body.chatId)));
            case "POST /api/send": {
                const message = await store.send(requireChatId(body.chatId), body.text, {
                    quoteId: str(body.quoteId),
                    origin: clientId(body.cid),
                    draftRev: Number.isInteger(body.draftRev) ? body.draftRev : null,
                });
                return sendJson(res, 200, { message });
            }
            case "POST /api/read":
                return sendJson(res, 200, { changed: await store.markRead(requireChatId(body.chatId)) });
            case "POST /api/unread":
                return sendJson(res, 200, { changed: await store.markUnread(requireChatId(body.chatId)) });
            case "POST /api/draft": {
                await store.ready;
                const text = typeof body.text === "string" ? body.text : "";
                const baseRev = Number.isInteger(body.baseRev) ? body.baseRev : -1;
                const result = store.updateDraft(requireChatId(body.chatId), text, str(body.quoteId), baseRev, clientId(body.cid));
                return sendJson(res, result.ok ? 200 : 409, result);
            }
            case "POST /api/open": {
                const ok = openExternal(body.url);
                return sendJson(res, ok ? 200 : 400, ok ? { ok } : { error: "Only http(s) and mailto links can be opened." });
            }
            default:
                throw new HttpError(404, "Not found.");
        }
    }

    function sendError(res, error, pathname) {
        if (res.headersSent) {
            res.end();
            return;
        }
        if (error instanceof HttpError) {
            sendJson(res, error.status, { error: error.message });
            return;
        }
        if (error?.name === "WorkIQError") {
            const throttled = error.status === 429 || error.code === "CircuitOpen" || error.code === "cooldown";
            let message = throttled ? "Microsoft Graph is rate-limiting these requests. Try again shortly." : error.message;
            let code = error.code || String(error.status || "");
            let status = throttled ? 429 : ["not_connected", "needs_auth", "denied", "ambiguous_server", "untrusted_server"].includes(error.code) ? 503 : 502;
            if (["empty", "too_long", "quote_rejected", "quote_unavailable"].includes(error.code)) status = 400;
            if (error.code === "send_pending") status = 409;
            if (pathname === "/api/send" && error.ambiguous) {
                status = 504;
                code = "unconfirmed";
                message = "Teams hasn't confirmed this message yet. Checking whether it was delivered — don't resend until this clears.";
            }
            sendJson(res, status, { error: message, code, retryAfterMs: error.retryAfterMs || 0 });
            return;
        }
        debug(`request failed: ${error?.stack ?? error}`);
        sendJson(res, 500, { error: "Something went wrong in the Teams canvas." });
    }

    async function handle(req, res) {
        const host = req.headers.host;
        if (host !== `127.0.0.1:${port}` && host !== `localhost:${port}`) {
            sendText(res, 421, "Misdirected request.");
            return;
        }
        const url = new URL(req.url ?? "/", `http://127.0.0.1:${port}`);
        try {
            if (req.method === "GET" && url.pathname === "/") {
                if (!tokenOk(url.searchParams.get("t"))) return sendText(res, 401, "Unauthorized.");
                const { body, csp } = await renderPage(url.searchParams.get("inst") ?? "");
                res.writeHead(200, {
                    "Content-Type": "text/html; charset=utf-8",
                    "Content-Security-Policy": csp,
                    "X-Frame-Options": "DENY",
                    "X-Content-Type-Options": "nosniff",
                    "Referrer-Policy": "no-referrer",
                    "Cache-Control": "no-store",
                });
                res.end(body);
                return;
            }
            if (req.method === "GET" && url.pathname === "/events") {
                if (!tokenOk(url.searchParams.get("t"))) return sendText(res, 401, "Unauthorized.");
                openEvents(req, res, url);
                return;
            }
            if (req.method === "GET" && url.pathname === "/api/blob") {
                if (!tokenOk(url.searchParams.get("t"))) return sendText(res, 401, "Unauthorized.");
                await serveImage(res, url.searchParams.get("p") ?? "");
                return;
            }
            if (url.pathname.startsWith("/api/")) {
                if (!tokenOk(req.headers["x-tl-token"])) return sendJson(res, 401, { error: "Unauthorized." });
                const origin = req.headers.origin;
                if (origin && origin !== `http://127.0.0.1:${port}` && origin !== `http://localhost:${port}`) {
                    return sendJson(res, 403, { error: "Forbidden." });
                }
                await handleApi(req, res, url);
                return;
            }
            if (url.pathname === "/favicon.ico") {
                res.writeHead(204);
                res.end();
                return;
            }
            sendText(res, 404, "Not found.");
        } catch (error) {
            sendError(res, error, url.pathname);
        }
    }

    async function ensureStarted() {
        clearTimeout(closeTimer);
        closeTimer = null;
        if (server) return;
        if (!starting) {
            starting = (async () => {
                const instance = createServer((req, res) => {
                    handle(req, res).catch((error) => {
                        debug(`unhandled request error: ${error?.stack ?? error}`);
                        if (!res.headersSent) sendText(res, 500, "Internal error.");
                        else res.end();
                    });
                });
                instance.keepAliveTimeout = 5000;
                await new Promise((resolve, reject) => {
                    instance.once("error", reject);
                    instance.listen(0, "127.0.0.1", () => {
                        instance.off("error", reject);
                        resolve();
                    });
                });
                port = instance.address().port;
                server = instance;
            })().finally(() => {
                starting = null;
            });
        }
        await starting;
    }

    function close() {
        clearTimeout(closeTimer);
        closeTimer = null;
        for (const client of clients.values()) {
            clearInterval(client.heartbeat);
            client.res.end();
        }
        clients.clear();
        updateWatching();
        if (server) {
            server.close();
            server.closeAllConnections?.();
            server = null;
            port = 0;
        }
    }

    return {
        async urlFor(instanceId) {
            await ensureStarted();
            return `http://127.0.0.1:${port}/?inst=${encodeURIComponent(instanceId)}&t=${token}`;
        },
        setView(instanceId, view) {
            views.set(instanceId, { chatId: view.chatId ?? null, filter: view.filter === "all" ? "all" : "unread" });
        },
        navigate(chatId, instanceId = null) {
            for (const [key, view] of views) {
                if (!instanceId || key === instanceId) views.set(key, { ...view, chatId });
            }
            if (instanceId && !views.has(instanceId)) views.set(instanceId, { chatId, filter: "all" });
            broadcast({ type: "navigate", chatId }, (client) => !instanceId || client.instanceId === instanceId);
        },
        dropInstance(instanceId) {
            views.delete(instanceId);
        },
        clientCount: () => clients.size,
        closeSoon(delayMs = 30_000) {
            clearTimeout(closeTimer);
            closeTimer = setTimeout(close, delayMs);
            closeTimer.unref?.();
        },
        close,
    };
}
