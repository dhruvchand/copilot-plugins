// WorkIQ transport for the Teams canvas. Every call runs through this session's
// own WorkIQ MCP connection (session.rpc.tools.execute), so the extension never
// handles Microsoft 365 tokens and inherits the user's connector sign-in.

const READ_TIMEOUT_MS = 45_000;
const WRITE_TIMEOUT_MS = 60_000;
const MAX_CONCURRENT = 2;
const SERVER_CACHE_MS = 5 * 60_000;
const SERVER_PREFIX = /^MCP server '[^']*':\s*/;
const REQUIRED_TOOLS = ["fetch", "do_action", "create_entity"];

export class WorkIQError extends Error {
    constructor(message, { status = 0, code = "", retryAfterMs = 0, transient = false, ambiguous = false } = {}) {
        super(message);
        this.name = "WorkIQError";
        this.status = status;
        this.code = code;
        this.retryAfterMs = retryAfterMs;
        this.transient = transient;
        // For writes: the request may have taken effect even though it failed
        // (timeout, lost connection, or a server error after Graph accepted it).
        this.ambiguous = ambiguous;
    }

    get throttled() {
        return this.status === 429 || this.code === "CircuitOpen";
    }
}

// Graph throttles per API and method, so cooldowns are tracked per family and
// direction: throttled history reads must not block sending, and vice versa.
export function bucketOf(path, kind = "read") {
    let family = "other";
    if (/^\/(?:me\/)?chats\/[^/?]+\/messages\/[^/?]+\/hostedContents/.test(path)) family = "images";
    else if (/^\/(?:me\/)?chats\/[^/?]+\/messages(?:[/?]|$)/.test(path)) family = "messages";
    else if (/^\/(?:me\/)?chats\/[^/?]+\/members/.test(path)) family = "members";
    else if (/^\/(?:me\/)?chats/.test(path)) family = "chats";
    else if (/^\/users\//.test(path)) family = "users";
    return kind === "write" ? `write:${family}` : family;
}

function parseLoose(value) {
    if (typeof value !== "string") return undefined;
    const text = value.replace(SERVER_PREFIX, "").trim();
    if (!text.startsWith("{")) return undefined;
    try {
        return JSON.parse(text);
    } catch {
        return undefined;
    }
}

function describe(error) {
    if (!error) return "";
    if (typeof error === "string") return error;
    return error.error?.message ?? error.message ?? JSON.stringify(error);
}

function entryError(entry, fallback = "WorkIQ request failed.") {
    const status = Number(entry?.statusCode) || 0;
    const message = describe(entry?.error) || (status ? `HTTP ${status}` : fallback);
    let code = entry?.errorCode ?? entry?.error?.error?.code ?? entry?.error?.code ?? "";
    // WorkIQ rejects paths outside its allow-list before calling Graph.
    if (!code && /access denied for the requested path/i.test(message)) code = "path_denied";
    const retryAfterMs = Number(entry?.retryAfterSeconds) > 0 ? Number(entry.retryAfterSeconds) * 1000 : 0;
    const transient = status === 429 || status >= 500 || code === "CircuitOpen";
    return new WorkIQError(message, { status, code, retryAfterMs, transient });
}

function isOk(entry, resultType) {
    const status = Number(entry?.statusCode) || 0;
    if (status) return status >= 200 && status < 300;
    return resultType === "success" && entry != null && !entry.error;
}

function normalize(raw) {
    if (typeof raw === "string") return { payload: parseLoose(raw), resultType: "success", message: raw };
    const payload = raw?.structuredContent ?? parseLoose(raw?.textResultForLlm) ?? parseLoose(raw?.error);
    const message = String(raw?.error ?? raw?.textResultForLlm ?? "").replace(SERVER_PREFIX, "");
    return { payload, resultType: raw?.resultType ?? "success", message };
}

function unexpected(resultType, message) {
    if (resultType === "denied" || resultType === "rejected") {
        return new WorkIQError("WorkIQ access was denied in this session.", { code: "denied" });
    }
    const timedOut = resultType === "timeout";
    return new WorkIQError(timedOut ? "WorkIQ took too long to respond." : message || "WorkIQ returned an unexpected response.", {
        code: timedOut ? "timeout" : "bad_response",
        transient: true,
        ambiguous: true,
    });
}

function hostOf(url) {
    try {
        return url ? new URL(url).host : null;
    } catch {
        return null;
    }
}

// Connectors (signed-in account, host-managed catalog, or GitHub's connector
// gateway) are trusted most; repository-configured or unknown servers least.
function rankOf(server) {
    if (server.source === "account" || server.source === "managed") return 0;
    const host = hostOf(server.url);
    if (host === "api.github.com" || host?.endsWith(".githubcopilot.com")) return 0;
    if (server.source === "user" || server.source === "builtin" || server.source === "plugin") return 1;
    return 2;
}

function looksLikeWorkIQ(server) {
    const label = `${server.name ?? ""} ${server.displayName ?? ""} ${server.sourcePlugin ?? ""}`;
    if (/work\s*-?\s*iq/i.test(label)) return true;
    return /^\s*WorkIQ\b/.test(server.serverMetadata?.instructions ?? "");
}

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

export function createWorkIQ(getSession, { debug = () => {}, readTimeoutMs = READ_TIMEOUT_MS, writeTimeoutMs = WRITE_TIMEOUT_MS } = {}) {
    let server = null;
    const cooldowns = new Map();
    let active = 0;
    const waiting = [];

    function acquire() {
        if (active < MAX_CONCURRENT) {
            active++;
            return Promise.resolve();
        }
        return new Promise((resolve) => waiting.push(resolve));
    }

    function release() {
        const next = waiting.shift();
        if (next) next();
        else active--;
    }

    function forgetServer() {
        server = null;
    }

    async function resolveServer(session) {
        if (server && Date.now() - server.at < SERVER_CACHE_MS) return server;
        let list;
        try {
            list = await session.rpc.mcp.list();
        } catch (error) {
            throw new WorkIQError(`Couldn't list MCP servers: ${error?.message ?? error}`, { code: "mcp_list", transient: true });
        }
        const candidates = (list?.servers ?? []).filter(looksLikeWorkIQ);
        const connected = candidates.filter((s) => s.status === "connected");
        if (!connected.length) {
            const pick = candidates.find((s) => s.status === "needs-auth") ?? candidates[0];
            if (!pick) {
                throw new WorkIQError("WorkIQ isn't connected in this session. Connect the WorkIQ connector, then retry.", { code: "not_connected" });
            }
            if (pick.status === "needs-auth") {
                throw new WorkIQError("WorkIQ needs you to sign in. Open Customize › Installed, find WorkIQ, and choose Sign in.", { code: "needs_auth" });
            }
            throw new WorkIQError(`WorkIQ is ${pick.status}${pick.error ? `: ${pick.error}` : ""}.`, {
                code: "not_connected",
                transient: pick.status === "pending",
            });
        }

        let metadata = null;
        try {
            metadata = (await session.rpc.tools.getCurrentMetadata())?.tools ?? null;
        } catch (error) {
            debug(`tool metadata unavailable: ${error?.message ?? error}`);
        }
        const toolsOf = (name) => {
            const map = new Map();
            for (const tool of metadata ?? []) {
                if (tool.mcpServerName === name && tool.mcpToolName && tool.name) map.set(tool.mcpToolName, tool.name);
            }
            return map;
        };
        // Only the most trusted tier is eligible, and its candidate must actually
        // expose WorkIQ's tools (when tool metadata is available). A lower-ranked
        // look-alike is never used as a fallback, and repository-provided or
        // unknown servers (rank 2) are never used at all.
        const ranked = connected.map((s) => ({ server: s, tools: toolsOf(s.name), rank: rankOf(s) }));
        const best = Math.min(...ranked.map((c) => c.rank));
        if (best > 1) {
            throw new WorkIQError(
                "The only WorkIQ server here comes from a repository or unknown source, so the Teams canvas won't use it. Connect the WorkIQ connector instead.",
                { code: "untrusted_server" },
            );
        }
        const top = ranked.filter((c) => c.rank === best && (!metadata || REQUIRED_TOOLS.every((tool) => c.tools.has(tool))));
        if (!top.length) {
            throw new WorkIQError("WorkIQ is connected, but its tools aren't available in this session yet. Try again shortly.", {
                code: "not_connected",
                transient: true,
            });
        }
        if (top.length > 1) {
            throw new WorkIQError(
                `More than one WorkIQ server is connected (${top.map((c) => c.server.name).join(", ")}). Disable the extra one so the Teams canvas knows which to use.`,
                { code: "ambiguous_server" },
            );
        }
        const pick = top[0];
        if (server?.name !== pick.server.name) {
            debug(
                `using WorkIQ server ${JSON.stringify({
                    name: pick.server.name,
                    source: pick.server.source ?? null,
                    rank: pick.rank,
                    plugin: pick.server.sourcePlugin ?? null,
                    host: hostOf(pick.server.url),
                    verifiedTools: Boolean(metadata),
                    candidates: connected.length,
                })}`,
            );
        }
        server = { name: pick.server.name, tools: pick.tools, at: Date.now() };
        return server;
    }

    async function execute(tool, args, timeoutMs) {
        const session = await getSession();
        const srv = await resolveServer(session);
        const name = srv.tools.get(tool) ?? `${srv.name}-${tool}`;
        await acquire();
        // Hold the slot until the RPC itself settles, even after a local timeout,
        // so abandoned calls still count against the concurrency limit.
        const call = Promise.resolve().then(() => session.rpc.tools.execute({ name, arguments: args }));
        call.then(release, release);
        let timer;
        const timeout = new Promise((_, reject) => {
            timer = setTimeout(() => {
                const error = new WorkIQError("WorkIQ took too long to respond.", { code: "timeout", transient: true, ambiguous: true });
                error.call = call;
                reject(error);
            }, timeoutMs);
        });
        try {
            return await Promise.race([call, timeout]);
        } catch (error) {
            if (error instanceof WorkIQError) throw error;
            forgetServer();
            throw new WorkIQError(`WorkIQ call failed: ${error?.message ?? error}`, { code: "rpc", transient: true, ambiguous: true });
        } finally {
            clearTimeout(timer);
        }
    }

    function cooling(bucket) {
        const until = cooldowns.get(bucket) ?? 0;
        return until > Date.now() ? until - Date.now() : 0;
    }

    function assertReady(paths, kind) {
        for (const path of paths) {
            const ms = cooling(bucketOf(path, kind));
            if (ms) {
                throw new WorkIQError("Microsoft Graph is rate-limiting these requests.", {
                    status: 429,
                    code: "cooldown",
                    retryAfterMs: ms,
                    transient: true,
                });
            }
        }
    }

    function noteError(path, error, kind) {
        if (error.status === 429) {
            const ms = Math.max(error.retryAfterMs || 30_000, 5_000);
            error.retryAfterMs = ms;
            cooldowns.set(bucketOf(path, kind), Date.now() + ms);
        } else if (error.code === "CircuitOpen") {
            error.retryAfterMs = 10_000;
            cooldowns.set(bucketOf(path, kind), Date.now() + 10_000);
        }
    }

    async function fetchMany(paths, { timeoutMs = readTimeoutMs } = {}) {
        assertReady(paths, "read");
        const { payload, resultType, message } = normalize(await execute("fetch", { entityUrls: paths }, timeoutMs));
        const results = payload?.results;
        if (!Array.isArray(results)) throw unexpected(resultType, message);
        return paths.map((path, i) => {
            const entry = results[i];
            if (isOk(entry, resultType)) return { ok: true, data: entry.data };
            const error = entryError(entry);
            noteError(path, error, "read");
            return { ok: false, error };
        });
    }

    async function fetchOne(path, { retries = 1, ...options } = {}) {
        for (let attempt = 0; ; attempt++) {
            let error;
            try {
                const [result] = await fetchMany([path], options);
                if (result.ok) return result.data;
                error = result.error;
            } catch (thrown) {
                error = thrown;
            }
            const retryable = error instanceof WorkIQError && error.transient && !error.throttled && error.code !== "cooldown";
            if (!retryable || attempt >= retries) throw error;
            await sleep(1500 * (attempt + 1));
        }
    }

    // Writes are never retried. When the local timeout fires first, the error
    // carries `late`: a promise for the call's eventual outcome, so callers can
    // reconcile instead of letting the user send a duplicate.
    async function write(tool, args, path) {
        assertReady([path], "write");
        const interpret = (raw) => {
            const { payload, resultType, message } = normalize(raw);
            const entry = Array.isArray(payload?.results) ? payload.results[0] : payload;
            if (entry && isOk(entry, resultType)) return entry.data ?? null;
            if (!entry) throw unexpected(resultType, message);
            const error = entryError(entry);
            // Throttled, circuit-broken, or unimplemented (501) calls never took
            // effect; other server errors may have been applied anyway.
            error.ambiguous = error.status >= 500 && error.status !== 501 && error.code !== "CircuitOpen";
            noteError(path, error, "write");
            throw error;
        };
        let raw;
        try {
            raw = await execute(tool, args, writeTimeoutMs);
        } catch (error) {
            if (error?.call) {
                error.late = error.call.then(interpret);
                error.late.catch(() => {});
                error.call = undefined;
            }
            throw error;
        }
        return interpret(raw);
    }

    async function blob(path) {
        assertReady([path], "read");
        const { payload, resultType, message } = normalize(await execute("fetch_blob", { path }, readTimeoutMs));
        if (!payload) throw unexpected(resultType, message);
        const status = Number(payload.statusCode) || 0;
        if (payload.base64Content && (!status || (status >= 200 && status < 300))) {
            return {
                bytes: Buffer.from(payload.base64Content, "base64"),
                contentType: payload.contentType ?? payload.mimeType ?? "",
            };
        }
        const error = entryError(payload);
        noteError(path, error, "read");
        throw error;
    }

    return {
        fetchMany,
        fetchOne,
        action: (path, body) => write("do_action", { actionUrl: path, jsonBody: body }, path),
        create: (parentPath, body) => write("create_entity", { parentUrl: parentPath, jsonBody: body }, parentPath),
        blob,
        cooling,
        forgetServer,
    };
}
