/**
 * window.storage polyfill
 * -----------------------------------------------------------------------------
 * ClientMaster.jsx persists data via an async key-value API exposed as
 * `window.storage` (this is the Claude Artifacts storage API). That object does
 * NOT exist in a normal browser / Vercel deployment, so we provide a drop-in
 * replacement backed by localStorage.
 *
 * The original component only ever uses the "personal" (shared = false) scope
 * and calls:
 *
 *   await window.storage.get(key, shared)     -> { key, value, shared } | null
 *   await window.storage.set(key, value, shared)
 *   await window.storage.delete(key, shared)
 *
 * `get` must return an object with a `.value` property (string) or null when the
 * key is absent. We mirror that contract exactly so ClientMaster.jsx works
 * unchanged. `list` is included for completeness/future use.
 *
 * Keys are namespaced per scope so a future "shared" usage can't collide with
 * personal data.
 */

const PREFIX = 'cm-storage:'; // namespace to avoid clashing with other localStorage keys

function physicalKey(key, shared) {
  return `${PREFIX}${shared ? 'shared' : 'personal'}:${key}`;
}

const memoryFallback = new Map(); // used if localStorage is unavailable (e.g. private mode)

function hasLocalStorage() {
  try {
    const t = '__cm_test__';
    window.localStorage.setItem(t, '1');
    window.localStorage.removeItem(t);
    return true;
  } catch (_) {
    return false;
  }
}

const useLS = typeof window !== 'undefined' && hasLocalStorage();

function rawGet(pk) {
  if (useLS) return window.localStorage.getItem(pk);
  return memoryFallback.has(pk) ? memoryFallback.get(pk) : null;
}
function rawSet(pk, value) {
  if (useLS) window.localStorage.setItem(pk, value);
  else memoryFallback.set(pk, value);
}
function rawDelete(pk) {
  if (useLS) window.localStorage.removeItem(pk);
  else memoryFallback.delete(pk);
}
function rawKeys() {
  if (useLS) {
    const out = [];
    for (let i = 0; i < window.localStorage.length; i++) out.push(window.localStorage.key(i));
    return out;
  }
  return Array.from(memoryFallback.keys());
}

const storage = {
  async get(key, shared = false) {
    const pk = physicalKey(key, shared);
    const value = rawGet(pk);
    if (value === null || value === undefined) return null;
    return { key, value, shared: !!shared };
  },

  async set(key, value, shared = false) {
    const pk = physicalKey(key, shared);
    // Persist exactly what the caller stored; the component always passes strings.
    const toStore = typeof value === 'string' ? value : JSON.stringify(value);
    rawSet(pk, toStore);
    return { key, value: toStore, shared: !!shared };
  },

  async delete(key, shared = false) {
    const pk = physicalKey(key, shared);
    rawDelete(pk);
    return { key, deleted: true, shared: !!shared };
  },

  async list(prefix = '', shared = false) {
    const scopePrefix = `${PREFIX}${shared ? 'shared' : 'personal'}:`;
    const keys = rawKeys()
      .filter((k) => k && k.startsWith(scopePrefix))
      .map((k) => k.slice(scopePrefix.length))
      .filter((k) => k.startsWith(prefix));
    return { keys, prefix, shared: !!shared };
  },
};

export function installStoragePolyfill() {
  if (typeof window === 'undefined') return;
  // Only install if the host environment hasn't already provided window.storage.
  if (!window.storage) {
    window.storage = storage;
  }
}

export default storage;
