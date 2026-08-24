/**
 * Reemplazo local de la API window.storage que usan los artefactos de Claude.ai.
 * Aquí no hay servidor: todo se guarda en el localStorage del navegador,
 * es decir, SOLO en este equipo y en este navegador. Si dos personas abren
 * la app en computadoras distintas, cada una tendrá su propia lista de
 * tickets y usuarios — no se sincronizan entre sí.
 *
 * Se mantiene la misma firma (get/set/delete/list) para que App.jsx no
 * necesite ningún cambio.
 */

function claveCompleta(key, shared) {
  return `${shared ? 'shared' : 'local'}:${key}`;
}

window.storage = {
  async get(key, shared = false) {
    const raw = window.localStorage.getItem(claveCompleta(key, shared));
    if (raw === null) return null;
    return { key, value: raw, shared };
  },

  async set(key, value, shared = false) {
    window.localStorage.setItem(claveCompleta(key, shared), value);
    return { key, value, shared };
  },

  async delete(key, shared = false) {
    const k = claveCompleta(key, shared);
    const existia = window.localStorage.getItem(k) !== null;
    window.localStorage.removeItem(k);
    return { key, deleted: existia, shared };
  },

  async list(prefix = '', shared = false) {
    const prefijoCompleto = claveCompleta(prefix, shared);
    const keys = [];
    for (let i = 0; i < window.localStorage.length; i++) {
      const k = window.localStorage.key(i);
      if (k.startsWith(prefijoCompleto)) {
        keys.push(k.slice(shared ? 'shared:'.length : 'local:'.length));
      }
    }
    return { keys, prefix, shared };
  },
};
