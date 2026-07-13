export const storage = {
  async get(key: string) {
    if (typeof window === "undefined") return null;
    const v = localStorage.getItem(key);
    return v !== null ? { key, value: v } : null;
  },
  async set(key: string, value: string) {
    if (typeof window === "undefined") return null;
    localStorage.setItem(key, value);
    return { key, value };
  },
  async delete(key: string) {
    if (typeof window === "undefined") return null;
    localStorage.removeItem(key);
    return { key, deleted: true };
  },
  async list(prefix: string = "") {
    if (typeof window === "undefined") return { keys: [] };
    const keys = Object.keys(localStorage).filter((k) => k.startsWith(prefix));
    return { keys };
  },
};
