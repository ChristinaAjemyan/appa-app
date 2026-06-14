const BASE = '/api/calc-result-storage';

export const calcStorage = {
  async get(key) {
    const res = await fetch(`${BASE}/${encodeURIComponent(key)}`);
    if (res.status === 404) return null;
    if (!res.ok) throw new Error(`Storage error: ${res.status}`);
    const data = await res.json();
    return { value: JSON.stringify(data.value) };
  },

  async set(key, jsonStr) {
    const value = JSON.parse(jsonStr);
    const res = await fetch(`${BASE}/${encodeURIComponent(key)}`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ value }),
    });
    if (!res.ok) throw new Error(`Storage save failed (${res.status})`);
  },

  async list(prefix) {
    const res = await fetch(`${BASE}?prefix=${encodeURIComponent(prefix)}`);
    if (!res.ok) return { keys: [] };
    return res.json();
  },
};
