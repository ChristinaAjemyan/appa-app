const BASE = '/api/calc-result-storage';

export const calcStorage = {
  async get(key) {
    const res = await fetch(`${BASE}/${encodeURIComponent(key)}`);
    if (!res.ok) return null;
    const data = await res.json();
    return { value: JSON.stringify(data.value) };
  },

  async set(key, jsonStr) {
    const value = JSON.parse(jsonStr);
    await fetch(`${BASE}/${encodeURIComponent(key)}`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ value }),
    });
  },

  async list(prefix) {
    const res = await fetch(`${BASE}?prefix=${encodeURIComponent(prefix)}`);
    if (!res.ok) return { keys: [] };
    return res.json();
  },
};
