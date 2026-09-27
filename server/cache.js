// Tiny LRU cache for parsed documents, keyed by real path + mtime + size so an
// edited file is re-read automatically.

export class DocCache {
  constructor(max, dispose = () => {}) {
    this.max = max;
    this.dispose = dispose;
    this.map = new Map();
  }

  async get(file, load) {
    const key = `${file.realPath}|${file.mtimeMs}|${file.size}`;
    if (this.map.has(key)) {
      const v = this.map.get(key);
      this.map.delete(key);
      this.map.set(key, v);
      return v;
    }
    // Drop stale versions of the same file.
    for (const [k, v] of this.map) {
      if (k.startsWith(file.realPath + "|")) {
        this.map.delete(k);
        Promise.resolve(v).then(this.dispose, () => {});
      }
    }
    const promise = load();
    this.map.set(key, promise);
    promise.catch(() => this.map.delete(key));
    while (this.map.size > this.max) {
      const [oldKey, oldVal] = this.map.entries().next().value;
      this.map.delete(oldKey);
      Promise.resolve(oldVal).then(this.dispose, () => {});
    }
    return promise;
  }
}
