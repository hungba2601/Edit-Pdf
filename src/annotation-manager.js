export class AnnotationManager {
  constructor() {
    this.pages = {}; // pageIndex -> array of annotation objects
    this.history = [];
    this.redoStack = [];
  }

  add(pageIdx, ann) {
    if (!this.pages[pageIdx]) this.pages[pageIdx] = [];
    ann.id = Date.now() + '_' + Math.random().toString(36).slice(2,7);
    this.pages[pageIdx].push(ann);
    this.history.push({ action: 'add', pageIdx, ann: { ...ann } });
    this.redoStack = [];
    return ann;
  }

  update(pageIdx, id, props) {
    const arr = this.pages[pageIdx] || [];
    const a = arr.find(x => x.id === id);
    if (a) {
      const old = { ...a };
      Object.assign(a, props);
      this.history.push({ action: 'update', pageIdx, id, old, updated: { ...a } });
      this.redoStack = [];
    }
  }

  remove(pageIdx, id) {
    const arr = this.pages[pageIdx] || [];
    const idx = arr.findIndex(x => x.id === id);
    if (idx !== -1) {
      const removed = arr.splice(idx, 1)[0];
      this.history.push({ action: 'remove', pageIdx, ann: removed, idx });
      this.redoStack = [];
      return removed;
    }
  }

  undo() {
    const h = this.history.pop();
    if (!h) return null;
    this.redoStack.push(h);
    if (h.action === 'add') {
      const arr = this.pages[h.pageIdx] || [];
      const i = arr.findIndex(x => x.id === h.ann.id);
      if (i !== -1) arr.splice(i, 1);
    } else if (h.action === 'remove') {
      if (!this.pages[h.pageIdx]) this.pages[h.pageIdx] = [];
      this.pages[h.pageIdx].splice(h.idx, 0, h.ann);
    } else if (h.action === 'update') {
      const arr = this.pages[h.pageIdx] || [];
      const a = arr.find(x => x.id === h.id);
      if (a) Object.assign(a, h.old);
    }
    return h;
  }

  redo() {
    const h = this.redoStack.pop();
    if (!h) return null;
    this.history.push(h);
    if (h.action === 'add') {
      if (!this.pages[h.pageIdx]) this.pages[h.pageIdx] = [];
      this.pages[h.pageIdx].push({ ...h.ann });
    } else if (h.action === 'remove') {
      const arr = this.pages[h.pageIdx] || [];
      const i = arr.findIndex(x => x.id === h.ann.id);
      if (i !== -1) arr.splice(i, 1);
    } else if (h.action === 'update') {
      const arr = this.pages[h.pageIdx] || [];
      const a = arr.find(x => x.id === h.id);
      if (a) Object.assign(a, h.updated);
    }
    return h;
  }

  clear() {
    this.pages = {};
    this.history = [];
    this.redoStack = [];
  }

  getPage(pageIdx) { return this.pages[pageIdx] || []; }
  getAll() { return this.pages; }
}
