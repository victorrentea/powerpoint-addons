// 🔤 Font bar, pane side: reports the font name and size of the selection to the
// server (which shows them in PowerPoint's title bar, font-bar.js) and applies
// what is typed there. Runs as long as the pane is open in the deck.
//
// The selection: the selected text if there is some; otherwise (a caret, or
// shapes selected as objects) all the text of the selected shapes.
const FontBar = (() => {
  const PANE = (crypto.randomUUID && crypto.randomUUID()) || String(Math.random()).slice(2);
  const POLL_MS = 600;   // the ribbon's own Font Size box changes the font without any event
  const TEXT_TYPES = new Set(["GeometricShape", "TextBox", "Placeholder", "Callout"]);
  let last = null;       // JSON of the last report, to send only changes
  let busy = false;

  function docName() {
    const url = Office.context.document.url || "";
    const base = decodeURIComponent(url.split(/[\\/]/).pop() || "");
    return base.replace(/\.(pptx|pptm|ppt|potx)$/i, "");
  }

  // One value when they all agree, null when mixed or none.
  function common(values) {
    const set = new Set(values.filter((v) => v !== null && v !== undefined && v !== ""));
    return set.size === 1 && !values.some((v) => v === null || v === undefined || v === "") ? [...set][0] : null;
  }

  // The text ranges the font bar shows and edits.
  async function targets(ctx) {
    const range = ctx.presentation.getSelectedTextRangeOrNullObject();
    range.load("length");
    await ctx.sync();
    if (!range.isNullObject && range.length > 0) return [range];
    const shapes = ctx.presentation.getSelectedShapes();
    shapes.load("items/type");
    await ctx.sync();
    const ranges = [];
    for (const shape of shapes.items.filter((s) => TEXT_TYPES.has(s.type))) {
      try {
        const r = shape.textFrame.textRange;
        r.load("length");
        await ctx.sync();
        if (r.length > 0) ranges.push(r);
      } catch { /* a placeholder that holds a picture has no text frame */ }
    }
    return ranges;
  }

  const SAMPLE_MAX = 24;
  const WORD_CHAR = /[\p{L}\p{N}_'’-]/u;

  // What the font list draws in each font: the selected text (first line), or
  // the word the caret is in or next to; null with no text selection.
  async function sample(ctx) {
    const sel = ctx.presentation.getSelectedTextRangeOrNullObject();
    sel.load("text,start,length");
    await ctx.sync();
    if (sel.isNullObject) return null;
    let text = sel.text;
    if (sel.length === 0) {
      const shapes = ctx.presentation.getSelectedShapes();
      shapes.load("items/type");
      await ctx.sync();
      const shape = shapes.items.find((s) => TEXT_TYPES.has(s.type));
      if (!shape) return null;
      const all = shape.textFrame.textRange;
      all.load("text");
      await ctx.sync();
      let a = sel.start, b = sel.start;
      while (a > 0 && WORD_CHAR.test(all.text[a - 1])) a--;
      while (b < all.text.length && WORD_CHAR.test(all.text[b])) b++;
      text = all.text.slice(a, b);
    }
    text = (text || "").split(/[\r\n\v]/).map((l) => l.trim()).find((l) => l) || "";
    if (!text) return null;
    return text.length > SAMPLE_MAX ? text.slice(0, SAMPLE_MAX - 1) + "…" : text;
  }

  async function read() {
    return PowerPoint.run(async (ctx) => {
      let text = null;
      try { text = await sample(ctx); } catch { /* a table cell, a chart: no sample */ }
      const ranges = await targets(ctx);
      ranges.forEach((r) => r.font.load("name,size"));
      await ctx.sync();
      return {
        name: common(ranges.map((r) => r.font.name)),
        size: common(ranges.map((r) => r.font.size)),
        sample: text,
      };
    });
  }

  async function report(force = false) {
    if (busy) return;
    busy = true;
    try {
      const font = await read();
      const body = JSON.stringify({ pane: PANE, doc: docName(), ...font });
      if (force || body !== last) {
        last = body;
        await fetch("/font/state", { method: "POST", body });
      }
    } catch (e) {
      // Mid-edit selections (a table cell, a chart) can refuse to load: next poll.
    } finally {
      busy = false;
    }
  }

  async function apply(change) {
    await PowerPoint.run(async (ctx) => {
      const ranges = await targets(ctx);
      for (const r of ranges) {
        if (change.name) r.font.name = change.name;
        if (change.size) r.font.size = change.size;
      }
      await ctx.sync();
    });
    await report(true);
  }

  function connect() {
    const events = new EventSource(`/font/events?pane=${PANE}`);
    events.onopen = () => report(true);   // a restarted server forgot us: report again
    events.onmessage = (e) => apply(JSON.parse(e.data)).catch((err) => console.error("🔤", err));
    // EventSource reconnects by itself after an error.
  }

  function start() {
    if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.5")) return;
    connect();
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => report());
    setInterval(report, POLL_MS);
  }

  return { start };
})();
