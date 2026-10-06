// 🧩 Moves shapes into an animated group without losing the group's animation.
// Office.js has no animation API, and ungroup + group gives the new group a new id, which
// the slide's <p:timing> doesn't know. So the pane exports the slide and this edits its XML:
// the shapes move inside the existing <p:grpSp>, which keeps its id, so every <p:spTgt>
// pointing at it still does. Pure DOM, no Office.js, so test/regroup.test.js runs it in Node.
(function (root) {
  const P = "http://schemas.openxmlformats.org/presentationml/2006/main";
  const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
  const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
  const SHAPES = new Set(["sp", "grpSp", "pic", "cxnSp", "graphicFrame", "contentPart"]);

  const elements = (node) => [...node.childNodes].filter((n) => n.nodeType === 1);
  const child = (node, ns, name) => elements(node).find((n) => n.namespaceURI === ns && n.localName === name);
  const all = (node, ns, name) => [...node.getElementsByTagNameNS(ns, name)];
  const isShape = (el) => (el.namespaceURI === P && SHAPES.has(el.localName))
    || (el.namespaceURI === MC && el.localName === "AlternateContent");

  // An mc:AlternateContent holds the same shape once per Choice/Fallback: edit each of them.
  function variants(el) {
    if (el.namespaceURI !== MC) return [el];
    return elements(el).flatMap((branch) => elements(branch).filter(isShape));
  }

  function cNvPr(el) {
    const nv = elements(variants(el)[0])[0];
    return nv && child(nv, P, "cNvPr");
  }

  function xfrm(shape) {
    if (shape.localName === "graphicFrame") return child(shape, P, "xfrm");
    const pr = child(shape, P, shape.localName === "grpSp" ? "grpSpPr" : "spPr");
    return pr && child(pr, A, "xfrm");
  }

  function readPoint(el, x, y) {
    return { x: Number(el.getAttribute(x)), y: Number(el.getAttribute(y)) };
  }

  function box(frame) {
    const off = readPoint(child(frame, A, "off"), "x", "y");
    const ext = readPoint(child(frame, A, "ext"), "cx", "cy");
    return { x: off.x, y: off.y, w: ext.x, h: ext.y, rot: Number(frame.getAttribute("rot") || 0) };
  }

  // What the shape covers on the slide: a rotated frame reaches past its unrotated box.
  function visibleBox(b) {
    const rad = (b.rot / 60000) * Math.PI / 180;
    const cos = Math.abs(Math.cos(rad)), sin = Math.abs(Math.sin(rad));
    const w = b.w * cos + b.h * sin, h = b.w * sin + b.h * cos;
    return { x: b.x + b.w / 2 - w / 2, y: b.y + b.h / 2 - h / 2, w, h };
  }

  function union(boxes) {
    const x0 = Math.min(...boxes.map((b) => b.x)), y0 = Math.min(...boxes.map((b) => b.y));
    const x1 = Math.max(...boxes.map((b) => b.x + b.w)), y1 = Math.max(...boxes.map((b) => b.y + b.h));
    return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
  }

  function setPoint(el, x, y, p) {
    el.setAttribute(x, String(Math.round(p.x)));
    el.setAttribute(y, String(Math.round(p.y)));
  }

  // Ids of every shape some animation targets, including trigger shapes of interactive sequences.
  function animatedIds(doc) {
    return new Set(all(doc, P, "spTgt").map((t) => t.getAttribute("spid")));
  }

  function fail(message) {
    throw new Error(message);
  }

  // A moved shape can't keep its own animation (PowerPoint only animates top-level shapes):
  // drop its effects, then any click/with step left empty by that.
  function dropAnimations(doc, id) {
    const timing = all(doc, P, "timing")[0];
    if (!timing) return 0;
    const targets = all(timing, P, "spTgt").filter((t) => t.getAttribute("spid") === id);
    let dropped = 0;
    for (const target of targets) {
      let node = target.parentNode;
      while (node && !(node.localName === "cond" && node.namespaceURI === P)
        && !(node.localName === "cTn" && node.getAttribute("presetClass"))) node = node.parentNode;
      if (!node) continue;
      if (node.localName === "cond") fail(`"${nameOf(doc, id)}" triggers an animation. Remove that trigger first.`);
      removeTimeNode(timing, node.parentNode);
      dropped++;
    }
    for (const name of ["bldP", "bldGraphic", "bldOleChart", "bldDgm"]) {
      for (const b of all(timing, P, name)) if (b.getAttribute("spid") === id) b.parentNode.removeChild(b);
    }
    if (timing.parentNode) renumberTimeNodes(timing);
    return dropped;
  }

  // Removes a <p:par>/<p:seq>, then each step it leaves empty (with-step, click-step, sequence).
  // An empty sequence isn't valid; with nothing left at all, the whole <p:timing> goes.
  function removeTimeNode(timing, node) {
    for (;;) {
      const list = node.parentNode;
      list.removeChild(node);
      if (elements(list).length) return;
      const owner = list.parentNode;
      if (owner.getAttribute("nodeType") === "tmRoot") return void timing.parentNode.removeChild(timing);
      node = owner.parentNode;
    }
  }

  // Keeps cTn ids dense (1..n), as PowerPoint writes them, after removing some.
  function renumberTimeNodes(timing) {
    const ids = new Map();
    all(timing, P, "cTn").forEach((cTn, i) => {
      ids.set(cTn.getAttribute("id"), String(i + 1));
      cTn.setAttribute("id", String(i + 1));
    });
    for (const tn of all(timing, P, "tn")) {
      if (ids.has(tn.getAttribute("val"))) tn.setAttribute("val", ids.get(tn.getAttribute("val")));
    }
  }

  function nameOf(doc, id) {
    const pr = all(doc, P, "cNvPr").find((c) => c.getAttribute("id") === id);
    return pr ? pr.getAttribute("name") : `shape ${id}`;
  }

  /**
   * `marks`: Map of marker name → original name, one per selected top-level shape.
   * Finds them by marker in the slide, picks the animated group among them, moves the
   * others into it and gives every shape its name back.
   * Returns { group, added, droppedAnimations }.
   */
  function addToGroup(doc, marks) {
    const tree = all(doc, P, "spTree")[0] || fail("The slide has no shapes.");
    const top = elements(tree).filter(isShape);
    const selected = top.filter((el) => marks.has(cNvPr(el)?.getAttribute("name")));
    const names = new Map(selected.map((el) => [el, marks.get(cNvPr(el).getAttribute("name"))]));
    for (const el of selected) for (const v of variants(el)) cNvPr(v).setAttribute("name", names.get(el));
    if (selected.length !== marks.size) fail("Couldn't find every selected shape in the exported slide.");
    if (selected.length < 2) fail("Select the animated group and the shapes to add to it.");

    const animated = animatedIds(doc);
    const groups = selected.filter((el) => el.localName === "grpSp" && el.namespaceURI === P);
    const animatedGroups = groups.filter((g) => animated.has(cNvPr(g).getAttribute("id")));
    if (!groups.length) fail("No group in the selection. Select the animated group too.");
    if (animatedGroups.length > 1) {
      fail(`${animatedGroups.length} animated groups are selected (${animatedGroups.map((g) => `"${names.get(g)}"`).join(", ")}). Keep only the one to add to.`);
    }
    if (!animatedGroups.length && groups.length > 1) fail("More than one group is selected, and none is animated. Select only one group.");
    const group = animatedGroups[0] || groups[0];
    const others = selected.filter((el) => el !== group);

    for (const el of others) {
      if (variants(el).some((v) => all(v, P, "ph").length)) fail(`"${names.get(el)}" is a placeholder; PowerPoint doesn't group placeholders.`);
      if (variants(el).some((v) => all(v, A, "tbl").length)) fail(`"${names.get(el)}" is a table; PowerPoint doesn't group tables.`);
      if (variants(el).some((v) => !xfrm(v))) fail(`"${names.get(el)}" has no position of its own; can't group it.`);
    }

    const gFrame = xfrm(group);
    if (gFrame.getAttribute("rot") || gFrame.getAttribute("flipH") || gFrame.getAttribute("flipV")) {
      fail(`"${names.get(group)}" is rotated or flipped; un-rotate it first.`);
    }
    const off = readPoint(child(gFrame, A, "off"), "x", "y");
    const ext = readPoint(child(gFrame, A, "ext"), "cx", "cy");
    const chOff = readPoint(child(gFrame, A, "chOff"), "x", "y");
    const chExt = readPoint(child(gFrame, A, "chExt"), "cx", "cy");
    // Children live in the group's own coordinate space, scaled when the group was resized.
    // Some tools write an all-zero frame, which PowerPoint reads as "no scaling".
    const sx = ext.x && chExt.x ? chExt.x / ext.x : 1, sy = ext.y && chExt.y ? chExt.y / ext.y : 1;
    const toSlide = (b) => ({ x: off.x + (b.x - chOff.x) / sx, y: off.y + (b.y - chOff.y) / sy, w: b.w / sx, h: b.h / sy });
    const g = ext.x && ext.y ? { x: off.x, y: off.y, w: ext.x, h: ext.y }
      : union(elements(group).filter(isShape).map((el) => toSlide(visibleBox(box(xfrm(variants(el)[0]))))));

    const bounds = union([g, ...others.map((el) => visibleBox(box(xfrm(variants(el)[0]))))]);

    // Slide → group space, then into the group: above it if it was above, below if below.
    const groupIndex = top.indexOf(group);
    const firstChild = elements(group).find(isShape) || null;
    let droppedAnimations = 0;
    for (const el of others) {
      for (const v of variants(el)) {
        const f = xfrm(v), b = box(f);
        setPoint(child(f, A, "off"), "x", "y", { x: chOff.x + (b.x - off.x) * sx, y: chOff.y + (b.y - off.y) * sy });
        setPoint(child(f, A, "ext"), "cx", "cy", { x: b.w * sx, y: b.h * sy });
      }
      droppedAnimations += dropAnimations(doc, cNvPr(el).getAttribute("id"));
      tree.removeChild(el);
      if (top.indexOf(el) < groupIndex) group.insertBefore(el, firstChild);
      else group.appendChild(el);
    }

    // Grow the group to cover the newcomers, keeping its scale so old children stay put.
    setPoint(child(gFrame, A, "off"), "x", "y", bounds);
    setPoint(child(gFrame, A, "ext"), "cx", "cy", { x: bounds.w, y: bounds.h });
    setPoint(child(gFrame, A, "chOff"), "x", "y", { x: chOff.x + (bounds.x - off.x) * sx, y: chOff.y + (bounds.y - off.y) * sy });
    setPoint(child(gFrame, A, "chExt"), "cx", "cy", { x: bounds.w * sx, y: bounds.h * sy });

    return {
      group: names.get(group),
      animated: animated.has(cNvPr(group).getAttribute("id")),
      added: others.map((el) => names.get(el)),
      droppedAnimations,
    };
  }

  const api = { addToGroup };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.Regroup = api;
})(this);
