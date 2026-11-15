"""Notebook: pages like Notion. Left = list (search, folder, pinned first). Right = the page: title, folder, tags, worker, a rich text editor, photos.
You can paste from the chat (headings, bullets, bold, tables stay the same). Everything saves by itself."""
import html as _html
import re
from datetime import datetime
from html.parser import HTMLParser

import streamlit as st

import attendance_db as db

IMG = ("png", "jpg", "jpeg", "webp", "gif")
MARK = "<!--rich-->"                                     # the text of a page starts with this when it was written in the rich editor

# ---------------------------------------------------------------- keep only safe, simple HTML
_OK = {"h1", "h2", "h3", "h4", "p", "br", "ul", "ol", "li", "strong", "em", "u", "s", "code", "pre", "blockquote", "table", "thead", "tbody", "tr", "th", "td", "a", "hr", "input"}
_DROP = {"script", "style", "iframe", "object", "embed", "head", "title", "svg", "noscript", "textarea", "select", "button"}
_VOID = {"br", "hr", "input"}
_ALIAS = {"b": "strong", "i": "em", "strike": "s", "del": "s"}


class _Clean(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.out, self.skip, self.stack = [], 0, []

    def handle_starttag(self, tag, attrs):
        tag = _ALIAS.get(tag, tag)
        if tag in _DROP:
            self.skip += 1
            return
        if tag == "span":
            a = dict(attrs)
            if "tex" in (a.get("class") or "").split() and a.get("data-tex"):
                self.out.append(f'<span class="tex" data-tex="{_html.escape(a["data-tex"][:2000], quote=True)}" data-display="{1 if a.get("data-display") == "1" else 0}">')
                self.stack.append("span")
            return
        if self.skip or tag not in _OK:
            return
        a = dict(attrs)
        extra = ""
        if tag == "a":
            h = (a.get("href") or "").strip()
            if re.match(r"(?i)^(https?:|mailto:)", h):
                extra = f' href="{_html.escape(h, quote=True)}"'
        elif tag == "input":
            if (a.get("type") or "").lower() != "checkbox":
                return
            extra = ' type="checkbox"' + (" checked" if "checked" in a else "")
        elif tag in ("td", "th"):
            for k in ("colspan", "rowspan"):
                if (a.get(k) or "").isdigit():
                    extra += f' {k}="{a[k]}"'
        self.out.append(f"<{tag}{extra}>")
        if tag not in _VOID:
            self.stack.append(tag)

    def handle_endtag(self, tag):
        tag = _ALIAS.get(tag, tag)
        if tag in _DROP:
            self.skip = max(0, self.skip - 1)
            return
        if tag == "span":
            if self.stack and self.stack[-1] == "span":
                self.stack.pop()
                self.out.append("</span>")
            return
        if self.skip or tag not in _OK or tag in _VOID:
            return
        if tag in self.stack:
            while self.stack:
                t = self.stack.pop()
                self.out.append(f"</{t}>")
                if t == tag:
                    break

    def handle_data(self, data):
        if not self.skip:
            self.out.append(_html.escape(data, quote=False))

    def close(self):
        super().close()
        while self.stack:
            self.out.append(f"</{self.stack.pop()}>")


def clean_html(s):
    p = _Clean()
    p.feed(s or "")
    p.close()
    return "".join(p.out)


def body_to_html(body):
    """A page saved by the rich editor starts with MARK. An older plain-text page becomes simple paragraphs."""
    body = body or ""
    if body.startswith(MARK):
        return body[len(MARK):]
    paras = [p for p in re.split(r"\n{2,}", body.replace("\r", "")) if p.strip()]
    return "".join("<p>" + _html.escape(p).replace("\n", "<br>") + "</p>" for p in paras)


def plain(body):
    """Text without tags (for search results and short previews)."""
    t = re.sub(r'data-tex="([^"]*)"[^>]*>', r'> \1 ', body_to_html(body))
    t = re.sub(r"<[^>]+>", " ", t)
    return re.sub(r"\s+", " ", _html.unescape(t)).strip()


RICH_CSS = r""".rt-wrap { border: 1px solid rgba(128,128,128,.35); border-radius: 10px; overflow: hidden; background: var(--st-background-color, #fff); }
.rt-bar { display: flex; flex-wrap: wrap; gap: 4px; align-items: center; padding: 6px 8px; background: var(--st-secondary-background-color, #f3f3f3); border-bottom: 1px solid rgba(128,128,128,.3); }
.rt-bar button { font: inherit; font-size: 13px; padding: 4px 9px; border-radius: 6px; border: 1px solid rgba(128,128,128,.4); background: var(--st-background-color, #fff); color: var(--st-text-color, #111); cursor: pointer; }
.rt-bar button:hover { border-color: var(--st-primary-color, #d9480f); }
.rt-st { margin-left: auto; font-size: 12px; opacity: .7; }
.rt-ed { min-height: 460px; max-height: 72vh; overflow: auto; padding: 16px 22px; outline: none; line-height: 1.65; font-size: 16px; color: var(--st-text-color, #111); }
.rt-ed h1 { font-size: 1.9em; margin: .7em 0 .3em; } .rt-ed h2 { font-size: 1.5em; margin: .7em 0 .3em; } .rt-ed h3 { font-size: 1.2em; margin: .6em 0 .3em; } .rt-ed h4 { font-size: 1.05em; margin: .6em 0 .3em; }
.rt-ed p { margin: .45em 0; } .rt-ed ul, .rt-ed ol { margin: .4em 0 .4em 1.4em; padding-left: .6em; } .rt-ed li { margin: .15em 0; }
.rt-ed blockquote { border-left: 4px solid rgba(128,128,128,.5); margin: .6em 0; padding: .1em 1em; opacity: .9; }
.rt-ed pre { background: rgba(128,128,128,.15); padding: 10px 12px; border-radius: 8px; overflow: auto; font-family: ui-monospace, Consolas, monospace; font-size: 14px; }
.rt-ed code { background: rgba(128,128,128,.18); padding: 1px 5px; border-radius: 4px; font-family: ui-monospace, Consolas, monospace; font-size: .92em; }
.rt-ed pre code { background: none; padding: 0; }
.rt-ed table { border-collapse: collapse; margin: .6em 0; } .rt-ed th, .rt-ed td { border: 1px solid rgba(128,128,128,.55); padding: 5px 10px; min-width: 60px; } .rt-ed th { background: rgba(128,128,128,.15); }
.rt-ed hr { border: none; border-top: 1px solid rgba(128,128,128,.5); margin: 1em 0; } .rt-ed a { color: var(--st-link-color, #1c7ed6); }
.rt-ed input[type=checkbox] { width: 16px; height: 16px; vertical-align: -2px; margin-right: 4px; cursor: pointer; }
.rt-ed span.tex { display: inline-block; cursor: pointer; vertical-align: middle; padding: 0 2px; border-radius: 4px; }
.rt-ed span.tex:hover { background: rgba(128,128,128,.15); }
.rt-ed span.tex[data-display="1"] { display: block; text-align: center; margin: .7em 0; }
.rt-ed math { font-family: "Latin Modern Math", "STIX Two Math", "Cambria Math", "Cambria", "Times New Roman", serif; font-size: 1.2em; }
.rt-main { display: flex; align-items: stretch; }
.rt-main .rt-ed { flex: 1 1 auto; min-width: 0; position: relative; }
.rt-toc { flex: 0 0 230px; max-height: 72vh; overflow: auto; margin: 12px 12px 12px 0; padding: 10px 6px; border: 1px solid rgba(128,128,128,.35); border-radius: 12px; font-size: 13px; align-self: flex-start; position: sticky; top: 8px; background: var(--st-secondary-background-color, #f6f6f6); }
.rt-toc-t { color: var(--st-primary-color, #d9480f); font-weight: 600; padding: 4px 10px 8px; }
.rt-toc-i { padding: 5px 10px; border-radius: 6px; cursor: pointer; opacity: .85; line-height: 1.3; overflow: hidden; text-overflow: ellipsis; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; }
.rt-toc-i:hover { background: rgba(128,128,128,.18); opacity: 1; }
.rt-toc-i.on { background: rgba(128,128,128,.22); font-weight: 600; opacity: 1; }
.rt-toc-i.l2 { padding-left: 20px; } .rt-toc-i.l3 { padding-left: 30px; } .rt-toc-i.l4 { padding-left: 40px; }
@media (max-width: 900px) { .rt-toc { display: none !important; } }
"""

RICH_JS = r"""export default function (component) {
  const { data, setStateValue, parentElement } = component;
  const root = parentElement;
  if (root.__rt_id === data.id) return;            // same page: do not rebuild (keeps the cursor)
  root.__rt_id = data.id;
  root.innerHTML = "";

  const ALLOW = new Set(["H1","H2","H3","H4","P","BR","UL","OL","LI","STRONG","EM","U","S","CODE","PRE","BLOCKQUOTE","TABLE","THEAD","TBODY","TR","TH","TD","A","HR","INPUT"]);
  const SKIP = new Set(["SCRIPT","STYLE","META","LINK","TITLE","HEAD","NOSCRIPT","SVG","IMG","BUTTON","SELECT","TEXTAREA","IFRAME"]);
  const BLOCK = new Set(["H1","H2","H3","H4","P","UL","OL","LI","PRE","BLOCKQUOTE","TABLE","TR","DIV","HR"]);

  function walk(n) {
    const frag = document.createDocumentFragment();
    if (n.nodeType === 3) { frag.appendChild(textFrag(n.nodeValue)); return frag; }
    if (n.nodeType !== 1) return frag;
    let tag = n.tagName.toUpperCase();
    const cl = n.classList;
    if (cl && (cl.contains("katex-display") || cl.contains("katex") || tag === "MATH" || tag === "MJX-CONTAINER")) {      // a formula from the chat
      const an = n.querySelector && n.querySelector('annotation[encoding="application/x-tex"]');
      if (an && an.textContent.trim()) {
        const disp = cl.contains("katex-display") || (n.parentElement && n.parentElement.classList && n.parentElement.classList.contains("katex-display")) || n.getAttribute("display") === "block" || n.getAttribute("display") === "true";
        frag.appendChild(texEl(an.textContent, disp)); return frag;
      }
      if (cl.contains("katex-display") || cl.contains("katex") || tag === "MATH") return frag;     // no source found: drop it (the text version follows)
    }
    if (tag === "SPAN" && cl && cl.contains("tex") && n.getAttribute("data-tex")) { frag.appendChild(texEl(n.getAttribute("data-tex"), n.getAttribute("data-display") === "1")); return frag; }
    if (SKIP.has(tag)) return frag;
    const kids = () => { const f = document.createDocumentFragment(); n.childNodes.forEach(c => f.appendChild(walk(c))); return f; };
    if (tag === "B") tag = "STRONG";
    if (tag === "I") tag = "EM";
    if (tag === "SPAN" || tag === "FONT") {
      const st = (n.getAttribute("style") || "").toLowerCase();
      let f = kids();
      const wrap = (t) => { const e = document.createElement(t); e.appendChild(f); f = document.createDocumentFragment(); f.appendChild(e); };
      if (/font-weight:\s*(bold|[6-9]00)/.test(st)) wrap("strong");
      if (/font-style:\s*italic/.test(st)) wrap("em");
      if (/text-decoration[^;]*underline/.test(st)) wrap("u");
      return f;
    }
    if (tag === "DIV" || tag === "SECTION" || tag === "ARTICLE") {
      const hasBlock = Array.from(n.children).some(c => BLOCK.has(c.tagName.toUpperCase()));
      if (hasBlock) return kids();
      const p = document.createElement("p"); p.appendChild(kids()); frag.appendChild(p); return frag;
    }
    if (!ALLOW.has(tag)) return kids();
    if (tag === "INPUT") {
      if ((n.getAttribute("type") || "").toLowerCase() !== "checkbox") return frag;
      const c = document.createElement("input"); c.type = "checkbox"; if (n.checked || n.hasAttribute("checked")) c.setAttribute("checked", ""); frag.appendChild(c); return frag;
    }
    const e = document.createElement(tag.toLowerCase());
    if (tag === "A") { const h = n.getAttribute("href") || ""; if (/^(https?:|mailto:)/i.test(h)) e.setAttribute("href", h); }
    if (tag === "TD" || tag === "TH") { ["colspan","rowspan"].forEach(a => { if (n.getAttribute(a)) e.setAttribute(a, n.getAttribute(a)); }); }
    e.appendChild(kids());
    frag.appendChild(e);
    return frag;
  }
  function cleanHTML(html) {
    const doc = new DOMParser().parseFromString(html, "text/html");
    const f = walk(doc.body);
    const d = document.createElement("div"); d.appendChild(f);
    return d.innerHTML;
  }

  const esc = s => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const attr = s => s.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const MATH_RE = /\\\[([\s\S]+?)\\\]|\\\(([\s\S]+?)\\\)|\$\$([\s\S]+?)\$\$/g;
  const mathHtml = (tex, disp) => '<span class="tex" data-tex="' + attr(tex.trim()) + '" data-display="' + (disp ? 1 : 0) + '"></span>';
  const inl0 = s => esc(s).replace(/`([^`]+)`/g, "<code>$1</code>").replace(/\*\*([^*]+)\*\*/g, "<strong>$1</strong>")
    .replace(/(^|[^*])\*([^*\s][^*]*)\*/g, "$1<em>$2</em>").replace(/\[([^\]]+)\]\((https?:[^)\s]+)\)/g, '<a href="$2">$1</a>');
  const inl = s => {
    let out = "", last = 0, m; MATH_RE.lastIndex = 0;
    while ((m = MATH_RE.exec(s))) {
      out += inl0(s.slice(last, m.index));
      out += mathHtml(m[1] !== undefined ? m[1] : m[2] !== undefined ? m[2] : m[3], m[2] === undefined);
      last = m.index + m[0].length;
    }
    return out + inl0(s.slice(last));
  };
  function texEl(tex, disp) {
    const s = document.createElement("span"); s.className = "tex"; s.setAttribute("data-tex", tex.trim()); s.setAttribute("data-display", disp ? "1" : "0"); return s;
  }
  function textFrag(str) {                                   // text that holds \( \) \[ \] $$ $$ formulas
    const f = document.createDocumentFragment(); let last = 0, m; MATH_RE.lastIndex = 0;
    while ((m = MATH_RE.exec(str))) {
      if (m.index > last) f.appendChild(document.createTextNode(str.slice(last, m.index)));
      f.appendChild(texEl(m[1] !== undefined ? m[1] : m[2] !== undefined ? m[2] : m[3], m[2] === undefined));
      last = m.index + m[0].length;
    }
    if (last < str.length) f.appendChild(document.createTextNode(str.slice(last)));
    return f;
  }
  function looksMd(t) { return /^(#{1,4}\s|\s*[-*+]\s|\s*\d+[.)]\s|```|\s*\|.*\|)/m.test(t) || /\*\*[^*\n]+\*\*/.test(t) || /\\\(|\\\[|\$\$/.test(t); }
  function md(text) {
    const lines = text.replace(/\r/g, "").split("\n"); const out = []; let i = 0;
    const isList = l => /^\s*([-*+]|\d+[.)])\s+/.test(l);
    function list(startI) {                       // nested lists by indentation
      const items = []; let j = startI;
      while (j < lines.length && isList(lines[j])) {
        const m = lines[j].match(/^(\s*)([-*+]|\d+[.)])\s+(.*)$/); items.push({ ind: m[1].replace(/\t/g, "    ").length, ord: /\d/.test(m[2]), txt: m[3] }); j++;
      }
      function build(from, ind) {
        let html = ""; let k = from; const ord = items[from].ord; html += ord ? "<ol>" : "<ul>";
        while (k < items.length && items[k].ind >= ind) {
          if (items[k].ind > ind) { const r = build(k, items[k].ind); html = html.replace(/<\/li>$/, "") + r.html + "</li>"; k = r.next; continue; }
          let t = items[k].txt; let cb = "";
          const c = t.match(/^\[( |x|X)\]\s+(.*)$/); if (c) { cb = '<input type="checkbox"' + (c[1] !== " " ? " checked" : "") + "> "; t = c[2]; }
          html += "<li>" + cb + inl(t) + "</li>"; k++;
        }
        html += ord ? "</ol>" : "</ul>"; return { html, next: k };
      }
      return { html: build(0, items[0].ind).html, next: j };
    }
    while (i < lines.length) {
      const l = lines[i]; let m;
      if (/^\s*(\\\[|\$\$)/.test(l) && !/(\\\[[\s\S]+\\\]|\$\$[\s\S]+\$\$)/.test(l)) {
        let acc = l; while (!/(\\\[[\s\S]+\\\]|\$\$[\s\S]+\$\$)/.test(acc) && i + 1 < lines.length) { i++; acc += "\n" + lines[i]; }
        i++; out.push("<p>" + inl(acc) + "</p>"); continue;
      }
      if (/^```/.test(l)) { const b = []; i++; while (i < lines.length && !/^```/.test(lines[i])) { b.push(lines[i]); i++; } i++; out.push("<pre>" + esc(b.join("\n")) + "</pre>"); continue; }
      if ((m = l.match(/^(#{1,4})\s+(.*)/))) { out.push("<h" + m[1].length + ">" + inl(m[2]) + "</h" + m[1].length + ">"); i++; continue; }
      if (/^\s*([-*_])\1\1+\s*$/.test(l)) { out.push("<hr>"); i++; continue; }
      if (/^\s*>/.test(l)) { const b = []; while (i < lines.length && /^\s*>/.test(lines[i])) { b.push(inl(lines[i].replace(/^\s*>\s?/, ""))); i++; } out.push("<blockquote>" + b.join("<br>") + "</blockquote>"); continue; }
      if (/^\s*\|.*\|\s*$/.test(l) && i + 1 < lines.length && /^\s*\|[\s:|-]+\|\s*$/.test(lines[i + 1])) {
        const cells = r => r.trim().replace(/^\||\|$/g, "").split("|").map(c => inl(c.trim()));
        let h = "<table><thead><tr>" + cells(l).map(c => "<th>" + c + "</th>").join("") + "</tr></thead><tbody>"; i += 2;
        while (i < lines.length && /^\s*\|.*\|\s*$/.test(lines[i])) { h += "<tr>" + cells(lines[i]).map(c => "<td>" + c + "</td>").join("") + "</tr>"; i++; }
        out.push(h + "</tbody></table>"); continue;
      }
      if (isList(l)) { const r = list(i); out.push(r.html); i = r.next; continue; }
      if (l.trim() === "") { i++; continue; }
      const b = [];
      while (i < lines.length && lines[i].trim() !== "" && !/^(#{1,4}\s|```|\s*>|\s*\|)/.test(lines[i]) && !isList(lines[i]) && !/^\s*([-*_])\1\1+\s*$/.test(lines[i])) { b.push(inl(lines[i])); i++; }
      if (b.length) out.push("<p>" + b.join("<br>") + "</p>"); else i++;
    }
    return out.join("");
  }
  function plainToHtml(t) { return t.replace(/\r/g, "").split(/\n{2,}/).map(p => "<p>" + p.split("\n").map(inl).join("<br>") + "</p>").join(""); }

  // ---- screen
  const wrap = document.createElement("div"); wrap.className = "rt-wrap";
  const bar = document.createElement("div"); bar.className = "rt-bar";
  const ed = document.createElement("div"); ed.className = "rt-ed"; ed.contentEditable = "true"; ed.spellcheck = true;
  ed.innerHTML = data.html || "<p><br></p>";
  function renderTex(el) {
    const tex = el.getAttribute("data-tex") || ""; const disp = el.getAttribute("data-display") === "1";
    el.contentEditable = "false"; el.title = "Double-click to change the formula";
    if (typeof katex !== "undefined" && katex) { try { el.innerHTML = katex.renderToString(tex, { output: "mathml", displayMode: disp, throwOnError: false }); } catch (e) { el.textContent = tex; } }
    else el.textContent = (disp ? "\\[" : "\\(") + tex + (disp ? "\\]" : "\\)");
  }
  function renderAll() { ed.querySelectorAll("span.tex").forEach(renderTex); }
  function serialize() { const c = ed.cloneNode(true); c.querySelectorAll("span.tex").forEach(s => { s.innerHTML = ""; s.removeAttribute("contenteditable"); s.removeAttribute("title"); }); return c.innerHTML; }
  renderAll();
  const status = document.createElement("span"); status.className = "rt-st";
  const BTN = [
    ["B", "Bold (Ctrl+B)", () => document.execCommand("bold")], ["I", "Italic (Ctrl+I)", () => document.execCommand("italic")], ["U", "Underline (Ctrl+U)", () => document.execCommand("underline")],
    ["H1", "Big heading", () => document.execCommand("formatBlock", false, "h1")], ["H2", "Heading", () => document.execCommand("formatBlock", false, "h2")], ["H3", "Small heading", () => document.execCommand("formatBlock", false, "h3")],
    ["¶", "Normal text", () => document.execCommand("formatBlock", false, "p")],
    ["• List", "Bullet list", () => document.execCommand("insertUnorderedList")], ["1. List", "Numbered list", () => document.execCommand("insertOrderedList")],
    ["☐ Task", "Checkbox", () => document.execCommand("insertHTML", false, '<input type="checkbox"> ')],
    ["❝", "Quote", () => document.execCommand("formatBlock", false, "blockquote")], ["―", "Line", () => document.execCommand("insertHorizontalRule")],
    ["▦ Table", "Insert a table 3x3", () => document.execCommand("insertHTML", false, "<table><thead><tr><th>&nbsp;</th><th>&nbsp;</th><th>&nbsp;</th></tr></thead><tbody><tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr><tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr></tbody></table><p><br></p>")],
    ["∑ Formula", "Insert a formula (LaTeX, example: \\frac{a}{b}). Double-click a formula to change it.", () => { const v = window.prompt("Formula (LaTeX), example: \\frac{300+116}{208}"); if (v && v.trim()) { document.execCommand("insertHTML", false, mathHtml(v, false) + " "); renderAll(); } }],
    ["⌫ Clear format", "Remove the formatting of the selected text", () => document.execCommand("removeFormat")],
    ["↶", "Undo (Ctrl+Z)", () => document.execCommand("undo")], ["↷", "Redo (Ctrl+Y)", () => document.execCommand("redo")],
  ];
  BTN.forEach(([t, tip, fn]) => { const b = document.createElement("button"); b.type = "button"; b.textContent = t; b.title = tip;
    b.addEventListener("mousedown", e => e.preventDefault()); b.addEventListener("click", () => { ed.focus(); fn(); schedule(0); }); bar.appendChild(b); });
  bar.appendChild(status);
  const main = document.createElement("div"); main.className = "rt-main";
  const toc = document.createElement("div"); toc.className = "rt-toc"; toc.style.display = "none";
  main.appendChild(ed); main.appendChild(toc);
  wrap.appendChild(bar); wrap.appendChild(main); root.appendChild(wrap);
  let tocOn = true, heads = [];
  function buildToc() {
    heads = Array.from(ed.querySelectorAll("h1,h2,h3,h4")).filter(h => h.textContent.trim());
    toc.innerHTML = "";
    if (!heads.length || !tocOn) { toc.style.display = "none"; return; }
    toc.style.display = "block";
    const t = document.createElement("div"); t.className = "rt-toc-t"; t.textContent = "On this page"; toc.appendChild(t);
    heads.forEach((h, i) => {
      const a = document.createElement("div"); a.className = "rt-toc-i l" + h.tagName.substring(1); a.textContent = h.textContent.trim(); a.title = h.textContent.trim(); a.dataset.i = i;
      a.addEventListener("mousedown", e => e.preventDefault());
      a.addEventListener("click", () => { ed.scrollTo({ top: Math.max(0, h.offsetTop - 14), behavior: "smooth" }); });
      toc.appendChild(a);
    });
    markActive();
  }
  function markActive() {
    if (!heads.length) return; let cur = 0;
    heads.forEach((h, i) => { if (h.offsetTop - 40 <= ed.scrollTop) cur = i; });
    toc.querySelectorAll(".rt-toc-i").forEach(a => a.classList.toggle("on", Number(a.dataset.i) === cur));
  }
  let tocTimer = null;
  function tocSoon() { clearTimeout(tocTimer); tocTimer = setTimeout(buildToc, 300); }
  ed.addEventListener("input", tocSoon); ed.addEventListener("scroll", markActive);
  const tb = document.createElement("button"); tb.type = "button"; tb.textContent = "☰ Outline"; tb.title = "Show or hide the list of headings";
  tb.addEventListener("mousedown", e => e.preventDefault()); tb.addEventListener("click", () => { tocOn = !tocOn; buildToc(); });
  bar.insertBefore(tb, bar.firstChild);

  buildToc();
  let timer = null, last = serialize();
  function out() { return serialize(); }
  function flush() { const v = out(); if (v !== last) { last = v; setStateValue("html", v); status.textContent = "saving…"; setTimeout(() => status.textContent = "✓ saved", 500); } }
  function schedule(ms) { clearTimeout(timer); timer = setTimeout(flush, ms === undefined ? 700 : ms); }
  ed.addEventListener("input", () => schedule());
  ed.addEventListener("blur", () => { clearTimeout(timer); flush(); });
  ed.addEventListener("change", () => schedule(0));
  ed.addEventListener("click", e => { if (e.target.tagName === "INPUT") { if (e.target.checked) e.target.setAttribute("checked", ""); else e.target.removeAttribute("checked"); schedule(0); } });
  ed.addEventListener("dblclick", e => {
    const t = e.target.closest && e.target.closest("span.tex"); if (!t) return;
    const v = window.prompt("Formula (LaTeX):", t.getAttribute("data-tex"));
    if (v !== null && v.trim()) { t.setAttribute("data-tex", v.trim()); renderTex(t); schedule(0); }
  });
  ed.addEventListener("paste", e => {
    const cd = e.clipboardData; if (!cd) return;
    const html = cd.getData("text/html"); const txt = cd.getData("text/plain");
    e.preventDefault();
    let ins = "";
    if (html && /<(h[1-6]|p|li|ul|ol|table|strong|b|em|i|pre|code|blockquote|br)\b/i.test(html)) ins = cleanHTML(html);
    else if (txt && looksMd(txt)) ins = md(txt);
    else ins = plainToHtml(txt || "");
    document.execCommand("insertHTML", false, ins);
    renderAll(); tocSoon();
    schedule(0);
  });
  return () => { clearTimeout(timer); clearTimeout(tocTimer); };
}
"""


def _sync():
    try:
        import cloud_sync
        if cloud_sync.enabled():
            cloud_sync.push()
    except Exception:
        pass


def _meta_save(pid):
    s = st.session_state
    folder = s.get(f"pg_fnew_{pid}", "").strip() or (s.get(f"pg_f_{pid}") if s.get(f"pg_f_{pid}") not in (None, "(no folder)") else "")
    db.page_save(pid, title=(s.get(f"pg_t_{pid}") or "").strip() or "Untitled", folder=folder, tags=s.get(f"pg_g_{pid}", ""),
                 worker_id=s.get(f"pg_w_{pid}", ""), pinned=1 if s.get(f"pg_p_{pid}") else 0)
    s["pg_saved"] = f"✓ Saved {datetime.now():%H:%M:%S}"
    _sync()


def _text_saved(pid):
    v = st.session_state.get(f"pg_rich_{pid}")
    h = (v.get("html") if isinstance(v, dict) else getattr(v, "html", None)) if v is not None else None
    if h is None:
        return
    db.page_save(pid, body=MARK + clean_html(h))
    st.session_state["pg_saved"] = f"✓ Saved {datetime.now():%H:%M:%S}"
    _sync()


def _body_saved_plain(pid):
    db.page_save(pid, body=st.session_state.get(f"pg_plain_{pid}", ""))
    st.session_state["pg_saved"] = f"✓ Saved {datetime.now():%H:%M:%S}"
    _sync()


def _open(pid):
    st.session_state["pg_open"] = pid


def _new(folder=""):
    pid = db.page_new("Untitled", folder)
    st.session_state["pg_open"] = pid
    _sync()


_RICH = None


def _rich():
    global _RICH
    if _RICH is None:
        import streamlit.components.v2 as v2
        import json as _json
        import os as _os
        kp = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), "katex.min.js")
        try:
            with open(kp, encoding="utf-8") as f_:
                kx = "const katex = (() => { try { const module = { exports: {} }; const exports = module.exports;\n" + f_.read() + "\n; return module.exports; } catch (e) { return null; } })();\n"
        except OSError:
            kx = "const katex = null;\n"
        js = kx + RICH_JS.replace('root.innerHTML = "";', 'root.innerHTML = ""; { const _s = document.createElement("style"); _s.textContent = ' + _json.dumps(RICH_CSS) + '; root.appendChild(_s); }', 1)
        _RICH = v2.component("hr_rich_editor", js=js, isolate_styles=False)
    return _RICH


def render():
    st.title("📓 Notebook")
    st.caption("Pages like Notion. You can **copy from the chat and paste here**: headings, bullets, bold and tables keep their look. "
               "Everything saves by itself.")
    left, right = st.columns([1, 2.8], gap="large")
    with left:
        q = st.text_input("🔎 Search", key="pg_q", placeholder="word, tag, worker ID")
        folders = db.page_folders()
        fsel = st.selectbox("Folder", ["All"] + folders, key="pg_fsel")
        st.button("➕ New page", key="pg_new", type="primary", width="stretch", on_click=_new, args=("" if fsel == "All" else fsel,))
        rows = db.page_list(q.strip(), "" if fsel == "All" else fsel)
        if not rows:
            st.caption("No pages yet. Press **New page**.")
        cur_folder = None
        for r in rows:
            if (r["folder"] or "") != cur_folder:
                cur_folder = r["folder"] or ""
                st.markdown(f"**📁 {cur_folder or 'No folder'}**")
            lab = ("📌 " if r["pinned"] else "") + (r["title"] or "Untitled") + (f"  · {r['worker_id']}" if r["worker_id"] else "")
            st.button(lab[:48], key=f"pg_b_open_{r['id']}", on_click=_open, args=(r["id"],), width="stretch",
                      type="primary" if st.session_state.get("pg_open") == r["id"] else "secondary")
    with right:
        pid = st.session_state.get("pg_open")
        pg = db.page_get(pid) if pid else None
        if not pg:
            st.info("Pick a page on the left, or press **New page**.")
            return
        s = st.session_state
        for k, v in (("t", pg["title"]), ("g", pg["tags"]), ("w", pg["worker_id"]), ("p", bool(pg["pinned"]))):
            s.setdefault(f"pg_{k}_{pid}", v)
        st.text_input("Title", key=f"pg_t_{pid}", on_change=_meta_save, args=(pid,), label_visibility="collapsed", placeholder="Title")
        c = st.columns([1.3, 1.3, 1.3, 1, 0.8])
        opts = ["(no folder)"] + db.page_folders()
        s.setdefault(f"pg_f_{pid}", pg["folder"] if pg["folder"] in opts else "(no folder)")
        c[0].selectbox("Folder", opts, key=f"pg_f_{pid}", on_change=_meta_save, args=(pid,))
        c[1].text_input("…or a new folder", key=f"pg_fnew_{pid}", on_change=_meta_save, args=(pid,), placeholder="Meetings")
        c[2].text_input("Tags (comma)", key=f"pg_g_{pid}", on_change=_meta_save, args=(pid,), placeholder="leader, safety")
        c[3].text_input("Worker ID", key=f"pg_w_{pid}", on_change=_meta_save, args=(pid,), placeholder="DMO-10078")
        c[4].checkbox("📌 Pin", key=f"pg_p_{pid}", on_change=_meta_save, args=(pid,))
        try:
            init_key = f"pg_init_{pid}"
            if init_key not in s:                                   # the text as it was when you opened the page (never changes while you type)
                s[init_key] = body_to_html(pg["body"])
            _rich()(key=f"pg_rich_{pid}", data={"id": f"page{pid}", "html": s[init_key]}, default={"html": s[init_key]},
                    on_html_change=lambda p=pid: _text_saved(p))
        except Exception as e:
            st.warning(f"The rich editor is not available ({type(e).__name__}). Using a plain box.")
            s.setdefault(f"pg_plain_{pid}", pg["body"])
            st.text_area("Text", key=f"pg_plain_{pid}", on_change=_body_saved_plain, args=(pid,), height=420, label_visibility="collapsed")
        st.caption(s.pop("pg_saved", None) or f"Last saved {pg['updated_at'][:16].replace('T', ' ')}")
        st.markdown("**📷 Photos and files**")
        up = st.file_uploader("Add photos", type=list(IMG) + ["pdf"], accept_multiple_files=True, key=f"pg_up_{pid}_{s.get('pg_up_ver', 0)}", label_visibility="collapsed")
        if up:
            for f in up:
                db.page_file_add(pid, f.name, f.getvalue())
            s["pg_up_ver"] = s.get("pg_up_ver", 0) + 1
            _sync()
            st.rerun()
        for fl in db.page_files(pid):
            cc = st.columns([6, 1])
            if fl["name"].lower().rsplit(".", 1)[-1] in IMG:
                cc[0].image(fl["path"], caption=fl["name"], width=360)
            else:
                cc[0].write("📎 " + fl["name"])
            if cc[1].button("🗑️", key=f"pg_fd_{fl['id']}"):
                db.page_file_delete(fl["id"])
                _sync()
                st.rerun()
        with st.expander("Delete this page"):
            if st.checkbox("Yes, delete this page and its photos", key=f"pg_del_ok_{pid}") and st.button("🗑️ Delete the page", key=f"pg_del_{pid}"):
                db.page_delete(pid)
                s.pop("pg_open", None)
                _sync()
                st.rerun()
