#!/usr/bin/env python3
"""Export OneNote for the web pages to PDF through the application's File > Print.

The tool drives a visible Chromium (Playwright) on an existing X11 display,
lets the user sign in manually, walks the notebook / section-group / section /
page (including subpage) hierarchy, and saves one PDF per page.

Printing deliberately goes through OneNote's own File > Print command.  The
browser runs with ``--kiosk-printing`` and "Save as PDF" as its default
destination, so the application's print job is written directly into a staging
directory and then moved to the output tree.  OneNote itself decides pagination
and page layout; this tool does not print the editor UI or re-render the notes.
"""

from __future__ import annotations

import argparse
import asyncio
import hashlib
import json
import os
import re
import sys
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Optional
from urllib.parse import parse_qs, urlparse
from urllib.parse import urlsplit

try:
    from playwright.async_api import BrowserContext, Frame, Page, async_playwright
except ImportError as exc:  # pragma: no cover - environment guard
    print(
        "缺少 playwright。请先运行 `uv sync`，再运行 "
        "`uv run playwright install chromium`。",
        file=sys.stderr,
    )
    raise SystemExit(2) from exc


NOTEBOOKS_URL = "https://www.onenote.com/notebooks"
DEFAULT_PROFILE = ".profile"
DEFAULT_OUTPUT = "output"
DEFAULT_PRINT_DIR = ".print-staging"
DEFAULT_PAPER = "A0"

# Chrome print-preview media sizes are in microns (portrait orientation).
PAPER_SIZES: dict[str, tuple[int, int]] = {
    "A0": (841_000, 1_189_000),
    "A1": (594_000, 841_000),
    "A2": (420_000, 594_000),
    "A3": (297_000, 420_000),
    "A4": (210_000, 297_000),
    "Letter": (215_900, 279_400),
    "Legal": (215_900, 355_600),
}

SIGN_IN_NAMES = re.compile(
    r"^\s*(Sign in|Sign in to OneNote|登录|登入)(\s+to\s+OneNote)?\s*$", re.I
)
APP_ENTRY_NAMES = re.compile(
    r"^\s*(Open OneNote|Go to OneNote|Open notebook|My notebooks|Notebooks|"
    r"打开 OneNote|转到 OneNote|打开笔记本|我的笔记本|笔记本)\s*$",
    re.I,
)

# ---------------------------------------------------------------------------
# Browser-side scripts
# ---------------------------------------------------------------------------

NAV_PROBE_JS = r"""
() => {
  const q = (sel) => {
    try { return document.querySelectorAll(sel).length; } catch (e) { return 0; }
  };
  return {
    sectionList: q('#NavPaneSectionList, .sectionList, [class*="sectionList"]'),
    sectionItems: q('#NavPaneSectionList div[class*="sectionListItem"], ' +
                    '#NavPaneSectionList div[class*="sectionGroup__groupItemWrap"]'),
    pageNodes: q('.pageNode, [class*="pageNode"]'),
    pageList: q('#PageList, [class*="pagesContainer"], [class*="pageListContainer"]'),
    canvas: q('#WACViewPanel, #PageContentContainer, .PageContentContainer, #OreoCanvas, .canvasContainer, [class*="canvasContainer"]'),
    treeitems: q('[role="treeitem"]'),
    fileTab: q('#FileTab, [id*="FileTab"]'),
  };
}
"""

LIST_ROWS_JS = r"""
() => {
  const out = [];
  const seen = new Set();
  for (const row of Array.from(document.querySelectorAll('div[role="row"], [role="listitem"]'))) {
    if (row.getAttribute('data-automationid') === 'row-header') continue;
    const cell = row.querySelector('[data-automationid="field-name"]');
    if (!cell) continue;
    const btn = cell.querySelector('button[role="link"], a[role="link"], a[href]');
    if (!btn) continue;
    const name = (btn.innerText || btn.textContent || '').trim();
    if (!name) continue;
    const id = row.getAttribute('data-automationid') || '';
    const key = id || name;
    if (seen.has(key)) continue;
    seen.add(key);
    out.push({
      id,
      name,
      pos: Math.round(row.getBoundingClientRect().top + (window.scrollY || 0)),
    });
  }
  return out;
}
"""

# ---------------------------------------------------------------------------
# DOM collection
# ---------------------------------------------------------------------------

COLLECT_SECTIONS_JS = r"""
(rootId) => {
  const findEl = (id) => {
    let el = document.getElementById(id);
    if (el) return el;
    for (const attr of ['data-page-id', 'data-section-id', 'data-id', 'data-automationid']) {
      for (const cand of Array.from(document.querySelectorAll('[' + attr + ']'))) {
        if (cand.getAttribute(attr) === id) return cand;
      }
    }
    return null;
  };
  const ITEM_SEL = 'div[class*="sectionListItem"], div[class*="sectionGroup__groupItemWrap"]';
  const inPageList = (el) => !!(el.closest && el.closest(
    '#PageList, [class*="pagesContainer"], [class*="pageListContainer"]'));
  const labelOf = (el) => (el && el.getAttribute && el.getAttribute('aria-label')) || '';
  const classOf = (el) => {
    if (!el) return '';
    const c = el.className;
    return String(c && c.baseVal !== undefined ? c.baseVal : (c || ''));
  };
  const isGroup = (el) => {
    const label = labelOf(el);
    return /sectionGroup/i.test(classOf(el)) ||
      /[,，]\s*Section\s+Group\b/i.test(label) ||
      /[,，]\s*分区组/.test(label);
  };

  let root = null;
  if (rootId) {
    const parent = findEl(rootId);
    if (!parent) return { items: [], missing: true };
    root = parent.querySelector('[role="group"]');
    if (!root) {
      let cur = parent;
      for (let i = 0; i < 3 && cur; i++) {
        const sib = cur.parentElement
          ? cur.parentElement.querySelector(':scope > [role="group"]')
          : null;
        if (sib) { root = sib; break; }
        const next = cur.nextElementSibling;
        if (next && next.getAttribute('role') === 'group') { root = next; break; }
        cur = cur.parentElement;
      }
    }
    if (!root) {
      const gc = parent.closest(
        '[class*="sectionGroupContainer"], [class*="sectionGroup"]');
      root = gc ? gc.querySelector('[role="group"]') : null;
    }
    if (!root) return { items: [], missing: true, noGroup: true };
  } else {
    root = document.querySelector('#NavPaneSectionList') ||
      document.querySelector('.sectionList') ||
      document.querySelector('[class*="sectionList"]') ||
      document.body;
  }

  const candidates = new Set(Array.from(root.querySelectorAll(ITEM_SEL)));
  for (const el of Array.from(root.querySelectorAll('[role="treeitem"]'))) {
    if (inPageList(el)) continue;
    if (el.closest(ITEM_SEL)) continue;
    if (el.closest('[class*="Breadcrumb"], [class*="breadcrumb"]')) continue;
    const label = labelOf(el);
    if (/[,，]\s*Section(\s+Group)?\b/i.test(label) ||
        /[,，]\s*分区组?/.test(label) ||
        /sectionListItem|sectionGroup/i.test(classOf(el))) {
      candidates.add(el);
    }
  }

  let scroller = null;
  {
    let p = root;
    while (p && p !== document.body) {
      const cs = getComputedStyle(p);
      if (/(auto|scroll)/.test(cs.overflowY) && p.scrollHeight > p.clientHeight + 4) {
        scroller = p;
        break;
      }
      p = p.parentElement;
    }
    if (!scroller) scroller = document.scrollingElement || document.documentElement;
  }
  const srect = scroller.getBoundingClientRect();

  const out = [];
  const seen = new Set();
  let missingId = 0;
  for (const node of Array.from(candidates)) {
    if (node.id && seen.has(node.id)) continue;
    if (node.offsetWidth === 0 && node.offsetHeight === 0) continue;
    let p = node.parentElement;
    let direct = true;
    while (p && p !== root) {
      if (p.getAttribute && p.getAttribute('role') === 'group') { direct = false; break; }
      p = p.parentElement;
    }
    if (!direct) continue;
    if (!node.id) { missingId += 1; continue; }
    seen.add(node.id);
    if (!rootId && (node.closest('[class*="Breadcrumb"]') ||
                    node.closest('[class*="breadcrumb"]'))) continue;
    seen.add(node.id);

    const nav = node.querySelector('.navItem, [class*="navItem"]');
    let label = labelOf(nav) || labelOf(node);
    if (!label) label = ((nav || node).innerText || '').trim().split('\n')[0];
    label = label
      .replace(/[,，]\s*Section\s+Group.*$/i, '')
      .replace(/[,，]\s*Section.*$/i, '')
      .replace(/[,，]\s*分区组.*$/, '')
      .replace(/[,，]\s*分区.*$/, '')
      .replace(/\s*[,，]\s*(Collapsed|Expanded|Selected|已折叠|已展开|已选择)\s*$/i, '')
      .replace(/\s+/g, ' ')
      .trim();

    const meta = nav || node;
    const rect = node.getBoundingClientRect();
    out.push({
      id: node.id,
      name: label || 'Untitled',
      kind: isGroup(node) ? 'group' : 'section',
      pos: Math.round(rect.top - srect.top + scroller.scrollTop),
      setsize: parseInt(meta.getAttribute('aria-setsize') || '0', 10) || 0,
      posinset: parseInt(meta.getAttribute('aria-posinset') || '0', 10) || 0,
    });
  }
  out.sort((a, b) => a.pos - b.pos);
  return { items: out, missing: false, missingId };
}
"""

COLLECT_PAGES_JS = r"""
() => {
  const LIST_SEL = '#PageList, [class*="pagesContainer"], ' +
    '[class*="pageListContainer"], [class*="pageList__"]';
  const listRoot = document.querySelector(LIST_SEL);
  const nodes = Array.from(document.querySelectorAll('.pageNode, [class*="pageNode"]'));
  if (listRoot) {
    for (const el of Array.from(listRoot.querySelectorAll('[role="treeitem"]'))) {
      if (!el.closest('.pageNode, [class*="pageNode"]')) nodes.push(el);
    }
  }

  let scroller = null;
  {
    let p = (nodes[0] && nodes[0].parentElement) || listRoot;
    while (p && p !== document.body) {
      const cs = getComputedStyle(p);
      if (/(auto|scroll)/.test(cs.overflowY) && p.scrollHeight > p.clientHeight + 4) {
        scroller = p;
        break;
      }
      p = p.parentElement;
    }
    if (!scroller) scroller = document.scrollingElement || document.documentElement;
  }
  const srect = scroller.getBoundingClientRect();

  const readLevel = (node, li, nav) => {
    for (const el of [node, li, nav]) {
      if (!el) continue;
      for (const attr of ['aria-level', 'data-level', 'data-depth', 'data-indent']) {
        const v = el.getAttribute(attr);
        if (v && /^\d+$/.test(v)) return { level: parseInt(v, 10), source: 'aria' };
      }
      const c = el.className;
      const cls = String(c && c.baseVal !== undefined ? c.baseVal : (c || ''));
      const m = cls.match(/(?:level|indent|depth)[-_]?(\d+)/i);
      if (m) return { level: parseInt(m[1], 10), source: 'class' };
    }
    let indent = 0;
    const wrap = node.querySelector('[class*="itemWrap"]');
    for (const el of [li, nav, wrap]) {
      if (!el) continue;
      const cs = getComputedStyle(el);
      indent = Math.max(indent,
        (parseFloat(cs.paddingLeft) || 0) + (parseFloat(cs.marginLeft) || 0));
    }
    return { level: -1, source: 'indent', indent };
  };

  const out = [];
  const seen = new Set();
  let missingId = 0;
  for (const node of nodes) {
    if (node.offsetWidth === 0 && node.offsetHeight === 0) continue;
    const li = node.querySelector('.pageListItem, [class*="pageListItem"]');
    const nav = node.querySelector('.navItem, [class*="navItem"]');
    if (!li && !nav && !listRoot) continue;
    const pid = node.id || node.getAttribute('data-page-id') ||
      node.getAttribute('data-id') || '';
    if (!pid) { missingId += 1; continue; }
    if (seen.has(pid)) continue;
    seen.add(pid);

    let label = (nav && nav.getAttribute('aria-label')) ||
      node.getAttribute('aria-label') || '';
    if (!label) label = ((nav || node).innerText || '').trim().split('\n')[0];
    label = label
      .replace(/[,，]\s*page\s+\d+\s+of\s+\d+\s*/gi, '')
      .replace(/[,，]\s*第\s*\d+\s*页[，,]\s*共\s*\d+\s*页\s*/g, '')
      .replace(/[,，]\s*Page\.?\s*(Selected|Select|Current).*$/i, '')
      .replace(/[,，]\s*(页面|页)[.。]?\s*(已选择|选择|当前).*$/i, '')
      .replace(/\s*[,，]\s*(Selected|Collapsed|Expanded|已选择|已折叠|已展开)\s*$/i, '')
      .replace(/\s+/g, ' ')
      .trim();

    const lv = readLevel(node, li, nav);
    let domDepth = 0;
    let q = node.parentElement;
    while (q && q !== document.body) {
      if (q.matches && q.matches('.pageNode')) domDepth += 1;
      q = q.parentElement;
    }
    const meta = nav || node;
    const rect = node.getBoundingClientRect();
    out.push({
      id: pid,
      title: label || 'Untitled Page',
      level: lv.level,
      levelSource: lv.source,
      indent: lv.indent || 0,
      domDepth,
      pos: Math.round(rect.top - srect.top + scroller.scrollTop),
      setsize: parseInt(meta.getAttribute('aria-setsize') || '0', 10) || 0,
      posinset: parseInt(meta.getAttribute('aria-posinset') || '0', 10) || 0,
    });
  }
  out.sort((a, b) => a.pos - b.pos);
  return { items: out, missingId };
}
"""

SCROLL_TO_TOP_JS = r"""
({ itemSelector, rootId, listSelector }) => {
  let scope = rootId ? document.getElementById(rootId) : null;
  if (!scope && listSelector) scope = document.querySelector(listSelector);
  let item = scope ? scope.querySelector(itemSelector) : null;
  if (!item) item = document.querySelector(itemSelector);
  if (!item) return { ok: false };
  let sc = null;
  let p = item.parentElement;
  while (p) {
    const cs = getComputedStyle(p);
    if (/(auto|scroll)/.test(cs.overflowY) && p.scrollHeight > p.clientHeight + 4) {
      sc = p;
      break;
    }
    p = p.parentElement;
  }
  if (!sc) sc = document.scrollingElement || document.documentElement;
  sc.scrollTop = 0;
  return { ok: true, atBottom: sc.scrollHeight <= sc.clientHeight + 2 };
}
"""

SCROLL_STEP_JS = r"""
({ itemSelector, rootId, listSelector }) => {
  let scope = rootId ? document.getElementById(rootId) : null;
  if (!scope && listSelector) scope = document.querySelector(listSelector);
  let item = scope ? scope.querySelector(itemSelector) : null;
  if (!item) item = document.querySelector(itemSelector);
  if (!item) return { ok: false };
  let sc = null;
  let p = item.parentElement;
  while (p) {
    const cs = getComputedStyle(p);
    if (/(auto|scroll)/.test(cs.overflowY) && p.scrollHeight > p.clientHeight + 4) {
      sc = p;
      break;
    }
    p = p.parentElement;
  }
  if (!sc) sc = document.scrollingElement || document.documentElement;
  const before = sc.scrollTop;
  const step = Math.max(120, Math.floor(sc.clientHeight * 0.8));
  sc.scrollTop = Math.min(sc.scrollHeight, before + step);
  const after = sc.scrollTop;
  return {
    ok: true,
    atBottom: after + sc.clientHeight >= sc.scrollHeight - 2,
    moved: after !== before,
  };
}
"""

SCROLL_INTO_VIEW_JS = r"""
(id) => {
  const findEl = (id) => {
    let el = document.getElementById(id);
    if (el) return el;
    for (const attr of ['data-page-id', 'data-section-id', 'data-id', 'data-automationid']) {
      for (const cand of Array.from(document.querySelectorAll('[' + attr + ']'))) {
        if (cand.getAttribute(attr) === id) return cand;
      }
    }
    return null;
  };
  const el = findEl(id);
  if (!el) return false;
  el.scrollIntoView({ block: 'center', inline: 'nearest' });
  return true;
}
"""

CLICK_ITEM_JS = r"""
(id) => {
  const findEl = (id) => {
    let el = document.getElementById(id);
    if (el) return el;
    for (const attr of ['data-page-id', 'data-section-id', 'data-id', 'data-automationid']) {
      for (const cand of Array.from(document.querySelectorAll('[' + attr + ']'))) {
        if (cand.getAttribute(attr) === id) return cand;
      }
    }
    return null;
  };
  const el = findEl(id);
  if (!el) return { ok: false, reason: 'not-found' };
  el.scrollIntoView({ block: 'center', inline: 'nearest' });
  const target = el.querySelector('.navItem, [class*="navItem"]') || el;
  target.click();
  return { ok: true };
}
"""

ELEMENT_PRESENT_JS = r"""
(id) => {
  const findEl = (id) => {
    let el = document.getElementById(id);
    if (el) return el;
    for (const attr of ['data-page-id', 'data-section-id', 'data-id', 'data-automationid']) {
      for (const cand of Array.from(document.querySelectorAll('[' + attr + ']'))) {
        if (cand.getAttribute(attr) === id) return cand;
      }
    }
    return null;
  };
  const el = findEl(id);
  if (!el) return false;
  const r = el.getBoundingClientRect();
  return r.width > 0 && r.height > 0;
}
"""

NESTED_GROUP_JS = r"""
(id) => {
  const findEl = (id) => {
    let el = document.getElementById(id);
    if (el) return el;
    for (const attr of ['data-page-id', 'data-section-id', 'data-id', 'data-automationid']) {
      for (const cand of Array.from(document.querySelectorAll('[' + attr + ']'))) {
        if (cand.getAttribute(attr) === id) return cand;
      }
    }
    return null;
  };
  const parent = findEl(id);
  if (!parent) return false;
  const group = parent.querySelector('[role="group"]');
  if (!group) return false;
  return !!group.querySelector(
    'div[class*="sectionListItem"], div[class*="sectionGroup__groupItemWrap"], ' +
    '[role="treeitem"]');
}
"""

GROUP_EXPANDED_JS = r"""
(id) => {
  const findEl = (id) => {
    let el = document.getElementById(id);
    if (el) return el;
    for (const attr of ['data-page-id', 'data-section-id', 'data-id', 'data-automationid']) {
      for (const cand of Array.from(document.querySelectorAll('[' + attr + ']'))) {
        if (cand.getAttribute(attr) === id) return cand;
      }
    }
    return null;
  };
  const parent = findEl(id);
  if (!parent) return true;
  const nav = parent.querySelector('.navItem, [class*="navItem"]') || parent;
  const v = parent.getAttribute('aria-expanded') || nav.getAttribute('aria-expanded');
  if (v === null) return true;
  return String(v).toLowerCase() === 'true';
}
"""

SELECTION_SUPPORTED_JS = r"""
() => {
  const nodes = Array.from(document.querySelectorAll('.pageNode, [class*="pageNode"]'));
  for (const n of nodes) {
    if (n.hasAttribute('aria-selected') || n.hasAttribute('aria-current')) return true;
    const li = n.querySelector('.pageListItem, [class*="pageListItem"]');
    const nav = n.querySelector('.navItem, [class*="navItem"]');
    if (li && (li.hasAttribute('aria-selected') || li.hasAttribute('aria-current'))) return true;
    if (nav && (nav.hasAttribute('aria-selected') || nav.hasAttribute('aria-current'))) return true;
    if (/(selected|active)/i.test(String(n.className || ''))) return true;
    if (li && /(selected|active)/i.test(String(li.className || ''))) return true;
  }
  return false;
}
"""

IS_SELECTED_JS = r"""
(id) => {
  const findEl = (id) => {
    let el = document.getElementById(id);
    if (el) return el;
    for (const attr of ['data-page-id', 'data-section-id', 'data-id', 'data-automationid']) {
      for (const cand of Array.from(document.querySelectorAll('[' + attr + ']'))) {
        if (cand.getAttribute(attr) === id) return cand;
      }
    }
    return null;
  };
  const node = findEl(id);
  if (!node) return { found: false, selected: false };
  const li = node.querySelector('.pageListItem, [class*="pageListItem"]');
  const nav = node.querySelector('.navItem, [class*="navItem"]');
  for (const el of [node, li, nav]) {
    if (!el) continue;
    if (el.getAttribute('aria-selected') === 'true' ||
        el.getAttribute('aria-current') === 'true') {
      return { found: true, selected: true };
    }
    const cls = String(el.className || '');
    if (/(selected|active|isSelected)/i.test(cls)) {
      return { found: true, selected: true };
    }
  }
  return { found: true, selected: false };
}
"""

PAGE_TITLE_JS = r"""
() => {
  const selectors = [
    '.TitleOutline',
    '#PageContentContainer .TitleOutline',
    '.pageTitle',
    '#pageTitle',
    '[data-automationid*="pageTitle"]',
    '[aria-label*="Page Title"]',
    '[aria-label*="页面标题"]',
    '#WACViewPanel [contenteditable="true"]',
    '[contenteditable="true"][role="textbox"]',
  ];
  for (const sel of selectors) {
    for (const el of Array.from(document.querySelectorAll(sel))) {
      const text = (el.innerText || el.textContent || el.value || '')
        .trim().split('\n')[0].trim();
      if (text) return { found: true, title: text };
    }
  }
  return { found: false, title: '' };
}
"""

LOCKED_SECTION_JS = r"""
() => {
  const texts = [
    'Section Password Protected',
    'This section is password protected',
    '分区受密码保护',
    '此分区受密码保护',
    '此分区已受密码保护',
    '需要密码',
  ];
  const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
  let node;
  while ((node = walker.nextNode())) {
    const content = node.textContent || '';
    if (!texts.some((t) => content.includes(t))) continue;
    const parent = node.parentElement;
    if (!parent) continue;
    const style = window.getComputedStyle(parent);
    if (style.display === 'none' || style.visibility === 'hidden' ||
        parseFloat(style.opacity) === 0) continue;
    if (parent.closest('.sectionList') ||
        parent.closest('.pagesContainer') ||
        parent.closest('#NavPaneSectionList') ||
        parent.closest('#PageList') ||
        parent.closest('[class*="navItem"]') ||
        parent.closest('[class*="pageListItem"]')) continue;
    return true;
  }
  return false;
}
"""

INSPECT_JS = r"""
() => {
  const count = (sel) => {
    try { return document.querySelectorAll(sel).length; } catch (e) { return -1; }
  };
  const roles = {};
  for (const el of Array.from(document.querySelectorAll('[role]'))) {
    const r = el.getAttribute('role') || '';
    roles[r] = (roles[r] || 0) + 1;
  }
  const levels = {};
  for (const el of Array.from(document.querySelectorAll('[aria-level]'))) {
    const v = el.getAttribute('aria-level') || '';
    levels[v] = (levels[v] || 0) + 1;
  }
  const labels = [];
  for (const el of Array.from(document.querySelectorAll(
      '.navItem, [role="treeitem"], .sectionListItem, .pageListItem'))) {
    const label = (el.getAttribute('aria-label') || el.innerText || '')
      .trim().replace(/\s+/g, ' ').slice(0, 120);
    if (label && !labels.includes(label)) labels.push(label);
    if (labels.length >= 12) break;
  }
  const iframes = Array.from(document.querySelectorAll('iframe')).map((f) => ({
    id: f.id || '',
    name: f.name || '',
    src: (f.getAttribute('src') || '').slice(0, 160),
  }));
  return {
    url: location.href,
    title: document.title || '',
    counts: {
      sectionList: count('.sectionList, [class*="sectionList"]'),
      sectionItems: count('div[class*="sectionListItem"], div[class*="sectionGroup__groupItemWrap"]'),
      pageNodes: count('.pageNode, [class*="pageNode"]'),
      pageList: count('#PageList, [class*="pagesContainer"], [class*="pageListContainer"]'),
      canvas: count('#WACViewPanel, #PageContentContainer, .PageContentContainer, #OreoCanvas, .canvasContainer, [class*="canvasContainer"]'),
      treeitems: count('[role="treeitem"]'),
      fileTab: count('#FileTab, [id*="FileTab"]'),
      printNamed: count('[aria-label*="Print"], [aria-label*="打印"]'),
      iframes: count('iframe'),
    },
    roles,
    levels,
    labels,
    iframes,
  };
}
"""


# ---------------------------------------------------------------------------
# Errors and small utilities
# ---------------------------------------------------------------------------


class ExportError(RuntimeError):
    """A recoverable, user-visible export failure."""


class StructureError(ExportError):
    """The OneNote DOM did not expose a structure we can trust."""


class _StopExport(Exception):
    """Internal: --max-pages reached."""


class _ImagePrompt(Exception):
    """Internal: OneNote asked whether to wait for images."""


def log(message: str) -> None:
    print(message, file=sys.stderr, flush=True)


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


_INVALID_FS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def slugify(text: str, max_len: int = 60) -> str:
    text = unicodedata.normalize("NFKC", text or "").strip()
    text = _INVALID_FS.sub("_", text)
    text = re.sub(r"\s+", " ", text).strip(" .")
    if not text:
        text = "Untitled"
    if len(text) > max_len:
        text = text[:max_len].rstrip(" .")
    return text or "Untitled"


def notebook_token(notebook: "NotebookRef", page_url: str) -> str:
    key = notebook.id
    if not key:
        for url in (notebook.url, page_url):
            try:
                key = parse_qs(urlparse(url).query).get("sourcedoc", [""])[0]
            except Exception:
                key = ""
            if key:
                break
    if not key:
        key = notebook.name or page_url
    return hashlib.sha1(key.encode("utf-8")).hexdigest()[:8]


def short_token(value: str, length: int = 8) -> str:
    raw = re.sub(r"[^0-9A-Za-z]", "", value or "")
    if len(raw) >= length:
        return raw[:length]
    return hashlib.sha1((value or "").encode("utf-8")).hexdigest()[:length]


def atomic_write_bytes(path: Path, data: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(f".{path.name}.{os.getpid()}.part")
    try:
        with open(tmp, "wb") as fh:
            fh.write(data)
            fh.flush()
            os.fsync(fh.fileno())
        os.replace(tmp, path)
    finally:
        try:
            tmp.unlink(missing_ok=True)
        except OSError:
            pass


def write_index(out_root: Path, manifest: "Manifest") -> None:
    pages = [
        rec for rec in manifest.data.get("pages", {}).values()
        if rec.get("status") == "ok"
    ]
    lines = ["# OneNote PDF 导出索引", ""]
    current = None
    for rec in sorted(pages, key=lambda r: (r.get("notebook", ""), r.get("trail", []))):
        trail = rec.get("trail") or []
        group = " / ".join(trail[:-1]) or rec.get("notebook", "")
        if group != current:
            current = group
            lines.extend([f"## {group}", ""])
        title = rec.get("title") or "Untitled"
        path = rec.get("path") or ""
        size = int(rec.get("bytes") or 0) / (1024 * 1024)
        lines.append(f"- [{title}](<{path}>) — {size:.1f} MB")
    lines.append("")
    (out_root / "INDEX.md").write_text("\n".join(lines), encoding="utf-8")


def atomic_write_json(path: Path, payload: dict[str, Any]) -> None:
    raw = json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")
    atomic_write_bytes(path, raw)


def normalize_title(text: str) -> str:
    text = unicodedata.normalize("NFKC", text or "")
    return re.sub(r"\s+", " ", text).strip().casefold()


def titles_match(actual: str, expected: str) -> bool:
    a = normalize_title(actual)
    e = normalize_title(expected)
    if a == e:
        return True
    if not a and e in {"untitled page", "untitled", "无标题页", "无标题"}:
        return True
    if not a or not e:
        return False
    return len(a) > 2 and (a in e or e in a)


def is_pdf_bytes(data: bytes) -> bool:
    return bool(data) and data[:5] == b"%PDF-"


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------


@dataclass
class NotebookRef:
    name: str
    id: str = ""
    url: str = ""
    source: str = "listing"  # listing | url


@dataclass
class NavItem:
    id: str
    name: str
    kind: str  # group | section
    pos: int = 0
    setsize: int = 0
    posinset: int = 0


@dataclass
class PageNode:
    id: str
    title: str
    level: int = 0
    indent: float = 0.0
    dom_depth: int = 0
    pos: int = 0
    setsize: int = 0
    posinset: int = 0
    parent: Optional["PageNode"] = None
    children: list["PageNode"] = field(default_factory=list)
    order: int = 0
    directory: Optional[Path] = None
    pdf_path: Optional[Path] = None

    def walk(self):
        yield self
        for child in self.children:
            yield from child.walk()


class Manifest:
    """Per-page resume state, keyed by OneNote page id."""

    def __init__(self, path: Path, output_root: Path):
        self.path = path
        self.output_root = output_root
        self.data: dict[str, Any] = {"version": 1, "pages": {}, "errors": []}
        if path.exists():
            try:
                loaded = json.loads(path.read_text(encoding="utf-8"))
            except (OSError, ValueError) as exc:
                raise ExportError(f"manifest 无法读取: {path}: {exc}") from exc
            if not isinstance(loaded, dict):
                raise ExportError(f"manifest 格式错误: {path}")
            self.data.update(loaded)
            self.data.setdefault("pages", {})
            self.data.setdefault("errors", [])
        self.save()

    def save(self) -> None:
        atomic_write_json(self.path, self.data)

    def page(self, page_id: str) -> Optional[dict[str, Any]]:
        entry = self.data["pages"].get(page_id)
        return entry if isinstance(entry, dict) else None

    def is_done(self, page_id: str, rel_path: str) -> bool:
        entry = self.page(page_id)
        if not entry or entry.get("status") != "ok":
            return False
        if entry.get("path") != rel_path:
            return False
        target = self.output_root / rel_path
        try:
            if not target.is_file() or target.stat().st_size < 6:
                return False
            with open(target, "rb") as fh:
                return is_pdf_bytes(fh.read(5))
        except OSError:
            return False

    def record_ok(
        self,
        node: PageNode,
        rel_path: str,
        size: int,
        notebook: str,
        trail: list[str],
    ) -> None:
        self.data["pages"][node.id] = {
            "status": "ok",
            "title": node.title,
            "level": node.level,
            "parent_id": node.parent.id if node.parent else None,
            "order": node.order,
            "path": rel_path,
            "bytes": size,
            "notebook": notebook,
            "trail": trail,
            "updated_at": utc_now(),
        }
        scope = "/".join(trail)
        self.data["errors"] = [
            item for item in self.data.get("errors", [])
            if item.get("scope") != scope
        ]
        self.save()

    def record_failed(
        self,
        node: PageNode,
        rel_path: str,
        error: str,
        notebook: str,
        trail: list[str],
    ) -> None:
        self.data["pages"][node.id] = {
            "status": "failed",
            "title": node.title,
            "level": node.level,
            "parent_id": node.parent.id if node.parent else None,
            "order": node.order,
            "path": rel_path,
            "notebook": notebook,
            "trail": trail,
            "error": error,
            "updated_at": utc_now(),
        }
        self.save()

    def record_error(self, scope: str, message: str) -> None:
        self.data["errors"].append(
            {"scope": scope, "message": message, "at": utc_now()}
        )
        self.save()


@dataclass
class ExportContext:
    output_root: Path
    manifest: Manifest
    notebook: str
    args: argparse.Namespace
    exported: int = 0
    skipped: int = 0
    failures: int = 0

    def error(self, scope: str, message: str) -> None:
        self.failures += 1
        log(f"[失败] {scope}: {message}")
        self.manifest.record_error(scope, message)


# ---------------------------------------------------------------------------
# Session / browser plumbing
# ---------------------------------------------------------------------------


class Session:
    def __init__(self, args: argparse.Namespace):
        self.args = args
        self._pw = None
        self.browser = None
        self.context: Optional[BrowserContext] = None
        self.page: Optional[Page] = None
        self._owned = False

    async def __aenter__(self) -> "Session":
        self._pw = await async_playwright().start()
        if getattr(self.args, "cdp_url", None):
            self.browser = await self._pw.chromium.connect_over_cdp(
                self.args.cdp_url
            )
            if not self.browser.contexts:
                raise ExportError(
                    "CDP 浏览器没有可用的 BrowserContext；请确认 Chrome 已启动。"
                )
            self.context = self.browser.contexts[0]
            self._owned = False
        else:
            display = os.environ.get("DISPLAY", "")
            if not display:
                os.environ["DISPLAY"] = ":0"
                display = ":0"
            profile = Path(self.args.profile).expanduser().resolve()
            profile.mkdir(parents=True, exist_ok=True)
            print_dir = Path(self.args.print_dir).expanduser().resolve()
            print_dir.mkdir(parents=True, exist_ok=True)
            self._configure_print_preferences(
                profile,
                print_dir,
                paper=str(getattr(self.args, "paper", DEFAULT_PAPER)),
                landscape=not bool(getattr(self.args, "portrait", False)),
            )
            launch_args = ["--start-maximized", "--kiosk-printing"]
            if display.startswith(":"):
                launch_args.append("--ozone-platform=x11")
            self.context = await self._pw.chromium.launch_persistent_context(
                str(profile),
                headless=False,
                no_viewport=True,
                accept_downloads=True,
                args=launch_args,
            )
            self._owned = True
        self.page = await self._pick_page()
        return self

    async def __aexit__(self, exc_type, exc, tb) -> None:
        try:
            if self.context is not None:
                if self._owned:
                    await self.context.close()
                elif self.browser is not None:
                    # Disconnects this Playwright client; leaves Chrome running.
                    await self.browser.close()
        except Exception:
            pass
        try:
            if self._pw is not None:
                await self._pw.stop()
        except Exception:
            pass

    @staticmethod
    def _configure_print_preferences(
        profile: Path,
        print_dir: Path,
        *,
        paper: str = DEFAULT_PAPER,
        landscape: bool = True,
    ) -> None:
        prefs_path = profile / "Default" / "Preferences"
        prefs_path.parent.mkdir(parents=True, exist_ok=True)
        try:
            prefs = json.loads(prefs_path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            prefs = {}
        app_state = {
            "recentDestinations": [
                {"id": "Save as PDF", "origin": "local", "account": ""}
            ],
            "selectedDestinationId": "Save as PDF",
            "version": 2,
            "isHeaderFooterEnabled": False,
            "isCssBackgroundEnabled": True,
        }
        dims = PAPER_SIZES.get(paper)
        if dims:
            width, height = dims
            app_state["mediaSize"] = {
                "width_microns": width,
                "height_microns": height,
                "imageable_area_left_microns": 0,
                "imageable_area_bottom_microns": 0,
                "imageable_area_right_microns": width,
                "imageable_area_top_microns": height,
                "is_default": False,
                "name": f"ISO_{paper}" if paper.startswith("A") else paper,
                "custom_display_name": paper,
            }
        app_state["isLandscapeEnabled"] = bool(landscape)
        prefs.setdefault("printing", {}).setdefault(
            "print_preview_sticky_settings", {}
        )["appState"] = json.dumps(app_state)
        prefs.setdefault("savefile", {})["default_directory"] = str(print_dir)
        prefs_path.write_text(json.dumps(prefs), encoding="utf-8")

    async def _pick_page(self) -> Page:
        assert self.context is not None
        pages = [p for p in self.context.pages if not p.is_closed()]
        if not pages:
            return await self.context.new_page()

        target_host = ""
        if getattr(self.args, "url", None):
            target_host = urlsplit(self.args.url).netloc.lower()

        def score(page: Page) -> int:
            frame_urls = [f.url.lower() for f in page.frames]
            frame_names = [(f.name or "").lower() for f in page.frames]
            urls = [page.url.lower()] + frame_urls
            value = 0
            if target_host and any(target_host in url for url in urls):
                value += 100
            if any("webapplicationframe" in name for name in frame_names):
                value += 40
            if any("onenoteframe.aspx" in url or "doc.aspx" in url
                   for url in frame_urls):
                value += 30
            if "onedrive.live.com" in page.url.lower():
                value += 20
            if "onenote" in page.url.lower():
                value += 5
            if "onenote.cloud.microsoft" in page.url.lower():
                value -= 5
            return value

        return max(pages, key=score)

    async def goto(self, url: str) -> None:
        assert self.page is not None
        await self.page.goto(url, wait_until="domcontentloaded", timeout=120_000)
        await self.page.wait_for_timeout(2_000)


# ---------------------------------------------------------------------------
# UI detection helpers
# ---------------------------------------------------------------------------


async def probe_frame(frame: Frame) -> dict[str, Any]:
    try:
        return await frame.evaluate(NAV_PROBE_JS)
    except Exception:
        return {
            "sectionList": 0,
            "sectionItems": 0,
            "pageNodes": 0,
            "pageList": 0,
            "canvas": 0,
            "treeitems": 0,
            "fileTab": 0,
        }


def _frame_score(info: dict[str, Any]) -> int:
    return (
        int(info.get("sectionList") or 0) * 20
        + int(info.get("sectionItems") or 0) * 5
        + int(info.get("pageNodes") or 0) * 5
        + int(info.get("pageList") or 0) * 2
        + int(info.get("canvas") or 0) * 2
    )


async def wait_for_onenote_ui(page: Page, timeout: float = 120.0) -> Frame:
    """Return the frame that owns the OneNote navigation/content UI."""
    deadline = asyncio.get_event_loop().time() + timeout
    while asyncio.get_event_loop().time() < deadline:
        best: Optional[Frame] = None
        best_info: dict[str, Any] = {}
        best_score = 0
        for frame in page.frames:
            url = (frame.url or "").lower()
            if not (
                frame.name == "WebApplicationFrame"
                or "onenoteframe.aspx" in url
                or "onenote.officeapps.live.com" in url
            ):
                continue
            info = await probe_frame(frame)
            score = _frame_score(info)
            if score > best_score:
                best, best_info, best_score = frame, info, score
        if (
            best is not None
            and best_score > 0
            and int(best_info.get("sectionItems") or 0) > 0
        ):
            return best
        await asyncio.sleep(1.0)
    raise StructureError(
        "未找到 OneNote 应用框架（WebApplicationFrame / onenoteframe.aspx）"
        "及其分区列表。请确认已在可见浏览器中登录并打开笔记本。"
    )


async def detect_state(page: Page) -> dict[str, Any]:
    """Classify the current page: notebook_open / notebook_list / marketing / login."""
    rows = 0
    nav_score = 0
    app_frame: Optional[Frame] = None
    for frame in page.frames:
        info = await probe_frame(frame)
        score = _frame_score(info)
        nav_score = max(nav_score, score)
        if score > 0 and app_frame is None:
            app_frame = frame
        try:
            rows = max(rows, int(await frame.evaluate(
                "() => document.querySelectorAll('[data-automationid=\"field-name\"]').length"
            )))
        except Exception:
            pass
    url = page.url.lower()
    is_login_host = "login.microsoftonline.com" in url or "login.live.com" in url
    return {
        "url": page.url,
        "notebook_open": nav_score > 0,
        "notebook_list": rows > 0,
        "login": is_login_host,
        "frame": app_frame,
    }


async def find_named_control(
    page: Page, pattern: re.Pattern[str]
) -> Optional[Any]:
    """Find a visible button/link/tab whose accessible name matches pattern."""
    frames: list[Frame] = []
    if page.main_frame not in frames:
        frames.append(page.main_frame)
    for frame in page.frames:
        if frame not in frames:
            frames.append(frame)
    for frame in frames:
        for role in ("button", "link", "tab", "menuitem"):
            try:
                locator = frame.get_by_role(role, name=pattern)
                count = await locator.count()
            except Exception:
                continue
            for index in range(min(count, 6)):
                candidate = locator.nth(index)
                try:
                    if await candidate.is_visible():
                        return candidate
                except Exception:
                    continue
    return None


async def ensure_authenticated(
    page: Page,
    timeout: float,
    *,
    auto_entry: bool = True,
) -> dict[str, Any]:
    """Wait until OneNote app UI is present; click Sign in / Open OneNote once."""
    deadline = asyncio.get_event_loop().time() + timeout
    clicked_signin = False
    clicked_entry = False
    last_note = 0.0
    while asyncio.get_event_loop().time() < deadline:
        state = await detect_state(page)
        if state["notebook_open"] or state["notebook_list"]:
            return state
        now = asyncio.get_event_loop().time()
        if auto_entry and not clicked_signin and not state["login"]:
            control = await find_named_control(page, SIGN_IN_NAMES)
            if control is not None:
                log("检测到登录页，点击 Sign in / 登录 ...")
                try:
                    await control.click()
                    clicked_signin = True
                    await page.wait_for_timeout(2_000)
                    continue
                except Exception as exc:
                    log(f"点击登录按钮失败: {exc}")
        if auto_entry and clicked_signin and not clicked_entry:
            control = await find_named_control(page, APP_ENTRY_NAMES)
            if control is not None:
                log("点击 OneNote 入口 ...")
                try:
                    await control.click()
                    clicked_entry = True
                    await page.wait_for_timeout(2_000)
                    continue
                except Exception as exc:
                    log(f"点击 OneNote 入口失败: {exc}")
        if now - last_note > 30:
            last_note = now
            log(f"等待登录/打开笔记本... 当前 URL: {page.url}")
        # New OneNote home (onenote.cloud.microsoft) can be authenticated
        # without exposing notebook rows or the editor.  Do not wait for a
        # structure that may not exist; the caller will report it explicitly.
        if (
            "onenote.cloud.microsoft" in state["url"].lower()
            and not state["login"]
            and not clicked_signin
            and not clicked_entry
            and await find_named_control(page, SIGN_IN_NAMES) is None
        ):
            return state
        await asyncio.sleep(2.0)
    raise ExportError(
        "等待 OneNote 应用超时。请在可见浏览器中完成登录并打开笔记本，"
        "然后重新运行（或增大 --login-timeout）。"
    )


# ---------------------------------------------------------------------------
# DOM collection
# ---------------------------------------------------------------------------


async def _add_items(
    seen: dict[str, dict[str, Any]],
    items: list[dict[str, Any]],
    order_ref: list[int],
) -> int:
    new = 0
    for item in items:
        key = item.get("id")
        if not key:
            continue
        if key not in seen:
            item["_order"] = order_ref[0]
            order_ref[0] += 1
            seen[key] = item
            new += 1
        else:
            # Keep the most useful level information if a node is re-rendered.
            old = seen[key]
            for field_name in ("level", "levelSource", "indent", "domDepth",
                               "setsize", "posinset", "pos"):
                if item.get(field_name) not in (None, 0, "", -1) or field_name == "pos":
                    old[field_name] = item.get(field_name)
    return new


async def collect_virtual(
    frame: Frame,
    collect_js: str,
    item_selector: str,
    *,
    root_id: Optional[str] = None,
    list_selector: Optional[str] = None,
    reset_top: bool = True,
    max_rounds: int = 300,
    strict: bool = True,
) -> list[dict[str, Any]]:
    """Scroll a virtualized list from top to bottom and collect every item."""
    if reset_top:
        try:
            await frame.evaluate(
                SCROLL_TO_TOP_JS,
                {
                    "itemSelector": item_selector,
                    "rootId": root_id,
                    "listSelector": list_selector,
                },
            )
        except Exception:
            pass

    seen: dict[str, dict[str, Any]] = {}
    order_ref = [0]
    reached_bottom = False
    stale = 0

    for _ in range(max_rounds):
        payload = await frame.evaluate(collect_js, root_id)
        if isinstance(payload, dict):
            items = payload.get("items") or []
        else:
            items = payload or []
        new = await _add_items(seen, items, order_ref)

        step = await frame.evaluate(
            SCROLL_STEP_JS,
            {
                "itemSelector": item_selector,
                "rootId": root_id,
                "listSelector": list_selector,
            },
        )
        if not step.get("ok"):
            break
        if step.get("atBottom"):
            await frame.wait_for_timeout(350)
            payload = await frame.evaluate(collect_js, root_id)
            items = payload.get("items") if isinstance(payload, dict) else payload
            await _add_items(seen, items or [], order_ref)
            reached_bottom = True
            break
        stale = stale + 1 if new == 0 else 0
        if stale >= 8:
            break
        await frame.wait_for_timeout(320)

    if not reached_bottom:
        declared = max(
            (int(item.get("setsize") or 0) for item in seen.values()), default=0
        )
        if declared and len(seen) >= declared:
            reached_bottom = True
    if not reached_bottom and strict:
        raise StructureError(
            "列表滚动未到达底部，可能漏掉页面/分区；已中止以免生成不完整结果。"
        )
    if not reached_bottom:
        log(
            f"警告：{list_selector or item_selector} 滚动未确认到底，"
            f"已收集 {len(seen)} 项；继续并由 setsize/标题校验兜底。"
        )

    return sorted(
        seen.values(), key=lambda item: (item.get("pos", 0), item.get("_order", 0))
    )


async def collect_sections(
    frame: Frame, parent_id: Optional[str] = None
) -> list[NavItem]:
    probe = await frame.evaluate(COLLECT_SECTIONS_JS, parent_id)
    if isinstance(probe, dict) and int(probe.get("missingId") or 0) > 0:
        raise StructureError(
            "分区/分区组列表中存在缺少稳定 id 的项；无法安全遍历，已中止。"
        )
    items = await collect_virtual(
        frame,
        COLLECT_SECTIONS_JS,
        'div[class*="sectionListItem"], div[class*="sectionGroup__groupItemWrap"], [role="treeitem"]',
        root_id=parent_id,
        list_selector=None if parent_id else '#NavPaneSectionList',
        reset_top=parent_id is None,
        max_rounds=200 if parent_id is None else 40,
        strict=False,
    )
    return [
        NavItem(
            id=item["id"],
            name=item.get("name") or "Untitled",
            kind=item.get("kind") or "section",
            pos=item.get("pos", 0),
            setsize=item.get("setsize", 0),
            posinset=item.get("posinset", 0),
        )
        for item in items
    ]


async def collect_pages(frame: Frame) -> list[PageNode]:
    probe = await frame.evaluate(COLLECT_PAGES_JS)
    if isinstance(probe, dict) and int(probe.get("missingId") or 0) > 0:
        raise StructureError(
            "页面列表中存在缺少稳定 id 的页面项；无法安全选择/续跑，已中止。"
        )
    items = await collect_virtual(
        frame,
        COLLECT_PAGES_JS,
        '.pageNode, [class*="pageNode"]',
        list_selector=(
            '#PageList, [class*="pagesContainer"], '
            '[class*="pageListContainer"], [class*="pageList__"]'
        ),
        reset_top=True,
        max_rounds=400,
        strict=False,
    )
    pages = [
        PageNode(
            id=item["id"],
            title=item.get("title") or "Untitled Page",
            level=int(item.get("level", -1)),
            indent=float(item.get("indent", 0) or 0),
            dom_depth=int(item.get("domDepth", 0) or 0),
            pos=int(item.get("pos", 0) or 0),
            setsize=int(item.get("setsize", 0) or 0),
            posinset=int(item.get("posinset", 0) or 0),
        )
        for item in items
    ]
    return pages


async def click_item(
    frame: Frame, item_id: str, *, is_page: bool = False
) -> bool:
    if is_page:
        item_selector = '.pageNode, [class*="pageNode"]'
        list_selector = (
            '#PageList, [class*="pagesContainer"], '
            '[class*="pageListContainer"], [class*="pageList__"]'
        )
    else:
        item_selector = (
            'div[class*="sectionListItem"], '
            'div[class*="sectionGroup__groupItemWrap"], [role="treeitem"]'
        )
        list_selector = '#NavPaneSectionList, .sectionList, [class*="sectionList"]'

    async def real_click() -> bool:
        try:
            node = frame.locator(f'[id="{item_id}"]')
            if not await node.count():
                return False
            inner = node.locator('.navItem, [class*="navItem"]')
            target = inner.first if await inner.count() else node.first
            await target.click(timeout=5_000)
            return True
        except Exception:
            return False

    for _ in range(2):
        if await real_click():
            return True
        try:
            result = await frame.evaluate(CLICK_ITEM_JS, item_id)
        except Exception as exc:
            log(f"点击元素失败 {item_id}: {exc}")
            result = None
        if result and result.get("ok"):
            return True

        # The list is virtualized: scroll from the top until the node is
        # rendered, then click it.
        try:
            await frame.evaluate(
                SCROLL_TO_TOP_JS,
                {
                    "itemSelector": item_selector,
                    "rootId": None,
                    "listSelector": list_selector,
                },
            )
        except Exception:
            pass
        for _ in range(300):
            if await element_present(frame, item_id):
                if await real_click():
                    return True
                try:
                    result = await frame.evaluate(CLICK_ITEM_JS, item_id)
                except Exception as exc:
                    log(f"点击元素失败 {item_id}: {exc}")
                    return False
                return bool(result and result.get("ok"))
            try:
                step = await frame.evaluate(
                    SCROLL_STEP_JS,
                    {
                        "itemSelector": item_selector,
                        "rootId": None,
                        "listSelector": list_selector,
                    },
                )
            except Exception:
                break
            if not step.get("ok") or step.get("atBottom"):
                break
            await frame.wait_for_timeout(150)
    return False


async def element_present(frame: Frame, item_id: str) -> bool:
    try:
        return bool(await frame.evaluate(ELEMENT_PRESENT_JS, item_id))
    except Exception:
        return False


async def has_nested_group(frame: Frame, group_id: str) -> bool:
    try:
        return bool(await frame.evaluate(NESTED_GROUP_JS, group_id))
    except Exception:
        return False


async def group_expanded(frame: Frame, group_id: str) -> bool:
    try:
        return bool(await frame.evaluate(GROUP_EXPANDED_JS, group_id))
    except Exception:
        return True


async def selection_supported(frame: Frame) -> bool:
    try:
        return bool(await frame.evaluate(SELECTION_SUPPORTED_JS))
    except Exception:
        return False


async def is_selected(frame: Frame, item_id: str) -> dict[str, Any]:
    try:
        return await frame.evaluate(IS_SELECTED_JS, item_id)
    except Exception:
        return {"found": False, "selected": False}


async def read_page_title(frame: Frame) -> dict[str, Any]:
    try:
        return await frame.evaluate(PAGE_TITLE_JS)
    except Exception:
        return {"found": False, "title": ""}


async def is_section_locked(frame: Frame) -> bool:
    try:
        return bool(await frame.evaluate(LOCKED_SECTION_JS))
    except Exception:
        return False


async def wait_for_page_list(frame: Frame, timeout: float = 30_000) -> None:
    try:
        await frame.wait_for_function(
            """() => !!document.querySelector(
                '#PageList, [class*="pagesContainer"], [class*="pageListContainer"], ' +
                '[class*="pageList__"], .pageNode, [class*="pageNode"]')""",
            timeout=timeout,
        )
    except Exception as exc:
        raise StructureError(
            "选中分区后未出现页面列表；无法确认该分区内容。"
        ) from exc


async def wait_for_content_frame(page: Page, timeout: float = 30_000) -> Frame:
    deadline = asyncio.get_event_loop().time() + timeout / 1000.0
    while asyncio.get_event_loop().time() < deadline:
        for frame in page.frames:
            try:
                has = await frame.evaluate(
                    """() => {
                        const sels = [
                          '#WACViewPanel', '#PageContentContainer',
                          '.PageContentContainer', '#OreoCanvas',
                          '.canvasContainer', '[class*="canvasContainer"]',
                          '.TitleOutline', '.pageTitle', '#pageTitle',
                          '[aria-label*="Page Title"]',
                        ];
                        for (const sel of sels) {
                          for (const el of document.querySelectorAll(sel)) {
                            const r = el.getBoundingClientRect();
                            const cs = getComputedStyle(el);
                            if (r.width > 0 && r.height > 0 &&
                                cs.display !== 'none' &&
                                cs.visibility !== 'hidden') return true;
                          }
                        }
                        return false;
                    }"""
                )
            except Exception:
                continue
            if has:
                return frame
        await asyncio.sleep(0.4)
    raise StructureError("未找到 OneNote 页面内容区域（canvas / page title）。")


# ---------------------------------------------------------------------------
# Page hierarchy
# ---------------------------------------------------------------------------


def _cluster_indents(values: list[float]) -> dict[int, int]:
    unique = sorted({round(v) for v in values})
    groups: list[list[int]] = []
    for value in unique:
        if groups and value - groups[-1][-1] <= 4:
            groups[-1].append(value)
        else:
            groups.append([value])
    return {value: index for index, group in enumerate(groups) for value in group}


def build_tree(pages: list[PageNode]) -> list[PageNode]:
    """Normalize subpage levels and build parent/child relationships."""
    if not pages:
        return []

    levels = [p.level for p in pages]
    have_levels = [v for v in levels if v >= 0]
    if have_levels and len(have_levels) != len(levels):
        raise StructureError(
            "页面层级信号不一致（部分页面缺少 aria-level/缩进），"
            "拒绝猜测子页面关系。"
        )

    if have_levels:
        base = min(have_levels)
        for page in pages:
            page.level -= base
    else:
        indents = [p.indent for p in pages]
        if max(indents) - min(indents) > 4:
            mapping = _cluster_indents(indents)
            for page in pages:
                page.level = mapping[round(page.indent)]
        elif any(p.dom_depth for p in pages):
            base = min(p.dom_depth for p in pages)
            for page in pages:
                page.level = p.dom_depth - base
        else:
            for page in pages:
                page.level = 0

    previous = 0
    for index, page in enumerate(pages):
        if index == 0:
            page.level = 0
        elif page.level > previous + 1:
            page.level = previous + 1
        previous = page.level

    stack: list[PageNode] = []
    roots: list[PageNode] = []
    for page in pages:
        while stack and stack[-1].level >= page.level:
            stack.pop()
        if stack:
            page.parent = stack[-1]
            stack[-1].children.append(page)
        else:
            roots.append(page)
        stack.append(page)

    # If the accessibility tree exposes sibling counts, use them to detect
    # silently missing pages (virtualization failure).
    if roots and all(page.level == 0 for page in pages):
        need = max((p.setsize or 0) for p in roots)
        if need and len(roots) < need:
            raise StructureError(
                f"页面列表只收集到 {len(roots)} 项，但 aria-setsize 声明 {need} 项；"
                "虚拟列表可能漏页，已中止。"
            )
    return roots


def assign_paths(roots: list[PageNode], base_dir: Path) -> None:
    used: set[str] = set()
    for index, node in enumerate(roots, start=1):
        node.order = index
        token_len = 8
        token = short_token(node.id, token_len)
        while token in used and token_len < len(re.sub(r"[^0-9A-Za-z]", "", node.id)):
            token_len += 4
            token = short_token(node.id, token_len)
        while token in used:
            token_len += 4
            token = short_token(node.id + str(token_len), token_len)
        used.add(token)
        name = f"{index:02d} {slugify(node.title)} [{token}]"
        node.directory = base_dir / name
        node.pdf_path = node.directory / f"{name}.pdf"
        assign_paths(node.children, node.directory)


def flatten(roots: list[PageNode]) -> list[PageNode]:
    out: list[PageNode] = []
    for root in roots:
        out.extend(root.walk())
    return out


# ---------------------------------------------------------------------------
# Printer: the application's own File > Print, saved by Chromium as PDF
# ---------------------------------------------------------------------------

class Printer:
    """Drive OneNote's File > Print and collect the PDF Chromium writes.

    Chromium must run with ``--kiosk-printing`` and a "Save as PDF" preference
    that writes into ``print_dir``.  OneNote itself decides pagination and
    layout; this class only triggers its print command and moves the result.
    """

    def __init__(
        self,
        page: Page,
        print_dir: Path,
        paper: str = DEFAULT_PAPER,
        portrait: bool = False,
    ):
        self.page = page
        self.print_dir = print_dir
        self.paper = paper if paper in PAPER_SIZES else DEFAULT_PAPER
        self.portrait = portrait
        self.print_dir.mkdir(parents=True, exist_ok=True)

    @staticmethod
    def _frames(page: Page) -> list[Frame]:
        frames: list[Frame] = []
        for frame in page.frames:
            if frame not in frames:
                frames.append(frame)
        return frames

    async def _first_visible(self, selectors: list[str]) -> Optional[Any]:
        for frame in self._frames(self.page):
            for selector in selectors:
                try:
                    locator = frame.locator(selector)
                    count = await locator.count()
                except Exception:
                    continue
                for index in range(min(count, 4)):
                    candidate = locator.nth(index)
                    try:
                        if await candidate.is_visible():
                            return candidate
                    except Exception:
                        continue
        return None

    async def _click(self, selectors: list[str], label: str, timeout: float = 25.0) -> None:
        deadline = asyncio.get_event_loop().time() + timeout
        while asyncio.get_event_loop().time() < deadline:
            locator = await self._first_visible(selectors)
            if locator is not None:
                try:
                    await locator.click(timeout=8_000)
                    return
                except Exception as exc:
                    log(f"点击{label}失败，重试: {exc}")
            await asyncio.sleep(0.4)
        raise ExportError(f"找不到 OneNote 的{label}。")

    def _clear_staging(self) -> None:
        try:
            for path in self.print_dir.iterdir():
                if path.is_file() and path.suffix.lower() == ".pdf":
                    path.unlink()
        except OSError as exc:
            raise ExportError(
                f"无法清空打印暂存目录 {self.print_dir}: {exc}"
            ) from exc

    async def _prompt_now(self) -> bool:
        for frame in self._frames(self.page):
            try:
                dialog = frame.locator("#WACDialogOuterContainer")
                if not await dialog.count() or not await dialog.first.is_visible():
                    continue
                text = await dialog.first.inner_text()
            except Exception:
                continue
            if any(word in text for word in ("图像", "图片", "image", "download")):
                return True
        return False

    async def _wait_new_pdf(self, timeout: float) -> tuple[Path, bytes]:
        deadline = asyncio.get_event_loop().time() + timeout
        sizes: dict[Path, int] = {}
        while asyncio.get_event_loop().time() < deadline:
            ready: list[tuple[float, Path, bytes]] = []
            try:
                entries = list(self.print_dir.iterdir())
            except OSError:
                entries = []
            for path in entries:
                if not path.is_file():
                    continue
                try:
                    size = path.stat().st_size
                    if size < 1024:
                        continue
                    if sizes.get(path) != size:
                        sizes[path] = size
                        continue
                    data = path.read_bytes()
                except OSError:
                    continue
                if is_pdf_bytes(data) and b"%%EOF" in data[-4096:]:
                    ready.append((path.stat().st_mtime, path, data))
            if ready:
                ready.sort(key=lambda item: item[0])
                if len(ready) > 1:
                    log(f"打印暂存目录出现 {len(ready)} 个 PDF，取最新一个。")
                _, path, data = ready[-1]
                try:
                    path.unlink()
                except OSError:
                    pass
                return path, data
            if await self._prompt_now():
                raise _ImagePrompt()
            await asyncio.sleep(0.3)
        raise ExportError(
            f"OneNote 打印后 {timeout:.0f}s 内没有在 {self.print_dir} 生成 PDF。"
        )

    async def _open_print_backstage(self) -> None:
        if await self._first_visible(["#FileMenuPrint"]) is not None:
            return
        await self._click(
            [
                "#FileMenuLauncherContainer",
                "#FileTab",
                "[role='tab'][aria-label='文件']",
                "[role='tab'][aria-label='File']",
                "button:has-text('文件')",
                "button:has-text('File')",
            ],
            "File / 文件 菜单",
        )
        await self._click(
            ["#FilePrintPage", "[data-unique-id='FilePrintPage']"],
            "Print / 打印 页面",
            timeout=15.0,
        )

    async def _print_button_enabled(self, timeout: float) -> bool:
        deadline = asyncio.get_event_loop().time() + timeout
        while asyncio.get_event_loop().time() < deadline:
            locator = await self._first_visible(["#FileMenuPrint"])
            if locator is not None:
                try:
                    if await locator.get_attribute("aria-disabled") != "true":
                        return True
                except Exception:
                    pass
            await asyncio.sleep(0.5)
        return False

    async def _click_print_action(self) -> None:
        locator = await self._first_visible(["#FileMenuPrint"])
        if locator is None:
            raise ExportError("找不到打印按钮（#FileMenuPrint）。")
        try:
            # Native kiosk printing may block this call until the PDF is
            # written; a timeout does not mean the print did not happen.
            await asyncio.wait_for(locator.click(timeout=5_000), timeout=15.0)
        except Exception as exc:
            log(f"打印命令点击等待结束: {exc}")

    async def _scroll_bottom_top(self) -> None:
        for frame in self._frames(self.page):
            try:
                info = await frame.evaluate(
                    """() => {
                        const el = document.querySelector('#WACViewPanel');
                        if (!el) return null;
                        window.__ompScrollEl = el;
                        return {height: el.scrollHeight, client: el.clientHeight};
                    }"""
                )
            except Exception:
                continue
            if not info:
                continue
            height = int(info.get("height") or 0)
            client = int(info.get("client") or 0)
            if height <= client + 20:
                return
            log(f"图像未下载完：滚动到底部再回到顶部（{height}px）")
            step = max(400, height // 100)
            pos = 0
            while pos < height:
                pos = min(height, pos + step)
                try:
                    await frame.evaluate(
                        "(pos) => { if (window.__ompScrollEl) "
                        "window.__ompScrollEl.scrollTop = pos; }",
                        pos,
                    )
                except Exception:
                    return
                await self.page.wait_for_timeout(80)
            for _ in range(40):
                try:
                    pending = await frame.evaluate(
                        "() => Array.from(document.images)"
                        ".filter((img) => !img.complete).length"
                    )
                except Exception:
                    pending = 0
                if not pending:
                    break
                await self.page.wait_for_timeout(500)
            pos = height
            while pos > 0:
                pos = max(0, pos - step)
                try:
                    await frame.evaluate(
                        "(pos) => { if (window.__ompScrollEl) "
                        "window.__ompScrollEl.scrollTop = pos; }",
                        pos,
                    )
                except Exception:
                    return
                await self.page.wait_for_timeout(80)
            await self.page.wait_for_timeout(300)
            return

    async def _apply_page_style(self) -> None:
        width, height = PAPER_SIZES[self.paper]
        if not self.portrait:
            width, height = height, width
        css = (
            f"@page {{ size: {width / 1000:g}mm {height / 1000:g}mm; "
            "margin: 0; }}"
        )
        for frame in self._frames(self.page):
            try:
                is_app = await frame.evaluate(
                    "() => !!document.querySelector('#WACViewPanel, #NavPaneSectionList')"
                )
            except Exception:
                continue
            if not is_app:
                continue
            await frame.evaluate(
                """(css) => {
                    let style = document.getElementById('__omp_print_page');
                    if (!style) {
                        style = document.createElement('style');
                        style.id = '__omp_print_page';
                        document.head.appendChild(style);
                    }
                    style.textContent = css;
                }""",
                css,
            )
            return

    async def recover(self) -> None:
        for _ in range(3):
            clicked = False
            for frame in self._frames(self.page):
                for selector in (
                    "#WACDialogOuterContainer button[aria-label='close']",
                    "#WACDialogOuterContainer button[aria-label='关闭']",
                    "#WACDialogOuterContainer button[aria-label='Close']",
                ):
                    try:
                        button = frame.locator(selector)
                        if await button.count() and await button.first.is_visible():
                            await button.first.click(timeout=3_000)
                            await self.page.wait_for_timeout(250)
                            clicked = True
                            break
                    except Exception:
                        continue
                if clicked:
                    break
            if clicked:
                continue
            for frame in self._frames(self.page):
                try:
                    locator = frame.locator("#FileMenuBackButton").first
                    if await locator.count() and await locator.is_visible():
                        await locator.click(timeout=5_000)
                        await self.page.wait_for_timeout(250)
                        clicked = True
                        break
                except Exception:
                    continue
            if not clicked:
                break
        try:
            await self.page.keyboard.press("Escape")
            await self.page.wait_for_timeout(200)
        except Exception:
            pass

    async def print_page(self, timeout: float = 180.0) -> bytes:
        last_error = ""
        per_attempt = max(90.0, timeout / 2.0)
        for attempt in range(1, 6):
            await self.recover()
            await self._apply_page_style()
            self._clear_staging()
            try:
                await self._open_print_backstage()
                if not await self._print_button_enabled(timeout=45.0):
                    last_error = "打印按钮一直不可用"
                    log(f"第 {attempt}/5 次：{last_error}")
                    await self._scroll_bottom_top()
                    continue
                await self._click_print_action()
                _, data = await self._wait_new_pdf(per_attempt)
                if not is_pdf_bytes(data):
                    raise ExportError("打印输出不是有效 PDF。")
                return data
            except _ImagePrompt:
                last_error = "OneNote 提示图像仍在下载"
                log(f"第 {attempt}/5 次：{last_error}，滚动后重试")
                await self.recover()
                await self._scroll_bottom_top()
            except ExportError as exc:
                last_error = str(exc)
                log(f"第 {attempt}/5 次打印失败: {last_error}")
                await self.recover()
                await self._scroll_bottom_top()
        raise ExportError(f"页面打印尝试 5 次仍失败: {last_error}")


async def print_current_page(
    page: Page,
    print_dir: Path,
    timeout: float = 60.0,
    paper: str = DEFAULT_PAPER,
    portrait: bool = False,
) -> bytes:
    """Print the currently displayed OneNote page through File > Print."""
    return await Printer(page, print_dir, paper, portrait).print_page(timeout)


# ---------------------------------------------------------------------------
# Export flow
# ---------------------------------------------------------------------------


async def list_notebooks(page: Page) -> list[NotebookRef]:
    deadline = asyncio.get_event_loop().time() + 8.0
    while asyncio.get_event_loop().time() < deadline:
        for frame in page.frames:
            try:
                rows = await frame.evaluate(LIST_ROWS_JS)
            except Exception:
                continue
            if rows:
                seen: set[str] = set()
                notebooks: list[NotebookRef] = []
                for row in sorted(rows, key=lambda r: r.get("pos", 0)):
                    key = row.get("id") or row.get("name") or ""
                    if not key or key in seen:
                        continue
                    seen.add(key)
                    notebooks.append(
                        NotebookRef(
                            name=row.get("name") or "Notebook",
                            id=row.get("id") or "",
                            source="listing",
                        )
                    )
                if notebooks:
                    return notebooks
        await asyncio.sleep(1.0)
    raise StructureError(
        "未发现笔记本列表（没有 [data-automationid=field-name] 行）。"
        "请用 --url 指定已打开的笔记本 URL，或先运行 login 并确认笔记本已显示。"
    )


async def open_notebook(session: Session, notebook: NotebookRef) -> bool:
    assert session.page is not None
    clicked = False
    for frame in session.page.frames:
        try:
            result = await frame.evaluate(
                """(nb) => {
                    const rows = Array.from(document.querySelectorAll('div[role="row"]'));
                    let row = null;
                    if (nb.id) {
                        row = rows.find((r) => r.getAttribute('data-automationid') === nb.id);
                    }
                    if (!row) {
                        row = rows.find((r) => {
                            const btn = r.querySelector('[data-automationid="field-name"] ' +
                                'button[role="link"], [data-automationid="field-name"] a');
                            return btn && (btn.innerText || '').trim() === nb.name;
                        });
                    }
                    if (!row) return false;
                    const btn = row.querySelector('[data-automationid="field-name"] ' +
                        'button[role="link"], [data-automationid="field-name"] a[href]');
                    if (!btn) return false;
                    btn.click();
                    return true;
                }""",
                {"id": notebook.id, "name": notebook.name},
            )
        except Exception:
            continue
        if result:
            clicked = True
            break
    if not clicked:
        log(f"无法点击笔记本: {notebook.name}")
        return False

    deadline = asyncio.get_event_loop().time() + 120.0
    while asyncio.get_event_loop().time() < deadline:
        for candidate in list(session.context.pages):
            if candidate.is_closed():
                continue
            try:
                frame = await wait_for_onenote_ui(candidate, timeout=2.0)
            except ExportError:
                continue
            session.page = candidate
            log(f"已打开笔记本，内容 frame: {frame.url}")
            return True
        await asyncio.sleep(1.0)
    log(f"打开笔记本后未出现 OneNote 导航: {notebook.name}")
    return False


async def read_notebook_name(page: Page, nav_frame: Optional[Frame] = None) -> str:
    frames: list[Frame] = []
    if nav_frame is not None:
        frames.append(nav_frame)
    frames.extend(f for f in page.frames if f not in frames)
    for frame in frames:
        try:
            data = await frame.evaluate(
                """() => {
                    const crumb = document.querySelector(
                        '[class*="Breadcrumb"], [class*="breadcrumb"]');
                    const items = crumb
                        ? Array.from(crumb.querySelectorAll('button, a, span'))
                            .map((e) => (e.innerText || '').trim())
                            .filter(Boolean)
                        : [];
                    return { items, title: document.title || '' };
                }"""
            )
        except Exception:
            continue
        items = data.get("items") or []
        if items:
            first = items[0]
            if not re.match(r"^(My notebooks|Notebooks|我的笔记本|笔记本)$", first, re.I):
                return first
        title = data.get("title") or ""
        title = re.split(r"\s+[-–—]\s+OneNote.*$", title, maxsplit=1, flags=re.I)[0]
        if title and title.lower() not in {"onenote", "microsoft onenote"}:
            return title
    try:
        title = await page.title()
    except Exception:
        title = ""
    title = re.split(r"\s+[-–—]\s+OneNote.*$", title, maxsplit=1, flags=re.I)[0]
    title = title.strip()
    if title and title.lower() not in {"onenote", "microsoft onenote"}:
        return title
    return "Notebook"


async def select_section_verified(nav_frame: Frame, item: NavItem) -> None:
    """Click a section, wait for selection, and require its page list to change."""
    before = await nav_frame.evaluate(
        "() => Array.from(document.querySelectorAll('#PageList .pageNode'))"
        ".map((el) => el.id)"
    )
    for _ in range(4):
        if not await click_item(nav_frame, item.id):
            await nav_frame.wait_for_timeout(500)
            continue
        selected = False
        for _ in range(50):
            state = await is_selected(nav_frame, item.id)
            if state.get("found") and state.get("selected"):
                selected = True
                break
            await nav_frame.wait_for_timeout(200)
        if not selected:
            continue
        if not before:
            await nav_frame.wait_for_timeout(600)
            return
        for _ in range(50):
            after = await nav_frame.evaluate(
                "() => Array.from(document.querySelectorAll('#PageList .pageNode'))"
                ".map((el) => el.id)"
            )
            if after and after != before:
                await nav_frame.wait_for_timeout(400)
                return
            await nav_frame.wait_for_timeout(200)
    raise ExportError(
        f"无法确认分区已切换：{item.name!r}；"
        "为避免把上一个分区的页面重复导出，已中止。"
    )


async def select_page_verified(
    nav_frame: Frame,
    page: Page,
    node: PageNode,
) -> Frame:
    support = await selection_supported(nav_frame)
    last_reason = ""
    for _ in range(3):
        if not await click_item(nav_frame, node.id, is_page=True):
            last_reason = "页面项不存在或不可点击"
            await page.wait_for_timeout(1_000)
            continue
        content = await wait_for_content_frame(page, timeout=30_000)
        await page.wait_for_timeout(700)
        if support:
            selected = await is_selected(nav_frame, node.id)
            if selected.get("found") and selected.get("selected"):
                return content
            last_reason = "页面项未显示选中状态"
        info = await read_page_title(content)
        if info.get("found") and titles_match(info.get("title", ""), node.title):
            return content
        last_reason = (
            f"标题不匹配（编辑器={info.get('title')!r}, 列表={node.title!r}）"
        )
        await page.wait_for_timeout(900)
    raise ExportError(f"无法确认已选中页面 {node.title!r}: {last_reason}")


async def navigate_back(
    page: Page, nav_frame: Frame, expected_parent_id: Optional[str] = None
) -> None:
    selectors = [
        "button[aria-label='Back']",
        "button[aria-label='后退']",
        "button[aria-label='返回']",
        "button[title='Back']",
        "button[title='后退']",
        "[class*='backButton'] button",
        "[class*='backButton']",
    ]
    for selector in selectors:
        try:
            locator = nav_frame.locator(selector).first
            if await locator.count() and await locator.is_visible():
                await locator.click()
                await nav_frame.wait_for_timeout(1_200)
                if expected_parent_id is None or await element_present(
                    nav_frame, expected_parent_id
                ):
                    return
        except Exception:
            continue

    # Breadcrumb fallback: return to the parent level.
    for selector in (
        "[class*='Breadcrumb'] button",
        "[class*='breadcrumb'] button",
    ):
        try:
            locator = nav_frame.locator(selector).first
            if await locator.count() and await locator.is_visible():
                await locator.click()
                await nav_frame.wait_for_timeout(1_200)
                if expected_parent_id is None or await element_present(
                    nav_frame, expected_parent_id
                ):
                    return
        except Exception:
            continue

    try:
        await page.go_back(wait_until="domcontentloaded")
        await nav_frame.wait_for_timeout(1_500)
        if expected_parent_id is None or await element_present(
            nav_frame, expected_parent_id
        ):
            return
    except Exception:
        pass

    raise StructureError(
        "进入分区组后无法返回上级导航；已中止，避免层级错乱。"
    )


async def export_section(
    nav_frame: Frame,
    page: Page,
    item: NavItem,
    out_dir: Path,
    trail: list[str],
    ctx: ExportContext,
) -> None:
    printer = Printer(
        page,
        Path(ctx.args.print_dir).expanduser().resolve(),
        ctx.args.paper,
        ctx.args.portrait,
    )
    await select_section_verified(nav_frame, item)

    if await is_section_locked(nav_frame):
        ctx.error("/".join(trail), "分区受密码保护，无法自动导出（可手动解锁后重跑）")
        return

    try:
        await wait_for_page_list(nav_frame)
    except StructureError as exc:
        ctx.error("/".join(trail), str(exc))
        return

    pages = await collect_pages(nav_frame)
    if not pages:
        log(f"  分区为空: {'/'.join(trail)}")
        return

    try:
        roots = build_tree(pages)
    except StructureError as exc:
        ctx.error("/".join(trail), str(exc))
        return

    section_dir = out_dir / slugify(item.name)
    assign_paths(roots, section_dir)

    for node in flatten(roots):
        if ctx.args.max_pages and ctx.exported >= ctx.args.max_pages:
            raise _StopExport()
        assert node.pdf_path is not None
        rel_path = node.pdf_path.relative_to(ctx.output_root).as_posix()
        if not ctx.args.no_resume and ctx.manifest.is_done(node.id, rel_path):
            ctx.skipped += 1
            log(f"  跳过（已完成）: {'/'.join(trail + [node.title])}")
            continue

        try:
            await select_page_verified(nav_frame, page, node)
            data = await printer.print_page(timeout=ctx.args.print_timeout)
            atomic_write_bytes(node.pdf_path, data)
            ctx.manifest.record_ok(
                node, rel_path, len(data), ctx.notebook, trail + [node.title]
            )
            ctx.exported += 1
            log(
                f"  [OK] {'/'.join(trail + [node.title])} -> "
                f"{rel_path} ({len(data)} bytes)"
            )
        except _StopExport:
            raise
        except Exception as exc:
            ctx.manifest.record_failed(
                node, rel_path, str(exc), ctx.notebook, trail + [node.title]
            )
            ctx.error("/".join(trail + [node.title]), str(exc))
            await printer.recover()


async def walk_sections(
    nav_frame: Frame,
    page: Page,
    out_dir: Path,
    *,
    parent_id: Optional[str],
    trail: list[str],
    seen: set[str],
    ctx: ExportContext,
) -> None:
    items = await collect_sections(nav_frame, parent_id)
    for _ in range(10):
        if items:
            break
        await nav_frame.wait_for_timeout(1_000)
        items = await collect_sections(nav_frame, parent_id)
    if parent_id is None and not items:
        probe = await probe_frame(nav_frame)
        if not (
            probe.get("sectionList")
            or probe.get("sectionItems")
            or probe.get("treeitems")
        ):
            raise StructureError(
                "未找到 OneNote 分区导航；拒绝把空列表当作成功。"
            )
        log("  笔记本没有分区（或当前层级为空）")
        return

    for item in items:
        if ctx.args.max_pages and ctx.exported >= ctx.args.max_pages:
            raise _StopExport()
        if item.id in seen:
            continue
        seen.add(item.id)

        if item.kind == "group":
            if ctx.args.section and not _section_filter_matches(ctx.args.section, item):
                # Still descend: a section inside may match.
                pass
            group_dir = out_dir / slugify(item.name)
            group_dir.mkdir(parents=True, exist_ok=True)
            nested = await has_nested_group(nav_frame, item.id)
            navigated = False

            if nested and not await group_expanded(nav_frame, item.id):
                await click_item(nav_frame, item.id)
                await nav_frame.wait_for_timeout(900)
            elif not nested:
                await click_item(nav_frame, item.id)
                await nav_frame.wait_for_timeout(1_000)
                nested = await has_nested_group(nav_frame, item.id)
                navigated = not await element_present(nav_frame, item.id)

            if nested:
                await walk_sections(
                    nav_frame,
                    page,
                    group_dir,
                    parent_id=item.id,
                    trail=trail + [item.name],
                    seen=seen,
                    ctx=ctx,
                )
            elif navigated:
                await walk_sections(
                    nav_frame,
                    page,
                    group_dir,
                    parent_id=None,
                    trail=trail + [item.name],
                    seen=seen,
                    ctx=ctx,
                )
            else:
                await nav_frame.wait_for_timeout(1_000)
                if await has_nested_group(nav_frame, item.id):
                    await walk_sections(
                        nav_frame,
                        page,
                        group_dir,
                        parent_id=item.id,
                        trail=trail + [item.name],
                        seen=seen,
                        ctx=ctx,
                    )
                else:
                    ctx.error(
                        "/".join(trail + [item.name]),
                        "分区组无法展开或进入；拒绝静默跳过",
                    )

            if navigated:
                await navigate_back(page, nav_frame, expected_parent_id=None)
                await wait_for_onenote_ui(page, timeout=30.0)
            continue

        if ctx.args.section and not _section_filter_matches(ctx.args.section, item):
            continue

        await export_section(
            nav_frame, page, item, out_dir, trail + [item.name], ctx
        )


def _section_filter_matches(needle: str, item: NavItem) -> bool:
    if not needle:
        return True
    needle = needle.casefold()
    return needle == item.id.casefold() or needle in item.name.casefold()


async def export_notebook(
    session: Session,
    notebook: NotebookRef,
    out_root: Path,
    manifest: Manifest,
    args: argparse.Namespace,
) -> ExportContext:
    assert session.page is not None
    nav_frame = await wait_for_onenote_ui(session.page, timeout=args.nav_timeout)
    name = notebook.name or await read_notebook_name(session.page, nav_frame)
    token = notebook_token(notebook, session.page.url)
    notebook_dir = out_root / f"{slugify(name, 50)} [{token}]"
    notebook_dir.mkdir(parents=True, exist_ok=True)

    ctx = ExportContext(
        output_root=out_root,
        manifest=manifest,
        notebook=name,
        args=args,
    )
    log(f"导出笔记本: {name} -> {notebook_dir}")
    log(f"导航 frame: {nav_frame.url}")

    seen: set[str] = set()
    try:
        await walk_sections(
            nav_frame,
            session.page,
            notebook_dir,
            parent_id=None,
            trail=[name],
            seen=seen,
            ctx=ctx,
        )
    except _StopExport:
        pass
    return ctx


async def run_export(args: argparse.Namespace) -> int:
    out_root = Path(args.output).expanduser().resolve()
    out_root.mkdir(parents=True, exist_ok=True)
    manifest = Manifest(out_root / "manifest.json", out_root)

    failures = 0
    exported = 0
    skipped = 0

    if args.cdp_url:
        log(
            "使用 --cdp-url：该 Chromium 必须已用 --kiosk-printing 启动，"
            "且“另存为 PDF”默认目录为 "
            f"{Path(args.print_dir).expanduser().resolve()}；"
            "本工具不会修改正在运行的浏览器偏好。"
        )
        log(
            "纸张与方向通过打印前注入 @page 生效；浏览器自身的"
            "默认打印机/保存目录沿用已有偏好。"
        )

    async with Session(args) as session:
        page = session.page
        assert page is not None
        if args.url:
            await session.goto(args.url)
        else:
            state = await detect_state(page)
            if not (state["notebook_open"] or state["notebook_list"]):
                await session.goto(NOTEBOOKS_URL)
        await ensure_authenticated(page, args.login_timeout)

        state = await detect_state(page)
        if state["notebook_open"]:
            notebooks = [
                NotebookRef(
                    name=await read_notebook_name(page),
                    url=page.url,
                    source="url",
                )
            ]
        else:
            notebooks = await list_notebooks(page)

        try:
            for notebook in notebooks:
                if args.max_pages and exported >= args.max_pages:
                    break
                if notebook.source == "listing":
                    await session.goto(NOTEBOOKS_URL)
                    await ensure_authenticated(page, args.login_timeout)
                    if not await open_notebook(session, notebook):
                        manifest.record_error(
                            notebook.name, "无法从笔记本列表打开该笔记本"
                        )
                        failures += 1
                        continue
                ctx = await export_notebook(
                    session, notebook, out_root, manifest, args
                )
                failures += ctx.failures
                exported += ctx.exported
                skipped += ctx.skipped
        except _StopExport:
            pass
        finally:
            manifest.save()
            write_index(out_root, manifest)

    log(
        f"完成: 新导出 {exported} 页，跳过 {skipped} 页，失败 {failures} 项。"
        f" manifest: {out_root / 'manifest.json'}"
    )
    return 1 if failures else 0


# ---------------------------------------------------------------------------
# login / inspect
# ---------------------------------------------------------------------------


async def run_login(args: argparse.Namespace) -> int:
    async with Session(args) as session:
        page = session.page
        assert page is not None
        if args.url:
            await session.goto(args.url)
        else:
            state = await detect_state(page)
            if not (state["notebook_open"] or state["notebook_list"]):
                await session.goto(NOTEBOOKS_URL)
        log("请在可见浏览器中完成 Microsoft 登录并打开一个笔记本。")
        log(f"profile: {Path(args.profile).expanduser().resolve()}")
        state = await ensure_authenticated(page, args.login_timeout)
        if state.get("notebook_open") or state.get("notebook_list"):
            log("已检测到 OneNote 应用界面。登录状态已保存在 profile 中。")
        else:
            log(
                "登录流程结束（未检测到笔记本编辑器）。"
                "请手动打开一个笔记本，或 export 时用 --url 指定笔记本 URL。"
            )
    return 0


async def run_inspect(args: argparse.Namespace) -> int:
    async with Session(args) as session:
        page = session.page
        assert page is not None
        if args.url:
            await session.goto(args.url)
        await page.wait_for_timeout(int(args.settle * 1000))

        state = await detect_state(page)
        print(f"page url: {page.url}")
        print(
            "state:",
            "notebook_open" if state["notebook_open"]
            else "notebook_list" if state["notebook_list"]
            else "login" if state["login"]
            else "unknown/marketing",
        )
        print(f"frames: {len(page.frames)}")
        for index, frame in enumerate(page.frames):
            try:
                info = await frame.evaluate(INSPECT_JS)
            except Exception as exc:
                print(f"\n[{index}] {frame.url}  (inspect failed: {exc})")
                continue
            parent = frame.parent_frame
            parent_url = parent.url if parent else "(main)"
            print(f"\n[{index}] url={info.get('url')}")
            print(f"    title={info.get('title')!r}")
            print(f"    parent={parent_url}")
            print(f"    counts={info.get('counts')}")
            print(f"    roles={info.get('roles')}")
            print(f"    aria-levels={info.get('levels')}")
            print(f"    iframes={info.get('iframes')}")
            labels = info.get("labels") or []
            if labels:
                print("    nav labels (titles only):")
                for label in labels:
                    print(f"      - {label}")
    return 0


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def build_parser() -> argparse.ArgumentParser:
    common = argparse.ArgumentParser(add_help=False)
    common.add_argument(
        "--cdp-url",
        default=None,
        help="连接已运行的 Chromium，例如 http://127.0.0.1:9223；不传则启动持久 profile",
    )
    common.add_argument(
        "--profile",
        default=DEFAULT_PROFILE,
        help=f"持久 Chromium profile 目录（默认 {DEFAULT_PROFILE}）",
    )
    common.add_argument(
        "--url",
        default=None,
        help=f"起始 URL（默认 {NOTEBOOKS_URL}；export 可直接传笔记本 URL）",
    )
    common.add_argument(
        "--print-dir",
        default=DEFAULT_PRINT_DIR,
        help=f"Chromium “另存为 PDF”的暂存目录（默认 {DEFAULT_PRINT_DIR}）",
    )
    common.add_argument(
        "--paper",
        default=DEFAULT_PAPER,
        choices=sorted(PAPER_SIZES),
        help=f"打印纸张（默认 {DEFAULT_PAPER}）",
    )
    common.add_argument(
        "--portrait",
        action="store_true",
        help="纵向打印（默认 A0 横向，适合超宽 OneNote 画布）",
    )

    parser = argparse.ArgumentParser(
        description="通过 OneNote Web 的 File > Print 导出每页 PDF"
    )
    sub = parser.add_subparsers(dest="command", required=True)

    login = sub.add_parser("login", parents=[common], help="打开 OneNote 并等待人工登录")
    login.add_argument("--login-timeout", type=float, default=900.0)

    export = sub.add_parser("export", parents=[common], help="导出笔记本为 PDF")
    export.add_argument("--output", default=DEFAULT_OUTPUT)
    export.add_argument("--login-timeout", type=float, default=900.0)
    export.add_argument("--nav-timeout", type=float, default=120.0)
    export.add_argument("--print-timeout", type=float, default=180.0)
    export.add_argument(
        "--max-pages",
        type=int,
        default=0,
        help="最多新导出多少页（0=不限；用于窄验证）",
    )
    export.add_argument(
        "--section",
        default="",
        help="只导出名称/ID 匹配该子串的分区（用于窄验证）",
    )
    export.add_argument(
        "--no-resume",
        action="store_true",
        help="忽略 manifest 中已成功的页面，重新导出",
    )

    inspect = sub.add_parser(
        "inspect", parents=[common], help="打印当前页面 frame / role / 选择器信息"
    )
    inspect.add_argument("--settle", type=float, default=4.0)
    return parser


def main(argv: Optional[list[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        if args.command == "login":
            return asyncio.run(run_login(args))
        if args.command == "export":
            return asyncio.run(run_export(args))
        if args.command == "inspect":
            return asyncio.run(run_inspect(args))
    except KeyboardInterrupt:
        log("已中断。")
        return 130
    except ExportError as exc:
        log(f"错误: {exc}")
        return 1
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
