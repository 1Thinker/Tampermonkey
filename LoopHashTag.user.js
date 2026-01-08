// ==UserScript==
// @name         Microsoft Loop Hashtag Tag Helper
// @author       https://www.appz.xyz
// @version      1.0
// @description  Autocomplete and manage #tags in Microsoft Loop, with import/export and search
// @match        https://loop.cloud.microsoft/*
// @downloadURL  https://github.com/1Thinker/Tampermonkey/raw/refs/heads/main/LoopHashTag.user.js
// @updateURL    https://github.com/1Thinker/Tampermonkey/raw/refs/heads/main/LoopHashTag.user.js
// @run-at       document-idle
// @grant        GM_getValue
// @grant        GM_setValue
// ==/UserScript==

(function () {
    'use strict';

    const STORAGE_KEY = 'tm_tag_helper_tags';
    const TAG_DROPDOWN_ID = 'tm-tag-helper-dropdown';
    const TAG_MANAGER_ID = 'tm-tag-helper-manager';
    const TAG_BUTTON_ID = 'tm-tag-helper-button';
    const DEBUG = false;

    function dbg() {}

    // ---------- Storage helpers ----------
    function loadTags() {
        try {
            const raw = GM_getValue(STORAGE_KEY, '[]');
            const arr = JSON.parse(raw);
            return Array.isArray(arr) ? arr : [];
        } catch (e) {
            console.error('Tag helper: failed to load tags', e);
            return [];
        }
    }

    function saveTags(tags) {
        try {
            GM_setValue(STORAGE_KEY, JSON.stringify([...new Set(tags)].sort()));
        } catch (e) {
            console.error('Tag helper: failed to save tags', e);
        }
    }

    function addTag(tag) {
        tag = tag.trim();
        if (!tag) return;
        const tags = loadTags();
        if (!tags.includes(tag)) {
            tags.push(tag);
            saveTags(tags);
        }
    }

    function removeTag(tag) {
        const tags = loadTags().filter(t => t !== tag);
        saveTags(tags);
    }

    // ---------- Import / Export helpers ----------
    function exportTagsAsText() {
        const tags = loadTags();
        return tags.join('\n');
    }

    function importTagsFromString(str) {
        if (!str) return;
        let imported = [];
        const trimmed = str.trim();

        if (trimmed.startsWith('[')) {
            try {
                const arr = JSON.parse(trimmed);
                if (Array.isArray(arr)) imported = arr;
            } catch (e) {
                console.warn('Tag helper: JSON import failed, falling back to lines', e);
            }
        }

        if (!imported.length) {
            imported = trimmed
                .split(/\r?\n/)
                .map(line => line.trim())
                .filter(Boolean);
        }

        imported = imported.map(t => t.replace(/^#/, '').trim()).filter(Boolean);
        if (!imported.length) return;

        const existing = loadTags();
        const merged = [...new Set([...existing, ...imported])].sort();
        saveTags(merged);
    }

    // ---------- Dropdown helpers ----------
    function createDropdown() {
        let dd = document.getElementById(TAG_DROPDOWN_ID);
        if (dd) return dd;

        dd = document.createElement('div');
        dd.id = TAG_DROPDOWN_ID;
        dd.style.position = 'fixed';
        dd.style.zIndex = '999999';
        dd.style.minWidth = '160px';
        dd.style.maxHeight = '200px';
        dd.style.overflowY = 'auto';
        dd.style.background = '#222';
        dd.style.color = '#fff';
        dd.style.border = '1px solid #555';
        dd.style.borderRadius = '4px';
        dd.style.fontSize = '12px';
        dd.style.padding = '4px 0';
        dd.style.boxShadow = '0 4px 8px rgba(0,0,0,0.4)';
        dd.style.display = 'none';
        dd.style.fontFamily = 'system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif';
        document.body.appendChild(dd);
        return dd;
    }

    let activeDropdownIndex = -1;
    let dropdownSelectCallback = null;

    function hideDropdown() {
        const dd = document.getElementById(TAG_DROPDOWN_ID);
        if (dd) dd.style.display = 'none';
        activeDropdownIndex = -1;
    }

    function showDropdown(items, rect, onSelect) {
        const dd = createDropdown();
        dd.innerHTML = '';
        dropdownSelectCallback = onSelect;
        activeDropdownIndex = items.length ? 0 : -1;

        if (!items.length) {
            hideDropdown();
            return;
        }

        items.forEach((tag, idx) => {
            const item = document.createElement('div');
            item.textContent = tag;
            item.style.padding = '4px 8px';
            item.style.cursor = 'pointer';
            item.dataset.index = idx.toString();

            item.addEventListener('mouseenter', () => {
                setActiveDropdownIndex(idx);
            });

            item.addEventListener('mousedown', (e) => {
                e.preventDefault(); // prevent editor blur
                if (dropdownSelectCallback) dropdownSelectCallback(tag);
            });

            dd.appendChild(item);
        });

        const top = Math.min(window.innerHeight - dd.offsetHeight - 10, rect.bottom + 4);
        const left = Math.min(window.innerWidth - 200, Math.max(0, rect.left));

        dd.style.top = `${top}px`;
        dd.style.left = `${left}px`;
        dd.style.display = 'block';
        setActiveDropdownIndex(0);
    }

    function setActiveDropdownIndex(idx) {
        const dd = document.getElementById(TAG_DROPDOWN_ID);
        if (!dd) return;
        const children = Array.from(dd.children);
        if (!children.length) return;

        // Clear all highlights
        children.forEach(ch => {
            ch.style.background = '';
        });

        activeDropdownIndex = idx;
        if (activeDropdownIndex >= 0 && children[activeDropdownIndex]) {
            children[activeDropdownIndex].style.background = '#555';
        }
    }

    function moveDropdownSelection(delta) {
        const dd = document.getElementById(TAG_DROPDOWN_ID);
        if (!dd) return;
        const children = Array.from(dd.children);
        if (!children.length) return;

        const len = children.length;
        let newIndex;
        if (activeDropdownIndex < 0) newIndex = 0;
        else newIndex = (activeDropdownIndex + delta + len) % len;

        setActiveDropdownIndex(newIndex);
    }

    function pickDropdownSelection() {
        const dd = document.getElementById(TAG_DROPDOWN_ID);
        if (!dd) return;
        const children = Array.from(dd.children);
        if (!children.length || activeDropdownIndex < 0) return;

        const tag = children[activeDropdownIndex].textContent;
        if (tag && dropdownSelectCallback) {
            dropdownSelectCallback(tag);
        }
    }

    // ---------- Editor helpers ----------
    function getActiveEditor() {
        const el = document.activeElement;
        if (!el) return null;
        if (el.tagName === 'TEXTAREA') return el;
        if (el.tagName === 'INPUT' && el.type === 'text') return el;
        if (el.isContentEditable) return el;
        return null;
    }

    function getCaretRectForContentEditable() {
        const sel = window.getSelection();
        if (!sel || !sel.rangeCount) return null;
        const range = sel.getRangeAt(0).cloneRange();
        if (!range.getClientRects) return null;
        range.collapse(true);
        const rects = range.getClientRects();
        if (rects.length) return rects[0];
        return null;
    }

    function getFallbackRect(el) {
        const r = el.getBoundingClientRect();
        return {
            top: r.bottom,
            bottom: r.bottom,
            left: r.left,
            right: r.left,
            width: 0,
            height: 0
        };
    }

    function getContentEditableRoot(node) {
        let n = node;
        while (n) {
            if (n.nodeType === Node.ELEMENT_NODE && n.isContentEditable) return n;
            n = n.parentNode;
        }
        return null;
    }

    function notifyContentEditableInput(root, dataText, inputType = 'insertText', phases) {
        if (!root) return { ok: true, prevented: false };
        const evtInit = { bubbles: true, cancelable: true, composed: true, inputType, data: dataText || '' };
        const targets = [root];
        if (document.activeElement && document.activeElement !== root) targets.push(document.activeElement);
        const toFire = phases && phases.length ? phases : ['beforeinput', 'input'];

        let ok = true;
        let prevented = false;
        targets.forEach(t => {
            toFire.forEach(type => {
                try {
                    const ev = new InputEvent(type, evtInit);
                    const res = t.dispatchEvent(ev);
                    ok = ok && res;
                    if (type === 'beforeinput' && res === false) prevented = true;
                } catch (e) {
                    ok = false;
                    if (type === 'input') {
                        t.dispatchEvent(new Event('input', { bubbles: true, cancelable: true, composed: true }));
                    }
                }
            });
        });
        return { ok, prevented };
    }

    // ---------- Tag state ----------
    let isTagMode = false;
    let currentTagText = '';
    let tagStartOffset = null; // textarea/input only

    function resetTagState() {
        isTagMode = false;
        currentTagText = '';
        tagStartOffset = null;
        hideDropdown();
    }

    function updateDropdownForCurrentTag(editor) {
        if (!isTagMode) {
            hideDropdown();
            return;
        }
        const allTags = loadTags();
        const q = currentTagText.toLowerCase();
        const matches = allTags
            .filter(t => t.toLowerCase().includes(q))
            .slice(0, 20);

        if (!matches.length) {
            hideDropdown();
            return;
        }

        let rect;
        if (editor && editor.isContentEditable) {
            rect = getCaretRectForContentEditable() || getFallbackRect(editor);
        } else if (editor) {
            rect = getFallbackRect(editor);
        } else {
            rect = { top: 40, bottom: 40, left: 40, right: 40, width: 0, height: 0 };
        }

        showDropdown(matches, rect, (tag) => {
            if (!editor) return;
            applyTagSelection(editor, tag);
        });
    }

    function applyTagSelection(editor, tag) {
        if (!editor) return;

        const insertText = '#' + tag + ' ';
        dbg('applyTagSelection', { tag, editorType: editor.tagName || 'contentEditable', isContentEditable: !!editor.isContentEditable });

        // textarea / input
        if (editor.tagName === 'TEXTAREA' || (editor.tagName === 'INPUT' && editor.type === 'text')) {
            const start = tagStartOffset;
            if (start == null) return;
            const end = editor.selectionStart;
            const before = editor.value.slice(0, start);
            const after = editor.value.slice(end);
            const newValue = before + insertText + after;
            const newCaret = before.length + insertText.length;
            editor.value = newValue;
            editor.selectionStart = editor.selectionEnd = newCaret;
            addTag(tag);
            resetTagState();
            return;
        }

        // contentEditable (Loop)
        if (editor.isContentEditable) {
            const sel = window.getSelection();
            if (!sel || !sel.rangeCount) {
                dbg('no selection range; bailing');
                addTag(tag);
                resetTagState();
                return;
            }

            const contentRoot = getContentEditableRoot(editor) || getContentEditableRoot(sel.anchorNode) || editor;
            let range = sel.getRangeAt(0);
            let container = range.endContainer;
            let offset = range.endOffset;
            dbg('initial caret', { containerType: container.nodeType, offset, textSnippet: container.textContent ? container.textContent.slice(0, 60) : '' });

            // Ensure we have a text node near the caret
            if (container.nodeType !== Node.TEXT_NODE) {
                if (container.childNodes && container.childNodes.length > 0 && offset > 0) {
                    let candidate = container.childNodes[offset - 1];
                    while (candidate && candidate.lastChild) candidate = candidate.lastChild;
                    if (candidate && candidate.nodeType === Node.TEXT_NODE) {
                        container = candidate;
                        offset = candidate.data.length;
                    }
                } else if (container.childNodes && container.childNodes[offset]) {
                    let candidate = container.childNodes[offset];
                    while (candidate && candidate.firstChild) candidate = candidate.firstChild;
                    if (candidate && candidate.nodeType === Node.TEXT_NODE) {
                        container = candidate;
                        offset = 0;
                    }
                }
                dbg('after node normalization', { containerType: container.nodeType, offset, textSnippet: container.textContent ? container.textContent.slice(0, 60) : '' });
            }

            if (container.nodeType !== Node.TEXT_NODE) {
                // Fallback: just insert at caret
                const textNode = document.createTextNode(insertText);
                range.insertNode(textNode);
                const newRange = document.createRange();
                newRange.setStartAfter(textNode);
                newRange.collapse(true);
                sel.removeAllRanges();
                sel.addRange(newRange);
                addTag(tag);
                dbg('fallback insertion path used');
                notifyContentEditableInput(getContentEditableRoot(textNode) || contentRoot, insertText);
                resetTagState();
                return;
            }

            const text = container.data;
            const beforeCaret = text.slice(0, offset);
            const afterCaret = text.slice(offset);
            const hashIndex = beforeCaret.lastIndexOf('#');

            // Build a backward range covering "#"+typed text based on the tracked tag text
            const targetLen = currentTagText.length + 1; // include '#'
            const rootForSearch = contentRoot || editor;
            const walker = document.createTreeWalker(rootForSearch, NodeFilter.SHOW_TEXT, null);
            const nodes = [];
            let n;
            while ((n = walker.nextNode())) nodes.push(n);
            const idx = nodes.indexOf(container);
            let startNode = container;
            let startOffset = offset;
            if (idx !== -1) {
                let remaining = targetLen;
                for (let i = idx; i >= 0 && remaining > 0; i--) {
                    const node = nodes[i];
                    const len = node.data.length;
                    if (i === idx) {
                        if (offset >= remaining) {
                            startNode = node;
                            startOffset = offset - remaining;
                            remaining = 0;
                        } else {
                            remaining -= offset;
                        }
                    } else {
                        if (len >= remaining) {
                            startNode = node;
                            startOffset = len - remaining;
                            remaining = 0;
                        } else {
                            remaining -= len;
                        }
                    }
                }
            }

            const replaceRange = document.createRange();
            replaceRange.setStart(startNode, startOffset);
            replaceRange.setEnd(container, offset);
            sel.removeAllRanges();
            sel.addRange(replaceRange);

            const beforeEvt = notifyContentEditableInput(contentRoot || editor, insertText, 'insertReplacementText', ['beforeinput']);
            if (beforeEvt.prevented) {
                dbg('beforeinput prevented by app; skipping manual replace');
                addTag(tag);
                resetTagState();
                return;
            }

            const newTextNode = document.createTextNode(insertText);
            replaceRange.deleteContents();
            replaceRange.insertNode(newTextNode);

            const newRange = document.createRange();
            newRange.setStartAfter(newTextNode);
            newRange.collapse(true);
            sel.removeAllRanges();
            sel.addRange(newRange);

            notifyContentEditableInput(getContentEditableRoot(newTextNode) || contentRoot || editor, insertText, 'insertReplacementText', ['input']);

            addTag(tag);
            resetTagState();
        }
    }

    // ---------- Tag Manager UI (button-only, with search) ----------
    function createTagManager() {
        let mgr = document.getElementById(TAG_MANAGER_ID);
        if (mgr) return mgr;

        mgr = document.createElement('div');
        mgr.id = TAG_MANAGER_ID;
        mgr.style.position = 'fixed';
        mgr.style.zIndex = '999999';
        mgr.style.top = '50%';
        mgr.style.left = '50%';
        mgr.style.transform = 'translate(-50%, -50%)';
        mgr.style.background = '#222';
        mgr.style.color = '#fff';
        mgr.style.border = '1px solid #555';
        mgr.style.borderRadius = '6px';
        mgr.style.padding = '12px';
        mgr.style.minWidth = '260px';
        mgr.style.maxWidth = '420px';
        mgr.style.maxHeight = '70vh';
        mgr.style.overflowY = 'auto';
        mgr.style.boxShadow = '0 6px 18px rgba(0,0,0,0.5)';
        mgr.style.fontFamily = 'system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif';
        mgr.style.display = 'none';

        const title = document.createElement('div');
        title.textContent = 'Tag Manager';
        title.style.fontWeight = '600';
        title.style.marginBottom = '4px';

        const info = document.createElement('div');
        info.textContent = 'Click ❌ to delete a tag. Esc or Close to dismiss.';
        info.style.fontSize = '11px';
        info.style.opacity = '0.8';
        info.style.marginBottom = '8px';

        const searchWrapper = document.createElement('div');
        searchWrapper.style.marginBottom = '6px';

        const searchInput = document.createElement('input');
        searchInput.id = TAG_MANAGER_ID + '-search';
        searchInput.type = 'text';
        searchInput.placeholder = 'Search tags...';
        searchInput.style.width = '100%';
        searchInput.style.boxSizing = 'border-box';
        searchInput.style.borderRadius = '4px';
        searchInput.style.border = '1px solid #555';
        searchInput.style.background = '#111';
        searchInput.style.color = '#eee';
        searchInput.style.fontSize = '11px';
        searchInput.style.padding = '4px 6px';
        searchInput.style.fontFamily = 'system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif';

        searchWrapper.appendChild(searchInput);

        const listLabel = document.createElement('div');
        listLabel.textContent = 'Current tags:';
        listLabel.style.fontSize = '12px';
        listLabel.style.margin = '4px 0 2px 0';

        const list = document.createElement('div');
        list.id = TAG_MANAGER_ID + '-list';
        list.style.marginBottom = '8px';

        const ioLabel = document.createElement('div');
        ioLabel.textContent = 'Import / Export:';
        ioLabel.style.fontSize = '12px';
        ioLabel.style.marginTop = '4px';
        ioLabel.style.marginBottom = '2px';

        const ioHelp = document.createElement('div');
        ioHelp.textContent = 'Export -> copy, Import -> paste newline or JSON ["tag1","tag2"].';
        ioHelp.style.fontSize = '11px';
        ioHelp.style.opacity = '0.7';
        ioHelp.style.marginBottom = '4px';

        const ioArea = document.createElement('textarea');
        ioArea.id = TAG_MANAGER_ID + '-io';
        ioArea.rows = 4;
        ioArea.style.width = '100%';
        ioArea.style.boxSizing = 'border-box';
        ioArea.style.background = '#111';
        ioArea.style.color = '#eee';
        ioArea.style.border = '1px solid #555';
        ioArea.style.borderRadius = '4px';
        ioArea.style.fontSize = '11px';
        ioArea.style.fontFamily = 'monospace';
        ioArea.style.padding = '4px';
        ioArea.placeholder = 'Export: click "Export". Import: paste tags then click "Import".';

        const ioButtonsRow = document.createElement('div');
        ioButtonsRow.style.display = 'flex';
        ioButtonsRow.style.justifyContent = 'flex-end';
        ioButtonsRow.style.gap = '6px';
        ioButtonsRow.style.marginTop = '4px';
        ioButtonsRow.style.marginBottom = '6px';

        const exportBtn = document.createElement('button');
        exportBtn.textContent = 'Export';
        exportBtn.style.padding = '3px 10px';
        exportBtn.style.borderRadius = '4px';
        exportBtn.style.border = '1px solid #777';
        exportBtn.style.background = '#333';
        exportBtn.style.color = '#fff';
        exportBtn.style.cursor = 'pointer';
        exportBtn.style.fontSize = '11px';

        exportBtn.addEventListener('click', () => {
            const txt = exportTagsAsText();
            ioArea.value = txt;
            if (navigator.clipboard && navigator.clipboard.writeText) {
                navigator.clipboard.writeText(txt).then(() => {
                    exportBtn.textContent = 'Copied!';
                    setTimeout(() => (exportBtn.textContent = 'Export'), 800);
                }).catch(() => {});
            }
        });

        const importBtn = document.createElement('button');
        importBtn.textContent = 'Import';
        importBtn.style.padding = '3px 10px';
        importBtn.style.borderRadius = '4px';
        importBtn.style.border = '1px solid #777';
        importBtn.style.background = '#333';
        importBtn.style.color = '#fff';
        importBtn.style.cursor = 'pointer';
        importBtn.style.fontSize = '11px';

        importBtn.addEventListener('click', () => {
            importTagsFromString(ioArea.value);
            renderTagManager();
            importBtn.textContent = 'Imported!';
            setTimeout(() => (importBtn.textContent = 'Import'), 800);
        });

        ioButtonsRow.appendChild(exportBtn);
        ioButtonsRow.appendChild(importBtn);

        const closeBtn = document.createElement('button');
        closeBtn.textContent = 'Close';
        closeBtn.style.marginTop = '4px';
        closeBtn.style.padding = '4px 10px';
        closeBtn.style.borderRadius = '4px';
        closeBtn.style.border = '1px solid #777';
        closeBtn.style.background = '#333';
        closeBtn.style.color = '#fff';
        closeBtn.style.cursor = 'pointer';
        closeBtn.addEventListener('click', () => {
            hideTagManager();
        });

        mgr.appendChild(title);
        mgr.appendChild(info);
        mgr.appendChild(searchWrapper);
        mgr.appendChild(listLabel);
        mgr.appendChild(list);
        mgr.appendChild(ioLabel);
        mgr.appendChild(ioHelp);
        mgr.appendChild(ioArea);
        mgr.appendChild(ioButtonsRow);
        mgr.appendChild(closeBtn);

        document.body.appendChild(mgr);

        searchInput.addEventListener('input', () => {
            renderTagManager();
        });

        return mgr;
    }

    function getTagFilterText() {
        const search = document.getElementById(TAG_MANAGER_ID + '-search');
        return (search && search.value) ? search.value.toLowerCase() : '';
    }

    function renderTagManager() {
        const mgr = createTagManager();
        const list = document.getElementById(TAG_MANAGER_ID + '-list');
        if (!list) return;

        list.innerHTML = '';

        const filter = getTagFilterText();
        let tags = loadTags();
        if (filter) {
            tags = tags.filter(t => t.toLowerCase().includes(filter));
        }

        if (!tags.length) {
            const empty = document.createElement('div');
            empty.textContent = 'No tags match. Create some by typing #tag in Loop.';
            empty.style.fontSize = '12px';
            list.appendChild(empty);
            return;
        }

        tags.forEach(tag => {
            const row = document.createElement('div');
            row.style.display = 'flex';
            row.style.alignItems = 'center';
            row.style.justifyContent = 'space-between';
            row.style.fontSize = '12px';
            row.style.padding = '2px 0';

            const span = document.createElement('span');
            span.textContent = '#' + tag;

            const del = document.createElement('button');
            del.textContent = '❌';
            del.style.border = 'none';
            del.style.background = 'transparent';
            del.style.cursor = 'pointer';
            del.style.fontSize = '12px';
            del.style.marginLeft = '8px';
            del.title = `Delete tag #${tag}`;
            del.setAttribute('aria-label', `Delete tag #${tag}`);

            del.addEventListener('click', () => {
                removeTag(tag);
                renderTagManager();
            });

            row.appendChild(span);
            row.appendChild(del);
            list.appendChild(row);
        });
    }

    function showTagManager() {
        renderTagManager();
        const mgr = createTagManager();
        mgr.style.display = 'block';
    }

    function hideTagManager() {
        const mgr = document.getElementById(TAG_MANAGER_ID);
        if (mgr) mgr.style.display = 'none';
    }

    function createTagButton() {
        let btn = document.getElementById(TAG_BUTTON_ID);
        if (btn) return btn;

        btn = document.createElement('button');
        btn.id = TAG_BUTTON_ID;
        btn.textContent = '# Tags';
        btn.title = 'Open Tag Manager';
        btn.style.position = 'fixed';
        btn.style.right = '10px';
        btn.style.bottom = '10px';
        btn.style.zIndex = '999999';
        btn.style.padding = '4px 10px';
        btn.style.fontSize = '11px';
        btn.style.borderRadius = '16px';
        btn.style.border = '1px solid #777';
        btn.style.background = 'rgba(30,30,30,0.85)';
        btn.style.color = '#fff';
        btn.style.cursor = 'pointer';
        btn.style.fontFamily = 'system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif';
        btn.addEventListener('click', showTagManager);
        document.body.appendChild(btn);
        return btn;
    }

    // ---------- Keyboard handling (capture phase) ----------
    function handleKeydown(e) {
        if (e._tmTagHandler) return;
        e._tmTagHandler = true;
        const key = e.key;
        const keyCode = e.keyCode || e.which;
        const dd = document.getElementById(TAG_DROPDOWN_ID);
        const dropdownVisible = dd && dd.style.display !== 'none';
        const mgr = document.getElementById(TAG_MANAGER_ID);
        const managerVisible = mgr && mgr.style.display === 'block';

        const isDownKey = key === 'ArrowDown' || key === 'Down' || keyCode === 40;
        const isUpKey = key === 'ArrowUp' || key === 'Up' || keyCode === 38;
        const isEnterLike = key === 'Enter' || key === 'Return' || keyCode === 13;
        const isSpaceLike = key === ' ' || keyCode === 32;
        const isTabLike = key === 'Tab' || keyCode === 9;

        if (managerVisible && key === 'Escape') {
            e.preventDefault();
            hideTagManager();
            return;
        }

        // Dropdown navigation / selection
        if (dropdownVisible && (isDownKey || isUpKey)) {
            e.preventDefault();
            e.stopPropagation();
            moveDropdownSelection(isDownKey ? 1 : -1);
            return;
        }

        if (dropdownVisible && (isEnterLike || isTabLike || isSpaceLike)) {
            e.preventDefault();
            e.stopPropagation();
            pickDropdownSelection();
            return;
        }

        if (dropdownVisible && key === 'Escape') {
            e.preventDefault();
            e.stopPropagation();
            resetTagState();
            return;
        }

        const editor = getActiveEditor();
        if (!editor) return;

        // Start tag mode on '#'
        if (key === '#') {
            isTagMode = true;
            currentTagText = '';

            if (editor.tagName === 'TEXTAREA' || (editor.tagName === 'INPUT' && editor.type === 'text')) {
                tagStartOffset = editor.selectionStart;
            } else {
                tagStartOffset = null; // not used for contentEditable
            }

            setTimeout(() => updateDropdownForCurrentTag(editor), 0);
            return;
        }

        if (!isTagMode) return;

        const isPunctuation = /[.,;:!?]/.test(key);

        // Finish tag (space / enter / tab / punctuation)
        if (isEnterLike || isSpaceLike || isTabLike || isPunctuation) {
            e.preventDefault();
            e.stopPropagation();

            const trimmed = currentTagText.trim();
            if (trimmed.length > 0) {
                // If this is a SPACE without dropdown, treat it like selecting our own tag
                if (isSpaceLike && !dropdownVisible) {
                    applyTagSelection(editor, trimmed);
                    // applyTagSelection already calls addTag + resetTagState
                    return;
                } else {
                    // Enter / Tab / punctuation: just remember the tag
                    addTag(trimmed);
                }
            }
            resetTagState();
            return;
        }

        if (key === 'Backspace') {
            if (currentTagText.length > 0) {
                currentTagText = currentTagText.slice(0, -1);
                setTimeout(() => updateDropdownForCurrentTag(editor), 0);
            } else {
                resetTagState();
            }
            return;
        }

        // Normal character while typing tag
        if (key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
            currentTagText += key;
            setTimeout(() => updateDropdownForCurrentTag(editor), 0);
        }
    }

    window.addEventListener('keydown', handleKeydown, true); // capture phase

    // Hide dropdown on click outside
    document.addEventListener('click', (e) => {
        const dd = document.getElementById(TAG_DROPDOWN_ID);
        if (dd && !dd.contains(e.target)) {
            hideDropdown();
        }

        const mgr = document.getElementById(TAG_MANAGER_ID);
        const mgrVisible = mgr && mgr.style.display === 'block';
        const btn = document.getElementById(TAG_BUTTON_ID);
        const clickedButton = btn && btn.contains(e.target);
        const path = typeof e.composedPath === 'function' ? e.composedPath() : [];
        const clickInsideMgr = mgr && (mgr.contains(e.target) || path.includes(mgr));
        if (mgrVisible && !clickInsideMgr && !clickedButton) {
            hideTagManager();
        }
    });

    // Create floating tag button on load
    createTagButton();

})();
