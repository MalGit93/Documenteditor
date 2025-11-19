(function() {
  const editorRoot = document.getElementById('editor-root');
  const blockActions = document.getElementById('block-actions');
  const textToolbar = document.getElementById('text-toolbar');
  const snackbar = document.getElementById('snackbar');
  const modalBackdrop = document.getElementById('modal-backdrop');
  const modalInner = document.getElementById('modal-inner');
  const jsonInput = document.getElementById('json-file-input');

  const DELETE_ICON_HTML = '<span class="delete-icon" contenteditable="false" role="button" tabindex="0" aria-label="Delete block">×</span>';
  const DRAFT_STORAGE_KEY = 'igaEbdDraft';
  const QUALITY_PLACEHOLDERS = [
    'New paragraph…',
    'Step 1…',
    'Step 2…',
    'Step 3…',
    'Q: …',
    'A: …',
    'Current state…',
    'Target state…',
    'Outcome / benefit…',
    'Who this is for…',
    'What we need them to do…',
    'Key dates…',
    'Add an explanatory caption…',
    'Your footnote text here.'
  ];

  const INTERNAL_MARKER_START = '<!--EBD_INTERNAL_START-->';
  const INTERNAL_MARKER_END = '<!--EBD_INTERNAL_END-->';

  let internalClipboardHTML = '';
  let currentTextRange = null;

  let undoStack = [];
  let redoStack = [];
  let lastSerialized = null;
  let inputSnapshotTimer = null;
  let suppressSnapshots = false;

  // For column resizing
  let tableResizeState = null;

  function persistDraft(serialized) {
    try {
      localStorage.setItem(DRAFT_STORAGE_KEY, serialized);
    } catch (err) {
      /* noop */
    }
  }

  function loadDraft() {
    try {
      return localStorage.getItem(DRAFT_STORAGE_KEY);
    } catch (err) {
      return null;
    }
  }

  function clearDraft() {
    try {
      localStorage.removeItem(DRAFT_STORAGE_KEY);
    } catch (err) {
      /* noop */
    }
  }

  function promptRestoreDraft() {
    const saved = loadDraft();
    if (!saved) return;
    if (lastSerialized && saved === lastSerialized) return;
    openModal({
      title: 'Restore draft?',
      body: '<p>We found a saved draft from your last session. Restore it?</p>',
      confirmLabel: 'Restore',
      cancelLabel: 'Discard',
      onConfirm: () => {
        closeModal();
        applyState(saved);
        showSnackbar('Draft restored');
      },
      onCancel: () => {
        clearDraft();
        closeModal();
      }
    });
  }

  function getCurrentTheme() {
    return document.body.classList.contains('theme-ft') ? 'ft' : 'iga';
  }

  function setTheme(theme) {
    document.body.classList.remove('theme-iga', 'theme-ft');
    if (theme === 'ft') {
      document.body.classList.add('theme-ft');
      document.getElementById('theme-select').value = 'ft';
    } else {
      document.body.classList.add('theme-iga');
      document.getElementById('theme-select').value = 'iga';
    }
  }

  function ensureFootnotesPage() {
    let footnotes = editorRoot.querySelector('.footnotes-page');
    if (!footnotes) {
      footnotes = document.createElement('div');
      footnotes.className = 'page-container footnotes-page';
      footnotes.dataset.pageType = 'footnotes';
      footnotes.innerHTML = `
        <div class="page-overlay"></div>
        <div class="page-body">
          <h3 class="section-heading"><span contenteditable="false">Footnotes</span></h3>
          <ol id="footnote-list"></ol>
        </div>
        <div class="page-footer">Footnotes</div>
      `;
      editorRoot.appendChild(footnotes);
    } else {
      editorRoot.appendChild(footnotes);
    }

    // Ensure existing lis are editable after reload/import
    const list = footnotes.querySelector('#footnote-list');
    if (list) {
      list.querySelectorAll('li').forEach(li => {
        if (!li.hasAttribute('contenteditable')) {
          li.contentEditable = 'true';
        }
      });
    }
  }

  function updatePageNumbers() {
    const pages = editorRoot.querySelectorAll('.page-container');
    let pageIndex = 1;
    pages.forEach(page => {
      const footer = page.querySelector('.page-footer');
      if (!footer) return;
      if (page.classList.contains('footnotes-page')) {
        footer.textContent = 'Footnotes';
      } else {
        footer.textContent = 'Page ' + pageIndex;
        page.dataset.pageNumber = String(pageIndex);
        pageIndex++;
      }
    });
  }

  function showSnackbar(message, duration = 2000) {
    snackbar.textContent = message;
    snackbar.classList.add('show');
    setTimeout(() => snackbar.classList.remove('show'), duration);
  }

  function clearBlockSelection() {
    editorRoot.querySelectorAll('.block-selected').forEach(el => el.classList.remove('block-selected'));
    hideBlockActions();
  }

  function getSelectedBlocks() {
    return Array.from(editorRoot.querySelectorAll('.block-selected')).filter(el =>
      el.closest('.page-container') &&
      !el.closest('.footnotes-page') &&
      el.dataset.deletable === 'true'
    );
  }

  function getLastSelectedBlock() {
    const blocks = getSelectedBlocks();
    return blocks.length ? blocks[blocks.length - 1] : null;
  }

  function isInEditor(node) {
    return node && editorRoot.contains(node);
  }

  function getActiveContentPageBody() {
    const sel = window.getSelection();
    if (sel && sel.rangeCount) {
      const node = sel.anchorNode;
      if (node) {
        const el = node.nodeType === 1 ? node : node.parentElement;
        const page = el && el.closest('.page-container');
        if (page && !page.classList.contains('footnotes-page')) {
          return page.querySelector('.page-body');
        }
      }
    }
    const pages = editorRoot.querySelectorAll('.page-container:not(.footnotes-page)');
    const last = pages[pages.length - 1];
    return last ? last.querySelector('.page-body') : null;
  }

  function cloneNodeWithoutEditorChrome(node) {
    const clone = node.cloneNode(true);
    clone.querySelectorAll('.delete-icon, .overflow-indicator, .page-controls').forEach(el => el.remove());
    if (clone.classList) {
      clone.classList.remove('block-selected');
    }
    clone.querySelectorAll('.block-selected').forEach(el => el.classList.remove('block-selected'));
    return clone;
  }

  function buildBlockData(blockEl) {
    if (!blockEl) return null;
    const clone = cloneNodeWithoutEditorChrome(blockEl);
    const dataset = { ...blockEl.dataset };
    return {
      tag: (blockEl.tagName || 'div').toLowerCase(),
      classes: Array.from(blockEl.classList || []).filter(cls => cls !== 'block-selected'),
      dataset,
      contentEditable: blockEl.getAttribute && blockEl.getAttribute('contenteditable') === 'true',
      html: clone.innerHTML.trim()
    };
  }

  function buildDocumentModel() {
    const pages = [];
    editorRoot.querySelectorAll('.page-container').forEach(page => {
      const pageType = page.dataset.pageType || 'content';
      if (pageType === 'footnotes') {
        const list = page.querySelector('#footnote-list');
        const footnotes = [];
        if (list) {
          list.querySelectorAll('li').forEach(li => {
            const clone = cloneNodeWithoutEditorChrome(li);
            clone.querySelectorAll('.footnote-return').forEach(btn => btn.remove());
            footnotes.push({
              dataset: { ...li.dataset },
              html: clone.innerHTML.trim()
            });
          });
        }
        pages.push({ type: pageType, footnotes });
        return;
      }
      const body = page.querySelector('.page-body');
      const blocks = [];
      if (body) {
        body.querySelectorAll('[data-deletable="true"]').forEach(blockEl => {
          const data = buildBlockData(blockEl);
          if (data) blocks.push(data);
        });
      }
      pages.push({ type: pageType, blocks });
    });
    return { pages };
  }

  function ensureBlockVisible(block) {
    if (!block) return;
    block.scrollIntoView({ behavior: 'smooth', block: 'center' });
  }

  function serializeState() {
    const theme = getCurrentTheme();
    clearBlockSelection();
    hideTextToolbar();
    hideBlockActions();
    const html = editorRoot.innerHTML;
    const model = buildDocumentModel();
    return JSON.stringify({ theme, html, model });
  }

  function checkPageOverflowState(pageBody) {
    if (!pageBody) return;
    const indicator = pageBody.querySelector('.overflow-indicator');
    if (!indicator) return;
    const contentHeight = pageBody.scrollHeight;
    const visibleHeight = pageBody.clientHeight;
    const delta = contentHeight - visibleHeight;
    const styles = getComputedStyle(pageBody);
    const scaleValue = parseFloat(styles.getPropertyValue('--page-scale')) || 1;
    const buffer = Math.max(24, 48 * scaleValue);
    const warn = delta > -buffer;
    if (warn) {
      pageBody.dataset.overflow = 'true';
      const pill = indicator.querySelector('span');
      if (pill) {
        const message = delta > 0
          ? 'Content exceeds page height. Consider splitting page.'
          : 'Content nearing page limit. Consider splitting page.';
        pill.title = message;
        pill.setAttribute('aria-label', message);
      }
    } else {
      pageBody.dataset.overflow = 'false';
    }
  }

  function checkAllPageOverflow() {
    editorRoot.querySelectorAll('.page-container:not(.footnotes-page) .page-body').forEach(checkPageOverflowState);
  }

  function captureSnapshot() {
    if (suppressSnapshots) return;
    if (lastSerialized === null) {
      lastSerialized = serializeState();
      return;
    }
    undoStack.push(lastSerialized);
    if (undoStack.length > 50) undoStack.shift();
    redoStack.length = 0;
  }

  function afterChange() {
    if (suppressSnapshots) return;
    repairOrphanContentBlocks();
    lastSerialized = serializeState();
    checkAllPageOverflow();
    persistDraft(lastSerialized);
  }

  function applyState(serialized) {
    suppressSnapshots = true;
    try {
      const state = JSON.parse(serialized);
      setTheme(state.theme || 'iga');
      if (state.model && Array.isArray(state.model.pages)) {
        renderDocumentModel(state.model);
      } else {
        editorRoot.innerHTML = state.html || '';
        ensureFootnotesPage();
        updatePageNumbers();
        renumberFootnotes();
        initAllTableResizers();
        repairOrphanContentBlocks();
        checkAllPageOverflow();
      }
    } catch (e) {
      console.error(e);
      showSnackbar('Failed to apply state');
    }
    suppressSnapshots = false;
    lastSerialized = serializeState();
    persistDraft(lastSerialized);
  }

  function undo() {
    if (!undoStack.length) {
      showSnackbar('Nothing to undo');
      return;
    }
    const current = serializeState();
    const prev = undoStack.pop();
    redoStack.push(current);
    applyState(prev);
  }

  function redo() {
    if (!redoStack.length) {
      showSnackbar('Nothing to redo');
      return;
    }
    const current = serializeState();
    const next = redoStack.pop();
    undoStack.push(current);
    applyState(next);
  }

  function initHistory() {
    undoStack = [];
    redoStack = [];
    lastSerialized = serializeState();
    checkAllPageOverflow();
    persistDraft(lastSerialized);
  }

  /* -------- Footnotes: debounced auto-renumber + observer -------- */

  let footnoteMutationTimer = null;

  function scheduleFootnoteRenumber() {
    if (footnoteMutationTimer) return;
    footnoteMutationTimer = setTimeout(() => {
      renumberFootnotes();
      afterChange();
      footnoteMutationTimer = null;
    }, 80);
  }

  const footnoteObserver = new MutationObserver(mutations => {
    let changed = false;

    for (const m of mutations) {
      if (m.type !== 'childList') continue;

      // Look at removed nodes
      m.removedNodes.forEach(node => {
        if (node.nodeType !== 1) return;

        // If a footnote ref or footnote li itself was removed
        if (node.matches && (node.matches('.footnote-ref') || node.matches('#footnote-list li'))) {
          changed = true;
          return;
        }

        // Or if a container containing them was removed
        if (
          node.querySelector &&
          (node.querySelector('.footnote-ref') || node.querySelector('#footnote-list li'))
        ) {
          changed = true;
        }
      });

      // Direct removals from the footnote list
      if (!changed && m.target && m.target.id === 'footnote-list' && m.removedNodes.length) {
        changed = true;
      }
    }

    if (changed) {
      scheduleFootnoteRenumber();
    }
  });

  footnoteObserver.observe(editorRoot, {
    childList: true,
    subtree: true
  });

  /* -------- Block actions UI -------- */

  function positionBlockActions() {
    const last = getLastSelectedBlock();
    if (!last) {
      hideBlockActions();
      return;
    }
    const rect = last.getBoundingClientRect();
    blockActions.style.left = (rect.left + rect.width / 2 + window.scrollX) + 'px';
    blockActions.style.top = (rect.top + window.scrollY - 8) + 'px';
    blockActions.classList.add('visible');
  }

  function showBlockActions() {
    if (getSelectedBlocks().length) {
      positionBlockActions();
    } else {
      hideBlockActions();
    }
  }

  function hideBlockActions() {
    blockActions.classList.remove('visible');
  }

  /* -------- Text toolbar / selection -------- */

  function selectionWithinEditor() {
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return false;
    const range = sel.getRangeAt(0);
    return isInEditor(range.commonAncestorContainer);
  }

  function isToolbarActiveElement() {
    const active = document.activeElement;
    return active && textToolbar.contains(active);
  }

  function updateTextToolbar() {
    if (!selectionWithinEditor()) {
      if (isToolbarActiveElement()) return;
      hideTextToolbar();
      return;
    }
    const sel = window.getSelection();
    if (sel.isCollapsed) {
      if (isToolbarActiveElement()) return;
      hideTextToolbar();
      return;
    }
    currentTextRange = sel.getRangeAt(0).cloneRange();
    const rect = currentTextRange.getBoundingClientRect();
    if (!rect || (!rect.width && !rect.height)) {
      hideTextToolbar();
      return;
    }
    textToolbar.style.left = (rect.left + rect.width / 2 + window.scrollX) + 'px';
    textToolbar.style.top = (rect.top + window.scrollY - 6) + 'px';
    textToolbar.classList.add('visible');
  }

  function hideTextToolbar() {
    textToolbar.classList.remove('visible');
    currentTextRange = null;
    const linkInput = document.getElementById('link-input');
    if (linkInput) linkInput.value = '';
  }

  function restoreSelection() {
    if (!currentTextRange) return;
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(currentTextRange);
  }

  function focusEditableFromRange(range) {
    if (!range) return;
    let node = range.startContainer;
    if (node && node.nodeType === Node.TEXT_NODE) {
      node = node.parentElement;
    }
    if (!node || !(node instanceof Element)) return;
    const editable = node.closest('[contenteditable="true"]');
    if (editable && typeof editable.focus === 'function') {
      try {
        editable.focus({ preventScroll: true });
      } catch (err) {
        editable.focus();
      }
    }
  }

  function execWithinSelection(callback, { hideToolbar = false } = {}) {
    if (!currentTextRange) return;
    focusEditableFromRange(currentTextRange);
    restoreSelection();
    callback();
    const sel = window.getSelection();
    if (sel && sel.rangeCount) {
      currentTextRange = sel.getRangeAt(0).cloneRange();
    } else {
      currentTextRange = null;
    }
    scheduleInputSnapshot();
    if (hideToolbar) hideTextToolbar();
    else updateTextToolbar();
  }

  function applyInlineCmd(cmd) {
    execWithinSelection(() => document.execCommand(cmd, false, null));
  }

  function toggleHighlight() {
    const command = document.queryCommandSupported('hiliteColor') ? 'hiliteColor' : 'backColor';
    execWithinSelection(() => {
      let currentValue = '';
      try {
        currentValue = document.queryCommandValue(command) || '';
      } catch (err) {
        currentValue = '';
      }
      const lower = typeof currentValue === 'string' ? currentValue.toLowerCase() : '';
      if (lower.includes('255') && lower.includes('243')) {
        document.execCommand(command, false, 'transparent');
      } else {
        document.execCommand(command, false, '#fff3b0');
      }
    }, { hideToolbar: true });
  }

  function applyParagraphStyle(style) {
    if (!currentTextRange) return;
    focusEditableFromRange(currentTextRange);
    restoreSelection();
    const sel = window.getSelection();
    if (!sel.rangeCount) return;
    const range = sel.getRangeAt(0);
    let el = range.startContainer;
    if (el.nodeType === Node.TEXT_NODE) el = el.parentElement;
    if (!el) return;
    let block = el.closest('.body-text, .section-heading, h1, h2, h3, p, .doc-title, .doc-subtitle');
    if (!block) block = el;

    function transform(tagName, className) {
      if (block.tagName && block.tagName.toLowerCase() === tagName.toLowerCase()) {
        block.className = className;
        return;
      }
      const replacement = document.createElement(tagName);
      replacement.className = className;
      replacement.innerHTML = block.innerHTML;
      block.parentNode.replaceChild(replacement, block);
    }

    if (style === 'h1') transform('h1', 'doc-title');
    else if (style === 'h2') transform('h2', 'section-heading');
    else if (style === 'h3') transform('h3', 'section-heading');
    else if (style === 'p') transform('p', 'body-text');

    scheduleInputSnapshot();
    hideTextToolbar();
  }

  function createLink(url) {
    if (!currentTextRange) return;
    if (!url) {
      hideTextToolbar();
      return;
    }
    if (!/^https?:\/\//i.test(url)) url = 'https://' + url;
    execWithinSelection(() => {
      document.execCommand('unlink');
      document.execCommand('createLink', false, url);
    }, { hideToolbar: true });
  }

  function removeLink() {
    if (!currentTextRange) return;
    execWithinSelection(() => document.execCommand('unlink'), { hideToolbar: true });
  }

  /* -------- Footnotes add + renumber -------- */

  function createFootnoteRefElement(id) {
    const sup = document.createElement('sup');
    sup.className = 'footnote-ref';
    sup.dataset.uniqueId = id;
    sup.contentEditable = 'false';
    sup.tabIndex = 0;
    sup.setAttribute('role', 'button');
    sup.setAttribute('aria-label', 'Jump to footnote');
    sup.textContent = '[?]';
    return sup;
  }

  function decorateFootnoteListItem(li, num) {
    if (!li) return;
    li.contentEditable = 'true';
    li.setAttribute('data-number', String(num));
    li.id = 'fn-note-' + num;
    let returnBtn = li.querySelector('.footnote-return');
    if (!returnBtn) {
      returnBtn = document.createElement('button');
      returnBtn.type = 'button';
      returnBtn.className = 'footnote-return';
      returnBtn.contentEditable = 'false';
      returnBtn.textContent = '↩';
      li.appendChild(returnBtn);
    }
    returnBtn.dataset.uniqueId = li.dataset.uniqueId;
    returnBtn.tabIndex = 0;
    returnBtn.setAttribute('aria-label', 'Jump to reference ' + num);
  }

  function cleanupOrphanFootnoteTargets() {
    editorRoot.querySelectorAll('.footnote-target').forEach(wrapper => {
      if (!wrapper.querySelector('.footnote-ref')) {
        while (wrapper.firstChild) {
          wrapper.parentNode.insertBefore(wrapper.firstChild, wrapper);
        }
        wrapper.remove();
      }
    });
  }

  function flashElement(el) {
    if (!el) return;
    el.classList.add('footnote-highlight');
    setTimeout(() => el.classList.remove('footnote-highlight'), 1200);
  }

  function scrollToFootnoteNote(uniqueId) {
    if (!uniqueId) return;
    const note = editorRoot.querySelector(`#footnote-list li[data-unique-id="${uniqueId}"]`);
    if (note) {
      note.scrollIntoView({ behavior: 'smooth', block: 'center' });
      flashElement(note);
      setTimeout(() => {
        try { note.focus(); } catch (err) { /* ignore */ }
      }, 200);
    }
  }

  function scrollToFootnoteRef(uniqueId) {
    if (!uniqueId) return;
    const ref = editorRoot.querySelector(`.footnote-ref[data-unique-id="${uniqueId}"]`);
    if (ref) {
      const wrapper = ref.closest('.footnote-target') || ref;
      wrapper.scrollIntoView({ behavior: 'smooth', block: 'center' });
      flashElement(wrapper);
      const range = document.createRange();
      range.setStartAfter(ref);
      range.setEndAfter(ref);
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
      currentTextRange = range.cloneRange();
    }
  }

  function addFootnote() {
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) {
      showSnackbar('Place the caret in body content first');
      return;
    }
    const range = sel.getRangeAt(0);
    const container = range.commonAncestorContainer.nodeType === 1
      ? range.commonAncestorContainer
      : range.commonAncestorContainer.parentElement;
    const page = container && container.closest('.page-container');
    if (!page || page.classList.contains('footnotes-page')) {
      showSnackbar('Footnotes can only be added in body pages');
      return;
    }

    captureSnapshot();
    const id = 'fn-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 8);
    const sup = createFootnoteRefElement(id);
    let seedText = 'Your footnote text here.';
    if (!range.collapsed) {
      const fragment = range.cloneContents();
      const text = fragment.textContent ? fragment.textContent.replace(/\s+/g, ' ').trim() : '';
      if (text) seedText = text;
      const wrapper = document.createElement('span');
      wrapper.className = 'footnote-target';
      wrapper.dataset.uniqueId = id;
      wrapper.appendChild(range.extractContents());
      wrapper.appendChild(sup);
      range.insertNode(wrapper);
      const after = document.createRange();
      after.setStartAfter(wrapper);
      after.setEndAfter(wrapper);
      sel.removeAllRanges();
      sel.addRange(after);
    } else {
      range.collapse(false);
      range.insertNode(sup);
      const after = document.createRange();
      after.setStartAfter(sup);
      after.setEndAfter(sup);
      sel.removeAllRanges();
      sel.addRange(after);
    }

    const list = editorRoot.querySelector('.footnotes-page #footnote-list');
    const li = document.createElement('li');
    li.dataset.uniqueId = id;
    li.contentEditable = 'true';
    li.textContent = seedText;
    list.appendChild(li);

    renumberFootnotes();
    afterChange();
    flashElement(sup);
  }

  function renumberFootnotes() {
    const footnotesPage = editorRoot.querySelector('.footnotes-page');
    if (!footnotesPage) return;
    const list = footnotesPage.querySelector('#footnote-list');
    if (!list) return;

    // Collect all refs from content pages
    const refNodes = [];
    const contentPages = editorRoot.querySelectorAll('.page-container:not(.footnotes-page)');
    contentPages.forEach(page => {
      page.querySelectorAll('.footnote-ref').forEach(ref => refNodes.push(ref));
    });

    // Map of existing lis by unique id
    const liMap = new Map();
    list.querySelectorAll('li[data-unique-id]').forEach(li => {
      liMap.set(li.dataset.uniqueId, li);
    });

    // If a li was deleted in the footnotes list, drop its refs
    refNodes.slice().forEach(ref => {
      if (!liMap.has(ref.dataset.uniqueId)) {
        const wrapper = ref.parentElement;
        ref.remove();
        if (wrapper && wrapper.classList && wrapper.classList.contains('footnote-target')) {
          cleanupOrphanFootnoteTargets();
        }
      }
    });

    // Re-scan (some refs may have been removed)
    const validRefIds = new Set();
    const updatedRefNodes = [];
    contentPages.forEach(page => {
      page.querySelectorAll('.footnote-ref').forEach(ref => {
        updatedRefNodes.push(ref);
        validRefIds.add(ref.dataset.uniqueId);
      });
    });

    // Remove any lis that no longer have a live ref
    list.querySelectorAll('li[data-unique-id]').forEach(li => {
      if (!validRefIds.has(li.dataset.uniqueId)) {
        li.remove();
      }
    });

    // Now assign compact numbers in DOM order of refs
    const assigned = new Map();   // uniqueId -> number
    const orderedLis = [];
    let counter = 1;

    const currentLis = new Map();
    list.querySelectorAll('li[data-unique-id]').forEach(li => {
      currentLis.set(li.dataset.uniqueId, li);
      // ensure editable
      if (!li.hasAttribute('contenteditable')) {
        li.contentEditable = 'true';
      }
    });

    updatedRefNodes.forEach(ref => {
      const id = ref.dataset.uniqueId;
      if (!currentLis.has(id)) return; // orphan ref already removed or no li
      if (!assigned.has(id)) {
        assigned.set(id, counter++);
      }
      const num = assigned.get(id);
      ref.textContent = '[' + num + ']';
      ref.contentEditable = 'false';
      ref.tabIndex = 0;
      ref.setAttribute('role', 'button');
      ref.setAttribute('aria-label', 'Jump to footnote ' + num);
      ref.id = 'fn-ref-' + num;
      const li = currentLis.get(id);
      if (li) {
        decorateFootnoteListItem(li, num);
      }
    });

    // Order lis by their assigned numbers
    assigned.forEach((num, id) => {
      const li = currentLis.get(id);
      if (li) orderedLis[num - 1] = li;
    });

    const desiredOrder = orderedLis.filter(Boolean);
    desiredOrder.forEach((li, index) => {
      const currentChild = list.children[index];
      if (currentChild === li) return;
      if (currentChild) {
        list.insertBefore(li, currentChild);
      } else {
        list.appendChild(li);
      }
    });

    Array.from(list.children).forEach(child => {
      if (!desiredOrder.includes(child)) {
        child.remove();
      }
    });

    cleanupOrphanFootnoteTargets();
  }

  /* -------- Page + block creation helpers -------- */

  function createOverflowIndicator() {
    const indicator = document.createElement('div');
    indicator.className = 'overflow-indicator';
    indicator.setAttribute('aria-hidden', 'true');
    const pill = document.createElement('span');
    pill.textContent = '!';
    pill.title = 'Consider splitting page';
    pill.setAttribute('aria-label', 'Consider splitting page');
    indicator.appendChild(pill);
    return indicator;
  }

  function decorateDeleteIcon(span) {
    if (!span) return;
    span.classList.add('delete-icon');
    span.contentEditable = 'false';
    span.setAttribute('role', 'button');
    span.setAttribute('tabindex', '0');
    span.setAttribute('aria-label', 'Delete block');
    span.textContent = '×';
  }

  function createTextBlockShell() {
    const p = document.createElement('p');
    p.className = 'body-text';
    p.dataset.deletable = 'true';
    p.contentEditable = 'true';
    return p;
  }

  const INLINE_ORPHAN_TAGS = new Set(['SPAN','STRONG','EM','B','I','A','SUP','SUB','BR']);

  function ensureDeleteIcon(node) {
    if (!node) return;
    const hasIcon = Array.from(node.children || []).some(child => child.classList && child.classList.contains('delete-icon'));
    if (!hasIcon) {
      const icon = document.createElement('span');
      decorateDeleteIcon(icon);
      node.appendChild(icon);
    }
  }

  function convertElementToEditableBlock(node) {
    if (!node || node.closest('.footnotes-page')) return;
    node.dataset.deletable = 'true';
    if (!node.classList.contains('body-text') && !node.classList.contains('section-heading') && node.tagName !== 'H1' && node.tagName !== 'H2' && node.tagName !== 'H3' && node.tagName !== 'H4' && node.tagName !== 'H5' && node.tagName !== 'H6') {
      node.classList.add('body-text');
    }
    if (!node.matches('table,thead,tbody,tr,td,th')) {
      if (!node.hasAttribute('contenteditable')) node.contentEditable = 'true';
    }
    ensureDeleteIcon(node);
  }

  function wrapOrphanNodeInBlock(node, body) {
    if (!body) return;
    const reference = node.nextSibling;
    const block = createTextBlockShell();
    const icon = document.createElement('span');
    decorateDeleteIcon(icon);

    block.appendChild(node);
    block.appendChild(icon);
    body.insertBefore(block, reference);
  }

  function repairOrphanContentBlocks() {
    const bodies = editorRoot.querySelectorAll('.page-container:not(.footnotes-page) .page-body');
    bodies.forEach(body => {
      const nodes = Array.from(body.childNodes);
      nodes.forEach(node => {
        if (node.nodeType === Node.TEXT_NODE) {
          if (!(node.textContent || '').trim()) {
            node.remove();
            return;
          }
          wrapOrphanNodeInBlock(node, body);
          return;
        }
        if (node.nodeType !== Node.ELEMENT_NODE) {
          node.remove();
          return;
        }
        if (node.dataset && node.dataset.deletable === 'true') return;
        if (node.classList.contains('overflow-indicator')) return;
        if (INLINE_ORPHAN_TAGS.has(node.tagName)) {
          wrapOrphanNodeInBlock(node, body);
        } else {
          convertElementToEditableBlock(node);
        }
      });
    });
  }

  function createPage(options = {}) {
    const { withDefaultContent = true } = options;
    const div = document.createElement('div');
    div.className = 'page-container';
    div.dataset.pageType = 'content';
    div.innerHTML = `
      <div class="page-controls">
        <button class="page-control-btn" data-move-page-up title="Move page up" aria-label="Move page up">↑</button>
        <button class="page-control-btn" data-move-page-down title="Move page down" aria-label="Move page down">↓</button>
        <button class="page-control-btn" data-dup-page title="Duplicate page" aria-label="Duplicate page">⧉</button>
        <button class="page-control-btn" data-del-page title="Delete page" aria-label="Delete page">✕</button>
      </div>
      <div class="page-overlay"></div>
      <div class="page-body"></div>
      <div class="page-footer"></div>
    `;
    const body = div.querySelector('.page-body');
    if (withDefaultContent) {
      const section = document.createElement('h3');
      section.className = 'section-heading';
      section.dataset.deletable = 'true';
      section.innerHTML = `<span contenteditable="true">New section title</span>${DELETE_ICON_HTML}`;
      const paragraph = document.createElement('p');
      paragraph.className = 'body-text';
      paragraph.dataset.deletable = 'true';
      paragraph.contentEditable = 'true';
      paragraph.innerHTML = `New paragraph…${DELETE_ICON_HTML}`;
      body.appendChild(section);
      body.appendChild(paragraph);
    }
    body.appendChild(createOverflowIndicator());
    return div;
  }

  function createBlockFromData(data = {}) {
    const tag = data.tag || 'div';
    const el = document.createElement(tag);
    (data.classes || []).forEach(cls => el.classList.add(cls));
    const dataset = data.dataset || {};
    Object.keys(dataset).forEach(key => {
      if (dataset[key] !== undefined) {
        el.dataset[key] = dataset[key];
      }
    });
    if (!el.dataset.deletable) {
      el.dataset.deletable = 'true';
    }
    if (data.contentEditable) {
      el.contentEditable = 'true';
    } else {
      el.removeAttribute('contenteditable');
    }
    el.innerHTML = data.html || '';
    ensureDeleteIcon(el);
    return el;
  }

  function hydrateFootnotesFromModel(pageData) {
    const list = editorRoot.querySelector('.footnotes-page #footnote-list');
    if (!list) return;
    list.innerHTML = '';
    const items = pageData && Array.isArray(pageData.footnotes) ? pageData.footnotes : [];
    items.forEach((item, index) => {
      const li = document.createElement('li');
      li.contentEditable = 'true';
      Object.entries((item && item.dataset) || {}).forEach(([key, value]) => {
        if (value !== undefined) {
          li.dataset[key] = value;
        }
      });
      li.innerHTML = (item && item.html) || '';
      list.appendChild(li);
      decorateFootnoteListItem(li, index + 1);
    });
  }

  function renderDocumentModel(model) {
    const pages = model && Array.isArray(model.pages) ? model.pages : [];
    Array.from(editorRoot.querySelectorAll('.page-container')).forEach(page => page.remove());
    const fragment = document.createDocumentFragment();
    let footnotesData = null;
    pages.forEach(pageData => {
      if (pageData.type === 'footnotes') {
        footnotesData = pageData;
        return;
      }
      const pageEl = createPage({ withDefaultContent: false });
      const body = pageEl.querySelector('.page-body');
      const overflow = body.querySelector('.overflow-indicator');
      (pageData.blocks || []).forEach(blockData => {
        const block = createBlockFromData(blockData);
        body.insertBefore(block, overflow);
      });
      fragment.appendChild(pageEl);
    });
    if (!fragment.childNodes.length) {
      fragment.appendChild(createPage());
    }
    editorRoot.appendChild(fragment);
    ensureFootnotesPage();
    if (footnotesData) {
      hydrateFootnotesFromModel(footnotesData);
    } else {
      const list = editorRoot.querySelector('.footnotes-page #footnote-list');
      if (list) list.innerHTML = '';
    }
    updatePageNumbers();
    renumberFootnotes();
    initAllTableResizers();
    repairOrphanContentBlocks();
    checkAllPageOverflow();
  }

  function createDocTitleElement(text) {
    const div = document.createElement('div');
    div.className = 'doc-title';
    div.dataset.deletable = 'true';
    div.contentEditable = 'true';
    div.innerHTML = `${text}${DELETE_ICON_HTML}`;
    return div;
  }

  function createDocSubtitleElement(text) {
    const p = document.createElement('p');
    p.className = 'doc-subtitle body-text';
    p.dataset.deletable = 'true';
    p.contentEditable = 'true';
    p.innerHTML = `${text}${DELETE_ICON_HTML}`;
    return p;
  }

  function createSectionHeadingElement(text) {
    const h = document.createElement('h3');
    h.className = 'section-heading';
    h.dataset.deletable = 'true';
    h.innerHTML = `<span contenteditable="true">${text}</span>${DELETE_ICON_HTML}`;
    return h;
  }

  function createBodyParagraphElement(text) {
    const p = document.createElement('p');
    p.className = 'body-text';
    p.dataset.deletable = 'true';
    p.contentEditable = 'true';
    p.innerHTML = `${text}${DELETE_ICON_HTML}`;
    return p;
  }

  function populateIgaBriefTemplate(body) {
    body.appendChild(createDocTitleElement('Executive Brief Title'));
    body.appendChild(createDocSubtitleElement('Short positioning subtitle / context line.'));
    body.appendChild(createSectionHeadingElement('At a glance'));
    body.appendChild(createAtAGlanceTable());
    body.appendChild(createSectionHeadingElement('Key messages'));
    body.appendChild(createBodyParagraphElement('Use this space to summarise the essentials in clear, direct language.'));
    body.appendChild(createSectionHeadingElement('Before / After / Impact'));
    body.appendChild(createBAITable());
    body.appendChild(createSectionHeadingElement('Actions'));
    body.appendChild(createStepsBlock());
    body.appendChild(createSectionHeadingElement('FAQs'));
    body.appendChild(createFaqBlock());
    body.appendChild(createSectionHeadingElement('Sources'));
    body.appendChild(createBodyParagraphElement('List supporting sources and references here.'));
  }

  function populateFtExplainerTemplate(body) {
    body.appendChild(createDocTitleElement('Explainer headline'));
    body.appendChild(createDocSubtitleElement('Optional kicker for FT-style tone.'));
    body.appendChild(createSectionHeadingElement('Context'));
    body.appendChild(createBodyParagraphElement('Set the context with a succinct summary and key framing.'));
    body.appendChild(createSectionHeadingElement('Key messages'));
    const cta = createCtaBox();
    const titleEl = cta.querySelector('.cta-title');
    const bodyEl = cta.querySelector('.cta-body');
    if (titleEl) titleEl.textContent = 'Key message';
    if (bodyEl) bodyEl.textContent = 'Use this callout to highlight the most important takeaway.';
    cta.dataset.customTitle = 'true';
    body.appendChild(cta);
    body.appendChild(createSectionHeadingElement('What changed'));
    body.appendChild(createBodyParagraphElement('Explain what has changed and why it matters for the audience.'));
    body.appendChild(createSectionHeadingElement('Implications'));
    body.appendChild(createBodyParagraphElement('Spell out what readers should do or think differently as a result.'));
    body.appendChild(createSectionHeadingElement('FAQs'));
    body.appendChild(createFaqBlock());
    body.appendChild(createSectionHeadingElement('Sources'));
    body.appendChild(createBodyParagraphElement('Include supporting evidence, links or appendices.'));
  }

  const TEMPLATE_BUILDERS = {
    'iga-brief': populateIgaBriefTemplate,
    'ft-explainer': populateFtExplainerTemplate
  };

  function applyTemplate(templateKey) {
    const builder = TEMPLATE_BUILDERS[templateKey];
    if (!builder) return;
    if (!window.confirm('Applying a template will replace the current content. Continue?')) return;
    captureSnapshot();
    Array.from(editorRoot.querySelectorAll('.page-container:not(.footnotes-page)')).forEach(page => page.remove());
    const newPage = createPage({ withDefaultContent: false });
    const footnotes = editorRoot.querySelector('.footnotes-page');
    editorRoot.insertBefore(newPage, footnotes);
    const body = newPage.querySelector('.page-body');
    builder(body);
    updatePageNumbers();
    renumberFootnotes();
    initAllTableResizers();
    checkAllPageOverflow();
    afterChange();
    ensureBlockVisible(newPage);
    showSnackbar('Template applied');
  }

  function addPage() {
    captureSnapshot();
    const footnotes = editorRoot.querySelector('.footnotes-page');
    const page = createPage();
    editorRoot.insertBefore(page, footnotes);
    updatePageNumbers();
    renumberFootnotes();
    afterChange();
    ensureBlockVisible(page);
  }

  function addSectionBlock() {
    const body = getActiveContentPageBody();
    if (!body) return;
    captureSnapshot();
    const h = document.createElement('h3');
    h.className = 'section-heading';
    h.dataset.deletable = 'true';
    h.innerHTML = `<span contenteditable="true">New section title</span>${DELETE_ICON_HTML}`;
    body.appendChild(h);
    afterChange();
    ensureBlockVisible(h);
  }

  function addTextBlock() {
    const body = getActiveContentPageBody();
    if (!body) return;
    captureSnapshot();
    const p = document.createElement('p');
    p.className = 'body-text';
    p.dataset.deletable = 'true';
    p.contentEditable = 'true';
    p.innerHTML = `New paragraph…${DELETE_ICON_HTML}`;
    body.appendChild(p);
    afterChange();
    ensureBlockVisible(p);
  }

  function createCtaBox() {
    const div = document.createElement('div');
    div.className = 'cta-box cta-blue page-break-avoider';
    div.dataset.deletable = 'true';
    div.dataset.generatedTitle = 'Info';
    div.innerHTML = `
      ${DELETE_ICON_HTML}
      <div class="cta-type-switch" contenteditable="false">
        <button data-cta="cta-blue" class="active">Info</button>
        <button data-cta="cta-green">Action</button>
        <button data-cta="cta-yellow">Caution</button>
        <button data-cta="cta-red">Critical</button>
      </div>
      <div class="cta-title" contenteditable="true">Info</div>
      <div class="cta-body" contenteditable="true">Key message text…</div>
    `;
    return div;
  }

  function addCtaBlock() {
    const body = getActiveContentPageBody();
    if (!body) return;
    captureSnapshot();
    const cta = createCtaBox();
    body.appendChild(cta);
    afterChange();
    ensureBlockVisible(cta);
  }

  function createStepsBlock() {
    const div = document.createElement('div');
    div.className = 'steps-block page-break-avoider';
    div.dataset.deletable = 'true';
    div.innerHTML = `
      ${DELETE_ICON_HTML}
      <div class="steps-title" contenteditable="true">Steps for implementation</div>
      <ol class="styled-steps">
        <li contenteditable="true">Step 1…</li>
        <li contenteditable="true">Step 2…</li>
        <li contenteditable="true">Step 3…</li>
      </ol>
      <button class="steps-add" contenteditable="false">+ Add step</button>
    `;
    return div;
  }

  function addStepsBlock() {
    const body = getActiveContentPageBody();
    if (!body) return;
    captureSnapshot();
    const block = createStepsBlock();
    body.appendChild(block);
    afterChange();
    ensureBlockVisible(block);
  }

  function createFaqBlock() {
    const div = document.createElement('div');
    div.className = 'faq-block page-break-avoider';
    div.dataset.deletable = 'true';
    div.innerHTML = `
      ${DELETE_ICON_HTML}
      <div class="faq-title" contenteditable="true">Frequently asked questions</div>
      <ol class="faq-list">
        <li>
          <div class="faq-q" contenteditable="true">Q: …</div>
          <div class="faq-a" contenteditable="true">A: …</div>
        </li>
      </ol>
      <button class="faq-add" contenteditable="false">+ Add FAQ</button>
    `;
    return div;
  }

  function addFaqBlock() {
    const body = getActiveContentPageBody();
    if (!body) return;
    captureSnapshot();
    const block = createFaqBlock();
    body.appendChild(block);
    afterChange();
    ensureBlockVisible(block);
  }

  function addQuoteBlock() {
    const body = getActiveContentPageBody();
    if (!body) return;
    captureSnapshot();
    const div = document.createElement('div');
    div.className = 'quote-block';
    div.dataset.deletable = 'true';
    div.innerHTML = `
      ${DELETE_ICON_HTML}
      <blockquote class="quote-body" contenteditable="true">“Quoted text…”</blockquote>
      <div class="quote-source" contenteditable="true">Source / speaker (optional)</div>
    `;
    body.appendChild(div);
    afterChange();
    ensureBlockVisible(div);
  }

  function addDividerBlock() {
    const body = getActiveContentPageBody();
    if (!body) return;
    captureSnapshot();
    const div = document.createElement('div');
    div.className = 'divider-block page-break-avoider';
    div.dataset.deletable = 'true';
    div.innerHTML = `
      <hr class="section-divider" />
      ${DELETE_ICON_HTML}
    `;
    body.appendChild(div);
    afterChange();
    ensureBlockVisible(div);
  }

  function addCitationBlock() {
    const body = getActiveContentPageBody();
    if (!body) return;
    captureSnapshot();
    const p = document.createElement('p');
    p.className = 'citation-block';
    p.dataset.deletable = 'true';
    p.contentEditable = 'true';
    p.innerHTML = `(Source: …)${DELETE_ICON_HTML}`;
    body.appendChild(p);
    afterChange();
    ensureBlockVisible(p);
  }

  function createAtAGlanceTable() {
    const div = document.createElement('div');
    div.className = 'table-block page-break-avoider';
    div.dataset.deletable = 'true';
    div.innerHTML = `
      ${DELETE_ICON_HTML}
      <table class="iga-table">
        <thead>
          <tr>
            <th contenteditable="true">Theme</th>
            <th contenteditable="true">Detail</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td contenteditable="true">Audience</td>
            <td contenteditable="true">Who this is for…</td>
          </tr>
          <tr>
            <td contenteditable="true">Objective</td>
            <td contenteditable="true">What we need them to do…</td>
          </tr>
          <tr>
            <td contenteditable="true">Timing</td>
            <td contenteditable="true">Key dates…</td>
          </tr>
        </tbody>
      </table>
      <div class="table-controls" contenteditable="false">
        <button data-table-action="add-row">+ Row</button>
        <button data-table-action="add-col">+ Col</button>
        <button data-table-action="del-row">- Row</button>
        <button data-table-action="del-col">- Col</button>
      </div>
    `;
    const table = div.querySelector('table');
    initTableResizing(table);
    return div;
  }

  function createBAITable() {
    const div = document.createElement('div');
    div.className = 'table-block page-break-avoider';
    div.dataset.deletable = 'true';
    div.innerHTML = `
      ${DELETE_ICON_HTML}
      <table class="iga-table">
        <thead>
          <tr>
            <th contenteditable="true">Before</th>
            <th contenteditable="true">After</th>
            <th contenteditable="true">Impact</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td contenteditable="true">Current state…</td>
            <td contenteditable="true">Target state…</td>
            <td contenteditable="true">Outcome / benefit…</td>
          </tr>
        </tbody>
      </table>
      <div class="table-controls" contenteditable="false">
        <button data-table-action="add-row">+ Row</button>
        <button data-table-action="add-col">+ Col</button>
        <button data-table-action="del-row">- Row</button>
        <button data-table-action="del-col">- Col</button>
      </div>
    `;
    const table = div.querySelector('table');
    initTableResizing(table);
    return div;
  }

  function addAtAGlanceTable() {
    const body = getActiveContentPageBody();
    if (!body) return;
    captureSnapshot();
    const block = createAtAGlanceTable();
    body.appendChild(block);
    afterChange();
    ensureBlockVisible(block);
  }

  function addBAITable() {
    const body = getActiveContentPageBody();
    if (!body) return;
    captureSnapshot();
    const block = createBAITable();
    body.appendChild(block);
    afterChange();
    ensureBlockVisible(block);
  }

  function openCustomTableModal() {
    openModal({
      title: 'Add custom table',
      body: `
        <label>Rows</label>
        <input type="number" id="tbl-rows" min="1" max="40" value="3" />
        <label>Columns</label>
        <input type="number" id="tbl-cols" min="1" max="8" value="3" />
        <label>Or paste TSV (overrides rows/cols)</label>
        <textarea id="tbl-tsv" rows="4" placeholder="cell1\tcell2\ncell3\tcell4"></textarea>
      `,
      onConfirm: () => {
        const rowsInput = modalInner.querySelector('#tbl-rows');
        const colsInput = modalInner.querySelector('#tbl-cols');
        const tsvInput = modalInner.querySelector('#tbl-tsv');
        const tsv = (tsvInput.value || '').trim();
        let rows = parseInt(rowsInput.value, 10) || 1;
        let cols = parseInt(colsInput.value, 10) || 1;

        let data = [];
        if (tsv) {
          const lines = tsv.split(/\r?\n/);
          lines.forEach(line => {
            const cells = line.split('\t');
            data.push(cells);
            if (cells.length > cols) cols = cells.length;
          });
          rows = data.length;
        }

        const body = getActiveContentPageBody();
        if (!body) {
          closeModal();
          return;
        }

        captureSnapshot();

        const block = document.createElement('div');
        block.className = 'table-block page-break-avoider';
        block.dataset.deletable = 'true';

        let html = `
          ${DELETE_ICON_HTML}
          <table class="iga-table">
            <tbody>
        `;
        for (let r = 0; r < rows; r++) {
          html += '<tr>';
          for (let c = 0; c < cols; c++) {
            const txt = data[r] && data[r][c] ? data[r][c] : ' ';
            html += `<td contenteditable="true">${txt}</td>`;
          }
          html += '</tr>';
        }
        html += `
            </tbody>
          </table>
          <div class="table-controls" contenteditable="false">
            <button data-table-action="add-row">+ Row</button>
            <button data-table-action="add-col">+ Col</button>
            <button data-table-action="del-row">- Row</button>
            <button data-table-action="del-col">- Col</button>
          </div>
        `;
        block.innerHTML = html;
        const table = block.querySelector('table');
        initTableResizing(table);
        body.appendChild(block);
        afterChange();
        ensureBlockVisible(block);
        closeModal();
      }
    });
  }

  function addImageBlockWithURL(url) {
    if (!url) return;
    const body = getActiveContentPageBody();
    if (!body) return;
    captureSnapshot();
    const div = document.createElement('div');
    div.className = 'image-block image-size-m page-break-avoider';
    div.dataset.deletable = 'true';
    div.innerHTML = `
      ${DELETE_ICON_HTML}
      <div class="image-size-controls" contenteditable="false">
        <button data-size="s">S</button>
        <button data-size="m" class="active">M</button>
        <button data-size="l">L</button>
      </div>
      <img src="${url}" alt="">
      <div class="image-caption" contenteditable="true">Add an explanatory caption…</div>
    `;
    body.appendChild(div);
    afterChange();
    ensureBlockVisible(div);
  }

  function openImageModal() {
    openModal({
      title: 'Add image',
      body: `
        <label>Image URL</label>
        <input type="url" id="img-url" placeholder="https://example.com/image.png" />
      `,
      onConfirm: () => {
        const input = modalInner.querySelector('#img-url');
        const url = input.value.trim();
        if (!url) {
          showSnackbar('Image URL required');
          return;
        }
        addImageBlockWithURL(url);
        closeModal();
      }
    });
  }

  /* -------- Modal helpers -------- */

  function openModal({ title, body, onConfirm, onCancel, confirmLabel = 'OK', cancelLabel = 'Cancel' }) {
    modalInner.innerHTML = `
      <h3>${title}</h3>
      ${body}
      <div class="modal-actions">
        <button type="button" data-modal-cancel>Cancel</button>
        <button type="button" class="primary" data-modal-confirm>OK</button>
      </div>
    `;
    modalBackdrop.classList.add('visible');
    const cancelBtn = modalInner.querySelector('[data-modal-cancel]');
    const confirmBtn = modalInner.querySelector('[data-modal-confirm]');
    if (cancelBtn) {
      cancelBtn.textContent = cancelLabel;
      cancelBtn.onclick = () => {
        if (typeof onCancel === 'function') {
          onCancel();
        } else {
          closeModal();
        }
      };
    }
    if (confirmBtn) {
      confirmBtn.textContent = confirmLabel;
      confirmBtn.onclick = () => {
        if (typeof onConfirm === 'function') {
          onConfirm();
        } else {
          closeModal();
        }
      };
    }
  }

  function closeModal() {
    modalBackdrop.classList.remove('visible');
    modalInner.innerHTML = '';
  }

  /* -------- Table controls + resizing -------- */

  function handleTableControl(btn) {
    const action = btn.dataset.tableAction;
    const block = btn.closest('.table-block');
    if (!block) return;
    const table = block.querySelector('table');
    if (!table) return;

    captureSnapshot();

    if (action === 'add-row') {
      const tbody = table.tBodies[0] || table.createTBody();
      const refRow = tbody.rows[tbody.rows.length - 1] || (table.tHead && table.tHead.rows[0]);
      const cols = refRow ? refRow.cells.length : 1;
      const row = tbody.insertRow(-1);
      for (let i = 0; i < cols; i++) {
        const cell = row.insertCell(-1);
        cell.contentEditable = 'true';
        cell.textContent = ' ';
      }
    } else if (action === 'add-col') {
      const rows = table.rows;
      for (let r = 0; r < rows.length; r++) {
        const cell = rows[r].insertCell(-1);
        cell.contentEditable = 'true';
        cell.textContent = ' ';
      }
    } else if (action === 'del-row') {
      const tbody = table.tBodies[0];
      if (tbody && tbody.rows.length > 0) tbody.deleteRow(-1);
    } else if (action === 'del-col') {
      const rows = table.rows;
      if (rows.length && rows[0].cells.length > 1) {
        for (let r = 0; r < rows.length; r++) {
          rows[r].deleteCell(-1);
        }
      }
    }

    const tbody = table.tBodies[0];
    if ((!tbody || !tbody.rows.length) && (!table.tHead || !table.tHead.rows.length)) {
      block.remove();
    }

    initTableResizing(table);
    afterChange();
  }

  function initTableResizing(table) {
    if (!table) return;
    // Remove existing handles
    table.querySelectorAll('.col-resize-handle').forEach(h => h.remove());

    const headerRow =
      (table.tHead && table.tHead.rows[0]) ||
      (table.tBodies[0] && table.tBodies[0].rows[0]);

    if (!headerRow) return;

    const cells = Array.from(headerRow.cells);
    if (cells.length <= 1) return;

    cells.forEach((cell, index) => {
      // Optional: skip last column handle if you prefer
      if (index === cells.length - 1) return;

      const handle = document.createElement('div');
      handle.className = 'col-resize-handle';
      handle.contentEditable = 'false';
      handle.addEventListener('mousedown', e => {
        e.preventDefault();
        e.stopPropagation();
        startColumnResize(table, index, e);
      });
      cell.appendChild(handle);
    });
  }

  function startColumnResize(table, colIndex, event) {
    const startX = event.clientX;
    const colCells = Array.from(table.rows)
      .map(row => row.cells[colIndex])
      .filter(Boolean);
    if (!colCells.length) return;

    const startWidths = colCells.map(cell => cell.offsetWidth);

    tableResizeState = {
      table,
      colCells,
      startX,
      startWidths
    };

    document.addEventListener('mousemove', handleColumnResizeMove);
    document.addEventListener('mouseup', stopColumnResize);
  }

  function handleColumnResizeMove(e) {
    if (!tableResizeState) return;
    const deltaX = e.clientX - tableResizeState.startX;

    tableResizeState.colCells.forEach((cell, i) => {
      const base = tableResizeState.startWidths[i];
      const next = Math.max(40, base + deltaX);
      cell.style.width = next + 'px';
    });

    e.preventDefault();
  }

  function stopColumnResize() {
    if (!tableResizeState) return;
    document.removeEventListener('mousemove', handleColumnResizeMove);
    document.removeEventListener('mouseup', stopColumnResize);
    tableResizeState = null;
    scheduleInputSnapshot();
  }

  function initAllTableResizers() {
    editorRoot.querySelectorAll('table.iga-table').forEach(initTableResizing);
  }

  /* -------- Paste / sanitize / clipboard -------- */

  function insertPastedImage(file) {
    const reader = new FileReader();
    reader.onload = e => addImageBlockWithURL(e.target.result);
    reader.readAsDataURL(file);
  }

  function sanitizeExternalHTML(html) {
    const allowedTags = new Set([
      'P','H1','H2','H3','H4','H5','H6','UL','OL','LI','A','STRONG','EM',
      'B','I','TABLE','TBODY','THEAD','TR','TD','TH','BLOCKQUOTE','BR','SPAN',
      'DIV','SECTION','ARTICLE','FIGURE','FIGCAPTION','SUP','SUB','IMG','HR'
    ]);
    const allowedAttrs = {
      'A': ['href'],
      'IMG': ['src','alt','title'],
      'TD': ['colspan','rowspan'],
      'TH': ['colspan','rowspan','scope'],
      'TABLE': [],
      'TR': [],
      'P': [],
      'BLOCKQUOTE': [],
      'SPAN': []
    };
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');

    function cleanse(node) {
      if (node.nodeType === Node.TEXT_NODE) return node.cloneNode(true);
      if (node.nodeType !== Node.ELEMENT_NODE) return document.createTextNode('');
      if (!allowedTags.has(node.tagName)) {
        const frag = document.createDocumentFragment();
        node.childNodes.forEach(child => {
          const c = cleanse(child);
          if (c) frag.appendChild(c);
        });
        return frag;
      }
      const el = document.createElement(node.tagName.toLowerCase());
      const allowed = allowedAttrs[node.tagName] || [];
      for (let attr of Array.from(node.attributes)) {
        const name = attr.name.toLowerCase();
        const value = attr.value;
        if (node.tagName === 'A' && name === 'href') {
          if (/^https?:\/\//i.test(value)) el.setAttribute('href', value);
          continue;
        }
        if (node.tagName === 'IMG' && name === 'src') {
          if (/^(https?:|data:image)/i.test(value)) el.setAttribute('src', value);
          continue;
        }
        if (allowed.includes(name)) {
          el.setAttribute(attr.name, value);
          continue;
        }
        if (name === 'class' || name === 'id' || name === 'title' || name === 'role' || name === 'tabindex') {
          el.setAttribute(attr.name, value);
          continue;
        }
        if (name === 'contenteditable') {
          const normalized = value === 'false' ? 'false' : 'true';
          el.setAttribute('contenteditable', normalized);
          continue;
        }
        if (name.startsWith('data-') || name.startsWith('aria-')) {
          el.setAttribute(attr.name, value);
          continue;
        }
      }
      node.childNodes.forEach(child => {
        const c = cleanse(child);
        if (c) el.appendChild(c);
      });
      return el;
    }

    const frag = document.createDocumentFragment();
    doc.body.childNodes.forEach(child => {
      const c = cleanse(child);
      if (c) frag.appendChild(c);
    });
    return frag;
  }

  function handlePaste(e) {
    if (!isInEditor(e.target)) return;
    const cd = e.clipboardData || window.clipboardData;
    if (!cd) return;

    const items = cd.items || [];
    for (let i = 0; i < items.length; i++) {
      const it = items[i];
      if (it.type && it.type.indexOf('image') === 0) {
        e.preventDefault();
        const file = it.getAsFile();
        if (file) insertPastedImage(file);
        return;
      }
    }

    const html = cd.getData('text/html');
    const text = cd.getData('text/plain');

    // Internal block copy/paste
    if (html && html.includes(INTERNAL_MARKER_START) && html.includes(INTERNAL_MARKER_END)) {
      e.preventDefault();
      const start = html.indexOf(INTERNAL_MARKER_START) + INTERNAL_MARKER_START.length;
      const end = html.indexOf(INTERNAL_MARKER_END);
      const blockHtml = html.slice(start, end);
      insertInternalBlocks(blockHtml);
      return;
    }

    if (html) {
      e.preventDefault();
      const safe = sanitizeExternalHTML(html);
      const sel = window.getSelection();
      if (!sel.rangeCount) return;
      captureSnapshot();
      const range = sel.getRangeAt(0);
      range.deleteContents();
      range.insertNode(safe);
      afterChange();
      return;
    }

    if (text) {
      e.preventDefault();
      const normalized = text.replace(/\r\n/g, '\n');
      const paragraphs = normalized
        .split(/\n{2,}/)
        .map(seg => seg.trim())
        .filter(seg => seg.length);

      if (!paragraphs.length) return;

      const escapeHtml = str => str.replace(/[&<>"']/g, ch => ({
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#39;'
      })[ch]);

      const sel = window.getSelection();
      if (sel && sel.rangeCount) {
        const range = sel.getRangeAt(0);
        const container = range.commonAncestorContainer.nodeType === Node.ELEMENT_NODE
          ? range.commonAncestorContainer
          : range.commonAncestorContainer.parentElement;
        const pageBody = container && container.closest('.page-body');

        if (pageBody) {
          captureSnapshot();
          const page = pageBody.closest('.page-container');
          if (page && page.classList.contains('footnotes-page')) {
            const inlineHtml = paragraphs
              .map(seg => escapeHtml(seg).replace(/\n/g, '<br>'))
              .join('<br><br>');
            document.execCommand('insertHTML', false, inlineHtml);
          } else {
            const blockHtml = paragraphs
              .map(seg => {
                const html = escapeHtml(seg).replace(/\n/g, '<br>');
                return `<p class="body-text" data-deletable="true" contenteditable="true">${html}${DELETE_ICON_HTML}</p>`;
              })
              .join('');
            document.execCommand('insertHTML', false, blockHtml);
          }
          afterChange();
          return;
        }
      }

      const body = getActiveContentPageBody();
      if (!body) return;
      captureSnapshot();
      paragraphs.forEach(seg => {
        const p = document.createElement('p');
        p.className = 'body-text';
        p.dataset.deletable = 'true';
        p.contentEditable = 'true';
        const lines = seg.split('\n');
        lines.forEach((line, index) => {
          if (index > 0) p.appendChild(document.createElement('br'));
          p.appendChild(document.createTextNode(line));
        });
        const icon = document.createElement('span');
        decorateDeleteIcon(icon);
        p.appendChild(icon);
        body.appendChild(p);
      });
      afterChange();
    }
  }

  function insertInternalBlocks(html) {
    const tmp = document.createElement('div');
    tmp.innerHTML = html;
    const blocks = Array.from(tmp.children);
    if (!blocks.length) return;

    const lastSelected = getLastSelectedBlock();
    const sel = window.getSelection();
    let insertAfter = lastSelected;

    if (!insertAfter && sel && sel.rangeCount) {
      const node = sel.anchorNode;
      const el = node && (node.nodeType === 1 ? node : node.parentElement);
      const blk = el && el.closest('[data-deletable="true"]');
      if (blk && !blk.closest('.footnotes-page')) insertAfter = blk;
    }

    const body = insertAfter ? insertAfter.closest('.page-body') : getActiveContentPageBody();
    if (!body) return;

    captureSnapshot();

    blocks.forEach(b => {
      if (b.dataset && b.dataset.deletable === 'true') {
        const clone = b.cloneNode(true);
        if (insertAfter) {
          insertAfter.insertAdjacentElement('afterend', clone);
          insertAfter = clone;
        } else {
          body.appendChild(clone);
        }
      }
    });

    renumberFootnotes();
    afterChange();
  }

  function copySelectedBlocks(isCut = false) {
    const blocks = getSelectedBlocks();
    if (!blocks.length) {
      showSnackbar('Select block(s) to copy');
      return;
    }
    const html = blocks.map(b => b.outerHTML).join('');
    internalClipboardHTML = html;
    const wrapped = INTERNAL_MARKER_START + html + INTERNAL_MARKER_END;
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(wrapped).catch(() => {});
    }
    if (isCut) {
      captureSnapshot();
      blocks.forEach(b => b.remove());
      renumberFootnotes();
      afterChange();
      clearBlockSelection();
    }
    showSnackbar(isCut ? 'Block(s) cut' : 'Block(s) copied', 1200);
  }

  function pasteBlocksAfterSelection() {
    if (!internalClipboardHTML) {
      showSnackbar('Clipboard is empty');
      return;
    }
    insertInternalBlocks(internalClipboardHTML);
  }

  function moveBlocksToNewPage() {
    const blocks = getSelectedBlocks();
    if (!blocks.length) {
      showSnackbar('Select block(s) to move');
      return;
    }

    captureSnapshot();

    const sorted = blocks.slice().sort((a, b) =>
      a.compareDocumentPosition(b) & Node.DOCUMENT_POSITION_FOLLOWING ? -1 : 1
    );

    const anchor = sorted[sorted.length - 1];
    const anchorPage = anchor.closest('.page-container');
    const newPage = createPage({ withDefaultContent: false });
    const footnotes = editorRoot.querySelector('.footnotes-page');

    if (anchorPage && !anchorPage.classList.contains('footnotes-page')) {
      anchorPage.insertAdjacentElement('afterend', newPage);
    } else {
      editorRoot.insertBefore(newPage, footnotes);
    }

    const body = newPage.querySelector('.page-body');
    sorted.forEach(block => body.appendChild(block));

    updatePageNumbers();
    renumberFootnotes();
    checkAllPageOverflow();
    afterChange();

    clearBlockSelection();
    sorted.forEach(block => block.classList.add('block-selected'));
    showBlockActions();
    if (sorted[0]) ensureBlockVisible(sorted[0]);
    showSnackbar('Moved blocks to new page');
  }

  function moveBlocks(direction) {
    const blocks = getSelectedBlocks();
    if (!blocks.length) return;
    captureSnapshot();

    const sorted = blocks.slice().sort((a, b) =>
      a.compareDocumentPosition(b) & Node.DOCUMENT_POSITION_FOLLOWING ? -1 : 1
    );
    const isUp = direction === 'up';

    if (isUp) {
      sorted.forEach(block => {
        const prev = block.previousElementSibling;
        if (prev && prev.dataset.deletable === 'true') prev.before(block);
      });
    } else {
      sorted.reverse().forEach(block => {
        const next = block.nextElementSibling;
        if (next && next.dataset.deletable === 'true') next.after(block);
      });
    }

    renumberFootnotes();
    afterChange();
    showBlockActions();
  }

  function duplicateBlocks() {
    const blocks = getSelectedBlocks();
    if (!blocks.length) return;
    captureSnapshot();
    let lastClone = null;
    blocks.forEach(block => {
      const clone = block.cloneNode(true);
      block.insertAdjacentElement('afterend', clone);
      lastClone = clone;
    });
    renumberFootnotes();
    afterChange();
    clearBlockSelection();
    if (lastClone) {
      lastClone.classList.add('block-selected');
      ensureBlockVisible(lastClone);
      showBlockActions();
    }
  }

  function deleteBlocks() {
    const blocks = getSelectedBlocks();
    if (!blocks.length) return;
    captureSnapshot();
    blocks.forEach(b => b.remove());
    renumberFootnotes();
    afterChange();
    clearBlockSelection();
  }

  function collectQualityIssues() {
    const issues = new Set();
    const placeholderHits = new Set();
    const scanNodes = editorRoot.querySelectorAll('[data-deletable="true"], td, th, .steps-block li, .faq-block .faq-q, .faq-block .faq-a, .image-caption');
    scanNodes.forEach(node => {
      const text = (node.textContent || '').trim();
      if (!text) return;
      QUALITY_PLACEHOLDERS.forEach(placeholder => {
        if (text.includes(placeholder)) placeholderHits.add(placeholder);
      });
    });
    if (placeholderHits.size) {
      issues.add('Replace placeholder copy: ' + Array.from(placeholderHits).join(', '));
    }

    editorRoot.querySelectorAll('.cta-box').forEach(box => {
      const titleEl = box.querySelector('.cta-title');
      const bodyEl = box.querySelector('.cta-body');
      const title = titleEl ? titleEl.textContent.trim() : '';
      const body = bodyEl ? bodyEl.textContent.trim() : '';
      const customTitle = box.dataset.customTitle === 'true';
      if ((!customTitle && title === 'Info') || body === 'Key message text…') {
        issues.add('Update CTA blocks so they no longer use default Info / Key message text.');
      }
    });

    editorRoot.querySelectorAll('.image-block').forEach(block => {
      const img = block.querySelector('img');
      const altText = img ? (img.getAttribute('alt') || '').trim() : '';
      const captionEl = block.querySelector('.image-caption');
      const captionText = captionEl ? captionEl.textContent.trim() : '';
      if (!altText) {
        issues.add('Provide alt text for all images.');
      }
      if (!captionText || captionText === 'Add an explanatory caption…') {
        issues.add('Add explanatory captions for images.');
      }
    });

    const unresolvedRefs = Array.from(editorRoot.querySelectorAll('.footnote-ref')).some(ref => (ref.textContent || '').includes('?'));
    if (unresolvedRefs) {
      issues.add('Resolve placeholder footnote references [?].');
    }

    return Array.from(issues);
  }

  function runQualityCheck(options = {}) {
    const { onContinue, showNoIssues = true } = options;
    const issues = collectQualityIssues();
    if (!issues.length) {
      if (showNoIssues) showSnackbar('No quality issues found');
      if (typeof onContinue === 'function') onContinue();
      return;
    }

    const listItems = issues.map(issue => `<li>${issue}</li>`).join('');
    openModal({
      title: 'Quality check',
      body: `<p>Fix these before finalising your brief:</p><ul>${listItems}</ul>`,
      confirmLabel: onContinue ? 'Continue anyway' : 'OK',
      cancelLabel: 'Close',
      onConfirm: () => {
        closeModal();
        if (typeof onContinue === 'function') onContinue();
      },
      onCancel: closeModal
    });
  }

  function showHelpModal() {
    const body = `
      <div class="help-modal-content">
        <p><strong>Selecting blocks:</strong> Click inside any block to select it. Use Ctrl/Cmd + click to multi-select and the floating menu to move, duplicate or delete.</p>
        <p><strong>CTA colours:</strong></p>
        <ul>
          <li><strong>Info</strong> — neutral context or updates.</li>
          <li><strong>Action</strong> — tasks and next steps.</li>
          <li><strong>Caution</strong> — risks or watch-outs.</li>
          <li><strong>Critical</strong> — urgent blockers or escalations.</li>
        </ul>
        <p><strong>Tables:</strong></p>
        <ul>
          <li><em>At-a-glance</em> covers audience, objective and timing.</li>
          <li><em>Before / After / Impact</em> captures the change story.</li>
          <li>Use custom tables for bespoke layouts.</li>
        </ul>
        <p><strong>Footnotes:</strong> Select text and press “+ Footnote” to wrap it. Click a reference to jump to its note and use the ↩ button to return.</p>
      </div>
    `;
    openModal({
      title: 'Quick help',
      body,
      confirmLabel: 'Got it',
      cancelLabel: 'Close',
      onConfirm: closeModal,
      onCancel: closeModal
    });
  }

  /* -------- Export / import -------- */

  function exportJSON() {
    const performExport = () => {
      const theme = getCurrentTheme();
      clearBlockSelection();
      hideTextToolbar();
      hideBlockActions();
      const model = buildDocumentModel();
      const payload = {
        version: '2.0.0',
        exportedAt: new Date().toISOString(),
        theme,
        pages: model.pages
      };
      const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'executive-brief.json';
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
      showSnackbar('JSON exported');
    };

    runQualityCheck({ onContinue: performExport });
  }

  function importJSONFromFile(file) {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const obj = JSON.parse(e.target.result);
        suppressSnapshots = true;
        if (Array.isArray(obj.pages)) {
          setTheme(obj.theme === 'ft' ? 'ft' : 'iga');
          renderDocumentModel({ pages: obj.pages });
          suppressSnapshots = false;
          initHistory();
          showSnackbar('JSON imported');
          return;
        }

        if (!obj.content) throw new Error('Missing content');
        const parser = new DOMParser();
        const doc = parser.parseFromString(obj.content, 'text/html');
        const main = doc.querySelector('main#editor-root') || doc.querySelector('main') || doc.body;
        if (!main) throw new Error('Invalid content wrapper');

        editorRoot.innerHTML = main.innerHTML;
        setTheme(obj.theme === 'ft' ? 'ft' : 'iga');
        ensureFootnotesPage();
        updatePageNumbers();
        renumberFootnotes();
        initAllTableResizers();
        suppressSnapshots = false;
        initHistory();
        showSnackbar('JSON imported');
      } catch (err) {
        console.error(err);
        showSnackbar('Import failed: ' + err.message);
        suppressSnapshots = false;
      }
    };
    reader.readAsText(file);
  }

  function htmlToPlainText(html) {
    if (!html) return '';
    const temp = document.createElement('div');
    temp.innerHTML = html;
    temp.querySelectorAll('br').forEach(br => br.replaceWith('\n'));
    return (temp.textContent || '')
      .replace(/\u00a0/g, ' ')
      .replace(/[\t ]+/g, ' ')
      .replace(/\s*\n\s*/g, '\n')
      .trim();
  }

  function extractListData(html) {
    if (!html) return null;
    const temp = document.createElement('div');
    temp.innerHTML = html;
    const list = temp.querySelector('ul, ol');
    if (!list) return null;
    const items = Array.from(list.querySelectorAll('li')).map(li => htmlToPlainText(li.innerHTML)).filter(Boolean);
    if (!items.length) return null;
    return { ordered: list.tagName === 'OL', items };
  }

  function extractTableRows(html) {
    if (!html) return [];
    const temp = document.createElement('div');
    temp.innerHTML = html;
    const table = temp.querySelector('table');
    if (!table) return [];
    const rows = [];
    table.querySelectorAll('tr').forEach(row => {
      const cells = Array.from(row.querySelectorAll('th, td')).map(cell => htmlToPlainText(cell.innerHTML));
      if (cells.length) rows.push(cells);
    });
    return rows;
  }

  function isTableBlock(data) {
    if (!data) return false;
    if ((data.tag || '').toLowerCase() === 'table') return true;
    return Array.isArray(data.classes) && data.classes.some(cls => cls && cls.includes('table'));
  }

  function getDocxHeading(block) {
    if (!window.docx || !block) return null;
    const { HeadingLevel } = window.docx;
    if (block.classes && block.classes.includes('doc-title')) return HeadingLevel.TITLE;
    if (block.classes && block.classes.includes('section-heading')) return HeadingLevel.HEADING_3;
    switch ((block.tag || '').toLowerCase()) {
      case 'h1':
        return HeadingLevel.HEADING_1;
      case 'h2':
        return HeadingLevel.HEADING_2;
      case 'h3':
        return HeadingLevel.HEADING_3;
      default:
        return null;
    }
  }

  function blockToDocxNodes(block) {
    if (!window.docx || !block) return null;
    const { Paragraph, Table, TableRow, TableCell } = window.docx;
    if (isTableBlock(block)) {
      const rows = extractTableRows(block.html);
      if (!rows.length) return null;
      return new Table({
        rows: rows.map(row => new TableRow({
          children: row.map(text => new TableCell({ children: [new Paragraph(text || '')] }))
        }))
      });
    }
    const listData = extractListData(block.html);
    if (listData) {
      return listData.items.map((item, index) => {
        const prefix = listData.ordered ? `${index + 1}. ` : '• ';
        return new Paragraph(prefix + item);
      });
    }
    const text = htmlToPlainText(block.html);
    if (!text) return null;
    const heading = getDocxHeading(block);
    return text
      .split('\n')
      .map(line => line.trim())
      .filter(Boolean)
      .map(line => heading ? new Paragraph({ text: line, heading }) : new Paragraph(line));
  }

  function getPdfStyle(block) {
    if (!block || !block.tag) return 'body';
    if (block.classes && block.classes.includes('doc-title')) return 'heading1';
    if (block.classes && block.classes.includes('section-heading')) return 'heading3';
    switch (block.tag.toLowerCase()) {
      case 'h1':
        return 'heading1';
      case 'h2':
        return 'heading2';
      case 'h3':
        return 'heading3';
      default:
        return 'body';
    }
  }

  function blockToPdfNodes(block) {
    if (!block) return null;
    if (isTableBlock(block)) {
      const rows = extractTableRows(block.html);
      if (!rows.length) return null;
      const widths = new Array(rows[0].length || 1).fill('*');
      return [{
        table: {
          headerRows: rows.length > 1 ? 1 : 0,
          widths,
          body: rows
        },
        layout: 'lightHorizontalLines',
        margin: [0, 6, 0, 6]
      }];
    }
    const listData = extractListData(block.html);
    if (listData) {
      return [{
        [listData.ordered ? 'ol' : 'ul']: listData.items,
        margin: [0, 4, 0, 4],
        style: 'body'
      }];
    }
    const text = htmlToPlainText(block.html);
    if (!text) return null;
    const style = getPdfStyle(block);
    return text
      .split('\n')
      .map(line => line.trim())
      .filter(Boolean)
      .map(line => ({ text: line, style, margin: [0, 4, 0, 4] }));
  }

  function exportDocx() {
    const performExport = () => {
      if (!window.docx) {
        showSnackbar('DOCX generator unavailable');
        return;
      }
      const model = buildDocumentModel();
      const pages = Array.isArray(model.pages) ? model.pages : [];
      const sections = [];
      const contentPages = pages.filter(page => page.type !== 'footnotes');
      contentPages.forEach(page => {
        const children = [];
        (page.blocks || []).forEach(block => {
          const nodes = blockToDocxNodes(block);
          if (Array.isArray(nodes)) {
            nodes.forEach(node => node && children.push(node));
          } else if (nodes) {
            children.push(nodes);
          }
        });
        if (children.length) {
          sections.push({ properties: {}, children });
        }
      });
      const footnotesPage = pages.find(page => page.type === 'footnotes');
      if (footnotesPage && Array.isArray(footnotesPage.footnotes) && footnotesPage.footnotes.length) {
        const heading = new window.docx.Paragraph({
          text: 'Footnotes',
          heading: window.docx.HeadingLevel.HEADING_2
        });
        const children = [heading];
        footnotesPage.footnotes.forEach((item, index) => {
          const text = htmlToPlainText(item.html);
          if (text) {
            children.push(new window.docx.Paragraph(`${index + 1}. ${text}`));
          }
        });
        sections.push({ properties: {}, children });
      }
      if (!sections.length) {
        sections.push({ properties: {}, children: [new window.docx.Paragraph('')] });
      }
      const docFile = new window.docx.Document({ sections });
      window.docx.Packer.toBlob(docFile)
        .then(blob => {
          const url = URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          a.download = 'executive-brief.docx';
          document.body.appendChild(a);
          a.click();
          a.remove();
          URL.revokeObjectURL(url);
          showSnackbar('DOCX exported');
        })
        .catch(err => {
          console.error(err);
          showSnackbar('DOCX export failed');
        });
    };

    runQualityCheck({ onContinue: performExport });
  }

  function exportEditablePdf() {
    const performExport = () => {
      if (!window.pdfMake) {
        showSnackbar('PDF generator unavailable');
        return;
      }
      const model = buildDocumentModel();
      const pages = Array.isArray(model.pages) ? model.pages : [];
      const content = [];
      const contentPages = pages.filter(page => page.type !== 'footnotes');
      contentPages.forEach((page, index) => {
        (page.blocks || []).forEach(block => {
          const nodes = blockToPdfNodes(block);
          if (Array.isArray(nodes)) {
            nodes.forEach(node => node && content.push(node));
          } else if (nodes) {
            content.push(nodes);
          }
        });
        if (index < contentPages.length - 1) {
          content.push({ text: '', pageBreak: 'after' });
        }
      });
      const footnotesPage = pages.find(page => page.type === 'footnotes');
      if (footnotesPage && Array.isArray(footnotesPage.footnotes) && footnotesPage.footnotes.length) {
        content.push({ text: 'Footnotes', style: 'heading3', margin: [0, 16, 0, 4] });
        footnotesPage.footnotes.forEach((item, index) => {
          const text = htmlToPlainText(item.html);
          if (text) {
            content.push({ text: `${index + 1}. ${text}`, style: 'footnote', margin: [0, 2, 0, 2] });
          }
        });
      }
      if (!content.length) {
        content.push({ text: '' });
      }
      const docDefinition = {
        info: { title: 'Executive Brief' },
        content,
        styles: {
          heading1: { fontSize: 20, bold: true, margin: [0, 12, 0, 6] },
          heading2: { fontSize: 16, bold: true, margin: [0, 10, 0, 4] },
          heading3: { fontSize: 14, bold: true, margin: [0, 8, 0, 4] },
          body: { fontSize: 11, margin: [0, 4, 0, 4] },
          footnote: { fontSize: 9, italics: true }
        },
        defaultStyle: { fontSize: 11 }
      };
      window.pdfMake.createPdf(docDefinition).download('executive-brief.pdf');
      showSnackbar('Editable PDF exported');
    };

    runQualityCheck({ onContinue: performExport });
  }

  function scheduleInputSnapshot() {
    if (suppressSnapshots || inputSnapshotTimer) return;
    captureSnapshot();
    inputSnapshotTimer = setTimeout(() => {
      afterChange();
      inputSnapshotTimer = null;
    }, 400);
  }

  /* -------- Global click handler -------- */

  document.addEventListener('click', e => {
    const target = e.target;

    const footnoteReturn = target.closest('.footnote-return');
    if (footnoteReturn) {
      e.preventDefault();
      e.stopPropagation();
      scrollToFootnoteRef(footnoteReturn.dataset.uniqueId);
      return;
    }

    const footnoteRef = target.closest('.footnote-ref');
    if (footnoteRef) {
      e.preventDefault();
      e.stopPropagation();
      scrollToFootnoteNote(footnoteRef.dataset.uniqueId);
      return;
    }

    // Delete icon for blocks (not used on footnotes page)
    if (target.classList.contains('delete-icon')) {
      const block = target.closest('[data-deletable="true"]');
      if (block && editorRoot.contains(block) && !block.closest('.footnotes-page')) {
        captureSnapshot();
        block.remove();
        renumberFootnotes();
        afterChange();
        clearBlockSelection();
      }
      return;
    }

    if (target.classList.contains('steps-add')) {
      const block = target.closest('.steps-block');
      const ol = block && block.querySelector('.styled-steps');
      if (ol) {
        captureSnapshot();
        const li = document.createElement('li');
        li.contentEditable = 'true';
        li.textContent = 'New step…';
        ol.appendChild(li);
        afterChange();
      }
      return;
    }

    if (target.classList.contains('faq-add')) {
      const block = target.closest('.faq-block');
      const ol = block && block.querySelector('.faq-list');
      if (ol) {
        captureSnapshot();
        const li = document.createElement('li');
        li.innerHTML = `
          <div class="faq-q" contenteditable="true">Q: …</div>
          <div class="faq-a" contenteditable="true">A: …</div>
        `;
        ol.appendChild(li);
        afterChange();
      }
      return;
    }

    if (target.closest('.cta-type-switch') && target.dataset.cta) {
      const box = target.closest('.cta-box');
      if (box) {
        captureSnapshot();
        box.classList.remove('cta-blue','cta-green','cta-yellow','cta-red');
        box.classList.add(target.dataset.cta);
        box.querySelectorAll('.cta-type-switch button').forEach(btn => btn.classList.remove('active'));
        target.classList.add('active');
        const titleEl = box.querySelector('.cta-title');
        if (titleEl && !box.dataset.customTitle) {
          const map = {
            'cta-blue':'Info',
            'cta-green':'Action',
            'cta-yellow':'Caution',
            'cta-red':'Critical'
          };
          titleEl.textContent = map[target.dataset.cta] || 'Info';
          box.dataset.generatedTitle = titleEl.textContent;
        }
        afterChange();
      }
      return;
    }

    if (target.closest('.table-controls') && target.dataset.tableAction) {
      handleTableControl(target);
      return;
    }

    if (target.closest('.image-size-controls') && target.dataset.size) {
      const block = target.closest('.image-block');
      if (block) {
        captureSnapshot();
        block.classList.remove('image-size-s','image-size-m','image-size-l');
        block.classList.add('image-size-' + target.dataset.size);
        block.querySelectorAll('.image-size-controls button')
          .forEach(btn => btn.classList.remove('active'));
        target.classList.add('active');
        afterChange();
      }
      return;
    }

    // Page controls
    if (target.dataset.movePageUp !== undefined ||
        target.dataset.movePageDown !== undefined ||
        target.dataset.dupPage !== undefined ||
        target.dataset.delPage !== undefined) {
      const page = target.closest('.page-container');
      if (!page || page.classList.contains('footnotes-page')) return;
      const footnotes = editorRoot.querySelector('.footnotes-page');
      const contentPages = editorRoot.querySelectorAll('.page-container:not(.footnotes-page)');

      if (target.dataset.movePageUp !== undefined) {
        captureSnapshot();
        const prev = page.previousElementSibling;
        if (prev && !prev.classList.contains('footnotes-page')) {
          editorRoot.insertBefore(page, prev);
        }
        updatePageNumbers();
        renumberFootnotes();
        afterChange();
      }

      if (target.dataset.movePageDown !== undefined) {
        captureSnapshot();
        const next = page.nextElementSibling;
        if (next && next !== footnotes) {
          editorRoot.insertBefore(next, page);
        }
        updatePageNumbers();
        renumberFootnotes();
        afterChange();
      }

      if (target.dataset.dupPage !== undefined) {
        captureSnapshot();
        const clone = page.cloneNode(true);
        clone.querySelectorAll('.block-selected').forEach(el => el.classList.remove('block-selected'));
        editorRoot.insertBefore(clone, page.nextElementSibling);
        updatePageNumbers();
        renumberFootnotes();
        initAllTableResizers();
        afterChange();
      }

      if (target.dataset.delPage !== undefined) {
        if (contentPages.length <= 1) {
          showSnackbar('Cannot delete last page');
        } else {
          captureSnapshot();
          page.remove();
          updatePageNumbers();
          renumberFootnotes();
          afterChange();
        }
      }
      return;
    }

    // Floating block actions
    if (target.closest('#block-actions') && target.dataset.action) {
      const action = target.dataset.action;
      if (action === 'move-up') moveBlocks('up');
      else if (action === 'move-down') moveBlocks('down');
      else if (action === 'duplicate') duplicateBlocks();
      else if (action === 'copy') copySelectedBlocks(false);
      else if (action === 'cut') copySelectedBlocks(true);
      else if (action === 'paste') pasteBlocksAfterSelection();
      else if (action === 'move-to-page') moveBlocksToNewPage();
      else if (action === 'delete') deleteBlocks();
      return;
    }

    // Text toolbar clicks (formatting)
    if (target.closest('#text-toolbar')) {
      const cmdBtn = target.closest('[data-cmd]');
      const actionBtn = target.closest('[data-action]');
      if (cmdBtn && cmdBtn.dataset.cmd) {
        applyInlineCmd(cmdBtn.dataset.cmd);
      } else if (actionBtn) {
        if (actionBtn.dataset.action === 'highlight') {
          toggleHighlight();
        } else if (actionBtn.dataset.action === 'link') {
          const input = document.getElementById('link-input');
          if (input) input.focus();
        } else if (actionBtn.dataset.action === 'unlink') {
          removeLink();
        }
      }
      e.stopPropagation();
      return;
    }

    // Click inside editor body
    if (isInEditor(target)) {
      const block = target.closest('[data-deletable="true"]');
      const inMenu = target.closest('#block-actions');
      if (block && !block.closest('.footnotes-page') && !inMenu) {
        if (e.metaKey || e.ctrlKey) {
          block.classList.toggle('block-selected');
        } else {
          clearBlockSelection();
          block.classList.add('block-selected');
        }
        showBlockActions();
      } else if (!inMenu && !target.closest('#text-toolbar')) {
        clearBlockSelection();
      }
    } else if (!target.closest('#block-actions') && !target.closest('#text-toolbar')) {
      clearBlockSelection();
      hideTextToolbar();
    }
  });

  /* -------- Toolbar inputs -------- */

  document.getElementById('paragraph-style').addEventListener('change', function() {
    if (this.value) applyParagraphStyle(this.value);
    this.value = '';
  });

  textToolbar.addEventListener('mousedown', e => {
    if (e.target.closest('input, textarea, select')) return;
    if (e.target.closest('button')) {
      e.preventDefault();
    }
  });

  document.getElementById('link-input').addEventListener('keydown', e => {
    if (e.key === 'Enter') {
      e.preventDefault();
      createLink(e.target.value.trim());
    } else if (e.key === 'Escape') {
      hideTextToolbar();
    }
  });

  document.addEventListener('mouseup', () => setTimeout(updateTextToolbar, 10));
  document.addEventListener('keyup', () => setTimeout(updateTextToolbar, 10));
  window.addEventListener('resize', () => checkAllPageOverflow());

  /* -------- Input handler (CTA + footnotes) -------- */

  editorRoot.addEventListener('input', e => {
    const target = e.target;

    // Track custom CTA title vs auto label
    if (target.classList && target.classList.contains('cta-title')) {
      const box = target.closest('.cta-box');
      if (box && target.textContent.trim() !== (box.dataset.generatedTitle || '').trim()) {
        box.dataset.customTitle = 'true';
      }
    }

    // Edits inside Footnotes page: keep references tidy
    if (target.closest && target.closest('.footnotes-page')) {
      renumberFootnotes();
    }

    scheduleInputSnapshot();
  });

  editorRoot.addEventListener('paste', handlePaste);

  /* -------- Top toolbar buttons -------- */

  document.querySelector('[data-add-page]').addEventListener('click', addPage);
  document.querySelector('[data-add-footnote]').addEventListener('click', addFootnote);

  document.getElementById('theme-select').addEventListener('change', function() {
    captureSnapshot();
    setTheme(this.value === 'ft' ? 'ft' : 'iga');
    afterChange();
  });

  document.getElementById('toggle-overlay').addEventListener('click', () => {
    document.body.classList.toggle('show-overlay');
  });

  document.getElementById('undo-btn').addEventListener('click', undo);
  document.getElementById('redo-btn').addEventListener('click', redo);
  const exportDocxBtn = document.getElementById('export-docx');
  if (exportDocxBtn) exportDocxBtn.addEventListener('click', exportDocx);
  const exportPdfBtn = document.getElementById('export-pdf');
  if (exportPdfBtn) exportPdfBtn.addEventListener('click', exportEditablePdf);
  document.getElementById('export-json').addEventListener('click', exportJSON);
  document.getElementById('import-json').addEventListener('click', () => jsonInput.click());
  document.getElementById('print-btn').addEventListener('click', () => window.print());

  jsonInput.addEventListener('change', function() {
    const file = this.files[0];
    if (file) importJSONFromFile(file);
    this.value = '';
  });

  const INSERT_BLOCK_ACTIONS = {
    section: addSectionBlock,
    text: addTextBlock,
    cta: addCtaBlock,
    steps: addStepsBlock,
    faq: addFaqBlock,
    quote: addQuoteBlock,
    divider: addDividerBlock,
    citation: addCitationBlock,
    'table-glance': addAtAGlanceTable,
    'table-bai': addBAITable,
    'table-custom': openCustomTableModal,
    image: openImageModal
  };

  document.querySelectorAll('[data-add]').forEach(btn => {
    btn.addEventListener('click', () => {
      const type = btn.dataset.add;
      const action = INSERT_BLOCK_ACTIONS[type];
      if (typeof action === 'function') action();
    });
  });

  document.querySelectorAll('[data-template]').forEach(btn => {
    btn.addEventListener('click', () => applyTemplate(btn.dataset.template));
  });

  document.getElementById('quality-check-btn').addEventListener('click', () => runQualityCheck({ showNoIssues: true }));
  document.getElementById('help-btn').addEventListener('click', showHelpModal);

  document.addEventListener('keydown', e => {
    const key = e.key;
    if ((key === 'Enter' || key === ' ') && e.target.classList) {
      if (e.target.classList.contains('delete-icon')) {
        e.preventDefault();
        e.stopPropagation();
        e.target.click();
        return;
      }
      if (e.target.classList.contains('footnote-ref')) {
        e.preventDefault();
        scrollToFootnoteNote(e.target.dataset.uniqueId);
        return;
      }
      if (e.target.classList.contains('footnote-return')) {
        e.preventDefault();
        scrollToFootnoteRef(e.target.dataset.uniqueId);
        return;
      }
    }

    if ((e.metaKey || e.ctrlKey) && !e.shiftKey && key.toLowerCase() === 'z') {
      e.preventDefault();
      undo();
    } else if (
      (e.metaKey || e.ctrlKey) &&
      ((e.shiftKey && key.toLowerCase() === 'z') || key.toLowerCase() === 'y')
    ) {
      e.preventDefault();
      redo();
    }
  });

  /* -------- Initialisation -------- */

  ensureFootnotesPage();
  updatePageNumbers();
  renumberFootnotes();
  initAllTableResizers();
  initHistory();
  promptRestoreDraft();
})();
