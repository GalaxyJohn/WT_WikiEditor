// ==UserScript==
// @name         War Thunder Wiki Tree Editor
// @namespace    wt-tree-editor
// @version      0.2.0
// @description  Browser-side editor for War Thunder Wiki aviation trees
// @match        https://wiki.warthunder.com/*
// @run-at       document-idle
// @grant        none
// ==/UserScript==

(function () {
  'use strict';

  const STORAGE_KEY = `wt-tree-editor:v1:${location.pathname}`;
  const LANGUAGE_KEY = 'wt-tree-editor:lang';
  const TOGGLE_HOTKEY = { ctrlKey: true, shiftKey: true, code: 'KeyE' };
  const UNDO_CODES = new Set(['KeyZ']);
  const REDO_CODES = new Set(['KeyY']);
  const COPY_CODES = new Set(['KeyC']);
  const PASTE_CODES = new Set(['KeyV']);
  const RANK_LABELS = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'XI'];
  const TREE_COLUMN_WIDTH = 132;
  const TREE_COLUMN_ADJUST = 12;
  const TREE_LAYOUT_GAP = 55;
  const TREE_MIN_SECTION_WIDTH = 120;
  const DRAG_THRESHOLD = 6;
  const MAX_UNDO_STEPS = 40;
  const STYLE_SELECTORS = {
    regular: '.wt-tree_item:not(.wt-tree_item--prem):not(.wt-tree_item--squad)',
    premium: '.wt-tree_item--prem',
    squadron: '.wt-tree_item--squad',
  };
  const UI_TEXT = {
    ko: {
      toolbar: {
        panelToggle: '조작기',
        shortHint: 'Ctrl+Shift+E | Ctrl+클릭 다중 선택',
        longHint: '조작기: Ctrl+Shift+E | Ctrl+클릭 다중 선택 | Ctrl+Z / Ctrl+Y',
        editIdle: '편집 모드',
        editActive: '편집 켜짐',
        undo: '이전',
        redo: '다시',
        showArrows: '화살표 보이기',
        hideArrows: '화살표 숨기기',
        reset: '저장 초기화',
        statusReady: '편집 가능',
        statusOff: '편집 모드 꺼짐',
        statusClipboard: '클립보드: {label}',
        statusGroupSelection: '폴더 선택: {count}',
      },
      menu: {
        editCard: '장비 삽입 / 수정',
        clearCell: '장비 삭제',
        copyCard: '장비 복사',
        pasteCard: '장비 붙여넣기',
        copyImageUrl: '이미지 URL 복사',
        groupSelected: '선택 장비 폴더화',
        ungroupCell: '폴더 해제',
        addRowBelow: '행 추가',
        duplicateRow: '행 복제',
        deleteRow: '행 삭제',
        addRankBelow: '랭크 추가',
        deleteRank: '랭크 삭제',
      },
      modal: {
        cardTitle: '장비 카드 편집',
        rankTitle: '랭크 추가',
        unitId: '유닛 ID',
        br: 'BR',
        name: '이름',
        imageUrl: '이미지 URL',
        linkUrl: '링크 URL',
        cardStyle: '카드 스타일',
        namePrefix: '이름 접두어',
        rankLabel: '랭크 이름',
        cancel: '취소',
        save: '저장',
        add: '추가',
      },
      style: {
        regular: '일반',
        premium: '골드 / 프리미엄',
        squadron: '스쿼드론',
      },
      prefix: {
        none: '없음',
        premium: '프리미엄 / 수입 (▄)',
        captured: '노획 / 수입 (␗)',
        event: '이벤트 (○)',
        squadron: '스쿼드론 (◌)',
        special: '특수 (◡)',
        legacyIsrael: '이스라엘 레거시 ()',
      },
      flash: {
        cannotDeleteLastRank: '마지막 랭크는 삭제할 수 없음',
        rankDeleted: '랭크 삭제됨',
        copied: '{label} 복사됨',
        pasted: '{label} 붙여넣음',
        imageCopied: '이미지 URL 복사됨',
        groupOnlyVehicles: '폴더화는 장비 카드끼리만 가능합니다',
        folderCreated: '폴더 생성됨 ({count}개)',
        folderReleased: '폴더 해제됨 ({count}개)',
        arrowsHidden: '화살표 숨김',
        arrowsShown: '화살표 표시',
        undoApplied: '이전 적용됨',
        redoApplied: '다시 적용됨',
        folderOpened: '폴더 열림',
        folderClosed: '폴더 닫힘',
        sameTreeOnly: '드래그 이동은 같은 트리 안에서만 가능합니다',
        folderOnlyVehicles: '폴더 안에는 장비 카드만 넣을 수 있음',
        folderSwapBlocked: '폴더 안 장비는 다른 폴더 셀과 직접 교환할 수 없음',
        vehicleMoved: '장비 이동됨',
      },
      confirm: {
        resetSaved: '저장된 커스텀 트리를 모두 지우고 페이지를 다시 불러올까요?',
      },
      content: {
        folder: '폴더',
        vehicle: '장비',
        folderLabel: '폴더: {name}',
        vehicleLabel: '장비: {name}',
        defaultFolderName: '커스텀 폴더',
      },
      rank: {
        newLabel: '새 랭크',
      },
      language: {
        ko: '한국어',
        en: 'English',
      },
    },
    en: {
      toolbar: {
        panelToggle: 'Editor',
        shortHint: 'Ctrl+Shift+E | Ctrl+Click multi-select',
        longHint: 'Editor: Ctrl+Shift+E | Ctrl+Click multi-select | Ctrl+Z / Ctrl+Y',
        editIdle: 'Edit Mode',
        editActive: 'Editing',
        undo: 'Undo',
        redo: 'Redo',
        showArrows: 'Show Arrows',
        hideArrows: 'Hide Arrows',
        reset: 'Reset Saved',
        statusReady: 'Ready to edit',
        statusOff: 'Edit mode off',
        statusClipboard: 'Clipboard: {label}',
        statusGroupSelection: 'Folder selection: {count}',
      },
      menu: {
        editCard: 'Insert / Edit Vehicle',
        clearCell: 'Delete Vehicle',
        copyCard: 'Copy Vehicle',
        pasteCard: 'Paste Vehicle',
        copyImageUrl: 'Copy Image URL',
        groupSelected: 'Group Selected Vehicles',
        ungroupCell: 'Unpack Folder',
        addRowBelow: 'Add Row',
        duplicateRow: 'Duplicate Row',
        deleteRow: 'Delete Row',
        addRankBelow: 'Add Rank',
        deleteRank: 'Delete Rank',
      },
      modal: {
        cardTitle: 'Edit Vehicle Card',
        rankTitle: 'Add Rank',
        unitId: 'Unit ID',
        br: 'BR',
        name: 'Name',
        imageUrl: 'Image URL',
        linkUrl: 'Link URL',
        cardStyle: 'Card Style',
        namePrefix: 'Name Prefix',
        rankLabel: 'Rank Label',
        cancel: 'Cancel',
        save: 'Save',
        add: 'Add',
      },
      style: {
        regular: 'Regular',
        premium: 'Gold / Premium',
        squadron: 'Squadron',
      },
      prefix: {
        none: 'None',
        premium: 'Premium / Imported (▄)',
        captured: 'Captured / Imported (␗)',
        event: 'Event (○)',
        squadron: 'Squadron (◌)',
        special: 'Special (◡)',
        legacyIsrael: 'Israel Legacy ()',
      },
      flash: {
        cannotDeleteLastRank: 'The last rank cannot be deleted',
        rankDeleted: 'Rank deleted',
        copied: '{label} copied',
        pasted: '{label} pasted',
        imageCopied: 'Image URL copied',
        groupOnlyVehicles: 'Only vehicle cards can be grouped',
        folderCreated: 'Folder created ({count})',
        folderReleased: 'Folder unpacked ({count})',
        arrowsHidden: 'Arrows hidden',
        arrowsShown: 'Arrows shown',
        undoApplied: 'Undo applied',
        redoApplied: 'Redo applied',
        folderOpened: 'Folder opened',
        folderClosed: 'Folder closed',
        sameTreeOnly: 'Dragging only works within the same tree',
        folderOnlyVehicles: 'Only vehicle cards can be moved into folders',
        folderSwapBlocked: 'Items inside folders cannot swap directly with another folder cell',
        vehicleMoved: 'Vehicle moved',
      },
      confirm: {
        resetSaved: 'Clear all saved custom trees and reload the page?',
      },
      content: {
        folder: 'Folder',
        vehicle: 'Vehicle',
        folderLabel: 'Folder: {name}',
        vehicleLabel: 'Vehicle: {name}',
        defaultFolderName: 'Custom Folder',
      },
      rank: {
        newLabel: 'New',
      },
      language: {
        ko: 'Korean',
        en: 'English',
      },
    },
  };
  const PREFIX_OPTIONS = [
    { value: '', labelKey: 'prefix.none' },
    { value: '▄', labelKey: 'prefix.premium' },
    { value: '␗', labelKey: 'prefix.captured' },
    { value: '○', labelKey: 'prefix.event' },
    { value: '◌', labelKey: 'prefix.squadron' },
    { value: '◡', labelKey: 'prefix.special' },
    { value: '', labelKey: 'prefix.legacyIsrael' },
  ];
  const MENU_ITEMS = [
    { action: 'edit-card', labelKey: 'menu.editCard' },
    { action: 'clear-cell', labelKey: 'menu.clearCell' },
    { action: 'copy-card', labelKey: 'menu.copyCard' },
    { action: 'paste-card', labelKey: 'menu.pasteCard' },
    { action: 'copy-image-url', labelKey: 'menu.copyImageUrl' },
    { action: 'group-selected', labelKey: 'menu.groupSelected' },
    { action: 'ungroup-cell', labelKey: 'menu.ungroupCell' },
    { action: 'add-row-below', labelKey: 'menu.addRowBelow' },
    { action: 'duplicate-row', labelKey: 'menu.duplicateRow' },
    { action: 'delete-row', labelKey: 'menu.deleteRow' },
    { action: 'add-rank-below', labelKey: 'menu.addRankBelow' },
    { action: 'delete-rank', labelKey: 'menu.deleteRank' },
  ];

  const state = {
    editMode: false,
    panelVisible: true,
    hoveredCell: null,
    hoveredRow: null,
    selectedCell: null,
    selectedRow: null,
    selectedRankRow: null,
    activeNode: null,
    selectedTree: null,
    contextCell: null,
    contextNode: null,
    contextKind: null,
    modalMode: null,
    pendingRankTarget: null,
    clipboard: null,
    groupSelection: new Set(),
    drag: null,
    skipClick: false,
    store: loadStore(),
    layoutQueue: new WeakSet(),
    saveTimers: new Map(),
    undoStacks: new Map(),
    redoStacks: new Map(),
    flashMessage: '',
    flashTimer: 0,
    language: loadLanguage(),
    ui: {},
  };

  init();

  function init() {
    if (!document.querySelector('#wt-unit-trees')) {
      return;
    }

    injectStyles();
    restoreSavedTrees();
    buildUi();
    bindEvents();
    syncGroupStates(document);
    scheduleLayoutForAll();
    positionHeaderTools();
    updateToolbar();
    requestAnimationFrame(() => positionHeaderTools());
  }

  function loadLanguage() {
    try {
      const language = localStorage.getItem(LANGUAGE_KEY);
      return language === 'en' ? 'en' : 'ko';
    } catch (error) {
      console.warn('[WTTE] Failed to load language', error);
    }
    return 'ko';
  }

  function saveLanguage(language) {
    try {
      localStorage.setItem(LANGUAGE_KEY, language);
    } catch (error) {
      console.warn('[WTTE] Failed to save language', error);
    }
  }

  function t(key, variables = {}) {
    const source = UI_TEXT[state.language] || UI_TEXT.ko;
    const template = key.split('.').reduce((value, part) => (value && typeof value === 'object' ? value[part] : undefined), source);
    if (typeof template !== 'string') {
      return key;
    }
    return template.replace(/\{(\w+)\}/g, (_, token) => String(variables[token] ?? ''));
  }

  function setLanguage(language) {
    if (!UI_TEXT[language] || state.language === language) {
      return;
    }
    state.language = language;
    saveLanguage(language);
    applyLanguageToUi();
  }

  function renderPrefixOptions() {
    const prefixSelect = state.ui.prefixSelect;
    if (!prefixSelect) {
      return;
    }

    const previousValue = prefixSelect.value;
    prefixSelect.innerHTML = '';
    PREFIX_OPTIONS.forEach((option) => {
      const element = document.createElement('option');
      element.value = option.value;
      element.textContent = t(option.labelKey);
      prefixSelect.appendChild(element);
    });
    if (PREFIX_OPTIONS.some((option) => option.value === previousValue)) {
      prefixSelect.value = previousValue;
    }
  }

  function applyLanguageToUi() {
    if (!state.ui.controls) {
      return;
    }

    state.ui.panelToggleButton.textContent = t('toolbar.panelToggle');
    state.ui.headerHint.textContent = t('toolbar.shortHint');
    state.ui.metaHint.textContent = t('toolbar.longHint');
    state.ui.resetButton.textContent = t('toolbar.reset');

    MENU_ITEMS.forEach((item) => {
      const button = state.ui.menuButtons[item.action];
      if (button) {
        button.textContent = t(item.labelKey);
      }
    });

    state.ui.cardTitle.textContent = t('modal.cardTitle');
    state.ui.cardLabels.unitId.textContent = t('modal.unitId');
    state.ui.cardLabels.br.textContent = t('modal.br');
    state.ui.cardLabels.name.textContent = t('modal.name');
    state.ui.cardLabels.imageUrl.textContent = t('modal.imageUrl');
    state.ui.cardLabels.linkUrl.textContent = t('modal.linkUrl');
    state.ui.cardLabels.style.textContent = t('modal.cardStyle');
    state.ui.cardLabels.prefix.textContent = t('modal.namePrefix');
    state.ui.cardButtons.cancel.textContent = t('modal.cancel');
    state.ui.cardButtons.submit.textContent = t('modal.save');

    state.ui.rankTitle.textContent = t('modal.rankTitle');
    state.ui.rankLabel.textContent = t('modal.rankLabel');
    state.ui.rankButtons.cancel.textContent = t('modal.cancel');
    state.ui.rankButtons.submit.textContent = t('modal.add');

    state.ui.styleOptions.regular.textContent = t('style.regular');
    state.ui.styleOptions.premium.textContent = t('style.premium');
    state.ui.styleOptions.squadron.textContent = t('style.squadron');
    renderPrefixOptions();

    state.ui.langButtons.forEach((button) => {
      const isActive = button.dataset.lang === state.language;
      button.classList.toggle('is-active', isActive);
      button.setAttribute('aria-pressed', isActive ? 'true' : 'false');
      button.title = t(`language.${button.dataset.lang}`);
      button.setAttribute('aria-label', t(`language.${button.dataset.lang}`));
    });

    updateToolbar();
  }

  function injectStyles() {
    if (document.getElementById('wtte-styles')) {
      return;
    }

    const style = document.createElement('style');
    style.id = 'wtte-styles';
    style.textContent = `
      .wtte-header-tools {
        display: flex;
        align-items: center;
        gap: 8px;
        font: 13px/1.2 system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        color: #f7f8fb;
        position: fixed;
        top: 16px;
        left: 16px;
        z-index: 100000;
      }
      .wtte-header-tools.wtte-floating {
        top: 16px;
        left: 16px;
      }
      .wtte-panel-toggle {
        border: 1px solid rgba(255, 255, 255, 0.18);
        border-radius: 999px;
        padding: 7px 12px;
        color: inherit;
        background: rgba(12, 16, 22, 0.92);
        box-shadow: 0 12px 28px rgba(0, 0, 0, 0.32);
        backdrop-filter: blur(10px);
        cursor: pointer;
        transition: background 120ms ease, opacity 120ms ease;
        white-space: nowrap;
      }
      .wtte-panel-toggle.is-on {
        background: #b78d37;
        color: #111;
        font-weight: 700;
      }
      .wtte-toolbar-panel {
        position: absolute;
        top: calc(100% + 10px);
        left: 0;
        display: grid;
        gap: 8px;
        min-width: min(520px, calc(100vw - 32px));
        padding: 10px 12px;
        border: 1px solid rgba(255, 255, 255, 0.18);
        border-radius: 14px;
        background: rgba(12, 16, 22, 0.92);
        box-shadow: 0 18px 45px rgba(0, 0, 0, 0.45);
        backdrop-filter: blur(10px);
      }
      .wtte-toolbar-panel[hidden] {
        display: none !important;
      }
      .wtte-header-tools.wtte-floating .wtte-toolbar-panel {
        position: static;
        margin-top: 8px;
      }
      .wtte-toolbar {
        display: flex;
        flex-wrap: wrap;
        align-items: center;
        gap: 6px;
      }
      .wtte-header-tools button,
      .wtte-menu button,
      .wtte-modal button,
      .wtte-modal input,
      .wtte-modal select {
        font: inherit;
      }
      .wtte-toolbar button {
        border: 0;
        border-radius: 10px;
        padding: 7px 11px;
        color: inherit;
        background: rgba(255, 255, 255, 0.1);
        cursor: pointer;
        transition: background 120ms ease, opacity 120ms ease;
        white-space: nowrap;
      }
      .wtte-toolbar button:hover,
      .wtte-menu button:hover,
      .wtte-modal button:hover {
        background: rgba(255, 255, 255, 0.18);
      }
      .wtte-toolbar button:disabled,
      .wtte-menu button:disabled {
        opacity: 0.45;
        cursor: default;
      }
      .wtte-toolbar .wtte-toggle.is-on {
        background: #b78d37;
        color: #111;
        font-weight: 700;
      }
      .wtte-lang-switch {
        display: inline-flex;
        align-items: center;
        gap: 4px;
      }
      .wtte-toolbar .wtte-lang {
        min-width: 40px;
        padding: 7px 9px;
        font-size: 16px;
        line-height: 1;
      }
      .wtte-toolbar .wtte-lang.is-active {
        background: rgba(255, 255, 255, 0.22);
        box-shadow: inset 0 0 0 1px rgba(255, 255, 255, 0.2);
      }
      .wtte-toolbar-meta {
        display: grid;
        gap: 3px;
        min-width: 270px;
      }
      .wtte-status {
        color: rgba(255, 255, 255, 0.78);
      }
      .wtte-hint {
        color: rgba(255, 255, 255, 0.62);
        font-size: 12px;
      }
      @media (max-width: 1320px) {
        .wtte-header-tools:not(.wtte-floating) .wtte-hint:first-of-type {
          display: none;
        }
      }
      .wtte-hover-cell {
        outline: 2px solid rgba(255, 204, 67, 0.95);
        outline-offset: -2px;
        background: rgba(255, 204, 67, 0.08);
      }
      .wtte-hover-row > td {
        background: rgba(72, 159, 255, 0.08);
      }
      .wtte-active-cell {
        outline: 3px solid rgba(89, 196, 255, 0.95);
        outline-offset: -3px;
        background: rgba(89, 196, 255, 0.1);
      }
      .wtte-active-row > td {
        box-shadow: inset 0 0 0 1px rgba(89, 196, 255, 0.55);
      }
      .wtte-group-cell {
        box-shadow: inset 0 0 0 2px rgba(133, 231, 169, 0.9);
        background: rgba(133, 231, 169, 0.1);
      }
      .wtte-drop-target {
        outline: 3px dashed rgba(255, 220, 103, 0.98);
        outline-offset: -3px;
        background: rgba(255, 220, 103, 0.12);
      }
      .wtte-drag-source {
        opacity: 0.45;
      }
      .wtte-drag-ghost {
        position: fixed;
        top: 0;
        left: 0;
        z-index: 100003;
        max-width: 180px;
        pointer-events: none;
        transform: translate3d(-9999px, -9999px, 0) scale(1.02);
        filter: drop-shadow(0 16px 32px rgba(0, 0, 0, 0.42));
        opacity: 0.92;
      }
      body.wtte-edit-mode #wt-unit-trees .wt-tree_item,
      body.wtte-edit-mode #wt-unit-trees .wt-tree_group-folder {
        cursor: grab;
      }
      .wt-tree_group.wtte-folder-closed > .wt-tree_group-items {
        display: none !important;
      }
      .wt-tree_group.wtte-folder-open > .wt-tree_group-items {
        display: block !important;
      }
      body.wtte-edit-mode #wt-unit-trees .wt-tree_item-link {
        pointer-events: none;
      }
      .wt-tree.wtte-hide-arrows .wt-tree_arrows {
        display: none !important;
      }
      .wtte-menu {
        position: fixed;
        z-index: 100001;
        min-width: 240px;
        padding: 6px;
        border: 1px solid rgba(255, 255, 255, 0.16);
        border-radius: 14px;
        background: rgba(12, 16, 22, 0.98);
        box-shadow: 0 18px 45px rgba(0, 0, 0, 0.5);
        backdrop-filter: blur(12px);
      }
      .wtte-menu[hidden] {
        display: none !important;
      }
      .wtte-menu button {
        display: block;
        width: 100%;
        padding: 10px 12px;
        border: 0;
        border-radius: 10px;
        text-align: left;
        color: #f7f8fb;
        background: transparent;
        cursor: pointer;
      }
      .wtte-menu button[hidden] {
        display: none !important;
      }
      .wtte-modal-layer {
        position: fixed;
        inset: 0;
        z-index: 100002;
        display: grid;
        place-items: center;
        padding: 20px;
        background: rgba(4, 8, 12, 0.68);
      }
      .wtte-modal-layer[hidden] {
        display: none !important;
      }
      .wtte-modal {
        width: min(560px, 100%);
        padding: 18px;
        border: 1px solid rgba(255, 255, 255, 0.16);
        border-radius: 18px;
        background: rgba(15, 20, 28, 0.98);
        color: #f7f8fb;
        box-shadow: 0 20px 60px rgba(0, 0, 0, 0.55);
      }
      .wtte-modal h3 {
        margin: 0 0 14px;
        font-size: 18px;
      }
      .wtte-grid {
        display: grid;
        gap: 12px;
        grid-template-columns: repeat(2, minmax(0, 1fr));
      }
      .wtte-grid .wtte-full {
        grid-column: 1 / -1;
      }
      .wtte-field {
        display: grid;
        gap: 6px;
      }
      .wtte-field label {
        color: rgba(255, 255, 255, 0.82);
        font-size: 12px;
      }
      .wtte-field input,
      .wtte-field select {
        width: 100%;
        padding: 10px 12px;
        border: 1px solid rgba(255, 255, 255, 0.12);
        border-radius: 12px;
        background: rgba(255, 255, 255, 0.06);
        color: #fff;
      }
      .wtte-field select option {
        color: #111;
        background: #fff;
      }
      .wtte-actions {
        display: flex;
        justify-content: flex-end;
        gap: 10px;
        margin-top: 16px;
      }
      .wtte-actions button {
        border: 0;
        border-radius: 12px;
        padding: 10px 14px;
        color: inherit;
        background: rgba(255, 255, 255, 0.1);
        cursor: pointer;
      }
      .wtte-actions .wtte-primary {
        background: #b78d37;
        color: #111;
        font-weight: 700;
      }
    `;

    document.head.appendChild(style);
  }

  function buildUi() {
    const controls = document.createElement('div');
    controls.className = 'wtte-header-tools';
    controls.innerHTML = `
      <button type="button" class="wtte-panel-toggle"></button>
      <div class="wtte-hint wtte-header-hint"></div>
      <div class="wtte-toolbar-panel">
        <div class="wtte-toolbar">
          <button type="button" class="wtte-toggle"></button>
          <button type="button" class="wtte-undo" disabled></button>
          <button type="button" class="wtte-redo" disabled></button>
          <button type="button" class="wtte-arrows" disabled></button>
          <button type="button" class="wtte-reset"></button>
          <div class="wtte-lang-switch">
            <button type="button" class="wtte-lang" data-lang="ko" aria-pressed="false">🇰🇷</button>
            <button type="button" class="wtte-lang" data-lang="en" aria-pressed="false">🇺🇸</button>
          </div>
        </div>
        <div class="wtte-toolbar-meta">
          <div class="wtte-status"></div>
          <div class="wtte-hint wtte-meta-hint"></div>
        </div>
      </div>
    `;

    const menu = document.createElement('div');
    menu.className = 'wtte-menu';
    menu.hidden = true;
    MENU_ITEMS.forEach((item) => {
      const button = document.createElement('button');
      button.type = 'button';
      button.dataset.action = item.action;
      menu.appendChild(button);
    });

    const modalLayer = document.createElement('div');
    modalLayer.className = 'wtte-modal-layer';
    modalLayer.hidden = true;
    modalLayer.innerHTML = `
      <div class="wtte-modal wtte-card-modal" data-modal="card" hidden>
        <h3 class="wtte-card-title"></h3>
        <form class="wtte-card-form">
          <div class="wtte-grid">
            <div class="wtte-field">
              <label for="wtte-unit-id" class="wtte-label-unit-id"></label>
              <input id="wtte-unit-id" name="unitId" required>
            </div>
            <div class="wtte-field">
              <label for="wtte-br" class="wtte-label-br"></label>
              <input id="wtte-br" name="br">
            </div>
            <div class="wtte-field wtte-full">
              <label for="wtte-name" class="wtte-label-name"></label>
              <input id="wtte-name" name="name" required>
            </div>
            <div class="wtte-field wtte-full">
              <label for="wtte-image" class="wtte-label-image"></label>
              <input id="wtte-image" name="imageUrl" placeholder="https://.../slots/example.png">
            </div>
            <div class="wtte-field wtte-full">
              <label for="wtte-link" class="wtte-label-link"></label>
              <input id="wtte-link" name="linkUrl" placeholder="/unit/example">
            </div>
            <div class="wtte-field">
              <label for="wtte-style" class="wtte-label-style"></label>
              <select id="wtte-style" name="style">
                <option value="regular"></option>
                <option value="premium"></option>
                <option value="squadron"></option>
              </select>
            </div>
            <div class="wtte-field">
              <label for="wtte-prefix" class="wtte-label-prefix"></label>
              <select id="wtte-prefix" name="prefix"></select>
            </div>
          </div>
          <div class="wtte-actions">
            <button type="button" class="wtte-card-cancel" data-close-modal></button>
            <button type="submit" class="wtte-card-submit wtte-primary"></button>
          </div>
        </form>
      </div>
      <div class="wtte-modal wtte-rank-modal" data-modal="rank" hidden>
        <h3 class="wtte-rank-title"></h3>
        <form class="wtte-rank-form">
          <div class="wtte-field">
            <label for="wtte-rank-label" class="wtte-rank-label-text"></label>
            <input id="wtte-rank-label" name="label" required>
          </div>
          <div class="wtte-actions">
            <button type="button" class="wtte-rank-cancel" data-close-modal></button>
            <button type="submit" class="wtte-rank-submit wtte-primary"></button>
          </div>
        </form>
      </div>
    `;

    const headerWrapper = document.querySelector('.layout-header_wrapper');
    const logo = headerWrapper?.querySelector('.layout-header_logo');
    if (logo) {
      logo.insertAdjacentElement('afterend', controls);
    } else {
      controls.classList.add('wtte-floating');
      document.body.appendChild(controls);
    }

    document.body.append(menu, modalLayer);

    const menuButtons = {};
    menu.querySelectorAll('button[data-action]').forEach((button) => {
      menuButtons[button.dataset.action] = button;
    });

    const toolbar = controls.querySelector('.wtte-toolbar');
    const panel = controls.querySelector('.wtte-toolbar-panel');
    state.ui = {
      controls,
      panel,
      toolbar,
      panelToggleButton: controls.querySelector('.wtte-panel-toggle'),
      headerHint: controls.querySelector('.wtte-header-hint'),
      toggleButton: toolbar.querySelector('.wtte-toggle'),
      undoButton: toolbar.querySelector('.wtte-undo'),
      redoButton: toolbar.querySelector('.wtte-redo'),
      arrowsButton: toolbar.querySelector('.wtte-arrows'),
      resetButton: toolbar.querySelector('.wtte-reset'),
      langButtons: Array.from(toolbar.querySelectorAll('.wtte-lang')),
      metaHint: controls.querySelector('.wtte-meta-hint'),
      status: controls.querySelector('.wtte-status'),
      menu,
      menuButtons,
      modalLayer,
      cardModal: modalLayer.querySelector('[data-modal="card"]'),
      rankModal: modalLayer.querySelector('[data-modal="rank"]'),
      cardForm: modalLayer.querySelector('.wtte-card-form'),
      rankForm: modalLayer.querySelector('.wtte-rank-form'),
      cardTitle: modalLayer.querySelector('.wtte-card-title'),
      rankTitle: modalLayer.querySelector('.wtte-rank-title'),
      cardLabels: {
        unitId: modalLayer.querySelector('.wtte-label-unit-id'),
        br: modalLayer.querySelector('.wtte-label-br'),
        name: modalLayer.querySelector('.wtte-label-name'),
        imageUrl: modalLayer.querySelector('.wtte-label-image'),
        linkUrl: modalLayer.querySelector('.wtte-label-link'),
        style: modalLayer.querySelector('.wtte-label-style'),
        prefix: modalLayer.querySelector('.wtte-label-prefix'),
      },
      rankLabel: modalLayer.querySelector('.wtte-rank-label-text'),
      cardButtons: {
        cancel: modalLayer.querySelector('.wtte-card-cancel'),
        submit: modalLayer.querySelector('.wtte-card-submit'),
      },
      rankButtons: {
        cancel: modalLayer.querySelector('.wtte-rank-cancel'),
        submit: modalLayer.querySelector('.wtte-rank-submit'),
      },
      styleOptions: {
        regular: modalLayer.querySelector('#wtte-style option[value="regular"]'),
        premium: modalLayer.querySelector('#wtte-style option[value="premium"]'),
        squadron: modalLayer.querySelector('#wtte-style option[value="squadron"]'),
      },
      prefixSelect: modalLayer.querySelector('#wtte-prefix'),
    };

    applyLanguageToUi();
  }

  function bindEvents() {
    state.ui.panelToggleButton.addEventListener('click', () => {
      state.panelVisible = !state.panelVisible;
      updateToolbar();
    });
    state.ui.toggleButton.addEventListener('click', () => toggleEditMode());
    state.ui.undoButton.addEventListener('click', () => undoActiveTree());
    state.ui.redoButton.addEventListener('click', () => redoActiveTree());
    state.ui.arrowsButton.addEventListener('click', () => toggleArrowVisibility());
    state.ui.resetButton.addEventListener('click', () => resetSavedTrees());
    state.ui.langButtons.forEach((button) => {
      button.addEventListener('click', () => setLanguage(button.dataset.lang));
    });

    state.ui.menu.addEventListener('click', (event) => {
      const button = event.target.closest('button[data-action]');
      if (!button || button.disabled) {
        return;
      }
      handleMenuAction(button.dataset.action);
    });

    state.ui.cardForm.addEventListener('submit', handleCardSubmit);
    state.ui.rankForm.addEventListener('submit', handleRankSubmit);

    state.ui.modalLayer.addEventListener('click', (event) => {
      if (event.target === state.ui.modalLayer || event.target.matches('[data-close-modal]')) {
        closeModal();
      }
    });

    document.addEventListener('mousedown', handlePointerDown, true);
    document.addEventListener('mousemove', handleMouseMove, true);
    document.addEventListener('mouseup', handlePointerUp, true);
    document.addEventListener('contextmenu', handleContextMenu, true);
    document.addEventListener('click', handleDocumentClick, true);
    document.addEventListener('keydown', handleKeyDown, true);
    document.addEventListener('scroll', hideMenu, true);
    document.addEventListener('dragstart', preventNativeDrag, true);
    document.addEventListener('click', handlePotentialTreeSwitch, true);
    window.addEventListener('resize', () => {
      scheduleLayoutForAll();
      positionHeaderTools();
    });
    window.addEventListener('blur', () => cleanupDrag());
  }

  function preventNativeDrag(event) {
    if (!state.editMode) {
      return;
    }
    if (event.target.closest('#wt-unit-trees')) {
      event.preventDefault();
    }
  }

  function handlePotentialTreeSwitch(event) {
    if (!event.target.closest?.('[data-tree-target], [data-country-id]')) {
      return;
    }
    requestAnimationFrame(() => {
      scheduleLayoutForAll();
      positionHeaderTools();
      updateToolbar();
    });
  }

  function handlePointerDown(event) {
    if (!state.editMode || state.modalMode || event.button !== 0) {
      return;
    }

    if (event.target.closest('.wtte-header-tools, .wtte-menu, .wtte-modal')) {
      return;
    }

    const cell = getTargetCell(event.target);
    if (!cell) {
      return;
    }

    const row = getTargetRow(event.target);
    setSelection({
      cell,
      row,
      rankRow: row ? row.closest('.wt-tree_rank') : null,
    });

    if (event.ctrlKey || event.metaKey) {
      return;
    }

    const contentNode = getTopLevelContent(event.target, cell);
    if (!contentNode) {
      setActiveContent(null);
      return;
    }

    setActiveContent(contentNode);

    state.drag = {
      phase: 'pending',
      sourceCell: cell,
      sourceTree: cell.closest('.unit-tree'),
      sourceNode: contentNode,
      sourceOrigin: getNodeOrigin(contentNode),
      ghost: null,
      targetInfo: null,
      targetMarker: null,
      startX: event.clientX,
      startY: event.clientY,
    };

    event.preventDefault();
  }

  function handleMouseMove(event) {
    if (state.drag) {
      if (state.drag.phase === 'pending') {
        const deltaX = event.clientX - state.drag.startX;
        const deltaY = event.clientY - state.drag.startY;
        if (Math.hypot(deltaX, deltaY) >= DRAG_THRESHOLD) {
          startDrag(event);
        }
      }
      if (state.drag && state.drag.phase === 'active') {
        updateDrag(event);
        return;
      }
    }

    if (!state.editMode || state.modalMode) {
      clearHoverState();
      return;
    }

    const cell = getTargetCell(event.target);
    const row = getTargetRow(event.target);
    setHoverState(cell, row);
  }

  function handlePointerUp(event) {
    if (!state.drag) {
      return;
    }

    const dragState = {
      phase: state.drag.phase,
      sourceCell: state.drag.sourceCell,
      sourceNode: state.drag.sourceNode,
      sourceOrigin: state.drag.sourceOrigin,
      targetInfo: state.drag.targetInfo,
    };
    cleanupDrag();

    if (dragState.phase !== 'active') {
      return;
    }

    if (dragState.targetInfo) {
      moveContentNode(dragState.sourceNode, dragState.sourceOrigin, dragState.targetInfo);
      state.skipClick = true;
    }

    updateToolbar();
  }

  function handleContextMenu(event) {
    if (!state.editMode) {
      return;
    }

    const cell = getTargetCell(event.target);
    const row = getTargetRow(event.target);
    const rankRow = getTargetRankRow(event.target);
    if (!cell && !row && !rankRow) {
      hideMenu();
      return;
    }

    event.preventDefault();
    const selectedCell = cell || getFirstCellFromRow(row);
    const selectedRow = row || (selectedCell ? selectedCell.parentElement : null);
    const resolvedRankRow = rankRow
      || (selectedCell ? selectedCell.closest('.wt-tree_rank') : null)
      || (selectedRow ? selectedRow.closest('.wt-tree_rank') : null);

    setSelection({
      cell: selectedCell,
      row: selectedRow,
      rankRow: resolvedRankRow,
    });

    state.contextCell = selectedCell;
    state.contextNode = getTopLevelContent(event.target, selectedCell) || getTopLevelContentFromCell(selectedCell);
    setActiveContent(state.contextNode);
    state.contextKind = getNodeKind(state.contextNode);

    showMenu(event.clientX, event.clientY);
  }

  function handleDocumentClick(event) {
    if (state.skipClick) {
      state.skipClick = false;
      event.preventDefault();
      return;
    }

    const folderGroup = !event.ctrlKey && !event.metaKey ? getFolderGroupFromTarget(event.target) : null;
    if (folderGroup) {
      const cell = getTargetCell(folderGroup);
      if (cell && state.editMode) {
        const row = cell.parentElement;
        setSelection({
          cell,
          row,
          rankRow: row ? row.closest('.wt-tree_rank') : null,
        });
      }
      setActiveContent(folderGroup);
      event.preventDefault();
      event.stopPropagation();
      hideMenu();
      toggleGroupOpenState(folderGroup);
      return;
    }

    if (!event.target.closest('.wt-tree_group, .wtte-header-tools, .wtte-menu, .wtte-modal')) {
      closeOpenFolders();
    }

    if (state.editMode) {
      const itemLink = event.target.closest('.wt-tree_item-link');
      if (itemLink && itemLink.closest('#wt-unit-trees')) {
        event.preventDefault();
      }
    }

    if (!event.target.closest('.wtte-menu')) {
      hideMenu();
    }

    if (!state.editMode || state.modalMode) {
      return;
    }

    if (event.target.closest('.wtte-header-tools, .wtte-modal, .wtte-menu')) {
      return;
    }

    const cell = getTargetCell(event.target);
    if (!cell) {
      setActiveContent(null);
      clearGroupSelection();
      updateToolbar();
      return;
    }

    const row = getTargetRow(event.target);
    setSelection({
      cell,
      row,
      rankRow: row ? row.closest('.wt-tree_rank') : null,
    });

    setActiveContent(getTopLevelContent(event.target, cell));

    if (event.ctrlKey || event.metaKey) {
      event.preventDefault();
      toggleGroupSelection(cell);
      return;
    }

    clearGroupSelection();
    updateToolbar();
  }

  function handleKeyDown(event) {
    if (matchesToggleHotkey(event)) {
      event.preventDefault();
      toggleEditMode();
      return;
    }

    if (event.key === 'Escape') {
      if (state.modalMode) {
        closeModal();
      } else {
        hideMenu();
        cleanupDrag();
      }
      return;
    }

    if (!state.editMode || isEditableTarget(event.target)) {
      return;
    }

    if (matchesUndoHotkey(event)) {
      event.preventDefault();
      undoActiveTree();
      return;
    }

    if (matchesRedoHotkey(event)) {
      event.preventDefault();
      redoActiveTree();
      return;
    }

    if ((event.ctrlKey || event.metaKey) && COPY_CODES.has(event.code)) {
      event.preventDefault();
      copySelectedContent();
      return;
    }

    if ((event.ctrlKey || event.metaKey) && PASTE_CODES.has(event.code)) {
      event.preventDefault();
      pasteIntoSelectedCell();
    }
  }

  function matchesToggleHotkey(event) {
    return event.ctrlKey === TOGGLE_HOTKEY.ctrlKey
      && event.shiftKey === TOGGLE_HOTKEY.shiftKey
      && event.code === TOGGLE_HOTKEY.code;
  }

  function matchesUndoHotkey(event) {
    return (event.ctrlKey || event.metaKey) && !event.shiftKey && UNDO_CODES.has(event.code);
  }

  function matchesRedoHotkey(event) {
    return (event.ctrlKey || event.metaKey)
      && (REDO_CODES.has(event.code) || (event.shiftKey && UNDO_CODES.has(event.code)));
  }

  function toggleEditMode(force) {
    state.editMode = typeof force === 'boolean' ? force : !state.editMode;
    document.body.classList.toggle('wtte-edit-mode', state.editMode);
    if (!state.editMode) {
      hideMenu();
      clearHoverState();
      clearSelection();
      clearGroupSelection();
      closeModal();
      cleanupDrag();
    }
    updateToolbar();
  }

  function updateToolbar() {
    const activeTree = getPrimaryTree();
    const undoCount = activeTree ? getUndoStack(activeTree).length : 0;
    const redoCount = activeTree ? getRedoStack(activeTree).length : 0;
    const arrowsHidden = Boolean(activeTree?.querySelector('.wt-tree')?.classList.contains('wtte-hide-arrows'));
    const groupCount = getGroupSelectionCells().length;

    state.ui.panel.hidden = !state.panelVisible;
    state.ui.panelToggleButton.classList.toggle('is-on', state.panelVisible);
    state.ui.toggleButton.classList.toggle('is-on', state.editMode);
    state.ui.toggleButton.textContent = state.editMode ? t('toolbar.editActive') : t('toolbar.editIdle');
    state.ui.undoButton.disabled = undoCount === 0;
    state.ui.undoButton.textContent = t('toolbar.undo');
    state.ui.redoButton.disabled = redoCount === 0;
    state.ui.redoButton.textContent = t('toolbar.redo');
    state.ui.arrowsButton.disabled = !activeTree;
    state.ui.arrowsButton.textContent = arrowsHidden ? t('toolbar.showArrows') : t('toolbar.hideArrows');

    if (state.flashMessage) {
      state.ui.status.textContent = state.flashMessage;
      return;
    }

    const parts = [];
    parts.push(state.editMode ? t('toolbar.statusReady') : t('toolbar.statusOff'));
    if (state.clipboard) {
      parts.push(t('toolbar.statusClipboard', { label: getClipboardLabel() }));
    }
    if (groupCount > 0) {
      parts.push(t('toolbar.statusGroupSelection', { count: groupCount }));
    }
    state.ui.status.textContent = parts.join(' | ');
  }

  function flashStatus(message) {
    clearTimeout(state.flashTimer);
    state.flashMessage = message;
    updateToolbar();
    state.flashTimer = setTimeout(() => {
      state.flashMessage = '';
      updateToolbar();
    }, 1800);
  }

  function setHoverState(cell, row) {
    if (state.hoveredCell && state.hoveredCell !== cell) {
      state.hoveredCell.classList.remove('wtte-hover-cell');
    }
    if (state.hoveredRow && state.hoveredRow !== row) {
      state.hoveredRow.classList.remove('wtte-hover-row');
    }

    state.hoveredCell = cell || null;
    state.hoveredRow = row || null;

    if (state.hoveredCell) {
      state.hoveredCell.classList.add('wtte-hover-cell');
    }
    if (state.hoveredRow) {
      state.hoveredRow.classList.add('wtte-hover-row');
    }
  }

  function clearHoverState() {
    if (state.hoveredCell) {
      state.hoveredCell.classList.remove('wtte-hover-cell');
    }
    if (state.hoveredRow) {
      state.hoveredRow.classList.remove('wtte-hover-row');
    }
    state.hoveredCell = null;
    state.hoveredRow = null;
  }

  function setActiveContent(node) {
    state.activeNode = node && node.isConnected ? node : null;
  }

  function setSelection({ cell, row, rankRow }) {
    if (state.selectedCell && state.selectedCell !== cell) {
      state.selectedCell.classList.remove('wtte-active-cell');
    }
    if (state.selectedRow && state.selectedRow !== row) {
      state.selectedRow.classList.remove('wtte-active-row');
    }

    state.selectedCell = cell || null;
    state.selectedRow = row || (cell ? cell.parentElement : null);
    state.selectedRankRow = rankRow || (state.selectedCell ? state.selectedCell.closest('.wt-tree_rank') : null);
    state.selectedTree = state.selectedCell
      ? state.selectedCell.closest('.unit-tree')
      : state.selectedRow
        ? state.selectedRow.closest('.unit-tree')
        : state.selectedRankRow
          ? state.selectedRankRow.closest('.unit-tree')
          : getVisibleUnitTree();

    if (state.selectedCell) {
      state.selectedCell.classList.add('wtte-active-cell');
    }
    if (state.selectedRow) {
      state.selectedRow.classList.add('wtte-active-row');
    }

    updateToolbar();
  }

  function clearSelection() {
    if (state.selectedCell) {
      state.selectedCell.classList.remove('wtte-active-cell');
    }
    if (state.selectedRow) {
      state.selectedRow.classList.remove('wtte-active-row');
    }

    state.selectedCell = null;
    state.selectedRow = null;
    state.selectedRankRow = null;
    state.activeNode = null;
    state.selectedTree = null;
    updateToolbar();
  }

  function toggleGroupSelection(cell) {
    if (!cell) {
      return;
    }

    const existing = getGroupSelectionCells();
    if (existing.length && existing[0].closest('.unit-tree') !== cell.closest('.unit-tree')) {
      clearGroupSelection();
    }

    if (state.groupSelection.has(cell)) {
      state.groupSelection.delete(cell);
      cell.classList.remove('wtte-group-cell');
    } else {
      state.groupSelection.add(cell);
      cell.classList.add('wtte-group-cell');
    }
    updateToolbar();
  }

  function clearGroupSelection() {
    state.groupSelection.forEach((cell) => {
      if (cell.isConnected) {
        cell.classList.remove('wtte-group-cell');
      }
    });
    state.groupSelection.clear();
  }

  function getGroupSelectionCells() {
    const cells = Array.from(state.groupSelection).filter((cell) => cell && cell.isConnected);
    state.groupSelection = new Set(cells);
    return cells;
  }

  function showMenu(x, y) {
    updateMenuButtons();
    state.ui.menu.hidden = false;
    const { innerWidth, innerHeight } = window;
    const rect = state.ui.menu.getBoundingClientRect();
    const maxX = innerWidth - rect.width - 12;
    const maxY = innerHeight - rect.height - 12;
    state.ui.menu.style.left = `${Math.max(12, Math.min(x, maxX))}px`;
    state.ui.menu.style.top = `${Math.max(12, Math.min(y, maxY))}px`;
  }

  function hideMenu() {
    state.ui.menu.hidden = true;
    clearContext();
  }

  function clearContext() {
    state.contextCell = null;
    state.contextNode = null;
    state.contextKind = null;
  }

  function updateMenuButtons() {
    const cell = ensureSelectedCell();
    const row = ensureSelectedRow();
    const contextNode = getContextContentNode();
    const hasContent = Boolean(getTopLevelContentFromCell(cell));
    const hasImage = Boolean(extractContentImageUrl(contextNode));
    const canPaste = Boolean(state.clipboard && cell);
    const canGroup = canGroupSelectedCells();
    const canUngroup = Boolean(cell && getDirectChild(cell, '.wt-tree_group'));

    setMenuActionState('edit-card', Boolean(cell), true);
    setMenuActionState('clear-cell', Boolean(cell && hasContent), true);
    setMenuActionState('copy-card', Boolean(contextNode), true);
    setMenuActionState('paste-card', canPaste, true);
    setMenuActionState('copy-image-url', hasImage, Boolean(contextNode));
    setMenuActionState('group-selected', canGroup, Boolean(cell));
    setMenuActionState('ungroup-cell', canUngroup, Boolean(cell));
    setMenuActionState('add-row-below', Boolean(row), true);
    setMenuActionState('duplicate-row', Boolean(row), true);
    setMenuActionState('delete-row', Boolean(row), true);
    setMenuActionState('add-rank-below', Boolean(state.selectedRankRow), true);
    setMenuActionState('delete-rank', canDeleteSelectedRank(), Boolean(state.selectedRankRow));
  }

  function setMenuActionState(action, enabled, visible) {
    const button = state.ui.menuButtons[action];
    if (!button) {
      return;
    }
    button.hidden = !visible;
    button.disabled = !enabled;
  }

  function handleMenuAction(action) {
    hideMenu();

    switch (action) {
      case 'edit-card':
        openCardModal();
        break;
      case 'clear-cell':
        clearSelectedCell();
        break;
      case 'copy-card':
        copySelectedContent();
        break;
      case 'paste-card':
        pasteIntoSelectedCell();
        break;
      case 'copy-image-url':
        void copyContextImageUrl();
        break;
      case 'group-selected':
        groupSelectedCells();
        break;
      case 'ungroup-cell':
        ungroupSelectedCell();
        break;
      case 'add-row-below':
        addRowBelow();
        break;
      case 'duplicate-row':
        duplicateRow();
        break;
      case 'delete-row':
        deleteRow();
        break;
      case 'add-rank-below':
        openRankModal();
        break;
      case 'delete-rank':
        deleteSelectedRank();
        break;
      default:
        break;
    }
  }

  function openCardModal() {
    const cell = ensureSelectedCell();
    if (!cell) {
      return;
    }

    const data = readCardData(cell, getContextContentNode());
    state.modalMode = 'card';
    state.ui.modalLayer.hidden = false;
    state.ui.cardModal.hidden = false;
    state.ui.rankModal.hidden = true;

    state.ui.cardForm.unitId.value = data.unitId || buildDefaultUnitId(data.name);
    state.ui.cardForm.name.value = data.name || '';
    state.ui.cardForm.br.value = data.br || '';
    state.ui.cardForm.imageUrl.value = data.imageUrl || '';
    state.ui.cardForm.linkUrl.value = data.linkUrl || '';
    state.ui.cardForm.style.value = data.style || 'regular';
    state.ui.cardForm.prefix.value = data.prefix || '';
    state.ui.cardForm.unitId.focus();
    state.ui.cardForm.unitId.select();
  }

  function openRankModal() {
    if (!state.selectedRankRow || !state.selectedRankRow.isConnected) {
      return;
    }

    state.pendingRankTarget = state.selectedRankRow;
    state.modalMode = 'rank';
    state.ui.modalLayer.hidden = false;
    state.ui.cardModal.hidden = true;
    state.ui.rankModal.hidden = false;
    state.ui.rankForm.label.value = guessNextRankLabel(state.selectedRankRow);
    state.ui.rankForm.label.focus();
    state.ui.rankForm.label.select();
  }

  function closeModal() {
    state.modalMode = null;
    state.pendingRankTarget = null;
    state.ui.modalLayer.hidden = true;
    state.ui.cardModal.hidden = true;
    state.ui.rankModal.hidden = true;
  }

  function handleCardSubmit(event) {
    event.preventDefault();
    const cell = ensureSelectedCell();
    if (!cell) {
      closeModal();
      return;
    }

    const formData = new FormData(state.ui.cardForm);
    const name = String(formData.get('name') || '').trim();
    const unitId = normalizeUnitId(String(formData.get('unitId') || '').trim() || buildDefaultUnitId(name));
    const cardData = {
      name,
      unitId,
      br: String(formData.get('br') || '').trim(),
      imageUrl: String(formData.get('imageUrl') || '').trim(),
      linkUrl: normalizeLinkUrl(String(formData.get('linkUrl') || '').trim(), unitId),
      style: String(formData.get('style') || 'regular'),
      prefix: String(formData.get('prefix') || ''),
    };

    const unitTree = cell.closest('.unit-tree');
    const targetNode = getContextContentNode();
    recordUndo(unitTree);

    if (targetNode?.matches('.wt-tree_group')) {
      updateGroupFolder(targetNode, cardData, cell);
      setActiveContent(targetNode);
    } else if (targetNode?.matches('.wt-tree_item')) {
      const replacement = buildCardElement(cardData, cell);
      targetNode.replaceWith(replacement);
      setActiveContent(replacement);
      if (isGroupChildItem(replacement)) {
        syncGroupAppearance(replacement.closest('.wt-tree_group'));
      }
    } else {
      setCellCard(cell, cardData);
      setActiveContent(getDirectChild(cell, '.wt-tree_item'));
    }

    finalizeTreeChange(unitTree);
    closeModal();
  }

  function handleRankSubmit(event) {
    event.preventDefault();
    if (!state.pendingRankTarget || !state.pendingRankTarget.isConnected) {
      closeModal();
      return;
    }

    const label = String(new FormData(state.ui.rankForm).get('label') || '').trim();
    if (!label) {
      return;
    }

    addRankBelow(state.pendingRankTarget, label);
    closeModal();
  }

  function clearSelectedCell() {
    const cell = ensureSelectedCell();
    if (!cell) {
      return;
    }

    const unitTree = cell.closest('.unit-tree');
    const targetNode = getContextContentNode() || getTopLevelContentFromCell(cell);
    recordUndo(unitTree);
    if (!targetNode || targetNode.parentElement === cell) {
      cell.innerHTML = '';
      setActiveContent(null);
    } else if (isGroupChildItem(targetNode)) {
      const group = targetNode.closest('.wt-tree_group');
      targetNode.remove();
      normalizeGroupAfterMutation(group);
      setActiveContent(null);
    }
    clearGroupSelection();
    finalizeTreeChange(unitTree);
  }

  function addRowBelow() {
    const row = ensureSelectedRow();
    if (!row) {
      return;
    }

    const unitTree = row.closest('.unit-tree');
    recordUndo(unitTree);

    const newRow = createBlankRowLike(row);
    row.insertAdjacentElement('afterend', newRow);
    clearGroupSelection();
    setSelection({
      cell: newRow.cells[0] || null,
      row: newRow,
      rankRow: newRow.closest('.wt-tree_rank'),
    });
    finalizeTreeChange(unitTree);
  }

  function duplicateRow() {
    const row = ensureSelectedRow();
    if (!row) {
      return;
    }

    const unitTree = row.closest('.unit-tree');
    recordUndo(unitTree);

    const newRow = row.cloneNode(true);
    stripEditorClasses(newRow);
    row.insertAdjacentElement('afterend', newRow);
    clearGroupSelection();
    setSelection({
      cell: newRow.cells[0] || null,
      row: newRow,
      rankRow: newRow.closest('.wt-tree_rank'),
    });
    finalizeTreeChange(unitTree);
  }

  function deleteRow() {
    const row = ensureSelectedRow();
    if (!row) {
      return;
    }

    const unitTree = row.closest('.unit-tree');
    const table = row.closest('table');
    recordUndo(unitTree);

    if (table.rows.length <= 1) {
      Array.from(row.cells).forEach((cell) => {
        cell.innerHTML = '';
      });
      clearGroupSelection();
      finalizeTreeChange(unitTree);
      return;
    }

    const fallback = row.previousElementSibling || row.nextElementSibling;
    row.remove();
    clearGroupSelection();
    setSelection({
      cell: fallback ? fallback.cells[0] || null : null,
      row: fallback || null,
      rankRow: fallback ? fallback.closest('.wt-tree_rank') : null,
    });
    finalizeTreeChange(unitTree);
  }

  function addRankBelow(rankRow, label) {
    const header = getRankHeader(rankRow);
    if (!header) {
      return;
    }

    const unitTree = rankRow.closest('.unit-tree');
    recordUndo(unitTree);

    const headerClone = header.cloneNode(true);
    const rankClone = rankRow.cloneNode(true);
    stripEditorClasses(headerClone);
    stripEditorClasses(rankClone);

    const labelNode = headerClone.querySelector('.wt-tree_r-header_label span');
    if (labelNode) {
      labelNode.textContent = label;
    }

    rankClone.querySelectorAll('td').forEach((cell) => {
      cell.innerHTML = '';
    });

    rankRow.insertAdjacentElement('afterend', rankClone);
    rankClone.insertAdjacentElement('beforebegin', headerClone);

    const newCell = rankClone.querySelector('td');
    clearGroupSelection();
    setSelection({
      cell: newCell,
      row: newCell ? newCell.parentElement : null,
      rankRow: rankClone,
    });
    finalizeTreeChange(unitTree);
  }

  function canDeleteSelectedRank() {
    const rankRow = state.selectedRankRow;
    if (!rankRow || !rankRow.isConnected || !getRankHeader(rankRow)) {
      return false;
    }
    const unitTree = rankRow.closest('.unit-tree');
    return getRankRows(unitTree).length > 1;
  }

  function deleteSelectedRank() {
    const rankRow = state.selectedRankRow;
    const header = getRankHeader(rankRow);
    const unitTree = rankRow?.closest('.unit-tree');
    if (!rankRow || !header || !unitTree) {
      return;
    }

    const rankRows = getRankRows(unitTree);
    if (rankRows.length <= 1) {
      flashStatus(t('flash.cannotDeleteLastRank'));
      return;
    }

    const fallbackRank = getNeighborRankRow(rankRow);
    recordUndo(unitTree);
    header.remove();
    rankRow.remove();
    clearGroupSelection();

    const fallbackCell = fallbackRank?.querySelector('td') || null;
    setSelection({
      cell: fallbackCell,
      row: fallbackCell ? fallbackCell.parentElement : null,
      rankRow: fallbackRank || null,
    });
    finalizeTreeChange(unitTree);
    flashStatus(t('flash.rankDeleted'));
  }

  function copySelectedContent() {
    const cell = ensureSelectedCell();
    const node = getContextContentNode() || getTopLevelContentFromCell(cell);
    if (!node) {
      return;
    }

    const clone = node.cloneNode(true);
    stripEditorClasses(clone);
    state.clipboard = {
      type: getNodeKind(clone),
      html: clone.outerHTML,
      name: getContentNodeName(clone),
    };
    flashStatus(t('flash.copied', { label: getClipboardLabel() }));
    updateToolbar();
  }

  function pasteIntoSelectedCell() {
    const cell = ensureSelectedCell();
    if (!cell || !state.clipboard) {
      return;
    }

    const parsed = htmlToElement(state.clipboard.html);
    if (!parsed) {
      return;
    }

    const unitTree = cell.closest('.unit-tree');
    recordUndo(unitTree);
    stripEditorClasses(parsed);
    cell.innerHTML = '';
    cell.appendChild(parsed);
    clearGroupSelection();
    finalizeTreeChange(unitTree);
    flashStatus(t('flash.pasted', { label: getClipboardLabel() }));
  }

  async function copyContextImageUrl() {
    const node = getContextContentNode();
    const imageUrl = extractContentImageUrl(node);
    if (!imageUrl) {
      return;
    }

    const copied = await copyTextToClipboard(imageUrl);
    if (copied) {
      flashStatus(t('flash.imageCopied'));
    }
  }

  function canGroupSelectedCells() {
    const cells = getGroupSelectionCells();
    if (cells.length < 2) {
      return false;
    }

    const tree = cells[0].closest('.unit-tree');
    return cells.every((cell) => cell.closest('.unit-tree') === tree && Boolean(getDirectChild(cell, '.wt-tree_item')));
  }

  function groupSelectedCells() {
    const cells = getGroupSelectionCells().slice();
    if (cells.length < 2) {
      return;
    }

    cells.sort(compareNodes);
    const unitTree = cells[0].closest('.unit-tree');
    const itemNodes = cells
      .map((cell) => getDirectChild(cell, '.wt-tree_item'))
      .filter(Boolean);

    if (itemNodes.length < 2 || itemNodes.length !== cells.length) {
      flashStatus(t('flash.groupOnlyVehicles'));
      return;
    }

    recordUndo(unitTree);
    const targetCell = cells[0];
    const group = buildGroupElement(itemNodes, targetCell);

    targetCell.innerHTML = '';
    targetCell.appendChild(group);
    cells.slice(1).forEach((cell) => {
      cell.innerHTML = '';
    });

    clearGroupSelection();
    setSelection({
      cell: targetCell,
      row: targetCell.parentElement,
      rankRow: targetCell.closest('.wt-tree_rank'),
    });
    finalizeTreeChange(unitTree);
    flashStatus(t('flash.folderCreated', { count: itemNodes.length }));
  }

  function ungroupSelectedCell() {
    const cell = ensureSelectedCell();
    const group = getDirectChild(cell, '.wt-tree_group');
    if (!cell || !group) {
      return;
    }

    const itemsContainer = group.querySelector('.wt-tree_group-items');
    const items = Array.from(itemsContainer?.children || []).filter((child) => child.matches('.wt-tree_item'));
    if (!items.length) {
      return;
    }

    const row = cell.parentElement;
    const table = row.closest('table');
    const unitTree = cell.closest('.unit-tree');
    const columnIndex = getCellIndex(row, cell);
    recordUndo(unitTree);

    const rows = [row];
    let insertionPoint = row;
    for (let index = 1; index < items.length; index += 1) {
      const newRow = createBlankRowLike(row);
      insertionPoint.insertAdjacentElement('afterend', newRow);
      insertionPoint = newRow;
      rows.push(newRow);
    }

    rows.forEach((targetRow, index) => {
      while (targetRow.cells.length < row.cells.length) {
        targetRow.appendChild(createCellLike(row.cells[targetRow.cells.length] || row.cells[0], false));
      }
      const targetCell = targetRow.cells[columnIndex] || targetRow.cells[targetRow.cells.length - 1];
      targetCell.innerHTML = '';
      targetCell.appendChild(items[index].cloneNode(true));
    });

    clearGroupSelection();
    setSelection({
      cell: table.rows[row.sectionRowIndex].cells[columnIndex] || null,
      row: table.rows[row.sectionRowIndex] || row,
      rankRow: row.closest('.wt-tree_rank'),
    });
    finalizeTreeChange(unitTree);
    flashStatus(t('flash.folderReleased', { count: items.length }));
  }

  function toggleArrowVisibility() {
    const unitTree = getPrimaryTree();
    const wtTree = unitTree?.querySelector('.wt-tree');
    if (!unitTree || !wtTree) {
      return;
    }

    wtTree.classList.toggle('wtte-hide-arrows');
    finalizeTreeChange(unitTree);
    flashStatus(wtTree.classList.contains('wtte-hide-arrows') ? t('flash.arrowsHidden') : t('flash.arrowsShown'));
  }

  function setCellCard(cell, data) {
    const card = buildCardElement(data, cell);
    cell.innerHTML = '';
    cell.appendChild(card);
  }

  function buildCardElement(data, contextCell) {
    const element = getTemplateForStyle(data.style, contextCell);
    element.classList.remove('wt-tree_item--prem', 'wt-tree_item--squad');

    if (data.style === 'premium') {
      element.classList.add('wt-tree_item--prem');
    }
    if (data.style === 'squadron') {
      element.classList.add('wt-tree_item--squad');
    }

    element.dataset.unitId = data.unitId;
    element.dataset.wtteCustom = '1';
    element.dataset.wtteName = data.name;
    element.dataset.wttePrefix = data.prefix;
    element.dataset.wtteStyle = data.style;
    element.dataset.wtteBr = data.br;
    element.dataset.wtteImageUrl = data.imageUrl;
    element.dataset.wtteLinkUrl = data.linkUrl;
    element.removeAttribute('data-unit-req');

    const icon = element.querySelector('.wt-tree_item-icon');
    if (icon) {
      icon.style.backgroundImage = data.imageUrl ? `url("${data.imageUrl}")` : '';
    }

    let link = element.querySelector('.wt-tree_item-link');
    if (!link) {
      link = document.createElement('a');
      element.appendChild(link);
    }
    link.className = 'wt-tree_item-link';
    link.setAttribute('href', data.linkUrl);

    let textContainer = element.querySelector('.wt-tree_item-text');
    if (!textContainer) {
      textContainer = document.createElement('div');
      textContainer.className = 'wt-tree_item-text';
      element.appendChild(textContainer);
    }
    textContainer.innerHTML = '';
    const span = document.createElement('span');
    span.textContent = formatDisplayName(data.name, data.prefix);
    textContainer.appendChild(span);

    let brNode = element.querySelector('.br');
    if (!brNode) {
      brNode = document.createElement('span');
      brNode.className = 'br';
      element.appendChild(brNode);
    }
    brNode.textContent = data.br;

    return element;
  }

  function buildGroupElement(itemNodes, contextCell) {
    const group = getGroupTemplate(contextCell);
    const itemsContainer = ensureGroupStructure(group);
    const folderInfo = extractItemData(itemNodes[0]);
    const groupName = buildGroupName(itemNodes);
    const groupId = `${buildDefaultUnitId(groupName)}_group`;
    const groupStyle = inferGroupStyleFromItems(itemNodes);

    Array.from(itemsContainer.children)
      .filter((child) => !child.matches('.wt-tree_group-canvas'))
      .forEach((child) => child.remove());

    itemNodes.forEach((itemNode) => {
      const clone = itemNode.cloneNode(true);
      stripEditorClasses(clone);
      clone.removeAttribute('data-unit-req');
      itemsContainer.appendChild(clone);
    });

    updateGroupFolder(group, {
      unitId: groupId,
      name: groupName,
      br: '',
      imageUrl: folderInfo.imageUrl,
      linkUrl: normalizeLinkUrl('', groupId),
      style: groupStyle,
      prefix: '',
    }, contextCell);

    applyGroupOpenState(group, false);
    group.dataset.wtteGroupSize = String(itemNodes.length);
    return group;
  }

  function updateGroupFolder(group, data, contextCell) {
    const itemsContainer = ensureGroupStructure(group);
    let folder = group.querySelector('.wt-tree_group-folder');
    if (!folder) {
      folder = document.createElement('div');
      folder.className = 'wt-tree_group-folder';
      group.insertBefore(folder, itemsContainer);
    }

    let inner = folder.querySelector('.wt-tree_group-folder_inner');
    if (!inner) {
      inner = document.createElement('div');
      inner.className = 'wt-tree_group-folder_inner';
      folder.appendChild(inner);
    }

    let icon = inner.querySelector('.wt-tree_item-icon');
    if (!icon) {
      icon = document.createElement('div');
      icon.className = 'wt-tree_item-icon';
      inner.appendChild(icon);
    }
    icon.style.backgroundImage = data.imageUrl ? `url("${data.imageUrl}")` : '';

    let textContainer = inner.querySelector('.wt-tree_item-text');
    if (!textContainer) {
      textContainer = document.createElement('div');
      textContainer.className = 'wt-tree_item-text';
      inner.appendChild(textContainer);
    }
    textContainer.innerHTML = '';
    const span = document.createElement('span');
    span.textContent = formatDisplayName(data.name, data.prefix);
    textContainer.appendChild(span);

    group.dataset.unitId = data.unitId;
    group.dataset.wtteCustom = '1';
    group.dataset.wtteName = data.name;
    group.dataset.wttePrefix = data.prefix;
    group.dataset.wtteStyle = data.style || 'regular';
    group.dataset.wtteImageUrl = data.imageUrl;
    group.dataset.wtteLinkUrl = data.linkUrl;
    group.removeAttribute('data-unit-req');
    group.classList.remove('wt-tree_group--prem', 'wt-tree_group--squad');
    if (data.style === 'premium') {
      group.classList.add('wt-tree_group--prem');
    }
    if (data.style === 'squadron') {
      group.classList.add('wt-tree_group--squad');
    }

    if (!group.querySelector('.wt-tree_group-items .wt-tree_group-canvas')) {
      const canvas = document.createElement('canvas');
      canvas.className = 'wt-tree_group-canvas';
      itemsContainer.insertBefore(canvas, itemsContainer.firstChild);
    }

    if (contextCell && !group.parentElement) {
      contextCell.appendChild(group);
    }
  }

  function ensureGroupStructure(group) {
    if (!group.classList.contains('wt-tree_group')) {
      group.className = 'wt-tree_group';
    }

    let itemsContainer = group.querySelector('.wt-tree_group-items');
    if (!itemsContainer) {
      itemsContainer = document.createElement('div');
      itemsContainer.className = 'wt-tree_group-items';
      group.appendChild(itemsContainer);
    }

    if (!itemsContainer.querySelector('.wt-tree_group-canvas')) {
      const canvas = document.createElement('canvas');
      canvas.className = 'wt-tree_group-canvas';
      itemsContainer.insertBefore(canvas, itemsContainer.firstChild);
    }

    return itemsContainer;
  }

  function getGroupItemsContainer(group) {
    return group?.querySelector('.wt-tree_group-items') || null;
  }

  function getGroupItemNodes(group) {
    return Array.from(getGroupItemsContainer(group)?.children || []).filter((child) => child.matches('.wt-tree_item'));
  }

  function isGroupChildItem(node) {
    return Boolean(node?.matches('.wt-tree_item') && node.parentElement?.matches('.wt-tree_group-items'));
  }

  function syncGroupAppearance(group) {
    if (!group?.isConnected) {
      return;
    }
    const data = extractGroupData(group);
    data.style = inferGroupStyle(group);
    updateGroupFolder(group, data, group.parentElement);
    group.dataset.wtteGroupSize = String(getGroupItemNodes(group).length);
  }

  function normalizeGroupAfterMutation(group) {
    if (!group?.isConnected) {
      return null;
    }

    const cell = group.parentElement;
    const items = getGroupItemNodes(group);
    if (!cell) {
      return null;
    }

    if (!items.length) {
      group.remove();
      return null;
    }

    if (items.length === 1) {
      const [item] = items;
      item.remove();
      group.remove();
      cell.innerHTML = '';
      cell.appendChild(item);
      return item;
    }

    syncGroupAppearance(group);
    return group;
  }

  function readCardData(cell, node = null) {
    if (node?.matches('.wt-tree_item')) {
      return extractItemData(node);
    }

    if (node?.matches('.wt-tree_group')) {
      return extractGroupData(node);
    }

    const directItem = getDirectChild(cell, '.wt-tree_item');
    if (directItem) {
      return extractItemData(directItem);
    }

    const group = getDirectChild(cell, '.wt-tree_group');
    if (group) {
      return extractGroupData(group);
    }

    const template = getTemplateForStyle('regular', cell);
    const sample = extractItemData(template);
    sample.unitId = buildDefaultUnitId('custom');
    sample.name = '';
    sample.br = '';
    sample.prefix = '';
    return sample;
  }

  function extractItemData(item) {
    const rawDisplayName = item.dataset.wtteName || getText(item, '.wt-tree_item-text span') || '';
    const parsedName = splitDisplayName(rawDisplayName);

    return {
      unitId: item.dataset.unitId || buildDefaultUnitId(parsedName.name || rawDisplayName),
      name: item.dataset.wtteName || parsedName.name,
      br: item.dataset.wtteBr || getText(item, '.br') || '',
      imageUrl: item.dataset.wtteImageUrl || readBackgroundImage(item.querySelector('.wt-tree_item-icon')),
      linkUrl: item.dataset.wtteLinkUrl || normalizeLinkUrl(item.querySelector('.wt-tree_item-link')?.getAttribute('href') || '', item.dataset.unitId || ''),
      style: item.dataset.wtteStyle || inferItemStyle(item),
      prefix: item.dataset.wttePrefix || parsedName.prefix,
    };
  }

  function extractGroupData(group) {
    const folder = group.querySelector('.wt-tree_group-folder_inner');
    const rawDisplayName = group.dataset.wtteName || folder?.querySelector('.wt-tree_item-text span')?.textContent?.trim() || '';
    const parsedName = splitDisplayName(rawDisplayName);
    return {
      unitId: group.dataset.unitId || buildDefaultUnitId(parsedName.name || 'custom_group'),
      name: group.dataset.wtteName || parsedName.name,
      br: '',
      imageUrl: group.dataset.wtteImageUrl || readBackgroundImage(folder?.querySelector('.wt-tree_item-icon')),
      linkUrl: group.dataset.wtteLinkUrl || normalizeLinkUrl('', group.dataset.unitId || ''),
      style: group.dataset.wtteStyle || inferGroupStyle(group),
      prefix: group.dataset.wttePrefix || parsedName.prefix,
    };
  }

  function splitDisplayName(displayName) {
    const normalized = String(displayName || '').trim();
    const option = PREFIX_OPTIONS.find((prefixOption) => prefixOption.value && normalized.startsWith(prefixOption.value));
    if (!option) {
      return { prefix: '', name: normalized };
    }
    return {
      prefix: option.value,
      name: normalized.slice(option.value.length).trimStart(),
    };
  }

  function inferItemStyle(item) {
    if (item.classList.contains('wt-tree_item--squad')) {
      return 'squadron';
    }
    if (item.classList.contains('wt-tree_item--prem')) {
      return 'premium';
    }
    return 'regular';
  }

  function inferGroupStyle(group) {
    return inferGroupStyleFromItems(getGroupItemNodes(group));
  }

  function inferGroupStyleFromItems(itemNodes) {
    const styles = itemNodes.map((itemNode) => extractItemData(itemNode).style);
    if (!styles.length) {
      return 'regular';
    }
    if (styles.every((style) => style === 'premium')) {
      return 'premium';
    }
    if (styles.every((style) => style === 'squadron')) {
      return 'squadron';
    }
    return 'regular';
  }

  function getTemplateForStyle(style, contextCell) {
    const scope = contextCell?.closest('.unit-tree') || document;
    const selector = STYLE_SELECTORS[style] || STYLE_SELECTORS.regular;
    const source = scope.querySelector(selector)
      || document.querySelector(selector)
      || document.querySelector('.wt-tree_item');

    if (source) {
      const clone = source.cloneNode(true);
      stripEditorClasses(clone);
      return clone;
    }

    const fallback = document.createElement('div');
    fallback.className = 'wt-tree_item';
    fallback.innerHTML = `
      <div class="wt-tree_item-icon"></div>
      <div class="wt-tree_item-text"><span></span></div>
      <a class="wt-tree_item-link" href="#"></a>
      <span class="br"></span>
    `;
    return fallback;
  }

  function getGroupTemplate(contextCell) {
    const scope = contextCell?.closest('.unit-tree') || document;
    const source = scope.querySelector('.wt-tree_group') || document.querySelector('.wt-tree_group');
    if (source) {
      const clone = source.cloneNode(true);
      stripEditorClasses(clone);
      const itemsContainer = ensureGroupStructure(clone);
      Array.from(itemsContainer.children)
        .filter((child) => !child.matches('.wt-tree_group-canvas'))
        .forEach((child) => child.remove());
      return clone;
    }

    const fallback = document.createElement('div');
    fallback.className = 'wt-tree_group';
    fallback.innerHTML = `
      <div class="wt-tree_group-folder">
        <div class="wt-tree_group-folder_inner">
          <div class="wt-tree_item-icon"></div>
          <div class="wt-tree_item-text"><span></span></div>
        </div>
      </div>
      <div class="wt-tree_group-items">
        <canvas class="wt-tree_group-canvas"></canvas>
      </div>
    `;
    return fallback;
  }

  function finalizeTreeChange(unitTree) {
    if (!unitTree) {
      return;
    }
    clearContext();
    syncGroupStates(unitTree);
    scheduleLayout(unitTree);
    scheduleTreeSave(unitTree);
    updateToolbar();
  }

  function scheduleLayoutForAll() {
    document.querySelectorAll('.unit-tree').forEach((tree) => scheduleLayout(tree));
  }

  function scheduleLayout(unitTree) {
    if (!unitTree || state.layoutQueue.has(unitTree)) {
      return;
    }

    state.layoutQueue.add(unitTree);
    requestAnimationFrame(() => {
      state.layoutQueue.delete(unitTree);
      updateTreeLayout(unitTree);
    });
  }

  function updateTreeLayout(unitTree) {
    const wtTree = unitTree.querySelector('.wt-tree');
    const wrapper = wtTree?.querySelector('.wt-tree_wrapper');
    const instance = wrapper?.querySelector('.wt-tree_instance');
    if (!wtTree || !wrapper || !instance) {
      return;
    }

    let leftWidth = 0;
    let rightWidth = 0;

    instance.querySelectorAll('.wt-tree_rank').forEach((rankRow) => {
      const sections = Array.from(rankRow.children).filter((child) => child.matches('div:not(.wt-tree_v-line)'));
      const left = sections[0];
      const right = sections[1];
      if (left) {
        leftWidth = Math.max(leftWidth, measureSection(left));
      }
      if (right) {
        rightWidth = Math.max(rightWidth, measureSection(right));
      }
    });

    leftWidth = Math.max(leftWidth, TREE_MIN_SECTION_WIDTH);
    rightWidth = Math.max(rightWidth, TREE_MIN_SECTION_WIDTH);

    wrapper.style.setProperty('--wt-tree-l-width', `${Math.ceil(leftWidth)}px`);
    wrapper.style.setProperty('--wt-tree-r-width', `${Math.ceil(rightWidth)}px`);
    wrapper.style.setProperty('--wt-tree-minwidth', `${Math.ceil(leftWidth + rightWidth + TREE_LAYOUT_GAP)}px`);
  }

  function measureSection(section) {
    const table = section.querySelector('table.wt-tree_rank-instance');
    if (!table) {
      return 0;
    }
    const columnCount = Math.max(...Array.from(table.rows).map((row) => row.cells.length), 0);
    return calculateTreeSectionWidth(columnCount);
  }

  function calculateTreeSectionWidth(columnCount) {
    if (!columnCount) {
      return 0;
    }
    return columnCount * TREE_COLUMN_WIDTH - TREE_COLUMN_ADJUST;
  }

  function scheduleTreeSave(unitTree) {
    const treeId = unitTree.dataset.treeId;
    if (!treeId) {
      return;
    }

    clearTimeout(state.saveTimers.get(treeId));
    const timer = setTimeout(() => {
      persistTree(unitTree);
      state.saveTimers.delete(treeId);
    }, 100);
    state.saveTimers.set(treeId, timer);
  }

  function persistTree(unitTree) {
    const treeId = unitTree.dataset.treeId;
    const snapshot = serializePersistedTreeState(unitTree);
    if (!treeId || !snapshot) {
      return;
    }

    state.store.version = 2;
    state.store.modifiedTrees[treeId] = snapshot;
    state.store.updatedAt = new Date().toISOString();
    saveStore();
  }

  function restoreSavedTrees() {
    if (!state.store || !state.store.modifiedTrees) {
      return;
    }

    Object.entries(state.store.modifiedTrees).forEach(([treeId, rawSnapshot]) => {
      const unitTree = document.querySelector(`.unit-tree[data-tree-id="${cssEscape(treeId)}"]`);
      const snapshot = normalizePersistedTreeState(rawSnapshot);
      if (!unitTree || !snapshot) {
        return;
      }
      applyPersistedTreeState(unitTree, snapshot);
    });
  }

  function resetSavedTrees() {
    const hasSavedData = Boolean(state.store && Object.keys(state.store.modifiedTrees || {}).length);
    if (!hasSavedData) {
      return;
    }

    const confirmed = window.confirm(t('confirm.resetSaved'));
    if (!confirmed) {
      return;
    }

    localStorage.removeItem(STORAGE_KEY);
    location.reload();
  }

  function loadStore() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) {
        return { version: 2, modifiedTrees: {}, updatedAt: null };
      }
      const parsed = JSON.parse(raw);
      if (parsed && typeof parsed === 'object' && parsed.modifiedTrees && typeof parsed.modifiedTrees === 'object') {
        return parsed;
      }
    } catch (error) {
      console.warn('[WTTE] Failed to load store', error);
    }
    return { version: 2, modifiedTrees: {}, updatedAt: null };
  }

  function saveStore() {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(state.store));
    } catch (error) {
      console.warn('[WTTE] Failed to save store', error);
    }
  }

  function recordUndo(unitTree) {
    if (!unitTree) {
      return;
    }

    const treeId = unitTree.dataset.treeId;
    const snapshot = serializeUndoTreeState(unitTree);
    if (!treeId || !snapshot) {
      return;
    }

    const stack = getUndoStack(unitTree);
    if (stack[stack.length - 1] === snapshot) {
      return;
    }

    getRedoStack(unitTree).length = 0;
    pushHistorySnapshot(stack, snapshot);
    updateToolbar();
  }

  function undoActiveTree() {
    const unitTree = getPrimaryTree();
    if (!unitTree) {
      return;
    }

    const stack = getUndoStack(unitTree);
    const snapshot = stack.pop();
    if (!snapshot) {
      return;
    }

    pushHistorySnapshot(getRedoStack(unitTree), serializeUndoTreeState(unitTree));

    if (!applyUndoTreeState(unitTree, snapshot)) {
      return;
    }

    hideMenu();
    clearHoverState();
    clearSelection();
    clearGroupSelection();
    cleanupDrag();
    scheduleLayout(unitTree);
    scheduleTreeSave(unitTree);
    flashStatus(t('flash.undoApplied'));
    updateToolbar();
  }

  function redoActiveTree() {
    const unitTree = getPrimaryTree();
    if (!unitTree) {
      return;
    }

    const stack = getRedoStack(unitTree);
    const snapshot = stack.pop();
    if (!snapshot) {
      return;
    }

    pushHistorySnapshot(getUndoStack(unitTree), serializeUndoTreeState(unitTree));

    if (!applyUndoTreeState(unitTree, snapshot)) {
      return;
    }

    hideMenu();
    clearHoverState();
    clearSelection();
    clearGroupSelection();
    cleanupDrag();
    scheduleLayout(unitTree);
    scheduleTreeSave(unitTree);
    flashStatus(t('flash.redoApplied'));
    updateToolbar();
  }

  function serializeUndoTreeState(unitTree) {
    const instance = getTreeInstance(unitTree);
    if (!instance) {
      return '';
    }
    const clone = instance.cloneNode(true);
    stripEditorClasses(clone);
    return clone.innerHTML;
  }

  function applyUndoTreeState(unitTree, snapshot) {
    const instance = getTreeInstance(unitTree);
    if (!instance || typeof snapshot !== 'string') {
      return false;
    }
    instance.innerHTML = snapshot;
    stripEditorClasses(instance);
    syncGroupStates(unitTree);
    return true;
  }

  function getUndoStack(unitTree) {
    const treeId = unitTree?.dataset?.treeId;
    if (!treeId) {
      return [];
    }
    if (!state.undoStacks.has(treeId)) {
      state.undoStacks.set(treeId, []);
    }
    return state.undoStacks.get(treeId);
  }

  function getRedoStack(unitTree) {
    const treeId = unitTree?.dataset?.treeId;
    if (!treeId) {
      return [];
    }
    if (!state.redoStacks.has(treeId)) {
      state.redoStacks.set(treeId, []);
    }
    return state.redoStacks.get(treeId);
  }

  function pushHistorySnapshot(stack, snapshot) {
    if (!snapshot || stack[stack.length - 1] === snapshot) {
      return;
    }
    stack.push(snapshot);
    if (stack.length > MAX_UNDO_STEPS) {
      stack.shift();
    }
  }

  function getPrimaryTree() {
    return getVisibleUnitTree() || (state.selectedTree?.isConnected ? state.selectedTree : null) || document.querySelector('.unit-tree');
  }

  function serializePersistedTreeState(unitTree) {
    const wtTree = unitTree.querySelector('.wt-tree');
    const instance = getTreeInstance(unitTree);
    if (!wtTree || !instance) {
      return null;
    }

    const clone = instance.cloneNode(true);
    stripEditorClasses(clone);
    return {
      version: 2,
      instanceHtml: clone.innerHTML,
      arrowsHidden: wtTree.classList.contains('wtte-hide-arrows'),
    };
  }

  function normalizePersistedTreeState(rawSnapshot) {
    if (!rawSnapshot) {
      return null;
    }

    if (typeof rawSnapshot === 'string') {
      const parsed = htmlToElement(rawSnapshot);
      if (!parsed || !parsed.classList.contains('wt-tree')) {
        return null;
      }
      return {
        version: 2,
        instanceHtml: parsed.querySelector('.wt-tree_instance')?.innerHTML || '',
        arrowsHidden: parsed.classList.contains('wtte-hide-arrows'),
      };
    }

    if (typeof rawSnapshot === 'object') {
      if (typeof rawSnapshot.instanceHtml === 'string') {
        return {
          version: 2,
          instanceHtml: rawSnapshot.instanceHtml,
          arrowsHidden: Boolean(rawSnapshot.arrowsHidden),
        };
      }
      if (typeof rawSnapshot.html === 'string') {
        return normalizePersistedTreeState(rawSnapshot.html);
      }
    }

    return null;
  }

  function applyPersistedTreeState(unitTree, snapshot) {
    const wtTree = unitTree.querySelector('.wt-tree');
    const instance = getTreeInstance(unitTree);
    if (!wtTree || !instance || !snapshot) {
      return false;
    }

    instance.innerHTML = snapshot.instanceHtml;
    wtTree.classList.toggle('wtte-hide-arrows', Boolean(snapshot.arrowsHidden));
    stripEditorClasses(instance);
    syncGroupStates(unitTree);
    return true;
  }

  function getTreeInstance(unitTree) {
    return unitTree?.querySelector('.wt-tree .wt-tree_instance') || null;
  }

  function getVisibleUnitTree() {
    return Array.from(document.querySelectorAll('.unit-tree')).find(isTreeVisible) || null;
  }

  function isTreeVisible(unitTree) {
    if (!unitTree) {
      return false;
    }
    const styles = window.getComputedStyle(unitTree);
    return styles.display !== 'none' && styles.visibility !== 'hidden';
  }

  function getTargetCell(node) {
    const cell = node?.closest?.('.wt-tree_rank-instance td');
    return cell && cell.closest('#wt-unit-trees') ? cell : null;
  }

  function getTargetRow(node) {
    const row = node?.closest?.('.wt-tree_rank-instance tr');
    return row && row.closest('#wt-unit-trees') ? row : null;
  }

  function getTargetRankRow(node) {
    const rankRow = node?.closest?.('.wt-tree_rank, .wt-tree_r-header');
    return rankRow && rankRow.closest('#wt-unit-trees')
      ? rankRow.matches('.wt-tree_r-header') ? rankRow.nextElementSibling : rankRow
      : null;
  }

  function getFolderGroupFromTarget(target) {
    const folder = target?.closest?.('.wt-tree_group-folder');
    const group = folder?.closest('.wt-tree_group');
    return group && group.closest('#wt-unit-trees') ? group : null;
  }

  function positionHeaderTools() {
    const controls = state.ui.controls;
    if (!controls || controls.classList.contains('wtte-floating')) {
      return;
    }

    const logo = document.querySelector('.layout-header_logo');
    const header = document.querySelector('.layout-header');
    if (!logo || !header) {
      return;
    }

    const logoRect = logo.getBoundingClientRect();
    const headerRect = header.getBoundingClientRect();
    const controlsHeight = controls.offsetHeight || 32;
    const logoWidth = logoRect.width || 24;
    const logoRight = logoRect.right > logoRect.left ? logoRect.right : logoRect.left + logoWidth;
    const centerY = logoRect.height
      ? logoRect.top + (logoRect.height / 2)
      : headerRect.top + (Math.min(headerRect.height, 56) / 2);
    const top = Math.max(8, Math.round(centerY - (controlsHeight / 2)));
    const left = Math.max(16, Math.round(logoRight + 12));

    controls.style.top = `${top}px`;
    controls.style.left = `${left}px`;
  }

  function syncGroupStates(scope) {
    const root = scope instanceof Element || scope instanceof Document ? scope : document;
    root.querySelectorAll('.wt-tree_group').forEach((group) => {
      const stored = group.dataset.wtteFolderOpen;
      const isOpen = stored === '1';
      applyGroupOpenState(group, isOpen);
    });
  }

  function applyGroupOpenState(group, isOpen) {
    const items = group?.querySelector('.wt-tree_group-items');
    if (!group || !items) {
      return;
    }
    group.dataset.wtteFolderOpen = isOpen ? '1' : '0';
    group.classList.toggle('wtte-folder-open', isOpen);
    group.classList.toggle('wtte-folder-closed', !isOpen);
    items.hidden = !isOpen;
  }

  function closeOpenFolders(exceptGroup = null) {
    document.querySelectorAll('.wt-tree_group[data-wtte-folder-open="1"]').forEach((group) => {
      if (group !== exceptGroup) {
        applyGroupOpenState(group, false);
      }
    });
  }

  function toggleGroupOpenState(group) {
    const nextState = group?.dataset?.wtteFolderOpen === '1' ? false : true;
    closeOpenFolders(nextState ? group : null);
    applyGroupOpenState(group, nextState);
    flashStatus(nextState ? t('flash.folderOpened') : t('flash.folderClosed'));
  }

  function getContentNodeName(node) {
    if (!node) {
      return '';
    }
    if (node.matches('.wt-tree_group')) {
      return node.querySelector('.wt-tree_group-folder_inner .wt-tree_item-text span')?.textContent?.trim() || '';
    }
    return node.querySelector('.wt-tree_item-text span')?.textContent?.trim() || '';
  }

  function describeContent(kind, name) {
    const safeKind = kind === 'folder' ? 'folder' : 'vehicle';
    const safeName = name || t(`content.${safeKind}`);
    return t(`content.${safeKind}Label`, { name: safeName });
  }

  function getClipboardLabel() {
    if (!state.clipboard) {
      return '';
    }
    return describeContent(state.clipboard.type, state.clipboard.name);
  }

  function getTopLevelContent(target, cell = getTargetCell(target)) {
    if (!cell) {
      return null;
    }

    const directItem = getDirectChild(cell, '.wt-tree_item');
    const directGroup = getDirectChild(cell, '.wt-tree_group');
    if (!target || !(target instanceof Element)) {
      return directGroup || directItem;
    }

    const item = target.closest('.wt-tree_item');
    if (item && (item.parentElement === cell || item.parentElement?.matches('.wt-tree_group-items'))) {
      return item;
    }

    const group = target.closest('.wt-tree_group');
    if (group && group.parentElement === cell) {
      return group;
    }

    return directGroup || directItem;
  }

  function getTopLevelContentFromCell(cell) {
    if (!cell || !cell.isConnected) {
      return null;
    }
    return getDirectChild(cell, '.wt-tree_group') || getDirectChild(cell, '.wt-tree_item');
  }

  function getContextContentNode() {
    if (state.contextNode && state.contextNode.isConnected) {
      return state.contextNode;
    }
    if (state.activeNode && state.activeNode.isConnected) {
      return state.activeNode;
    }
    const cell = state.contextCell?.isConnected ? state.contextCell : ensureSelectedCell();
    return getTopLevelContentFromCell(cell);
  }

  function getNodeKind(node) {
    if (!node) {
      return '';
    }
    return node.matches('.wt-tree_group') ? 'folder' : 'equipment';
  }

  function describeContentNode(node) {
    return describeContent(getNodeKind(node), getContentNodeName(node));
  }

  function extractContentImageUrl(node) {
    if (!node) {
      return '';
    }
    const icon = node.matches('.wt-tree_group')
      ? node.querySelector('.wt-tree_group-folder_inner .wt-tree_item-icon')
      : node.querySelector('.wt-tree_item-icon');
    return readBackgroundImage(icon);
  }

  function ensureSelectedCell() {
    if (state.selectedCell && state.selectedCell.isConnected) {
      return state.selectedCell;
    }
    const row = ensureSelectedRow();
    if (!row) {
      return null;
    }
    const firstCell = getFirstCellFromRow(row);
    if (firstCell) {
      setSelection({
        cell: firstCell,
        row,
        rankRow: row.closest('.wt-tree_rank'),
      });
    }
    return firstCell;
  }

  function ensureSelectedRow() {
    if (state.selectedRow && state.selectedRow.isConnected) {
      return state.selectedRow;
    }
    if (state.selectedCell && state.selectedCell.isConnected) {
      return state.selectedCell.parentElement;
    }
    return null;
  }

  function getFirstCellFromRow(row) {
    return row?.cells?.[0] || null;
  }

  function getRankHeader(rankRow) {
    let pointer = rankRow?.previousElementSibling || null;
    while (pointer) {
      if (pointer.matches('.wt-tree_r-header')) {
        return pointer;
      }
      if (pointer.matches('.wt-tree_rank')) {
        break;
      }
      pointer = pointer.previousElementSibling;
    }
    return null;
  }

  function getRankRows(unitTree) {
    return Array.from(unitTree?.querySelectorAll('.wt-tree_rank') || []);
  }

  function getNeighborRankRow(rankRow) {
    let pointer = rankRow?.nextElementSibling || null;
    while (pointer) {
      if (pointer.matches('.wt-tree_rank')) {
        return pointer;
      }
      pointer = pointer.nextElementSibling;
    }

    pointer = rankRow?.previousElementSibling || null;
    while (pointer) {
      if (pointer.matches('.wt-tree_rank')) {
        return pointer;
      }
      pointer = pointer.previousElementSibling;
    }

    return null;
  }

  function guessNextRankLabel(rankRow) {
    const current = getRankHeader(rankRow)?.querySelector('.wt-tree_r-header_label span')?.textContent?.trim() || '';
    const index = RANK_LABELS.indexOf(current);
    if (index >= 0 && RANK_LABELS[index + 1]) {
      return RANK_LABELS[index + 1];
    }
    return current ? `${current}+` : t('rank.newLabel');
  }

  function createBlankRowLike(row) {
    const newRow = row.cloneNode(false);
    Array.from(row.cells).forEach((cell) => {
      newRow.appendChild(createCellLike(cell, false));
    });
    return newRow;
  }

  function createCellLike(sourceCell, keepContent) {
    const cell = sourceCell ? sourceCell.cloneNode(keepContent) : document.createElement('td');
    stripEditorClasses(cell);
    if (!keepContent) {
      cell.innerHTML = '';
    }
    return cell;
  }

  function getCellIndex(row, cell) {
    return Array.from(row.cells).indexOf(cell);
  }

  function getNodeOrigin(node) {
    const cell = getTargetCell(node);
    if (!node || !cell) {
      return null;
    }

    if (isGroupChildItem(node)) {
      const group = node.closest('.wt-tree_group');
      return {
        type: 'group-item',
        cell,
        group,
        index: getGroupItemNodes(group).indexOf(node),
      };
    }

    return {
      type: 'cell',
      cell,
    };
  }

  function getDirectCellContent(cell) {
    return getDirectChild(cell, '.wt-tree_group') || getDirectChild(cell, '.wt-tree_item');
  }

  function getOpenGroupTarget(target, cell) {
    const group = target?.closest?.('.wt-tree_group');
    if (!group || group.parentElement !== cell || group.dataset.wtteFolderOpen !== '1') {
      return null;
    }
    return group;
  }

  function getDropTargetInfo(target) {
    const cell = getTargetCell(target);
    if (!cell || cell.closest('.unit-tree') !== state.drag?.sourceTree) {
      return null;
    }

    const openGroup = getOpenGroupTarget(target, cell);
    if (openGroup && state.drag?.sourceNode?.matches('.wt-tree_item')) {
      return {
        type: 'group',
        cell,
        group: openGroup,
        marker: cell,
      };
    }

    return {
      type: 'cell',
      cell,
      node: getDirectCellContent(cell),
      marker: cell,
    };
  }

  function setCellContent(cell, node) {
    cell.innerHTML = '';
    if (node) {
      cell.appendChild(node);
    }
  }

  function insertItemIntoGroup(group, item, index = null) {
    const itemsContainer = ensureGroupStructure(group);
    const items = getGroupItemNodes(group);
    const reference = Number.isInteger(index) && index >= 0 ? items[index] || null : null;
    if (reference) {
      itemsContainer.insertBefore(item, reference);
    } else {
      itemsContainer.appendChild(item);
    }
  }

  function moveContentNode(sourceNode, sourceOrigin, targetInfo) {
    if (!sourceNode?.isConnected || !sourceOrigin || !targetInfo) {
      return;
    }

    const sourceTree = sourceOrigin.cell.closest('.unit-tree');
    if (!sourceTree || targetInfo.cell.closest('.unit-tree') !== sourceTree) {
      flashStatus(t('flash.sameTreeOnly'));
      return;
    }

    if (targetInfo.type === 'group' && sourceOrigin.type === 'group-item' && targetInfo.group === sourceOrigin.group) {
      return;
    }

    if (targetInfo.type === 'cell' && sourceOrigin.type === 'cell' && targetInfo.cell === sourceOrigin.cell) {
      return;
    }

    if (targetInfo.type === 'group') {
      if (!sourceNode.matches('.wt-tree_item')) {
        flashStatus(t('flash.folderOnlyVehicles'));
        return;
      }

      recordUndo(sourceTree);

      if (sourceOrigin.type === 'cell') {
        setCellContent(sourceOrigin.cell, null);
      } else {
        sourceNode.remove();
      }

      insertItemIntoGroup(targetInfo.group, sourceNode);
      if (sourceOrigin.type === 'group-item') {
        normalizeGroupAfterMutation(sourceOrigin.group);
      }
      syncGroupAppearance(targetInfo.group);
      clearGroupSelection();
      setSelection({
        cell: targetInfo.cell,
        row: targetInfo.cell.parentElement,
        rankRow: targetInfo.cell.closest('.wt-tree_rank'),
      });
      setActiveContent(sourceNode);
      finalizeTreeChange(sourceTree);
      flashStatus(t('flash.vehicleMoved'));
      return;
    }

    const targetCell = targetInfo.cell;
    const targetNode = targetInfo.node && targetInfo.node !== sourceNode ? targetInfo.node : getDirectCellContent(targetCell);
    if (sourceOrigin.type === 'group-item' && targetNode?.matches('.wt-tree_group')) {
      flashStatus(t('flash.folderSwapBlocked'));
      return;
    }

    recordUndo(sourceTree);

    if (sourceOrigin.type === 'cell') {
      const sourceCell = sourceOrigin.cell;
      if (targetNode) {
        targetNode.remove();
      }
      sourceNode.remove();
      setCellContent(targetCell, sourceNode);
      setCellContent(sourceCell, targetNode || null);
    } else {
      sourceNode.remove();
      if (targetNode) {
        targetNode.remove();
        setCellContent(targetCell, sourceNode);
        if (targetNode.matches('.wt-tree_item')) {
          insertItemIntoGroup(sourceOrigin.group, targetNode, sourceOrigin.index);
        }
      } else {
        setCellContent(targetCell, sourceNode);
      }
      normalizeGroupAfterMutation(sourceOrigin.group);
    }

    clearGroupSelection();
    setSelection({
      cell: targetCell,
      row: targetCell.parentElement,
      rankRow: targetCell.closest('.wt-tree_rank'),
    });
    setActiveContent(sourceNode);
    finalizeTreeChange(sourceTree);
    flashStatus(t('flash.vehicleMoved'));
  }

  function startDrag(event) {
    if (!state.drag || !state.drag.sourceNode?.isConnected) {
      cleanupDrag();
      return;
    }

    hideMenu();
    clearHoverState();
    state.drag.phase = 'active';
    state.drag.ghost = createDragGhost(state.drag.sourceNode);
    document.body.appendChild(state.drag.ghost);
    state.drag.sourceNode.classList.add('wtte-drag-source');
    document.body.style.userSelect = 'none';
    updateDrag(event);
  }

  function updateDrag(event) {
    if (!state.drag || state.drag.phase !== 'active' || !state.drag.ghost) {
      return;
    }

    const x = event.clientX;
    const y = event.clientY;
    state.drag.ghost.style.transform = `translate3d(${x + 14}px, ${y + 14}px, 0) scale(1.02)`;
    updateDropTarget(x, y);
  }

  function updateDropTarget(x, y) {
    if (!state.drag) {
      return;
    }

    clearDropTarget();
    const target = document.elementFromPoint(x, y);
    const targetInfo = getDropTargetInfo(target);
    if (!targetInfo) {
      return;
    }

    state.drag.targetInfo = targetInfo;
    state.drag.targetMarker = targetInfo.marker;
    if (targetInfo.marker && (targetInfo.type !== 'cell' || targetInfo.cell !== state.drag.sourceCell || state.drag.sourceOrigin?.type !== 'cell')) {
      targetInfo.marker.classList.add('wtte-drop-target');
    }
  }

  function clearDropTarget() {
    if (state.drag?.targetMarker?.isConnected) {
      state.drag.targetMarker.classList.remove('wtte-drop-target');
    }
    if (state.drag) {
      state.drag.targetInfo = null;
      state.drag.targetMarker = null;
    }
  }

  function cleanupDrag() {
    clearDropTarget();
    if (state.drag?.sourceNode?.isConnected) {
      state.drag.sourceNode.classList.remove('wtte-drag-source');
    }
    if (state.drag?.ghost?.isConnected) {
      state.drag.ghost.remove();
    }
    state.drag = null;
    document.body.style.userSelect = '';
  }

  function createDragGhost(node) {
    const clone = node.cloneNode(true);
    stripEditorClasses(clone);
    const wrapper = document.createElement('div');
    wrapper.className = 'wtte-drag-ghost';
    wrapper.appendChild(clone);
    return wrapper;
  }

  function buildGroupName(itemNodes) {
    const names = itemNodes
      .map((itemNode) => extractItemData(itemNode).name)
      .filter(Boolean);
    if (!names.length) {
      return t('content.defaultFolderName');
    }
    if (names.length === 2) {
      return `${names[0]} / ${names[1]}`;
    }
    return `${names[0]} +${names.length - 1}`;
  }

  function compareNodes(a, b) {
    if (a === b) {
      return 0;
    }
    return a.compareDocumentPosition(b) & Node.DOCUMENT_POSITION_FOLLOWING ? -1 : 1;
  }

  function stripEditorClasses(root) {
    if (!(root instanceof Element)) {
      return;
    }

    root.classList.remove(
      'wtte-hover-cell',
      'wtte-hover-row',
      'wtte-active-cell',
      'wtte-active-row',
      'wtte-group-cell',
      'wtte-drop-target',
      'wtte-drag-source'
    );
    root.querySelectorAll(
      '.wtte-hover-cell, .wtte-hover-row, .wtte-active-cell, .wtte-active-row, .wtte-group-cell, .wtte-drop-target, .wtte-drag-source'
    ).forEach((node) => {
      node.classList.remove(
        'wtte-hover-cell',
        'wtte-hover-row',
        'wtte-active-cell',
        'wtte-active-row',
        'wtte-group-cell',
        'wtte-drop-target',
        'wtte-drag-source'
      );
    });
  }

  function getDirectChild(parent, selector) {
    if (!parent) {
      return null;
    }
    return Array.from(parent.children).find((child) => child.matches(selector)) || null;
  }

  function getText(scope, selector) {
    return scope.querySelector(selector)?.textContent?.trim() || '';
  }

  function readBackgroundImage(node) {
    if (!node) {
      return '';
    }
    const raw = node.style.backgroundImage || window.getComputedStyle(node).backgroundImage || '';
    const match = raw.match(/url\(["']?(.*?)["']?\)/);
    return match ? match[1] : '';
  }

  function formatDisplayName(name, prefix) {
    const normalizedName = String(name || '').trim();
    if (!prefix) {
      return normalizedName;
    }
    return normalizedName.startsWith(prefix) ? normalizedName : `${prefix}${normalizedName}`;
  }

  function normalizeUnitId(value) {
    return String(value || '').replace(/\s+/g, '_').replace(/[^a-zA-Z0-9_\-]/g, '_');
  }

  function buildDefaultUnitId(name) {
    const source = String(name || 'custom_unit').trim().toLowerCase();
    const normalized = normalizeUnitId(source.replace(/\s+/g, '_'));
    return normalized || `custom_unit_${Date.now()}`;
  }

  function normalizeLinkUrl(value, unitId) {
    if (!value) {
      return unitId ? `/unit/${unitId}` : '#';
    }
    return value;
  }

  function htmlToElement(html) {
    const template = document.createElement('template');
    template.innerHTML = String(html || '').trim();
    return template.content.firstElementChild;
  }

  async function copyTextToClipboard(text) {
    try {
      if (navigator.clipboard?.writeText) {
        await navigator.clipboard.writeText(text);
        return true;
      }
    } catch (error) {
      console.warn('[WTTE] Clipboard API failed, using fallback', error);
    }

    const textarea = document.createElement('textarea');
    textarea.value = text;
    textarea.setAttribute('readonly', '');
    textarea.style.position = 'fixed';
    textarea.style.opacity = '0';
    document.body.appendChild(textarea);
    textarea.select();
    const copied = document.execCommand('copy');
    textarea.remove();
    return copied;
  }

  function isEditableTarget(target) {
    if (!(target instanceof Element)) {
      return false;
    }
    return Boolean(target.closest('input, textarea, select, [contenteditable="true"]'));
  }

  function cssEscape(value) {
    if (window.CSS && typeof window.CSS.escape === 'function') {
      return window.CSS.escape(value);
    }
    return String(value).replace(/"/g, '\\"');
  }
})();
