// ==UserScript==
// @name         War Thunder Wiki Tree Editor
// @namespace    wt-tree-editor
// @version      0.4.0
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
  const ATTRIBUTE_FIELDS = [
    { key: 'imageUrl', labelToken: 'attribute.imageUrl' },
    { key: 'name', labelToken: 'attribute.name' },
    { key: 'unitId', labelToken: 'attribute.unitId' },
    { key: 'br', labelToken: 'attribute.br' },
    { key: 'style', labelToken: 'attribute.style' },
    { key: 'prefix', labelToken: 'attribute.prefix' },
    { key: 'linkUrl', labelToken: 'attribute.linkUrl' },
    { key: 'isFolder', labelToken: 'attribute.isFolder' },
  ];
  const ATTRIBUTE_FIELD_KEYS = ATTRIBUTE_FIELDS.map((field) => field.key);
  const ATTRIBUTE_DATA_KEYS = ATTRIBUTE_FIELD_KEYS.filter((key) => key !== 'isFolder');
  const EXTRA_LABELS = {
    ko: {
      'attribute.imageUrl': '이미지 URL',
      'attribute.name': '이름',
      'attribute.unitId': '유닛 ID',
      'attribute.br': 'BR',
      'attribute.style': '카드 스타일',
      'attribute.prefix': '이름 접두어',
      'attribute.linkUrl': '링크 URL',
      'attribute.isFolder': '폴더 여부',
      'menu.resetAttributeCopy': '복사 초기화',
      'menu.selectAllAttributes': '전체 선택',
      'menu.clearAllAttributes': '전체 선택 해제',
      'menu.addRowAbove': '위쪽에 행 추가',
      'menu.addRowBelow': '아래쪽에 행 추가',
      'menu.addRowBothAbove': '위쪽에 행 추가 (양쪽)',
      'menu.addRowBothBelow': '아래쪽에 행 추가 (양쪽)',
      'menu.addColumnLeft': '왼쪽에 열 추가',
      'menu.addColumnRight': '오른쪽에 열 추가',
      'menu.addRankAbove': '위쪽에 랭크 추가',
      'menu.addRankBelow': '아래쪽에 랭크 추가',
      'menu.addCustomTreeBlank': '백지 트리 (1랭크)',
      'menu.addCustomTreeFull': '기본 트리 (9랭크)',
      'menu.duplicateTree': '트리 복제',
      'menu.deleteTree': '트리 삭제',
      'menuGroup.createTree': '트리 추가',
      'confirm.deleteTree': '이 커스텀 트리를 삭제할까요?',
      'flash.treeDuplicated': '트리를 복제했습니다',
      'flash.treeDeleted': '트리를 삭제했습니다',
    },
    en: {
      'attribute.imageUrl': 'Image URL',
      'attribute.name': 'Name',
      'attribute.unitId': 'Unit ID',
      'attribute.br': 'BR',
      'attribute.style': 'Card Style',
      'attribute.prefix': 'Name Prefix',
      'attribute.linkUrl': 'Link URL',
      'attribute.isFolder': 'Folder Type',
      'menu.resetAttributeCopy': 'Reset Copy',
      'menu.selectAllAttributes': 'Select All',
      'menu.clearAllAttributes': 'Clear All',
      'menu.addRowAbove': 'Add Row Above',
      'menu.addRowBelow': 'Add Row Below',
      'menu.addRowBothAbove': 'Add Row Above (Both)',
      'menu.addRowBothBelow': 'Add Row Below (Both)',
      'menu.addColumnLeft': 'Add Column Left',
      'menu.addColumnRight': 'Add Column Right',
      'menu.addRankAbove': 'Add Rank Above',
      'menu.addRankBelow': 'Add Rank Below',
      'menu.addCustomTreeBlank': 'Blank Tree (1 Rank)',
      'menu.addCustomTreeFull': 'Full Tree (9 Ranks)',
      'menu.duplicateTree': 'Duplicate Tree',
      'menu.deleteTree': 'Delete Tree',
      'menuGroup.createTree': 'Create Tree',
      'confirm.deleteTree': 'Delete this custom tree?',
      'flash.treeDuplicated': 'Tree duplicated',
      'flash.treeDeleted': 'Tree deleted',
    },
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
        resetTree: '현재 트리 초기화',
        import: '불러오기',
        export: '내보내기',
        reset: '저장 초기화',
        statusReady: '편집 가능',
        statusOff: '편집 모드 꺼짐',
        statusClipboard: '클립보드: {label}',
        statusGroupSelection: '폴더 선택: {count}',
      },
      menu: {
        editCard: '장비 삽입 / 수정',
        clearCell: '장비 삭제',
        copyAttributes: '속성 복사',
        pasteAttributes: '속성 붙여넣기',
        copyCard: '장비 복사',
        pasteCard: '장비 붙여넣기',
        copyImageUrl: '이미지 URL 복사',
        groupSelected: '선택 장비 폴더화',
        ungroupCell: '폴더 해제',
        addRowBelow: '행 추가',
        duplicateRow: '행 복제',
        deleteRow: '행 삭제',
        addColumnRight: '열 추가',
        duplicateColumn: '열 복제',
        deleteColumn: '열 삭제',
        addRankBelow: '랭크 추가',
        deleteRank: '랭크 삭제',
        hideTab: '숨기기',
        renameTab: '이름 바꾸기',
      },
      menuGroup: {
        properties: '속성',
        equipment: '장비',
        folder: '폴더화',
        row: '행',
        column: '열',
        rank: '랭크',
        tab: '국가 탭',
        show: '보이기',
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
        copyAttributes: '속성 복사',
        pasteAttributes: '속성 붙여넣기',
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
        currentTreeReset: '현재 트리 초기화됨',
        importApplied: '{count}개 트리 불러오기 완료',
        exportReady: '{count}개 트리 내보내기 완료',
        importFailed: '불러오기 실패',
        importPathMismatch: '다른 위키 페이지용 파일은 불러올 수 없음',
        attributesCopied: '속성 복사됨',
        attributesPasted: '속성 붙여넣음',
        noAttributes: '붙여넣을 속성이 없음',
        tabsUpdated: '국가 탭 설정 저장됨',
        customTreeAdded: '커스텀 트리 추가됨',
        cannotHideLastTab: '마지막 보이는 국가는 숨길 수 없음',
      },
      confirm: {
        resetCurrentTree: '현재 보고 있는 트리만 초기화할까요?',
        resetSaved: '저장된 커스텀 트리를 모두 지우고 페이지를 다시 불러올까요?',
      },
      prompt: {
        renameTab: '새 국가 이름을 입력하세요',
      },
      content: {
        folder: '폴더',
        vehicle: '장비',
        folderLabel: '폴더: {name}',
        vehicleLabel: '장비: {name}',
        defaultFolderName: '커스텀 폴더',
        customTreeName: '커스텀 트리 {index}',
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
        resetTree: 'Reset Tree',
        import: 'Import',
        export: 'Export',
        reset: 'Reset Saved',
        statusReady: 'Ready to edit',
        statusOff: 'Edit mode off',
        statusClipboard: 'Clipboard: {label}',
        statusGroupSelection: 'Folder selection: {count}',
      },
      menu: {
        editCard: 'Insert / Edit Vehicle',
        clearCell: 'Delete Vehicle',
        copyAttributes: 'Copy Attributes',
        pasteAttributes: 'Paste Attributes',
        copyCard: 'Copy Vehicle',
        pasteCard: 'Paste Vehicle',
        copyImageUrl: 'Copy Image URL',
        groupSelected: 'Group Selected Vehicles',
        ungroupCell: 'Unpack Folder',
        addRowBelow: 'Add Row',
        duplicateRow: 'Duplicate Row',
        deleteRow: 'Delete Row',
        addColumnRight: 'Add Column',
        duplicateColumn: 'Duplicate Column',
        deleteColumn: 'Delete Column',
        addRankBelow: 'Add Rank',
        deleteRank: 'Delete Rank',
        hideTab: 'Hide',
        renameTab: 'Rename',
      },
      menuGroup: {
        properties: 'Attributes',
        equipment: 'Vehicle',
        folder: 'Folder',
        row: 'Row',
        column: 'Column',
        rank: 'Rank',
        tab: 'Country Tab',
        show: 'Show',
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
        copyAttributes: 'Copy Attributes',
        pasteAttributes: 'Paste Attributes',
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
        currentTreeReset: 'Current tree reset',
        importApplied: 'Imported {count} tree(s)',
        exportReady: 'Exported {count} tree(s)',
        importFailed: 'Import failed',
        importPathMismatch: 'This file belongs to a different wiki page',
        attributesCopied: 'Attributes copied',
        attributesPasted: 'Attributes pasted',
        noAttributes: 'No attributes to paste',
        tabsUpdated: 'Tab settings saved',
        customTreeAdded: 'Custom tree added',
        cannotHideLastTab: 'Cannot hide the last visible tab',
      },
      confirm: {
        resetCurrentTree: 'Reset only the currently visible tree?',
        resetSaved: 'Clear all saved custom trees and reload the page?',
      },
      prompt: {
        renameTab: 'Enter a new country name',
      },
      content: {
        folder: 'Folder',
        vehicle: 'Vehicle',
        folderLabel: 'Folder: {name}',
        vehicleLabel: 'Vehicle: {name}',
        defaultFolderName: 'Custom Folder',
        customTreeName: 'Custom Tree {index}',
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
  const ACTION_ITEMS = {
    'edit-card': 'menu.editCard',
    'clear-cell': 'menu.clearCell',
    'copy-attributes': 'menu.copyAttributes',
    'paste-attributes': 'menu.pasteAttributes',
    'reset-attribute-copy': 'menu.resetAttributeCopy',
    'copy-card': 'menu.copyCard',
    'paste-card': 'menu.pasteCard',
    'copy-image-url': 'menu.copyImageUrl',
    'group-selected': 'menu.groupSelected',
    'ungroup-cell': 'menu.ungroupCell',
    'add-row-above': 'menu.addRowAbove',
    'add-row-below': 'menu.addRowBelow',
    'add-row-both-above': 'menu.addRowBothAbove',
    'add-row-both-below': 'menu.addRowBothBelow',
    'duplicate-row': 'menu.duplicateRow',
    'delete-row': 'menu.deleteRow',
    'add-column-left': 'menu.addColumnLeft',
    'add-column-right': 'menu.addColumnRight',
    'duplicate-column': 'menu.duplicateColumn',
    'delete-column': 'menu.deleteColumn',
    'add-rank-above': 'menu.addRankAbove',
    'add-rank-below': 'menu.addRankBelow',
    'delete-rank': 'menu.deleteRank',
    'add-custom-tree-blank': 'menu.addCustomTreeBlank',
    'add-custom-tree-full': 'menu.addCustomTreeFull',
    'duplicate-tree': 'menu.duplicateTree',
    'delete-tree': 'menu.deleteTree',
    'hide-tab': 'menu.hideTab',
    'rename-tab': 'menu.renameTab',
  };
  const TREE_MENU_GROUPS = [
    { id: 'properties', labelKey: 'menuGroup.properties', actions: ['copy-attributes', 'paste-attributes'] },
    { id: 'equipment', labelKey: 'menuGroup.equipment', actions: ['edit-card', 'clear-cell', 'copy-card', 'paste-card', 'copy-image-url'] },
    { id: 'folder', labelKey: 'menuGroup.folder', actions: ['group-selected', 'ungroup-cell'] },
    {
      id: 'row',
      labelKey: 'menuGroup.row',
      actions: [
        { action: 'add-row-above', labelToken: 'menu.addRowAbove' },
        { action: 'add-row-below', labelToken: 'menu.addRowBelow' },
        { action: 'add-row-both-above', labelToken: 'menu.addRowBothAbove' },
        { action: 'add-row-both-below', labelToken: 'menu.addRowBothBelow' },
        { action: 'duplicate-row' },
        { action: 'delete-row' },
      ],
    },
    {
      id: 'column',
      labelKey: 'menuGroup.column',
      actions: [
        { action: 'add-column-left', labelToken: 'menu.addColumnLeft' },
        { action: 'add-column-right', labelToken: 'menu.addColumnRight' },
        { action: 'duplicate-column' },
        { action: 'delete-column' },
      ],
    },
    {
      id: 'rank',
      labelKey: 'menuGroup.rank',
      actions: [
        { action: 'add-rank-above', labelToken: 'menu.addRankAbove' },
        { action: 'add-rank-below', labelToken: 'menu.addRankBelow' },
        { action: 'delete-rank' },
      ],
    },
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
    contextTabId: '',
    contextMenuMode: 'tree',
    modalMode: null,
    modalTargetNode: null,
    pendingRankTarget: null,
    pendingRankDirection: 'below',
    clipboard: null,
    attributeClipboard: null,
    attributeSelection: new Set(),
    groupSelection: new Set(),
    selectionDrag: null,
    drag: null,
    tabDrag: null,
    skipClick: false,
    store: loadStore(),
    layoutQueue: new WeakSet(),
    saveTimers: new Map(),
    undoStacks: new Map(),
    redoStacks: new Map(),
    flashMessage: '',
    flashTimer: 0,
    language: loadLanguage(),
    originalTrees: {},
    menuAnchor: null,
    ui: {},
  };

  init();

  function init() {
    if (!document.querySelector('#wt-unit-trees')) {
      return;
    }

    injectStyles();
    restoreStoredTabs();
    state.originalTrees = captureOriginalTreeStates();
    restoreSavedTrees();
    buildUi();
    bindEvents();
    ensureAddTabButton();
    syncGroupStates(document);
    document.querySelectorAll('.unit-tree').forEach((tree) => syncArrowVisibility(tree));
    activateTreeTab(getInitialActiveTreeId());
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
    const fallback = EXTRA_LABELS[state.language]?.[key] || EXTRA_LABELS.en[key];
    const resolved = typeof template === 'string' ? template : fallback;
    if (typeof resolved !== 'string') {
      return key;
    }
    return resolved.replace(/\{(\w+)\}/g, (_, token) => String(variables[token] ?? ''));
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
    state.ui.resetTreeButton.textContent = t('toolbar.resetTree');
    state.ui.importButton.textContent = t('toolbar.import');
    state.ui.exportButton.textContent = t('toolbar.export');
    state.ui.resetButton.textContent = t('toolbar.reset');
    renderContextMenu();

    state.ui.cardTitle.textContent = t('modal.cardTitle');
    state.ui.cardLabels.unitId.textContent = t('modal.unitId');
    state.ui.cardLabels.br.textContent = t('modal.br');
    state.ui.cardLabels.name.textContent = t('modal.name');
    state.ui.cardLabels.imageUrl.textContent = t('modal.imageUrl');
    state.ui.cardLabels.linkUrl.textContent = t('modal.linkUrl');
    state.ui.cardLabels.style.textContent = t('modal.cardStyle');
    state.ui.cardLabels.prefix.textContent = t('modal.namePrefix');
    state.ui.cardButtons.copyAttributes.textContent = t('modal.copyAttributes');
    state.ui.cardButtons.pasteAttributes.textContent = t('modal.pasteAttributes');
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
    updateCardAttributeButtons();

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
      .wtte-tab-drag-source {
        opacity: 0.55;
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
      #wt-tree-tabs .navtabs_wrapper,
      #wt-tree-tabs .arrow-scroll_wrapper {
        overflow-x: auto !important;
        overflow-y: hidden !important;
        scrollbar-width: thin;
      }
      #wt-tree-tabs .navtabs_item {
        flex: 0 0 auto;
      }
      .wtte-menu {
        position: absolute;
        z-index: 100001;
        width: 124px;
        min-width: 124px;
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
      .wtte-menu-group {
        position: relative;
      }
      .wtte-menu-group + .wtte-menu-group {
        margin-top: 4px;
      }
      .wtte-menu-trigger,
      .wtte-menu button {
        display: block;
        width: 100%;
        padding: 8px 10px;
        border: 0;
        border-radius: 10px;
        text-align: left;
        color: #f7f8fb;
        background: transparent;
        cursor: pointer;
        white-space: normal;
      }
      .wtte-menu-trigger {
        position: relative;
        padding-right: 24px;
      }
      .wtte-menu-trigger::after {
        content: '›';
        position: absolute;
        top: 50%;
        right: 12px;
        transform: translateY(-50%);
        color: rgba(255, 255, 255, 0.65);
      }
      .wtte-menu-submenu {
        position: absolute;
        top: 0;
        left: calc(100% - 6px);
        width: 160px;
        min-width: 160px;
        max-height: calc(100vh - 24px);
        overflow-y: auto;
        padding: 6px;
        border: 1px solid rgba(255, 255, 255, 0.16);
        border-radius: 14px;
        background: rgba(12, 16, 22, 0.99);
        box-shadow: 0 18px 45px rgba(0, 0, 0, 0.5);
        backdrop-filter: blur(12px);
        display: none;
      }
      .wtte-menu-group.wtte-open-up > .wtte-menu-submenu {
        top: auto;
        bottom: 0;
      }
      .wtte-menu.wtte-open-left .wtte-menu-submenu {
        left: auto;
        right: calc(100% - 6px);
      }
      .wtte-menu-group:hover > .wtte-menu-submenu,
      .wtte-menu-group:focus-within > .wtte-menu-submenu {
        display: block;
      }
      .wtte-menu-group[hidden],
      .wtte-menu button[hidden] {
        display: none !important;
      }
      .wtte-menu-checklist {
        display: grid;
        gap: 4px;
        margin: 6px 0;
      }
      .wtte-menu-check {
        display: flex;
        align-items: center;
        justify-content: flex-start;
        gap: 8px;
        padding: 6px 8px;
        border-radius: 10px;
        color: #f7f8fb;
        cursor: pointer;
        text-align: left;
      }
      .wtte-menu-check:hover {
        background: rgba(255, 255, 255, 0.12);
      }
      .wtte-menu-check input {
        margin: 0;
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
        align-items: center;
        justify-content: space-between;
        gap: 10px;
        margin-top: 16px;
      }
      .wtte-actions-left,
      .wtte-actions-right {
        display: flex;
        align-items: center;
        gap: 10px;
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
      .wtte-tab-hidden {
        display: none !important;
      }
      .wtte-add-tab .navtabs_item-label {
        font-weight: 700;
      }
      .wtte-add-tab {
        user-select: none;
      }
      .wtte-country-label {
        margin-left: 0.35em;
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
          <button type="button" class="wtte-reset-tree" disabled></button>
          <button type="button" class="wtte-import"></button>
          <button type="button" class="wtte-export" disabled></button>
          <button type="button" class="wtte-reset"></button>
          <div class="wtte-lang-switch">
            <button type="button" class="wtte-lang" data-lang="ko" aria-pressed="false">🇰🇷</button>
            <button type="button" class="wtte-lang" data-lang="en" aria-pressed="false">🇺🇸</button>
          </div>
        </div>
        <input type="file" class="wtte-import-input" accept="application/json,.json" hidden>
        <div class="wtte-toolbar-meta">
          <div class="wtte-status"></div>
          <div class="wtte-hint wtte-meta-hint"></div>
        </div>
      </div>
    `;

    const menu = document.createElement('div');
    menu.className = 'wtte-menu';
    menu.hidden = true;

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
            <div class="wtte-actions-left">
              <button type="button" class="wtte-card-copy-attributes"></button>
              <button type="button" class="wtte-card-paste-attributes"></button>
            </div>
            <div class="wtte-actions-right">
              <button type="button" class="wtte-card-cancel" data-close-modal></button>
              <button type="submit" class="wtte-card-submit wtte-primary"></button>
            </div>
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
      resetTreeButton: toolbar.querySelector('.wtte-reset-tree'),
      importButton: toolbar.querySelector('.wtte-import'),
      exportButton: toolbar.querySelector('.wtte-export'),
      resetButton: toolbar.querySelector('.wtte-reset'),
      langButtons: Array.from(toolbar.querySelectorAll('.wtte-lang')),
      importInput: controls.querySelector('.wtte-import-input'),
      metaHint: controls.querySelector('.wtte-meta-hint'),
      status: controls.querySelector('.wtte-status'),
      menu,
      menuButtons: {},
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
        copyAttributes: modalLayer.querySelector('.wtte-card-copy-attributes'),
        pasteAttributes: modalLayer.querySelector('.wtte-card-paste-attributes'),
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

    renderContextMenu();
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
    state.ui.resetTreeButton.addEventListener('click', () => resetCurrentTree());
    state.ui.importButton.addEventListener('click', () => openImportDialog());
    state.ui.exportButton.addEventListener('click', () => exportSavedTrees());
    state.ui.resetButton.addEventListener('click', () => resetSavedTrees());
    state.ui.importInput.addEventListener('change', handleImportFileSelection);
    state.ui.langButtons.forEach((button) => {
      button.addEventListener('click', () => setLanguage(button.dataset.lang));
    });

    state.ui.menu.addEventListener('click', (event) => {
      const button = event.target.closest('button[data-action]');
      if (!button || button.disabled) {
        return;
      }
      handleMenuAction(button.dataset.action, { treeId: button.dataset.treeId || '' });
    });
    state.ui.menu.addEventListener('change', handleMenuChange);

    state.ui.cardForm.addEventListener('submit', handleCardSubmit);
    state.ui.rankForm.addEventListener('submit', handleRankSubmit);
    state.ui.cardButtons.copyAttributes.addEventListener('click', () => copyModalAttributes());
    state.ui.cardButtons.pasteAttributes.addEventListener('click', () => pasteModalAttributes());

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
    document.addEventListener('dragstart', preventNativeDrag, true);
    document.addEventListener('click', handlePotentialTreeSwitch, true);
    document.addEventListener('scroll', handleDocumentScroll, true);
    window.addEventListener('resize', () => {
      scheduleLayoutForAll();
      positionHeaderTools();
      updateContextMenuPlacement({ clampToViewport: true, refreshLayout: true });
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
      const tab = getTargetTreeTab(event.target);
      const countryButton = event.target.closest?.('[data-country-id]');
      if (tab && !tab.classList.contains('wtte-add-tab')) {
        activateTreeTab(tab.dataset.treeTarget);
      } else if (countryButton?.dataset.countryId && !getTreeTabById(countryButton.dataset.countryId)?.classList.contains('wtte-tab-hidden')) {
        activateTreeTab(countryButton.dataset.countryId);
      }
      scheduleLayoutForAll();
      positionHeaderTools();
      updateToolbar();
    });
  }

  function getTargetTreeTab(target) {
    const tab = target?.closest?.('#wt-tree-tabs .navtabs_item[data-tree-target]');
    return tab instanceof Element ? tab : null;
  }

  function updateAddTabVisibility() {
    const addTab = getTreeTabById('wtte-add-tab');
    if (!addTab) {
      return;
    }
    addTab.style.display = state.editMode ? '' : 'none';
  }

  function startTabPointerInteraction(tab, event) {
    if (!tab || tab.classList.contains('wtte-add-tab')) {
      return;
    }
    state.tabDrag = {
      phase: 'pending',
      tab,
      startX: event.clientX,
      startY: event.clientY,
    };
  }

  function updateTabDrag(event) {
    if (!state.tabDrag) {
      return;
    }

    if (state.tabDrag.phase === 'pending') {
      const deltaX = event.clientX - state.tabDrag.startX;
      const deltaY = event.clientY - state.tabDrag.startY;
      if (Math.hypot(deltaX, deltaY) < DRAG_THRESHOLD) {
        return;
      }
      state.tabDrag.phase = 'active';
      state.tabDrag.tab.classList.add('wtte-tab-drag-source');
    }

    const targetTab = getTargetTreeTab(document.elementFromPoint(event.clientX, event.clientY));
    if (!targetTab || targetTab === state.tabDrag.tab || targetTab.classList.contains('wtte-add-tab')) {
      return;
    }

    const wrapper = getTreeTabsWrapper();
    if (!wrapper) {
      return;
    }

    const targetRect = targetTab.getBoundingClientRect();
    const insertAfter = event.clientX > targetRect.left + (targetRect.width / 2);
    wrapper.insertBefore(state.tabDrag.tab, insertAfter ? targetTab.nextElementSibling : targetTab);
  }

  function finishTabDrag(event) {
    if (!state.tabDrag) {
      return;
    }

    const { phase, tab } = state.tabDrag;
    tab.classList.remove('wtte-tab-drag-source');
    state.tabDrag = null;
    if (phase === 'active') {
      reorderTreeDom(collectCurrentTabState().order);
      ensureAddTabButton();
      persistTabState(true);
      state.skipClick = true;
      return;
    }

    if (event) {
      activateTreeTab(tab.dataset.treeTarget);
      state.skipClick = true;
    }
  }

  function handlePointerDown(event) {
    if (!state.editMode || state.modalMode || event.button !== 0) {
      return;
    }

    if (event.target.closest('.wtte-header-tools, .wtte-menu, .wtte-modal')) {
      return;
    }

    const treeTab = getTargetTreeTab(event.target);
    if (treeTab) {
      startTabPointerInteraction(treeTab, event);
      event.preventDefault();
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

    if (event.shiftKey || event.ctrlKey || event.metaKey) {
      startSelectionDrag(cell, event);
      event.preventDefault();
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
    if (state.tabDrag) {
      updateTabDrag(event);
      return;
    }

    if (state.selectionDrag) {
      updateSelectionDrag(event);
      return;
    }

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
    if (state.tabDrag) {
      finishTabDrag(event);
      return;
    }

    if (state.selectionDrag) {
      finishSelectionDrag();
      return;
    }

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

    const treeTab = getTargetTreeTab(event.target);
    const countryButton = event.target.closest?.('[data-country-id]');
    const tabsRoot = getTreeTabsRoot();
    if (treeTab || countryButton || tabsRoot?.contains(event.target)) {
      event.preventDefault();
      if (treeTab?.classList.contains('wtte-add-tab')) {
        state.contextMenuMode = 'add-tab';
        state.contextTabId = '';
      } else {
        state.contextMenuMode = 'tab';
        state.contextTabId = treeTab && !treeTab.classList.contains('wtte-add-tab')
          ? treeTab.dataset.treeTarget
          : countryButton?.dataset.countryId || '';
      }
      showMenu(event.pageX, event.pageY);
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
    state.contextMenuMode = 'tree';
    state.contextTabId = '';

    showMenu(event.pageX, event.pageY);
  }

  function handleDocumentClick(event) {
    if (state.skipClick) {
      state.skipClick = false;
      event.preventDefault();
      return;
    }

    const profileLink = event.target.closest('.layout-header_user a');
    if (profileLink) {
      event.preventDefault();
      event.stopPropagation();
      toggleEditMode();
      return;
    }

    const treeTab = getTargetTreeTab(event.target);
    if (treeTab) {
      event.preventDefault();
      event.stopPropagation();
      hideMenu();
      if (treeTab.classList.contains('wtte-add-tab')) {
        if (state.editMode) {
          state.contextMenuMode = 'add-tab';
          state.contextTabId = '';
          const rect = treeTab.getBoundingClientRect();
          showMenu(window.scrollX + rect.left + (rect.width / 2), window.scrollY + rect.bottom + 6);
        }
      } else {
        activateTreeTab(treeTab.dataset.treeTarget);
      }
      positionHeaderTools();
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

  function handleDocumentScroll() {
    syncContextMenuVisibility();
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
    if (event.code === 'Backspace') {
      event.preventDefault();
      deleteSelectionRange();
      return;
    }
    if (event.code === 'Delete') {
      event.preventDefault();
      deleteSelectionContentsOrRows();
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
    updateAddTabVisibility();
    updateToolbar();
  }

  function updateToolbar() {
    const activeTree = getPrimaryTree();
    const undoCount = activeTree ? getUndoStack(activeTree).length : 0;
    const redoCount = activeTree ? getRedoStack(activeTree).length : 0;
    const arrowsHidden = Boolean(activeTree?.querySelector('.wt-tree')?.classList.contains('wtte-hide-arrows'));
    const groupCount = getGroupSelectionCells().length;
    const hasSavedOrPendingChanges = Boolean(state.saveTimers.size || Object.keys(state.store.modifiedTrees || {}).length || hasTabMetadataChanges());

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
    state.ui.resetTreeButton.disabled = !isTreeModified(activeTree);
    state.ui.exportButton.disabled = !hasSavedOrPendingChanges;
    state.ui.resetButton.disabled = !hasSavedOrPendingChanges;

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

    setGroupSelectionState(cell, !state.groupSelection.has(cell));
    updateToolbar();
  }

  function setGroupSelectionState(cell, selected) {
    if (!cell) {
      return;
    }
    if (selected) {
      state.groupSelection.add(cell);
      cell.classList.add('wtte-group-cell');
    } else {
      state.groupSelection.delete(cell);
      cell.classList.remove('wtte-group-cell');
    }
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

  function startSelectionDrag(cell, event) {
    const existing = getGroupSelectionCells();
    if (existing.length && existing[0].closest('.unit-tree') !== cell.closest('.unit-tree')) {
      clearGroupSelection();
    }

    if (event.shiftKey && !event.ctrlKey && !event.metaKey) {
      clearGroupSelection();
      state.selectionDrag = {
        kind: 'rect',
        anchorCell: cell,
      };
      applyRectangleSelection(cell);
      return;
    }

    state.selectionDrag = {
      kind: 'path',
      mode: state.groupSelection.has(cell) ? 'remove' : 'add',
      visited: new Set(),
      startX: event.clientX,
      startY: event.clientY,
    };
    applySelectionDragCell(cell);
  }

  function applySelectionDragCell(cell) {
    if (!state.selectionDrag || state.selectionDrag.kind !== 'path' || !cell || state.selectionDrag.visited.has(cell)) {
      return;
    }
    state.selectionDrag.visited.add(cell);
    setGroupSelectionState(cell, state.selectionDrag.mode !== 'remove');
    updateToolbar();
  }

  function getSelectionGrid(unitTree) {
    if (!unitTree) {
      return { orderedCells: [], positions: new Map() };
    }

    const rankRows = getRankRows(unitTree);
    const rankEntries = rankRows.map((rankRow) => {
      const tables = getRankRowTables(rankRow);
      return {
        rankRow,
        tables,
        height: Math.max(1, ...tables.map((table) => table?.rows?.length || 0), 1),
      };
    });
    const sectionCount = Math.max(0, ...rankEntries.map((entry) => entry.tables.length));
    const sectionWidths = Array.from({ length: sectionCount }, () => 1);

    rankEntries.forEach(({ tables }) => {
      tables.forEach((table, sectionIndex) => {
        sectionWidths[sectionIndex] = Math.max(sectionWidths[sectionIndex], getTableColumnCount(table) || 1);
      });
    });

    const sectionOffsets = [];
    let columnOffset = 0;
    sectionWidths.forEach((width, sectionIndex) => {
      sectionOffsets[sectionIndex] = columnOffset;
      columnOffset += Math.max(1, width);
    });

    const orderedCells = [];
    const positions = new Map();
    let rowOffset = 0;
    rankEntries.forEach(({ tables, height }) => {
      tables.forEach((table, sectionIndex) => {
        if (!table) {
          return;
        }
        Array.from(table.rows).forEach((row, rowIndex) => {
          Array.from(row.cells).forEach((cell, columnIndex) => {
            const info = {
              cell,
              row: rowOffset + rowIndex,
              column: sectionOffsets[sectionIndex] + columnIndex,
            };
            orderedCells.push(info);
            positions.set(cell, info);
          });
        });
      });
      rowOffset += height;
    });

    return { orderedCells, positions };
  }

  function getRectangleSelectionCells(anchorCell, targetCell) {
    const unitTree = anchorCell?.closest('.unit-tree');
    if (!unitTree || targetCell?.closest('.unit-tree') !== unitTree) {
      return anchorCell ? [anchorCell] : [];
    }

    const { orderedCells, positions } = getSelectionGrid(unitTree);
    const anchor = positions.get(anchorCell);
    const target = positions.get(targetCell);
    if (!anchor || !target) {
      return anchorCell ? [anchorCell] : [];
    }

    const rowMin = Math.min(anchor.row, target.row);
    const rowMax = Math.max(anchor.row, target.row);
    const columnMin = Math.min(anchor.column, target.column);
    const columnMax = Math.max(anchor.column, target.column);

    return orderedCells
      .filter((info) => info.row >= rowMin && info.row <= rowMax && info.column >= columnMin && info.column <= columnMax)
      .map((info) => info.cell);
  }

  function getDistanceToRect(clientX, clientY, rect) {
    const dx = clientX < rect.left ? rect.left - clientX : clientX > rect.right ? clientX - rect.right : 0;
    const dy = clientY < rect.top ? rect.top - clientY : clientY > rect.bottom ? clientY - rect.bottom : 0;
    return (dx * dx) + (dy * dy);
  }

  function getClosestCellInTree(unitTree, clientX, clientY) {
    let bestCell = null;
    let bestDistance = Number.POSITIVE_INFINITY;
    unitTree?.querySelectorAll('.wt-tree_rank-instance td').forEach((cell) => {
      const rect = cell.getBoundingClientRect();
      const distance = getDistanceToRect(clientX, clientY, rect);
      if (distance < bestDistance) {
        bestDistance = distance;
        bestCell = cell;
      }
    });
    return bestCell;
  }

  function getSelectionDragTargetCell(event) {
    const directCell = getTargetCell(event.target) || getTargetCell(document.elementFromPoint(event.clientX, event.clientY));
    if (directCell) {
      return directCell;
    }
    const unitTree = state.selectionDrag?.anchorCell?.closest('.unit-tree');
    return unitTree ? getClosestCellInTree(unitTree, event.clientX, event.clientY) : null;
  }

  function applyRectangleSelection(targetCell) {
    if (!state.selectionDrag || state.selectionDrag.kind !== 'rect' || !state.selectionDrag.anchorCell) {
      return;
    }

    const nextCells = getRectangleSelectionCells(state.selectionDrag.anchorCell, targetCell);
    clearGroupSelection();
    nextCells.forEach((cell) => setGroupSelectionState(cell, true));
    setSelection({
      cell: targetCell,
      row: targetCell?.parentElement || null,
      rankRow: targetCell ? targetCell.closest('.wt-tree_rank') : null,
    });
    updateToolbar();
  }

  function updateSelectionDrag(event) {
    const cell = getSelectionDragTargetCell(event);
    if (!cell) {
      return;
    }
    if (state.selectionDrag?.kind === 'rect') {
      applyRectangleSelection(cell);
      return;
    }
    if (state.selectionDrag?.kind === 'path') {
      applySelectionDragCell(cell);
    }
  }

  function finishSelectionDrag() {
    state.selectionDrag = null;
    state.skipClick = true;
    updateToolbar();
  }

  function renderContextMenu() {
    if (!state.ui.menu) {
      return;
    }

    const menu = state.ui.menu;
    menu.innerHTML = '';
    state.ui.menuButtons = {};
    if (state.contextMenuMode === 'add-tab') {
      getAddTabMenuGroups()
        .flatMap((group) => group.actions)
        .map((item) => normalizeMenuActionItem(item))
        .filter((item) => item && !item.hidden)
        .forEach((item) => {
          const button = createMenuActionButton(item);
          menu.appendChild(button);
          state.ui.menuButtons[item.action] = button;
        });
      return;
    }
    const groups = state.contextMenuMode === 'tab' ? getTabMenuGroups() : TREE_MENU_GROUPS;

    groups.forEach((group) => {
      const visibleActions = group.actions
        .map((item) => normalizeMenuActionItem(item))
        .filter((item) => item && !item.hidden);
      if (!visibleActions.length) {
        return;
      }

      const groupNode = document.createElement('div');
      groupNode.className = 'wtte-menu-group';
      groupNode.dataset.groupId = group.id;

      const trigger = document.createElement('button');
      trigger.type = 'button';
      trigger.className = 'wtte-menu-trigger';
      trigger.setAttribute('aria-haspopup', 'true');
      trigger.textContent = t(group.labelKey);
      groupNode.appendChild(trigger);

      const submenu = document.createElement('div');
      submenu.className = 'wtte-menu-submenu';
      if (state.contextMenuMode === 'tree' && group.id === 'properties') {
        renderAttributeMenuGroup(submenu);
      } else {
        visibleActions.forEach((item) => {
          const button = createMenuActionButton(item);
          submenu.appendChild(button);
          if (!state.ui.menuButtons[item.action]) {
            state.ui.menuButtons[item.action] = button;
          }
        });
      }
      groupNode.appendChild(submenu);
      menu.appendChild(groupNode);
    });

    if (state.contextMenuMode === 'tree') {
      updateMenuButtons();
    }
  }

  function normalizeMenuActionItem(item) {
    if (typeof item === 'string') {
      return { action: item };
    }
    if (item && typeof item === 'object' && typeof item.action === 'string') {
      return item;
    }
    return null;
  }

  function createMenuActionButton(item) {
    const button = document.createElement('button');
    button.type = 'button';
    button.dataset.action = item.action;
    if (item.treeId) {
      button.dataset.treeId = item.treeId;
    }
    button.disabled = Boolean(item.disabled);
    const labelKey = item.labelToken || ACTION_ITEMS[item.action];
    button.textContent = item.label ?? (labelKey ? t(labelKey) : item.action);
    return button;
  }

  function renderAttributeMenuGroup(submenu) {
    const actionItems = [
      { action: 'copy-attributes' },
      { action: 'reset-attribute-copy' },
      { action: 'paste-attributes' },
    ];

    actionItems.forEach((item) => {
      const button = createMenuActionButton(item);
      submenu.appendChild(button);
      if (!state.ui.menuButtons[item.action]) {
        state.ui.menuButtons[item.action] = button;
      }
    });

    const checklist = document.createElement('div');
    checklist.className = 'wtte-menu-checklist';
    ATTRIBUTE_FIELDS.forEach((field) => {
      const label = document.createElement('label');
      label.className = 'wtte-menu-check';

      const input = document.createElement('input');
      input.type = 'checkbox';
      input.dataset.attributeKey = field.key;
      input.checked = state.attributeSelection.has(field.key);

      const text = document.createElement('span');
      text.textContent = t(field.labelToken);

      label.appendChild(input);
      label.appendChild(text);
      checklist.appendChild(label);
    });
    submenu.appendChild(checklist);

    const toggleButton = createMenuActionButton({
      action: 'toggle-all-attributes',
    });
    submenu.appendChild(toggleButton);
    state.ui.menuButtons['toggle-all-attributes'] = toggleButton;
    updateAttributeMenuControls();
  }

  function getViewportMetrics() {
    const viewport = window.visualViewport;
    const clientWidth = viewport?.width || document.documentElement.clientWidth || window.innerWidth;
    const clientHeight = viewport?.height || document.documentElement.clientHeight || window.innerHeight;
    const pageLeft = window.scrollX + (viewport?.offsetLeft || 0);
    const pageTop = window.scrollY + (viewport?.offsetTop || 0);
    return {
      clientWidth,
      clientHeight,
      pageLeft,
      pageTop,
      pageRight: pageLeft + clientWidth,
      pageBottom: pageTop + clientHeight,
    };
  }

  function clampNumber(value, minimum, maximum) {
    if (maximum < minimum) {
      return minimum;
    }
    return Math.min(Math.max(value, minimum), maximum);
  }

  function measureMenuNode(node) {
    if (!node) {
      return new DOMRect(0, 0, 0, 0);
    }

    const previousDisplay = node.style.display;
    const previousVisibility = node.style.visibility;
    node.style.visibility = 'hidden';
    node.style.display = 'block';
    const rect = node.getBoundingClientRect();
    node.style.display = previousDisplay;
    node.style.visibility = previousVisibility;
    return rect;
  }

  function positionSubmenuGroup(groupNode, metrics = getViewportMetrics()) {
    const submenu = groupNode?.querySelector('.wtte-menu-submenu');
    if (!submenu) {
      return;
    }

    const rect = measureMenuNode(submenu);
    const groupRect = groupNode.getBoundingClientRect();
    const spaceBelow = Math.max(0, metrics.clientHeight - groupRect.top - 12);
    const spaceAbove = Math.max(0, groupRect.bottom - 12);
    const openUp = rect.height > spaceBelow && spaceAbove > spaceBelow;
    groupNode.classList.toggle('wtte-open-up', openUp);
    submenu.style.maxHeight = `${Math.max(0, Math.min(openUp ? spaceAbove : spaceBelow, metrics.clientHeight - 24))}px`;
  }

  function updateContextMenuLayout(metrics = getViewportMetrics()) {
    const menu = state.ui.menu;
    if (!menu || menu.hidden) {
      return;
    }

    const submenuWidth = Math.max(
      0,
      ...Array.from(menu.querySelectorAll('.wtte-menu-submenu')).map((submenu) => measureMenuNode(submenu).width || submenu.offsetWidth || 0),
    );
    const menuRect = menu.getBoundingClientRect();
    const spaceRight = Math.max(0, metrics.clientWidth - menuRect.right - 12);
    const spaceLeft = Math.max(0, menuRect.left - 12);
    menu.classList.toggle('wtte-open-left', submenuWidth > spaceRight && spaceLeft >= spaceRight);
    menu.querySelectorAll('.wtte-menu-group').forEach((groupNode) => positionSubmenuGroup(groupNode, metrics));
  }

  function handleMenuChange(event) {
    const input = event.target.closest('input[data-attribute-key]');
    if (!input) {
      return;
    }
    if (input.checked) {
      state.attributeSelection.add(input.dataset.attributeKey);
    } else {
      state.attributeSelection.delete(input.dataset.attributeKey);
    }
    updateAttributeMenuControls();
  }

  function areAllAttributeFieldsSelected() {
    return ATTRIBUTE_FIELD_KEYS.every((key) => state.attributeSelection.has(key));
  }

  function updateAttributeMenuControls() {
    if (!state.ui.menu || state.contextMenuMode !== 'tree') {
      return;
    }

    const cell = ensureSelectedCell();
    const canCopy = Boolean(cell && state.attributeSelection.size);
    const canPaste = Boolean(state.attributeClipboard);
    const canReset = Boolean(state.attributeClipboard || state.attributeSelection.size);
    const toggleButton = state.ui.menu.querySelector('button[data-action="toggle-all-attributes"]');

    state.ui.menu.querySelectorAll('button[data-action="copy-attributes"]').forEach((button) => {
      button.disabled = !canCopy;
    });
    state.ui.menu.querySelectorAll('button[data-action="paste-attributes"]').forEach((button) => {
      button.disabled = !canPaste;
    });
    state.ui.menu.querySelectorAll('button[data-action="reset-attribute-copy"]').forEach((button) => {
      button.disabled = !canReset;
    });
    if (toggleButton) {
      toggleButton.textContent = areAllAttributeFieldsSelected() ? t('menu.clearAllAttributes') : t('menu.selectAllAttributes');
    }
  }

  function getTabMenuGroups() {
    const hiddenTabs = getTreeTabItems({ includeHidden: true, includeAdd: false })
      .filter((tab) => tab.classList.contains('wtte-tab-hidden'))
      .map((tab) => ({
        action: 'show-tab',
        treeId: tab.dataset.treeTarget,
        label: getTreeTabLabel(tab),
      }));

    return [
      {
        id: 'tab',
        labelKey: 'menuGroup.tab',
        actions: [
          { action: 'rename-tab', hidden: !state.contextTabId, disabled: !state.contextTabId },
          { action: 'duplicate-tree', hidden: !state.contextTabId, disabled: !state.contextTabId },
          { action: 'delete-tree', hidden: !state.contextTabId || !isCustomTreeId(state.contextTabId), disabled: !isCustomTreeId(state.contextTabId) },
          { action: 'hide-tab', hidden: !state.contextTabId, disabled: !canHideTreeTab(state.contextTabId) },
        ],
      },
      {
        id: 'show',
        labelKey: 'menuGroup.show',
        actions: hiddenTabs,
      },
    ];
  }

  function getAddTabMenuGroups() {
    return [
      {
        id: 'create-tree',
        labelKey: 'menuGroup.createTree',
        actions: [
          { action: 'add-custom-tree-blank' },
          { action: 'add-custom-tree-full' },
        ],
      },
    ];
  }

  function updateContextMenuPlacement({ clampToViewport = false, refreshLayout = false } = {}) {
    const menu = state.ui.menu;
    if (!menu || menu.hidden) {
      return;
    }

    const metrics = getViewportMetrics();
    const rect = menu.getBoundingClientRect();
    const width = rect.width || menu.offsetWidth || 0;
    const height = rect.height || menu.offsetHeight || 0;
    let left = state.menuAnchor?.left ?? metrics.pageLeft + 12;
    let top = state.menuAnchor?.top ?? metrics.pageTop + 12;

    if (clampToViewport) {
      left = clampNumber(left, metrics.pageLeft + 12, metrics.pageRight - width - 12);
      top = clampNumber(top, metrics.pageTop + 12, metrics.pageBottom - height - 12);
      state.menuAnchor = { left, top };
    }

    menu.style.left = `${left}px`;
    menu.style.top = `${top}px`;
    if (refreshLayout) {
      updateContextMenuLayout(metrics);
    }
  }

  function showMenu(pageX, pageY) {
    renderContextMenu();
    if (!state.ui.menu.children.length) {
      hideMenu();
      return;
    }
    updateMenuButtons();
    state.menuAnchor = { left: pageX, top: pageY };
    state.ui.menu.hidden = false;
    state.ui.menu.classList.remove('wtte-open-left');
    state.ui.menu.style.left = '-9999px';
    state.ui.menu.style.top = '-9999px';
    updateContextMenuPlacement({ clampToViewport: true, refreshLayout: true });
  }

  function hideMenu() {
    state.ui.menu.hidden = true;
    state.menuAnchor = null;
    clearContext();
  }

  function syncContextMenuVisibility() {
    const menu = state.ui.menu;
    if (!menu || menu.hidden) {
      return;
    }
    const { clientWidth, clientHeight } = getViewportMetrics();
    const rect = menu.getBoundingClientRect();
    if (rect.bottom <= 0 || rect.right <= 0 || rect.top >= clientHeight || rect.left >= clientWidth) {
      hideMenu();
    }
  }

  function clearContext() {
    state.contextCell = null;
    state.contextNode = null;
    state.contextKind = null;
    state.contextTabId = '';
    state.contextMenuMode = 'tree';
  }

  function updateMenuButtons() {
    if (state.contextMenuMode !== 'tree') {
      return;
    }
    const cell = ensureSelectedCell();
    const row = ensureSelectedRow();
    const contextNode = getContextContentNode();
    const hasContent = Boolean(getTopLevelContentFromCell(cell));
    const hasImage = Boolean(extractContentImageUrl(contextNode));
    const canPaste = Boolean(state.clipboard && cell);
    const canPasteAttributes = Boolean(state.attributeClipboard && cell);
    const canGroup = canGroupSelectedCells();
    const canUngroup = Boolean(cell && getDirectChild(cell, '.wt-tree_group'));
    const hasAttributeSelection = state.attributeSelection.size > 0;

    setMenuActionState('edit-card', Boolean(cell), true);
    setMenuActionState('clear-cell', Boolean(cell && hasContent), true);
    setMenuActionState('copy-attributes', Boolean(cell && hasAttributeSelection), true);
    setMenuActionState('paste-attributes', canPasteAttributes, true);
    setMenuActionState('reset-attribute-copy', Boolean(state.attributeClipboard || hasAttributeSelection), true);
    setMenuActionState('copy-card', Boolean(contextNode), true);
    setMenuActionState('paste-card', canPaste, true);
    setMenuActionState('copy-image-url', hasImage, Boolean(contextNode));
    setMenuActionState('group-selected', canGroup, Boolean(cell));
    setMenuActionState('ungroup-cell', canUngroup, Boolean(cell));
    setMenuActionState('add-row-above', Boolean(row), true);
    setMenuActionState('add-row-below', Boolean(row), true);
    setMenuActionState('add-row-both-above', Boolean(row), true);
    setMenuActionState('add-row-both-below', Boolean(row), true);
    setMenuActionState('duplicate-row', Boolean(row), true);
    setMenuActionState('delete-row', Boolean(row), true);
    setMenuActionState('add-column-left', Boolean(cell), true);
    setMenuActionState('add-column-right', Boolean(cell), true);
    setMenuActionState('duplicate-column', Boolean(cell), true);
    setMenuActionState('delete-column', Boolean(cell), true);
    setMenuActionState('add-rank-above', Boolean(state.selectedRankRow), true);
    setMenuActionState('add-rank-below', Boolean(state.selectedRankRow), true);
    setMenuActionState('delete-rank', canDeleteSelectedRank(), Boolean(state.selectedRankRow));
    updateAttributeMenuControls();
  }

  function setMenuActionState(action, enabled, visible) {
    state.ui.menu.querySelectorAll(`button[data-action="${action}"]`).forEach((button) => {
      button.hidden = !visible;
      button.disabled = !enabled;
    });
    state.ui.menu.querySelectorAll('.wtte-menu-group').forEach((group) => {
      const hasVisibleAction = Array.from(group.querySelectorAll('.wtte-menu-submenu > button')).some((button) => !button.hidden);
      group.hidden = !hasVisibleAction;
    });
  }

  function handleMenuAction(action, payload = {}) {
    const contextTabId = state.contextTabId;
    const persistentActions = new Set([
      'copy-attributes',
      'paste-attributes',
      'reset-attribute-copy',
      'toggle-all-attributes',
    ]);
    if (!persistentActions.has(action)) {
      hideMenu();
    }

    switch (action) {
      case 'edit-card':
        openCardModal();
        break;
      case 'clear-cell':
        clearSelectedCell();
        break;
      case 'copy-attributes':
        copySelectedAttributes();
        break;
      case 'reset-attribute-copy':
        resetAttributeCopy();
        break;
      case 'paste-attributes':
        pasteSelectedAttributes();
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
      case 'add-row-above':
        addRowAbove();
        break;
      case 'add-row-below':
        addRowBelow();
        break;
      case 'add-row-both-above':
        addRowAboveBothSections();
        break;
      case 'add-row-both-below':
        addRowBelowBothSections();
        break;
      case 'duplicate-row':
        duplicateRow();
        break;
      case 'delete-row':
        deleteRow();
        break;
      case 'add-column-left':
        addColumnLeft();
        break;
      case 'add-column-right':
        addColumnRight();
        break;
      case 'duplicate-column':
        duplicateColumn();
        break;
      case 'delete-column':
        deleteColumn();
        break;
      case 'add-rank-above':
        openRankModal('above');
        break;
      case 'add-rank-below':
        openRankModal('below');
        break;
      case 'delete-rank':
        deleteSelectedRank();
        break;
      case 'add-custom-tree-blank':
        addCustomTree('blank');
        break;
      case 'add-custom-tree-full':
        addCustomTree('full');
        break;
      case 'duplicate-tree':
        duplicateTree(contextTabId);
        break;
      case 'delete-tree':
        deleteTree(contextTabId);
        break;
      case 'hide-tab':
        hideTreeTab(contextTabId);
        break;
      case 'rename-tab':
        renameTreeTab(contextTabId);
        break;
      case 'show-tab':
        showTreeTab(payload.treeId || '');
        break;
      case 'toggle-all-attributes':
        toggleAllAttributeFields();
        break;
      default:
        break;
    }

    if (persistentActions.has(action) && !state.ui.menu.hidden) {
      updateMenuButtons();
    }
  }

  function openCardModal() {
    const cell = ensureSelectedCell();
    if (!cell) {
      return;
    }

    state.modalTargetNode = getContextContentNode() || getTopLevelContentFromCell(cell);
    const data = readCardData(cell, state.modalTargetNode);
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
    updateCardAttributeButtons();
    state.ui.cardForm.unitId.focus();
    state.ui.cardForm.unitId.select();
  }

  function openRankModal(direction = 'below') {
    if (!state.selectedRankRow || !state.selectedRankRow.isConnected) {
      return;
    }

    state.pendingRankTarget = state.selectedRankRow;
    state.pendingRankDirection = direction === 'above' ? 'above' : 'below';
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
    state.modalTargetNode = null;
    state.pendingRankTarget = null;
    state.pendingRankDirection = 'below';
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

    const targetNode = state.modalTargetNode?.isConnected ? state.modalTargetNode : (getContextContentNode() || getTopLevelContentFromCell(cell));
    const unitTree = cell.closest('.unit-tree');
    recordUndo(unitTree);
    applyCardDataToCell(cell, { ...readCardData(cell, targetNode), ...readCardFormData() }, targetNode);
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

    addRankRelative(state.pendingRankTarget, label, state.pendingRankDirection);
    closeModal();
  }

  function readCardFormData() {
    const formData = new FormData(state.ui.cardForm);
    const name = String(formData.get('name') || '').trim();
    const unitId = normalizeUnitId(String(formData.get('unitId') || '').trim() || buildDefaultUnitId(name));
    return {
      name,
      unitId,
      br: String(formData.get('br') || '').trim(),
      imageUrl: String(formData.get('imageUrl') || '').trim(),
      linkUrl: normalizeLinkUrl(String(formData.get('linkUrl') || '').trim(), unitId),
      style: String(formData.get('style') || 'regular'),
      prefix: String(formData.get('prefix') || ''),
    };
  }

  function fillCardForm(data) {
    if (!data) {
      return;
    }
    state.ui.cardForm.unitId.value = data.unitId || '';
    state.ui.cardForm.name.value = data.name || '';
    state.ui.cardForm.br.value = data.br || '';
    state.ui.cardForm.imageUrl.value = data.imageUrl || '';
    state.ui.cardForm.linkUrl.value = data.linkUrl || '';
    state.ui.cardForm.style.value = data.style || 'regular';
    state.ui.cardForm.prefix.value = data.prefix || '';
  }

  function updateCardAttributeButtons() {
    if (!state.ui.cardButtons) {
      return;
    }
    state.ui.cardButtons.pasteAttributes.disabled = !getClipboardAttributeKeys(state.attributeClipboard, { includeFolder: false }).length;
  }

  function cloneCardData(data) {
    return {
      unitId: data.unitId || '',
      name: data.name || '',
      br: data.br || '',
      imageUrl: data.imageUrl || '',
      linkUrl: data.linkUrl || '',
      style: data.style || 'regular',
      prefix: data.prefix || '',
      requiredId: data.requiredId || '',
    };
  }

  function copySelectedAttributes() {
    const cell = ensureSelectedCell();
    const keys = getSelectedAttributeKeys();
    if (!cell || !keys.length) {
      return;
    }
    const sourceNode = getContextContentNode() || getTopLevelContentFromCell(cell);
    state.attributeClipboard = buildAttributeClipboard(keys, readCardData(cell, sourceNode), sourceNode);
    updateCardAttributeButtons();
    updateAttributeMenuControls();
    flashStatus(t('flash.attributesCopied'));
  }

  function pasteSelectedAttributes() {
    if (!state.attributeClipboard) {
      flashStatus(t('flash.noAttributes'));
      return;
    }
    const cells = getSelectedCellsForBatchAction();
    if (!cells.length) {
      return;
    }

    const unitTree = cells[0].closest('.unit-tree');
    recordUndo(unitTree);
    cells.forEach((cell) => {
      applyAttributeClipboardToCell(cell, state.attributeClipboard);
    });
    setSelection({
      cell: cells[0],
      row: cells[0].parentElement,
      rankRow: cells[0].closest('.wt-tree_rank'),
    });
    clearGroupSelection();
    finalizeTreeChange(unitTree);
    flashStatus(t('flash.attributesPasted'));
  }

  function copyModalAttributes() {
    const formData = readCardFormData();
    const keys = getSelectedAttributeKeys({ includeFolder: false, fallbackAll: true });
    state.attributeClipboard = buildAttributeClipboard(keys, formData, null, 'equipment');
    updateCardAttributeButtons();
    updateAttributeMenuControls();
    flashStatus(t('flash.attributesCopied'));
  }

  function pasteModalAttributes() {
    const keys = getClipboardAttributeKeys(state.attributeClipboard, { includeFolder: false });
    if (!state.attributeClipboard || !keys.length) {
      flashStatus(t('flash.noAttributes'));
      return;
    }
    fillCardForm(mergeCardDataByKeys(readCardFormData(), state.attributeClipboard.data, keys));
    flashStatus(t('flash.attributesPasted'));
  }

  function resetAttributeCopy() {
    state.attributeClipboard = null;
    state.attributeSelection.clear();
    state.ui.menu.querySelectorAll('input[data-attribute-key]').forEach((input) => {
      input.checked = false;
    });
    updateCardAttributeButtons();
    updateAttributeMenuControls();
  }

  function toggleAllAttributeFields() {
    if (areAllAttributeFieldsSelected()) {
      state.attributeSelection.clear();
    } else {
      ATTRIBUTE_FIELD_KEYS.forEach((key) => state.attributeSelection.add(key));
    }
    state.ui.menu.querySelectorAll('input[data-attribute-key]').forEach((input) => {
      input.checked = state.attributeSelection.has(input.dataset.attributeKey);
    });
    updateAttributeMenuControls();
  }

  function getSelectedAttributeKeys({ includeFolder = true, fallbackAll = false } = {}) {
    const keys = ATTRIBUTE_FIELD_KEYS.filter((key) => state.attributeSelection.has(key) && (includeFolder || key !== 'isFolder'));
    if (keys.length || !fallbackAll) {
      return keys;
    }
    return ATTRIBUTE_FIELD_KEYS.filter((key) => includeFolder || key !== 'isFolder');
  }

  function getClipboardAttributeKeys(clipboard, { includeFolder = true } = {}) {
    if (!clipboard?.keys?.length) {
      return [];
    }
    return clipboard.keys.filter((key) => includeFolder || key !== 'isFolder');
  }

  function buildAttributeClipboard(keys, data, sourceNode, sourceKind = '') {
    const resolvedKind = sourceKind || getNodeKind(sourceNode) || 'equipment';
    return {
      keys: keys.slice(),
      kind: resolvedKind,
      data: cloneCardData(data),
      nodeHtml: sourceNode?.outerHTML || '',
    };
  }

  function mergeCardDataByKeys(targetData, sourceData, keys) {
    const nextData = cloneCardData(targetData);
    keys.forEach((key) => {
      if (ATTRIBUTE_DATA_KEYS.includes(key)) {
        nextData[key] = sourceData[key] || '';
      }
    });
    return nextData;
  }

  function cloneClipboardNode(clipboard) {
    const parsed = htmlToElement(clipboard?.nodeHtml || '');
    if (!parsed) {
      return null;
    }
    stripEditorClasses(parsed);
    return parsed;
  }

  function buildFallbackGroupFromCell(cell, data) {
    const targetNode = getTopLevelContentFromCell(cell);
    if (targetNode?.matches('.wt-tree_group')) {
      const clone = targetNode.cloneNode(true);
      stripEditorClasses(clone);
      return clone;
    }
    const item = targetNode?.matches('.wt-tree_item')
      ? targetNode.cloneNode(true)
      : buildCardElement(data, cell);
    stripEditorClasses(item);
    return buildGroupElement([item], cell);
  }

  function applyAttributeClipboardToCell(cell, clipboard) {
    const keys = getClipboardAttributeKeys(clipboard);
    if (!cell || !keys.length) {
      return;
    }

    const targetNode = getTopLevelContentFromCell(cell);
    const currentData = readCardData(cell, targetNode);
    const nextData = mergeCardDataByKeys(currentData, clipboard.data, keys);

    if (keys.includes('isFolder')) {
      if (clipboard.kind === 'folder') {
        const folderGroup = targetNode?.matches('.wt-tree_group')
          ? targetNode
          : buildFallbackGroupFromCell(cell, nextData);
        if (folderGroup?.matches('.wt-tree_group')) {
          if (!folderGroup.parentElement) {
            cell.innerHTML = '';
            cell.appendChild(folderGroup);
          }
          updateGroupFolder(folderGroup, nextData, cell);
          folderGroup.dataset.wtteGroupSize = String(getGroupItemNodes(folderGroup).length);
          setActiveContent(folderGroup);
          return;
        }
      }

      setCellCard(cell, nextData);
      setActiveContent(getDirectChild(cell, '.wt-tree_item'));
      return;
    }

    applyCardDataToCell(cell, nextData);
  }

  function applyCardDataToCell(cell, cardData, targetNode = null) {
    if (!cell) {
      return;
    }
    const resolvedTarget = targetNode && targetNode.isConnected ? targetNode : getTopLevelContentFromCell(cell);
    if (resolvedTarget?.matches('.wt-tree_group')) {
      updateGroupFolder(resolvedTarget, cardData, cell);
      setActiveContent(resolvedTarget);
      return;
    }
    if (resolvedTarget?.matches('.wt-tree_item')) {
      const replacement = buildCardElement(cardData, cell);
      resolvedTarget.replaceWith(replacement);
      setActiveContent(replacement);
      if (isGroupChildItem(replacement)) {
        syncGroupAppearance(replacement.closest('.wt-tree_group'));
      }
      return;
    }
    setCellCard(cell, cardData);
    setActiveContent(getDirectChild(cell, '.wt-tree_item'));
  }

  function getSelectedCellsForBatchAction() {
    const cells = getGroupSelectionCells();
    if (cells.length) {
      return cells.slice().sort(compareNodes);
    }
    const cell = ensureSelectedCell();
    return cell ? [cell] : [];
  }

  function deleteSelectionContentsOrRows() {
    const cells = getSelectedCellsForBatchAction();
    if (!cells.length) {
      return;
    }

    if (cells.some((cell) => Boolean(getTopLevelContentFromCell(cell)))) {
      clearSelectedCell({ wholeCell: true });
      return;
    }

    deleteRowsForCells(cells);
  }

  function deleteSelectionRange() {
    const cells = getSelectedCellsForBatchAction();
    if (!cells.length) {
      return;
    }
    if (cells.length < 2) {
      deleteSelectionContentsOrRows();
      return;
    }

    const unitTree = cells[0].closest('.unit-tree');
    const rankRows = getRankRows(unitTree);
    const affectedRankRows = new Set(cells.map((cell) => cell.closest('.wt-tree_rank')).filter(Boolean));
    const fullRankRows = new Set(Array.from(affectedRankRows).filter((rankRow) => {
      const tables = getRankRowTables(rankRow).filter(Boolean);
      return tables.length && tables.every((table) => {
        const selectedRows = new Set(cells
          .filter((cell) => cell.closest('.wt-tree_rank') === rankRow && cell.parentElement?.closest('table') === table)
          .map((cell) => cell.parentElement));
        return selectedRows.size >= table.rows.length;
      });
    }));

    recordUndo(unitTree);
    let nextCell = null;

    if (fullRankRows.size >= rankRows.length && rankRows.length) {
      const keepRankRow = rankRows[0];
      rankRows.slice(1).forEach((rankRow) => {
        getRankHeader(rankRow)?.remove();
        rankRow.remove();
      });
      configureCustomTreeSkeleton(unitTree, {
        rankCount: 1,
        rowCount: 1,
        leftColumns: getTableColumnCount(getRankRowTables(keepRankRow)[0]) || 5,
        rightColumns: getTableColumnCount(getRankRowTables(keepRankRow)[1]) || 2,
      });
      nextCell = getRankRows(unitTree)[0]?.querySelector('td') || null;
    } else {
      fullRankRows.forEach((rankRow) => {
        getRankHeader(rankRow)?.remove();
        rankRow.remove();
      });

      const rowsByTable = new Map();
      cells.forEach((cell) => {
        if (!cell.isConnected) {
          return;
        }
        const rankRow = cell.closest('.wt-tree_rank');
        if (rankRow && fullRankRows.has(rankRow)) {
          return;
        }
        const row = cell.parentElement;
        const table = row?.closest('table');
        if (!row || !table) {
          return;
        }
        if (!rowsByTable.has(table)) {
          rowsByTable.set(table, new Set());
        }
        rowsByTable.get(table).add(row);
      });

      rowsByTable.forEach((rowSet, table) => {
        const rows = Array.from(table.rows);
        const selectedRows = rows.filter((row) => rowSet.has(row));
        if (!selectedRows.length) {
          return;
        }

        if (selectedRows.length >= rows.length) {
          const keepRow = rows[0];
          rows.slice(1).forEach((row) => row.remove());
          Array.from(keepRow?.cells || []).forEach((cell) => {
            cell.innerHTML = '';
          });
          nextCell = nextCell || keepRow?.cells?.[0] || null;
          return;
        }

        const firstRowIndex = rows.findIndex((row) => rowSet.has(row));
        selectedRows
          .slice()
          .sort((leftRow, rightRow) => rows.indexOf(rightRow) - rows.indexOf(leftRow))
          .forEach((row) => row.remove());

        const remainingRows = Array.from(table.rows);
        const fallbackRow = remainingRows[Math.min(firstRowIndex, remainingRows.length - 1)] || remainingRows[remainingRows.length - 1] || null;
        nextCell = nextCell || fallbackRow?.cells?.[0] || null;
      });
    }

    nextCell = nextCell || getRankRows(unitTree)[0]?.querySelector('td') || null;
    clearGroupSelection();
    setSelection({
      cell: nextCell,
      row: nextCell?.parentElement || null,
      rankRow: nextCell ? nextCell.closest('.wt-tree_rank') : getRankRows(unitTree)[0] || null,
    });
    finalizeTreeChange(unitTree);
  }

  function deleteRowsForCells(cells) {
    if (!cells.length) {
      return;
    }

    const unitTree = cells[0].closest('.unit-tree');
    const rowsByTable = new Map();
    cells.forEach((cell) => {
      const row = cell.parentElement;
      const table = row?.closest('table');
      if (!row || !table) {
        return;
      }
      if (!rowsByTable.has(table)) {
        rowsByTable.set(table, new Set());
      }
      rowsByTable.get(table).add(row);
    });

    if (!rowsByTable.size) {
      return;
    }

    recordUndo(unitTree);
    let nextCell = null;

    rowsByTable.forEach((rowSet, table) => {
      const rows = Array.from(table.rows);
      const selectedRows = rows.filter((row) => rowSet.has(row));
      if (!selectedRows.length) {
        return;
      }

      if (selectedRows.length >= rows.length) {
        const keepRow = rows[0];
        rows.slice(1).forEach((row) => row.remove());
        Array.from(keepRow?.cells || []).forEach((cell) => {
          cell.innerHTML = '';
        });
        nextCell = nextCell || keepRow?.cells?.[0] || null;
        return;
      }

      const firstRowIndex = rows.findIndex((row) => rowSet.has(row));
      selectedRows
        .slice()
        .sort((leftRow, rightRow) => rows.indexOf(rightRow) - rows.indexOf(leftRow))
        .forEach((row) => row.remove());

      const remainingRows = Array.from(table.rows);
      const fallbackRow = remainingRows[Math.min(firstRowIndex, remainingRows.length - 1)] || remainingRows[remainingRows.length - 1] || null;
      nextCell = nextCell || fallbackRow?.cells?.[0] || null;
    });

    clearGroupSelection();
    setSelection({
      cell: nextCell,
      row: nextCell?.parentElement || null,
      rankRow: nextCell ? nextCell.closest('.wt-tree_rank') : null,
    });
    finalizeTreeChange(unitTree);
  }

  function clearSelectedCell(options = {}) {
    const cells = getSelectedCellsForBatchAction();
    if (!cells.length) {
      return;
    }

    const cell = cells[0];
    const unitTree = cell.closest('.unit-tree');
    recordUndo(unitTree);
    cells.forEach((targetCell, index) => {
      const targetNode = index === 0 && cells.length === 1 && !options.wholeCell
        ? (getContextContentNode() || getTopLevelContentFromCell(targetCell))
        : getTopLevelContentFromCell(targetCell);
      if (!targetNode || targetNode.parentElement === targetCell) {
        targetCell.innerHTML = '';
        setActiveContent(null);
      } else if (isGroupChildItem(targetNode)) {
        const group = targetNode.closest('.wt-tree_group');
        targetNode.remove();
        normalizeGroupAfterMutation(group);
        setActiveContent(null);
      }
    });
    setSelection({
      cell,
      row: cell.parentElement,
      rankRow: cell.closest('.wt-tree_rank'),
    });
    clearGroupSelection();
    finalizeTreeChange(unitTree);
  }

  function addRowAbove() {
    insertRowRelative('above');
  }

  function addRowBelow() {
    insertRowRelative('below');
  }

  function addRowAboveBothSections() {
    insertRowRelativeInRankSections('above');
  }

  function addRowBelowBothSections() {
    insertRowRelativeInRankSections('below');
  }

  function insertRowRelative(direction) {
    const row = ensureSelectedRow();
    if (!row) {
      return;
    }

    const unitTree = row.closest('.unit-tree');
    recordUndo(unitTree);

    const newRow = createBlankRowLike(row);
    ensureRowColumnCount(newRow, getTableColumnCount(row.closest('table')), row.cells[row.cells.length - 1] || state.selectedCell);
    row.insertAdjacentElement(direction === 'above' ? 'beforebegin' : 'afterend', newRow);
    clearGroupSelection();
    setSelection({
      cell: newRow.cells[0] || null,
      row: newRow,
      rankRow: newRow.closest('.wt-tree_rank'),
    });
    finalizeTreeChange(unitTree);
  }

  function insertRowRelativeInRankSections(direction) {
    const row = ensureSelectedRow();
    const table = row?.closest('table');
    const rankRow = row?.closest('.wt-tree_rank');
    if (!row || !table || !rankRow) {
      return;
    }

    const unitTree = row.closest('.unit-tree');
    const rowIndex = Array.from(table.rows).indexOf(row);
    const tables = getRankRowTables(rankRow).filter(Boolean);
    if (rowIndex < 0 || !tables.length) {
      return;
    }

    recordUndo(unitTree);
    tables.forEach((currentTable) => {
      while (currentTable.rows.length <= rowIndex) {
        const seedRow = currentTable.rows[currentTable.rows.length - 1] || currentTable.rows[0] || document.createElement('tr');
        const paddedRow = createBlankRowLike(seedRow);
        ensureRowColumnCount(paddedRow, getTableColumnCount(currentTable) || getTableColumnCount(table) || 1, seedRow.cells?.[0] || row.cells?.[0] || null);
        currentTable.appendChild(paddedRow);
      }
      const referenceRow = currentTable.rows[rowIndex];
      const newRow = createBlankRowLike(referenceRow);
      ensureRowColumnCount(newRow, getTableColumnCount(currentTable) || getTableColumnCount(table) || 1, referenceRow.cells?.[0] || row.cells?.[0] || null);
      referenceRow.insertAdjacentElement(direction === 'above' ? 'beforebegin' : 'afterend', newRow);
    });

    const targetTable = getRankRowTables(rankRow)[getTableSectionIndex(table)] || table;
    const targetRow = targetTable.rows[direction === 'above' ? rowIndex : rowIndex + 1] || targetTable.rows[rowIndex] || null;
    clearGroupSelection();
    setSelection({
      cell: targetRow?.cells?.[0] || null,
      row: targetRow,
      rankRow,
    });
    finalizeTreeChange(unitTree);
  }

  function addColumnLeft() {
    insertColumnRelative('left');
  }

  function addColumnRight() {
    insertColumnRelative('right');
  }

  function insertColumnRelative(direction) {
    const selection = getSelectedTableContext();
    if (!selection) {
      return;
    }

    const {
      cell,
      rowIndex,
      table,
      unitTree,
      columnIndex,
      sectionTables,
    } = selection;

    const targetIndex = direction === 'left' ? columnIndex : columnIndex + 1;
    recordUndo(unitTree);
    sectionTables.forEach((currentTable) => {
      normalizeTableColumnCount(currentTable, targetIndex, cell);
      Array.from(currentTable.rows).forEach((currentRow) => {
        const sourceCell = currentRow.cells[Math.min(columnIndex, Math.max(0, currentRow.cells.length - 1))] || cell;
        const newCell = createCellLike(sourceCell, false);
        insertCellAtIndex(currentRow, targetIndex, newCell);
      });
    });

    clearGroupSelection();
    setSelectionToColumnCell(table, rowIndex, targetIndex);
    finalizeTreeChange(unitTree);
  }

  function duplicateColumn() {
    const selection = getSelectedTableContext();
    if (!selection) {
      return;
    }

    const {
      cell,
      rowIndex,
      table,
      unitTree,
      columnIndex,
      sectionTables,
    } = selection;

    recordUndo(unitTree);
    const targetIndex = columnIndex + 1;
    sectionTables.forEach((currentTable) => {
      normalizeTableColumnCount(currentTable, targetIndex, cell);
      Array.from(currentTable.rows).forEach((currentRow) => {
        const sourceCell = currentRow.cells[columnIndex] || currentRow.cells[currentRow.cells.length - 1] || cell;
        const clone = createCellLike(sourceCell, true);
        insertCellAtIndex(currentRow, targetIndex, clone);
      });
    });

    clearGroupSelection();
    setSelectionToColumnCell(table, rowIndex, targetIndex);
    finalizeTreeChange(unitTree);
  }

  function deleteColumn() {
    const selection = getSelectedTableContext();
    if (!selection) {
      return;
    }

    const {
      cell,
      rowIndex,
      table,
      unitTree,
      columnIndex,
      sectionTables,
    } = selection;

    recordUndo(unitTree);
    const tableColumnCount = Math.max(...sectionTables.map((currentTable) => getTableColumnCount(currentTable)), 0);

    if (tableColumnCount <= 1) {
      sectionTables.forEach((currentTable) => {
        Array.from(currentTable.rows).forEach((currentRow) => {
          const targetCell = currentRow.cells[0];
          if (targetCell) {
            targetCell.innerHTML = '';
          }
        });
      });
      clearGroupSelection();
      setSelectionToColumnCell(table, rowIndex, 0);
      finalizeTreeChange(unitTree);
      return;
    }

    sectionTables.forEach((currentTable) => {
      normalizeTableColumnCount(currentTable, columnIndex + 1, cell);
      Array.from(currentTable.rows).forEach((currentRow) => {
        currentRow.cells[columnIndex]?.remove();
        if (!currentRow.cells.length) {
          currentRow.appendChild(createCellLike(cell, false));
        }
      });
    });

    clearGroupSelection();
    setSelectionToColumnCell(table, rowIndex, Math.max(0, Math.min(columnIndex, getTableColumnCount(table) - 1)));
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
    ensureRowColumnCount(newRow, getTableColumnCount(row.closest('table')), row.cells[row.cells.length - 1] || state.selectedCell);
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

  function addRankRelative(rankRow, label, direction = 'below') {
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

    if (direction === 'above') {
      header.insertAdjacentElement('beforebegin', headerClone);
      headerClone.insertAdjacentElement('afterend', rankClone);
    } else {
      rankRow.insertAdjacentElement('afterend', rankClone);
      rankClone.insertAdjacentElement('beforebegin', headerClone);
    }

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
    const cells = getSelectedCellsForBatchAction();
    if (!cells.length || !state.clipboard) {
      return;
    }

    const unitTree = cells[0].closest('.unit-tree');
    recordUndo(unitTree);
    cells.forEach((cell) => {
      const parsed = htmlToElement(state.clipboard.html);
      if (!parsed) {
        return;
      }
      stripEditorClasses(parsed);
      cell.innerHTML = '';
      cell.appendChild(parsed);
    });
    setSelection({
      cell: cells[0],
      row: cells[0].parentElement,
      rankRow: cells[0].closest('.wt-tree_rank'),
    });
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
    const rowIndex = Array.from(table?.rows || []).indexOf(row);
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
      const clone = items[index].cloneNode(true);
      if (index === 0 && !clone.getAttribute('data-unit-req') && group.getAttribute('data-unit-req')) {
        clone.setAttribute('data-unit-req', group.getAttribute('data-unit-req'));
      }
      targetCell.appendChild(clone);
    });

    clearGroupSelection();
    setSelection({
      cell: table.rows[rowIndex]?.cells[columnIndex] || null,
      row: table.rows[rowIndex] || row,
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
    if (data.requiredId) {
      element.setAttribute('data-unit-req', data.requiredId);
    } else {
      element.removeAttribute('data-unit-req');
    }

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
    const groupRequiredId = folderInfo.requiredId || '';

    Array.from(itemsContainer.children)
      .filter((child) => !child.matches('.wt-tree_group-canvas'))
      .forEach((child) => child.remove());

    itemNodes.forEach((itemNode, index) => {
      const clone = itemNode.cloneNode(true);
      stripEditorClasses(clone);
      if (index === 0 && groupRequiredId) {
        clone.removeAttribute('data-unit-req');
      }
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
      requiredId: groupRequiredId,
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
    if (data.requiredId) {
      group.setAttribute('data-unit-req', data.requiredId);
    } else {
      group.removeAttribute('data-unit-req');
    }
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
      if (!item.getAttribute('data-unit-req') && group.getAttribute('data-unit-req')) {
        item.setAttribute('data-unit-req', group.getAttribute('data-unit-req'));
      }
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
    sample.requiredId = '';
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
      requiredId: item.getAttribute('data-unit-req') || '',
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
      requiredId: group.getAttribute('data-unit-req') || '',
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
    syncArrowVisibility(unitTree);
    scheduleLayout(unitTree);
    scheduleTreeSave(unitTree);
    updateToolbar();
  }

  function syncArrowVisibility(unitTree) {
    const wtTree = unitTree?.querySelector('.wt-tree');
    if (!wtTree) {
      return;
    }
    const hideArrows = wtTree.classList.contains('wtte-hide-arrows');
    wtTree.querySelectorAll('.wt-tree_arrows').forEach((node) => {
      node.hidden = hideArrows;
      if (hideArrows) {
        node.style.display = 'none';
      } else {
        node.style.removeProperty('display');
      }
    });
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

    state.store.version = 3;
    if (arePersistedTreeStatesEqual(snapshot, state.originalTrees[treeId])) {
      delete state.store.modifiedTrees[treeId];
    } else {
      state.store.modifiedTrees[treeId] = snapshot;
    }
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
    flushPendingTreeSaves();
    const hasSavedData = Boolean(
      state.store
      && (Object.keys(state.store.modifiedTrees || {}).length || hasTabMetadataChanges())
    );
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

  function resetCurrentTree() {
    const unitTree = getPrimaryTree();
    const treeId = unitTree?.dataset?.treeId;
    const originalState = treeId ? state.originalTrees[treeId] : null;
    if (!unitTree || !originalState || !isTreeModified(unitTree)) {
      return;
    }

    const confirmed = window.confirm(t('confirm.resetCurrentTree'));
    if (!confirmed) {
      return;
    }

    recordUndo(unitTree);
    applyPersistedTreeState(unitTree, originalState);
    delete state.store.modifiedTrees[treeId];
    state.store.updatedAt = new Date().toISOString();
    saveStore();
    hideMenu();
    clearHoverState();
    clearGroupSelection();
    clearSelection();
    cleanupDrag();
    scheduleLayout(unitTree);
    updateToolbar();
    flashStatus(t('flash.currentTreeReset'));
  }

  function openImportDialog() {
    const input = state.ui.importInput;
    if (!input) {
      return;
    }
    input.value = '';
    input.click();
  }

  async function handleImportFileSelection(event) {
    const input = event.target;
    const file = input?.files?.[0];
    if (!file) {
      return;
    }

    try {
      const parsed = JSON.parse(await file.text());
      const importedStore = normalizeImportedStore(parsed);
      if (!importedStore) {
        flashStatus(t('flash.importFailed'));
        return;
      }
      if (importedStore.pathname && importedStore.pathname !== location.pathname) {
        flashStatus(t('flash.importPathMismatch'));
        return;
      }
      applyImportedStore(importedStore);
      flashStatus(t('flash.importApplied', { count: Object.keys(importedStore.modifiedTrees).length }));
    } catch (error) {
      console.warn('[WTTE] Failed to import store', error);
      flashStatus(t('flash.importFailed'));
    } finally {
      if (input) {
        input.value = '';
      }
    }
  }

  function exportSavedTrees() {
    const exportPayload = buildExportPayload();
    const treeCount = Object.keys(exportPayload.modifiedTrees).length;
    const hasTabData = hasTabMetadataChanges(exportPayload.tabs);
    if (!treeCount && !hasTabData) {
      return;
    }

    const blob = new Blob([JSON.stringify(exportPayload, null, 2)], { type: 'application/json' });
    const objectUrl = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = objectUrl;
    link.download = buildExportFilename();
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(objectUrl), 0);
    flashStatus(t('flash.exportReady', { count: treeCount || Object.keys(exportPayload.tabs.custom || {}).length }));
  }

  function createDefaultTabStore() {
    return {
      order: [],
      hidden: [],
      labels: {},
      custom: {},
    };
  }

  function createDefaultStore() {
    return {
      version: 3,
      modifiedTrees: {},
      tabs: createDefaultTabStore(),
      updatedAt: null,
    };
  }

  function normalizeTabStore(rawTabs) {
    const fallback = createDefaultTabStore();
    if (!rawTabs || typeof rawTabs !== 'object') {
      return fallback;
    }
    return {
      order: Array.isArray(rawTabs.order) ? rawTabs.order.filter((value) => typeof value === 'string') : [],
      hidden: Array.isArray(rawTabs.hidden) ? rawTabs.hidden.filter((value) => typeof value === 'string') : [],
      labels: rawTabs.labels && typeof rawTabs.labels === 'object' ? { ...rawTabs.labels } : {},
      custom: rawTabs.custom && typeof rawTabs.custom === 'object' ? { ...rawTabs.custom } : {},
    };
  }

  function normalizeStore(rawStore) {
    const fallback = createDefaultStore();
    if (!rawStore || typeof rawStore !== 'object') {
      return fallback;
    }
    return {
      version: typeof rawStore.version === 'number' ? rawStore.version : 3,
      modifiedTrees: rawStore.modifiedTrees && typeof rawStore.modifiedTrees === 'object' ? { ...rawStore.modifiedTrees } : {},
      tabs: normalizeTabStore(rawStore.tabs),
      updatedAt: typeof rawStore.updatedAt === 'string' ? rawStore.updatedAt : null,
    };
  }

  function cloneTabStore(tabStore = state.store.tabs) {
    const normalized = normalizeTabStore(tabStore);
    return {
      order: [...normalized.order],
      hidden: [...normalized.hidden],
      labels: { ...normalized.labels },
      custom: { ...normalized.custom },
    };
  }

  function loadStore() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) {
        return createDefaultStore();
      }
      return normalizeStore(JSON.parse(raw));
    } catch (error) {
      console.warn('[WTTE] Failed to load store', error);
    }
    return createDefaultStore();
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

  function captureOriginalTreeStates() {
    return Array.from(document.querySelectorAll('.unit-tree')).reduce((result, unitTree) => {
      const treeId = unitTree.dataset.treeId;
      const snapshot = serializePersistedTreeState(unitTree);
      if (treeId && snapshot) {
        result[treeId] = snapshot;
      }
      return result;
    }, {});
  }

  function arePersistedTreeStatesEqual(left, right) {
    if (!left || !right) {
      return !left && !right;
    }
    return left.instanceHtml === right.instanceHtml && Boolean(left.arrowsHidden) === Boolean(right.arrowsHidden);
  }

  function collectModifiedTreeSnapshots() {
    return Array.from(document.querySelectorAll('.unit-tree')).reduce((result, unitTree) => {
      const treeId = unitTree.dataset.treeId;
      const snapshot = serializePersistedTreeState(unitTree);
      if (treeId && snapshot && !arePersistedTreeStatesEqual(snapshot, state.originalTrees[treeId])) {
        result[treeId] = snapshot;
      }
      return result;
    }, {});
  }

  function isTreeModified(unitTree) {
    const treeId = unitTree?.dataset?.treeId;
    const snapshot = unitTree ? serializePersistedTreeState(unitTree) : null;
    return Boolean(treeId && snapshot && !arePersistedTreeStatesEqual(snapshot, state.originalTrees[treeId]));
  }

  function buildExportPayload() {
    flushPendingTreeSaves();
    return {
      type: 'wt-tree-editor-export',
      version: 3,
      pathname: location.pathname,
      exportedAt: new Date().toISOString(),
      modifiedTrees: collectModifiedTreeSnapshots(),
      tabs: cloneTabStore(state.store.tabs),
    };
  }

  function buildExportFilename() {
    const pageSlug = location.pathname.replace(/^\/+|\/+$/g, '').replace(/[^\w-]+/g, '-') || 'wiki';
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    return `wtte-${pageSlug}-${timestamp}.json`;
  }

  function flushPendingTreeSaves() {
    state.saveTimers.forEach((timer, treeId) => {
      clearTimeout(timer);
      const unitTree = document.querySelector(`.unit-tree[data-tree-id="${cssEscape(treeId)}"]`);
      if (unitTree) {
        persistTree(unitTree);
      }
    });
    state.saveTimers.clear();
  }

  function normalizeImportedStore(rawStore) {
    if (!rawStore || typeof rawStore !== 'object' || !rawStore.modifiedTrees || typeof rawStore.modifiedTrees !== 'object') {
      return null;
    }

    const modifiedTrees = Object.entries(rawStore.modifiedTrees).reduce((result, [treeId, snapshot]) => {
      const normalizedSnapshot = normalizePersistedTreeState(snapshot);
      if (normalizedSnapshot) {
        result[treeId] = normalizedSnapshot;
      }
      return result;
    }, {});

    return {
      version: 3,
      pathname: typeof rawStore.pathname === 'string' ? rawStore.pathname : '',
      modifiedTrees,
      tabs: normalizeTabStore(rawStore.tabs),
      updatedAt: typeof rawStore.updatedAt === 'string' ? rawStore.updatedAt : new Date().toISOString(),
    };
  }

  function applyImportedStore(importedStore) {
    if (!importedStore) {
      return;
    }

    state.saveTimers.forEach((timer) => clearTimeout(timer));
    state.saveTimers.clear();
    hideMenu();
    clearHoverState();
    clearSelection();
    clearGroupSelection();
    cleanupDrag();
    closeModal();
    state.store = normalizeStore({
      version: importedStore.version,
      modifiedTrees: importedStore.modifiedTrees,
      tabs: importedStore.tabs,
      updatedAt: importedStore.updatedAt,
    });
    restoreStoredTabs();
    state.originalTrees = captureOriginalTreeStates();
    restoreSavedTrees();
    state.undoStacks.clear();
    state.redoStacks.clear();
    activateTreeTab(getInitialActiveTreeId());
    saveStore();
    updateToolbar();
  }

  function hasTabMetadataChanges(tabStore = state.store.tabs) {
    const normalized = normalizeTabStore(tabStore);
    return Boolean(
      normalized.order.length
      || normalized.hidden.length
      || Object.keys(normalized.labels).length
      || Object.keys(normalized.custom).length
    );
  }

  function getTabsStore() {
    state.store.tabs = normalizeTabStore(state.store.tabs);
    return state.store.tabs;
  }

  function getTreeTabsRoot() {
    return document.querySelector('#wt-tree-tabs');
  }

  function getTreeTabsWrapper() {
    return document.querySelector('#wt-tree-tabs .navtabs_wrapper, #wt-tree-tabs .arrow-scroll_wrapper');
  }

  function getTreeInstancesContainer() {
    return document.querySelector('#wt-unit-trees .unit-trees_instances');
  }

  function getTreeTabItems({ includeHidden = true, includeAdd = false } = {}) {
    return Array.from(document.querySelectorAll('#wt-tree-tabs .navtabs_item[data-tree-target]')).filter((tab) => {
      if (!includeAdd && tab.classList.contains('wtte-add-tab')) {
        return false;
      }
      if (!includeHidden && tab.classList.contains('wtte-tab-hidden')) {
        return false;
      }
      return true;
    });
  }

  function getTreeTabById(treeId) {
    return document.querySelector(`#wt-tree-tabs .navtabs_item[data-tree-target="${cssEscape(treeId)}"]`);
  }

  function getCountryFilterButtons() {
    return Array.from(document.querySelectorAll('#wt-filter-country-items [data-country-id]'));
  }

  function getCountryFilterButton(treeId) {
    return document.querySelector(`#wt-filter-country-items [data-country-id="${cssEscape(treeId)}"]`);
  }

  function getCountryFilterDefaultLabel(button) {
    if (!button) {
      return '';
    }
    if (!button.dataset.wtteDefaultLabel) {
      const text = Array.from(button.childNodes)
        .filter((node) => node.nodeType === Node.TEXT_NODE)
        .map((node) => node.textContent || '')
        .join(' ')
        .replace(/\s+/g, ' ')
        .trim();
      button.dataset.wtteDefaultLabel = text || button.textContent.trim();
    }
    return button.dataset.wtteDefaultLabel || '';
  }

  function setCountryFilterButtonLabel(treeId, label) {
    const button = getCountryFilterButton(treeId);
    if (!button) {
      return;
    }
    const fallbackLabel = getCountryFilterDefaultLabel(button);

    Array.from(button.childNodes).forEach((node) => {
      if (node.nodeType === Node.TEXT_NODE) {
        node.remove();
      }
    });

    let labelNode = button.querySelector('.wtte-country-label');
    if (!labelNode) {
      labelNode = document.createElement('span');
      labelNode.className = 'wtte-country-label';
      button.appendChild(labelNode);
    }

    labelNode.textContent = label || fallbackLabel;
  }

  function syncCountryFilterButtons() {
    getCountryFilterButtons().forEach((button) => {
      const treeId = button.dataset.countryId;
      const tab = getTreeTabById(treeId);
      const fallbackLabel = getCountryFilterDefaultLabel(button);
      button.hidden = Boolean(tab?.classList.contains('wtte-tab-hidden'));
      setCountryFilterButtonLabel(treeId, getTreeTabLabel(tab) || fallbackLabel);
    });
  }

  function getUnitTreeById(treeId) {
    return document.querySelector(`.unit-tree[data-tree-id="${cssEscape(treeId)}"]`);
  }

  function getTreeTabLabel(tabOrId) {
    const tab = typeof tabOrId === 'string' ? getTreeTabById(tabOrId) : tabOrId;
    return tab?.querySelector('.navtabs_item-label')?.textContent?.trim() || '';
  }

  function setTreeTabLabel(treeId, label) {
    const tab = getTreeTabById(treeId);
    const labelNode = tab?.querySelector('.navtabs_item-label');
    if (labelNode) {
      labelNode.textContent = label;
    }
    setCountryFilterButtonLabel(treeId, label);
  }

  function isCustomTreeId(treeId) {
    return Boolean(getTabsStore().custom?.[treeId]);
  }

  function getVisibleTreeTabItems() {
    return getTreeTabItems({ includeHidden: false, includeAdd: false });
  }

  function getNextCustomTreeId() {
    const existing = new Set([
      ...getTreeTabItems({ includeHidden: true, includeAdd: false }).map((tab) => tab.dataset.treeTarget),
      ...Object.keys(getTabsStore().custom || {}),
    ]);
    let index = 1;
    while (existing.has(`wtte-custom-${index}`)) {
      index += 1;
    }
    return `wtte-custom-${index}`;
  }

  function getNextCustomTreeLabel() {
    const customCount = Object.keys(getTabsStore().custom || {}).length + 1;
    return t('content.customTreeName', { index: customCount });
  }

  function createTabElement(treeId, label, options = {}) {
    const source = getTreeTabItems({ includeHidden: true, includeAdd: false })[0] || document.querySelector('#wt-tree-tabs .navtabs_item');
    const tab = source ? source.cloneNode(true) : document.createElement('div');
    tab.className = 'navtabs_item';
    tab.dataset.treeTarget = treeId;
    tab.classList.remove('active', 'wtte-tab-hidden', 'wtte-add-tab', 'wtte-tab-drag-source');
    if (options.custom) {
      tab.dataset.wtteCustomTree = '1';
    } else {
      delete tab.dataset.wtteCustomTree;
    }
    let labelNode = tab.querySelector('.navtabs_item-label');
    if (!labelNode) {
      labelNode = document.createElement('div');
      labelNode.className = 'navtabs_item-label';
      tab.appendChild(labelNode);
    }
    labelNode.textContent = label;
    return tab;
  }

  function createCustomTreeTemplate(baseTree) {
    const clone = baseTree.cloneNode(true);
    stripEditorClasses(clone);
    clone.querySelectorAll('td').forEach((cell) => {
      cell.innerHTML = '';
    });
    clone.querySelector('.wt-tree')?.classList.remove('wtte-hide-arrows');
    return serializePersistedTreeState(clone);
  }

  function normalizeTableShape(table, rowCount, columnCount) {
    if (!table || rowCount < 1 || columnCount < 1) {
      return;
    }

    const seedRow = table.rows[0] || document.createElement('tr');
    while (table.rows.length < rowCount) {
      const newRow = createBlankRowLike(seedRow);
      ensureRowColumnCount(newRow, columnCount, seedRow.cells?.[0] || null);
      table.appendChild(newRow);
    }
    while (table.rows.length > rowCount) {
      table.rows[table.rows.length - 1].remove();
    }

    Array.from(table.rows).forEach((row) => {
      ensureRowColumnCount(row, columnCount, row.cells[row.cells.length - 1] || seedRow.cells?.[0] || null);
      while (row.cells.length > columnCount) {
        row.cells[row.cells.length - 1].remove();
      }
      Array.from(row.cells).forEach((cell) => {
        cell.innerHTML = '';
      });
    });
  }

  function getRankPairs(unitTree) {
    return getRankRows(unitTree).map((rankRow) => ({
      rankRow,
      header: getRankHeader(rankRow),
    }));
  }

  function configureCustomTreeSkeleton(unitTree, { rankCount = 1, rowCount = 1, leftColumns = 5, rightColumns = 2 } = {}) {
    let rankPairs = getRankPairs(unitTree);
    if (!rankPairs.length) {
      return;
    }

    while (rankPairs.length < rankCount) {
      const sourcePair = rankPairs[rankPairs.length - 1] || rankPairs[0];
      const rankClone = sourcePair.rankRow.cloneNode(true);
      const headerClone = sourcePair.header?.cloneNode(true) || null;
      stripEditorClasses(rankClone);
      if (headerClone) {
        stripEditorClasses(headerClone);
      }
      sourcePair.rankRow.insertAdjacentElement('afterend', rankClone);
      if (headerClone) {
        rankClone.insertAdjacentElement('beforebegin', headerClone);
      }
      rankPairs = getRankPairs(unitTree);
    }

    rankPairs = getRankPairs(unitTree);
    rankPairs.forEach(({ rankRow, header }, index) => {
      if (index >= rankCount) {
        header?.remove();
        rankRow.remove();
        return;
      }

      const labelNode = header?.querySelector('.wt-tree_r-header_label span');
      if (labelNode) {
        labelNode.textContent = RANK_LABELS[index] || String(index + 1);
      }

      const tables = getRankRowTables(rankRow);
      normalizeTableShape(tables[0], rowCount, leftColumns);
      if (tables[1]) {
        normalizeTableShape(tables[1], rowCount, rightColumns);
      }
      tables.slice(2).forEach((table) => {
        normalizeTableShape(table, rowCount, getTableColumnCount(table) || 1);
      });
    });
  }

  function createPresetCustomTreeTemplate(baseTree, mode = 'blank') {
    const clone = baseTree.cloneNode(true);
    stripEditorClasses(clone);
    configureCustomTreeSkeleton(clone, mode === 'full'
      ? { rankCount: 9, rowCount: 1, leftColumns: 5, rightColumns: 2 }
      : { rankCount: 1, rowCount: 1, leftColumns: 5, rightColumns: 2 });
    clone.querySelector('.wt-tree')?.classList.remove('wtte-hide-arrows');
    return serializePersistedTreeState(clone);
  }

  function resolveCustomTreeTemplate(meta = {}) {
    const snapshot = normalizePersistedTreeState(meta.template);
    if (snapshot) {
      return snapshot;
    }
    const baseTree = getUnitTreeById(meta.baseTreeId) || document.querySelector('.unit-tree');
    return baseTree ? createCustomTreeTemplate(baseTree) : null;
  }

  function createCustomTreeDom(treeId, meta = {}) {
    if (getUnitTreeById(treeId)) {
      return getUnitTreeById(treeId);
    }

    const baseTree = getUnitTreeById(meta.baseTreeId) || document.querySelector('.unit-tree');
    const container = getTreeInstancesContainer();
    if (!baseTree || !container) {
      return null;
    }

    const unitTree = baseTree.cloneNode(true);
    stripEditorClasses(unitTree);
    unitTree.dataset.treeId = treeId;
    unitTree.dataset.wtteCustomTree = '1';
    unitTree.style.display = 'none';
    const template = resolveCustomTreeTemplate(meta);
    if (template) {
      applyPersistedTreeState(unitTree, template);
    } else {
      unitTree.querySelectorAll('td').forEach((cell) => {
        cell.innerHTML = '';
      });
    }
    container.appendChild(unitTree);
    return unitTree;
  }

  function ensureCustomTree(treeId, meta = {}) {
    const wrapper = getTreeTabsWrapper();
    if (!wrapper) {
      return null;
    }

    let tab = getTreeTabById(treeId);
    if (!tab) {
      tab = createTabElement(treeId, getTabsStore().labels[treeId] || meta.label || getNextCustomTreeLabel(), { custom: true });
      wrapper.appendChild(tab);
    }
    tab.dataset.wtteCustomTree = '1';
    createCustomTreeDom(treeId, meta);
    return tab;
  }

  function removeAllCustomTrees() {
    document.querySelectorAll('#wt-tree-tabs .navtabs_item[data-wtte-custom-tree="1"]').forEach((tab) => {
      tab.remove();
    });
    document.querySelectorAll('.unit-tree[data-wtte-custom-tree="1"]').forEach((tree) => {
      tree.remove();
    });
  }

  function applyStoredTabLabels() {
    Object.entries(getTabsStore().labels).forEach(([treeId, label]) => {
      if (typeof label === 'string' && label.trim()) {
        setTreeTabLabel(treeId, label);
      }
    });
  }

  function reorderTreeDom(order = []) {
    const wrapper = getTreeTabsWrapper();
    const instances = getTreeInstancesContainer();
    if (!wrapper || !instances) {
      return;
    }

    const tabs = new Map(getTreeTabItems({ includeHidden: true, includeAdd: false }).map((tab) => [tab.dataset.treeTarget, tab]));
    const trees = new Map(Array.from(document.querySelectorAll('.unit-tree')).map((tree) => [tree.dataset.treeId, tree]));
    const orderedIds = [...new Set([
      ...order.filter((treeId) => tabs.has(treeId) && trees.has(treeId)),
      ...Array.from(tabs.keys()).filter((treeId) => !order.includes(treeId)),
    ])];

    orderedIds.forEach((treeId) => {
      wrapper.appendChild(tabs.get(treeId));
      instances.appendChild(trees.get(treeId));
    });
  }

  function applyStoredHiddenTabs() {
    const hidden = new Set(getTabsStore().hidden || []);
    getTreeTabItems({ includeHidden: true, includeAdd: false }).forEach((tab) => {
      const treeId = tab.dataset.treeTarget;
      const isHidden = hidden.has(treeId);
      tab.classList.toggle('wtte-tab-hidden', isHidden);
      const unitTree = getUnitTreeById(treeId);
      if (unitTree && isHidden && tab.classList.contains('active')) {
        unitTree.style.display = 'none';
      }
    });
    syncCountryFilterButtons();
  }

  function ensureAddTabButton() {
    const wrapper = getTreeTabsWrapper();
    if (!wrapper) {
      return;
    }
    let addTab = getTreeTabById('wtte-add-tab');
    if (!addTab) {
      addTab = createTabElement('wtte-add-tab', '+');
      addTab.classList.add('wtte-add-tab');
    }
    wrapper.appendChild(addTab);
    updateAddTabVisibility();
  }

  function getInitialActiveTreeId() {
    const activeVisibleTab = getTreeTabItems({ includeHidden: false, includeAdd: false }).find((tab) => tab.classList.contains('active'));
    if (activeVisibleTab) {
      return activeVisibleTab.dataset.treeTarget;
    }
    const visibleTree = getVisibleUnitTree();
    if (visibleTree && !getTreeTabById(visibleTree.dataset.treeId)?.classList.contains('wtte-tab-hidden')) {
      return visibleTree.dataset.treeId;
    }
    return getVisibleTreeTabItems()[0]?.dataset.treeTarget || getTreeTabItems({ includeHidden: true, includeAdd: false })[0]?.dataset.treeTarget || '';
  }

  function activateTreeTab(treeId) {
    if (!treeId) {
      return;
    }

    const tabs = getTreeTabItems({ includeHidden: true, includeAdd: false });
    const targetTab = tabs.find((tab) => tab.dataset.treeTarget === treeId && !tab.classList.contains('wtte-tab-hidden'));
    const activeTreeId = targetTab?.dataset.treeTarget || getVisibleTreeTabItems()[0]?.dataset.treeTarget;
    if (!activeTreeId) {
      return;
    }

    tabs.forEach((tab) => {
      const isActive = tab.dataset.treeTarget === activeTreeId;
      tab.classList.toggle('active', isActive);
    });

    document.querySelectorAll('.unit-tree').forEach((unitTree) => {
      unitTree.style.display = unitTree.dataset.treeId === activeTreeId ? '' : 'none';
    });

    state.selectedTree = getUnitTreeById(activeTreeId);
    scheduleLayoutForAll();
    updateToolbar();
  }

  function collectCurrentTabState() {
    const custom = { ...getTabsStore().custom };
    getTreeTabItems({ includeHidden: true, includeAdd: false }).forEach((tab) => {
      const treeId = tab.dataset.treeTarget;
      if (tab.dataset.wtteCustomTree === '1' && !custom[treeId]) {
        custom[treeId] = {
          baseTreeId: document.querySelector('.unit-tree:not([data-wtte-custom-tree="1"])')?.dataset?.treeId || '',
          template: state.originalTrees[treeId] || serializePersistedTreeState(getUnitTreeById(treeId)),
        };
      }
    });
    return {
      order: getTreeTabItems({ includeHidden: true, includeAdd: false }).map((tab) => tab.dataset.treeTarget),
      hidden: getTreeTabItems({ includeHidden: true, includeAdd: false })
        .filter((tab) => tab.classList.contains('wtte-tab-hidden'))
        .map((tab) => tab.dataset.treeTarget),
      labels: getTreeTabItems({ includeHidden: true, includeAdd: false }).reduce((result, tab) => {
        result[tab.dataset.treeTarget] = getTreeTabLabel(tab);
        return result;
      }, {}),
      custom,
    };
  }

  function persistTabState(showFlash = false) {
    state.store.tabs = collectCurrentTabState();
    state.store.version = 3;
    state.store.updatedAt = new Date().toISOString();
    saveStore();
    updateToolbar();
    if (showFlash) {
      flashStatus(t('flash.tabsUpdated'));
    }
  }

  function restoreStoredTabs() {
    state.store = normalizeStore(state.store);
    removeAllCustomTrees();
    Object.entries(getTabsStore().custom).forEach(([treeId, meta]) => {
      ensureCustomTree(treeId, meta);
    });
    applyStoredTabLabels();
    reorderTreeDom(getTabsStore().order || []);
    applyStoredHiddenTabs();
    ensureAddTabButton();
    syncCountryFilterButtons();
  }

  function canHideTreeTab(treeId) {
    if (!treeId) {
      return false;
    }
    const tab = getTreeTabById(treeId);
    if (!tab || tab.classList.contains('wtte-tab-hidden')) {
      return false;
    }
    return getVisibleTreeTabItems().length > 1;
  }

  function hideTreeTab(treeId) {
    if (!canHideTreeTab(treeId)) {
      flashStatus(t('flash.cannotHideLastTab'));
      return;
    }
    getTreeTabById(treeId)?.classList.add('wtte-tab-hidden');
    if (getTreeTabById(treeId)?.classList.contains('active')) {
      activateTreeTab(getVisibleTreeTabItems().find((tab) => tab.dataset.treeTarget !== treeId)?.dataset.treeTarget || '');
    }
    syncCountryFilterButtons();
    persistTabState(true);
  }

  function showTreeTab(treeId) {
    const tab = getTreeTabById(treeId);
    if (!tab) {
      return;
    }
    tab.classList.remove('wtte-tab-hidden');
    syncCountryFilterButtons();
    persistTabState(true);
  }

  function renameTreeTab(treeId) {
    const currentLabel = getTreeTabLabel(treeId);
    const nextLabel = window.prompt(t('prompt.renameTab'), currentLabel);
    if (!nextLabel || !nextLabel.trim()) {
      return;
    }
    setTreeTabLabel(treeId, nextLabel.trim());
    persistTabState(true);
  }

  function addCustomTree(mode = 'blank') {
    const baseTree = getPrimaryTree() || document.querySelector('.unit-tree');
    if (!baseTree) {
      return;
    }
    const template = mode === 'full'
      ? createPresetCustomTreeTemplate(baseTree, 'full')
      : createPresetCustomTreeTemplate(baseTree, 'blank');
    const treeId = createCustomTreeFromSnapshot(baseTree, template, getNextCustomTreeLabel());
    if (!treeId) {
      return;
    }
    flashStatus(t('flash.customTreeAdded'));
  }

  function createCustomTreeFromSnapshot(baseTree, template, label) {
    const treeId = getNextCustomTreeId();
    getTabsStore().custom[treeId] = {
      baseTreeId: baseTree.dataset.treeId,
      template,
    };
    getTabsStore().labels[treeId] = label;
    ensureCustomTree(treeId, getTabsStore().custom[treeId]);
    state.originalTrees[treeId] = template;
    reorderTreeDom([...collectCurrentTabState().order, treeId]);
    ensureAddTabButton();
    activateTreeTab(treeId);
    persistTabState();
    return treeId;
  }

  function duplicateTree(treeId) {
    const sourceTree = getUnitTreeById(treeId);
    if (!sourceTree) {
      return;
    }
    const baseLabel = getTreeTabLabel(treeId) || t('content.customTreeName', { index: '' }).trim();
    const duplicateLabel = state.language === 'ko' ? `${baseLabel} 복제` : `${baseLabel} Copy`;
    const snapshot = serializePersistedTreeState(sourceTree);
    if (!snapshot) {
      return;
    }
    createCustomTreeFromSnapshot(sourceTree, snapshot, duplicateLabel);
    flashStatus(t('flash.treeDuplicated'));
  }

  function deleteTree(treeId) {
    if (!isCustomTreeId(treeId)) {
      return;
    }
    if (!window.confirm(t('confirm.deleteTree'))) {
      return;
    }

    const tab = getTreeTabById(treeId);
    const unitTree = getUnitTreeById(treeId);
    const fallbackTab = getVisibleTreeTabItems().find((item) => item.dataset.treeTarget !== treeId)
      || getTreeTabItems({ includeHidden: true, includeAdd: false }).find((item) => item.dataset.treeTarget !== treeId);

    tab?.remove();
    unitTree?.remove();
    delete getTabsStore().custom[treeId];
    delete getTabsStore().labels[treeId];
    delete state.store.modifiedTrees[treeId];
    delete state.originalTrees[treeId];
    state.undoStacks.delete(treeId);
    state.redoStacks.delete(treeId);
    clearTimeout(state.saveTimers.get(treeId));
    state.saveTimers.delete(treeId);
    getTabsStore().order = getTabsStore().order.filter((value) => value !== treeId);
    getTabsStore().hidden = getTabsStore().hidden.filter((value) => value !== treeId);

    if (!getVisibleTreeTabItems().length) {
      fallbackTab?.classList.remove('wtte-tab-hidden');
    }

    ensureAddTabButton();
    syncCountryFilterButtons();
    persistTabState();
    activateTreeTab(fallbackTab?.dataset.treeTarget || getInitialActiveTreeId());
    flashStatus(t('flash.treeDeleted'));
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
    syncArrowVisibility(unitTree);
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

  function getSelectedTableContext() {
    const cell = ensureSelectedCell();
    const row = cell?.parentElement || null;
    const table = row?.closest('table') || null;
    const unitTree = cell?.closest('.unit-tree') || null;
    const columnIndex = row && cell ? getCellIndex(row, cell) : -1;
    const rowIndex = table && row ? Array.from(table.rows).indexOf(row) : -1;
    const sectionIndex = table ? getTableSectionIndex(table) : -1;
    const sectionTables = unitTree && sectionIndex >= 0 ? getSectionTables(unitTree, sectionIndex) : [];
    if (!cell || !row || !table || !unitTree || columnIndex < 0 || rowIndex < 0 || sectionIndex < 0 || !sectionTables.length) {
      return null;
    }
    return {
      cell,
      row,
      table,
      unitTree,
      rowIndex,
      columnIndex,
      sectionIndex,
      sectionTables,
    };
  }

  function getTableSectionIndex(table) {
    const rankRow = table?.closest('.wt-tree_rank');
    if (!rankRow) {
      return -1;
    }
    const sections = Array.from(rankRow.children).filter((child) => child.matches('div:not(.wt-tree_v-line)'));
    return sections.findIndex((section) => section.contains(table));
  }

  function getSectionTables(unitTree, sectionIndex) {
    return getRankRows(unitTree)
      .map((rankRow) => {
        return getRankRowTables(rankRow)[sectionIndex] || null;
      })
      .filter(Boolean);
  }

  function getRankRowTables(rankRow) {
    return Array.from(rankRow?.children || [])
      .filter((child) => child.matches('div:not(.wt-tree_v-line)'))
      .map((section) => section.querySelector('table.wt-tree_rank-instance') || null);
  }

  function setSelectionToColumnCell(table, rowIndex, columnIndex) {
    const row = table?.rows?.[rowIndex] || table?.rows?.[0] || null;
    const cell = row?.cells?.[columnIndex] || row?.cells?.[row?.cells?.length - 1] || null;
    setSelection({
      cell,
      row,
      rankRow: row ? row.closest('.wt-tree_rank') : null,
    });
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

  function ensureRowColumnCount(row, minimumCount, templateCell) {
    while (row.cells.length < minimumCount) {
      row.appendChild(createCellLike(row.cells[row.cells.length - 1] || templateCell, false));
    }
  }

  function normalizeTableColumnCount(table, minimumCount, templateCell) {
    const targetCount = Math.max(getTableColumnCount(table), minimumCount);
    Array.from(table?.rows || []).forEach((row) => {
      ensureRowColumnCount(row, targetCount, row.cells[row.cells.length - 1] || templateCell);
    });
    return targetCount;
  }

  function insertCellAtIndex(row, index, cell) {
    if (!row || !cell) {
      return;
    }
    const referenceCell = row.cells[index];
    if (referenceCell) {
      referenceCell.insertAdjacentElement('beforebegin', cell);
    } else {
      row.appendChild(cell);
    }
  }

  function getTableColumnCount(table) {
    return Math.max(...Array.from(table?.rows || []).map((row) => row.cells.length), 0);
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
    if (state.tabDrag?.tab?.isConnected) {
      state.tabDrag.tab.classList.remove('wtte-tab-drag-source');
    }
    state.tabDrag = null;
    state.selectionDrag = null;
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
