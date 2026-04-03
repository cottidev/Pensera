const STORAGE_KEY = "pensera-journal-state";
const THEME_META_SELECTOR = 'meta[name="theme-color"]';
const THEME_COLORS = {
  light: "#f3f0e8",
  dark: "#14181f",
};
const MIN_SIDEBAR_WIDTH = 300;
const MAX_SIDEBAR_WIDTH = 520;
const MIN_EDITOR_PANEL_WIDTH = 760;
const HIGHLIGHT_COLORS = {
  light: ["#fff29a", "#c7f1c9", "#bfe3ff", "#ffd6ea"],
  dark: ["#8c7421", "#285a3d", "#204d74", "#74365d"],
};
const TEXT_COLORS = {
  light: ["#b84c4c", "#2d6a4f", "#2f6fed"],
  dark: ["#ef8c8c", "#6fc79a", "#8ab4ff"],
};
const URL_PATTERN = /(?:https?:\/\/|www\.)[^\s<]+[^\s<.,:;"')\]\}]/i;
const URL_GLOBAL_PATTERN = /(?:https?:\/\/|www\.)[^\s<]+[^\s<.,:;"')\]\}]/gi;

const FONT_MAP = {
  inter: "Inter",
  "dm-sans": "DM Sans",
  manrope: "Manrope",
  "space-grotesk": "Space Grotesk",
  serif: "Instrument Serif",
  crimson: "Crimson Text",
  baskerville: "Libre Baskerville",
  "eb-garamond": "EB Garamond",
  cormorant: "Cormorant Garamond",
};

const FONT_SIZE_OPTIONS = [
  "16px",
  "18px",
  "20px",
  "22px",
  "24px",
  "26px",
  "28px",
  "30px",
];
const PAGE_WIDTH_OPTIONS = ["740px", "860px", "980px"];
const DEFAULT_STATE = {
  entries: [],
  groups: [],
  selectedId: null,
  activeGroupId: "all",
  searchTerm: "",
  entrySort: "recent",
  fontKey: "crimson",
  fontSize: "18px",
  lineHeight: "1.82",
  pageWidth: "860px",
  sidebarWidth: 320,
  theme: "light",
  focusMode: false,
};

const entryList = document.querySelector("[data-entry-list]");
const groupList = document.querySelector("[data-group-list]");
const newEntryButton = document.querySelector("[data-new-entry]");
const openGroupFormButton = document.querySelector("[data-open-group-form]");
const importJsonButton = document.querySelector("[data-import-json]");
const exportJsonButton = document.querySelector("[data-export-json]");
const importFileInput = document.querySelector("[data-import-file]");
const imageUploadInput = document.querySelector("[data-image-upload-file]");
const groupForm = document.querySelector("[data-group-form]");
const groupInput = document.querySelector("[data-group-input]");
const groupSubmitButton = document.querySelector("[data-group-submit]");
const cancelGroupButton = document.querySelector("[data-cancel-group]");
const searchInput = document.querySelector("[data-search]");
const entrySortDropdown = document.querySelector("[data-entry-sort-dropdown]");
const entrySortTrigger = document.querySelector("[data-entry-sort-trigger]");
const entrySortLabel = document.querySelector("[data-entry-sort-label]");
const entrySortMenu = document.querySelector("[data-entry-sort-menu]");
const entrySortOptions = document.querySelectorAll("[data-entry-sort-option]");
const titleInput = document.querySelector("[data-title]");
const editor = document.querySelector("[data-editor]");
const currentTime = document.querySelector("[data-current-time]");
const activeGroupLabel = document.querySelector("[data-active-group]");
const entryDate = document.querySelector("[data-entry-date]");
const saveStatus = document.querySelector("[data-save-status]");
const editorStats = document.querySelector("[data-editor-stats]");
const lineCount = document.querySelector("[data-line-count]");
const wordCount = document.querySelector("[data-word-count]");
const characterCount = document.querySelector("[data-character-count]");
const entryTemplate = document.querySelector("#entry-item-template");
const groupTemplate = document.querySelector("#group-item-template");
const pickerEntryTemplate = document.querySelector("#picker-entry-template");
const groupFlyout = document.querySelector("[data-group-flyout]");
const groupFlyoutTitle = document.querySelector("[data-group-flyout-title]");
const groupFlyoutActions = document.querySelectorAll(
  "[data-group-flyout-action]",
);
const commandButtons = document.querySelectorAll("[data-command]");
const clearHighlightToolbarButton = document.querySelector(
  "[data-clear-highlight-toolbar]",
);
const exportPdfToolbarButton = document.querySelector(
  "[data-export-pdf-toolbar]",
);
const duplicateEntryButton = document.querySelector("[data-duplicate-entry]");
const fontSelect = document.querySelector("[data-font-select]");
const fontSizeSelect = document.querySelector("[data-font-size-select]");
const pageWidthSelect = document.querySelector("[data-page-width-select]");
const pageWidthDropdown = document.querySelector("[data-page-width-dropdown]");
const pageWidthTrigger = document.querySelector("[data-page-width-trigger]");
const pageWidthMenu = document.querySelector("[data-page-width-menu]");
const pageWidthLabel = document.querySelector("[data-page-width-label]");
const pageWidthOptions = document.querySelectorAll("[data-page-width-option]");
const linkToolbarButton = document.querySelector("[data-link-toolbar]");
const toggleChecklistButton = document.querySelector("[data-toggle-checklist]");
const imageUploadTrigger = document.querySelector(
  "[data-image-upload-trigger]",
);
const imageTools = document.querySelector("[data-image-tools]");
const imageSizeRange = document.querySelector("[data-image-size-range]");
const editorContextMenu = document.querySelector("[data-editor-context-menu]");
const highlightColorButtons = document.querySelectorAll(
  "[data-highlight-color]",
);
const textColorButtons = document.querySelectorAll("[data-text-color]");
const clearHighlightButton = document.querySelector("[data-clear-highlight]");
const clearTextColorButton = document.querySelector("[data-clear-text-color]");
const entryPicker = document.querySelector("[data-entry-picker]");
const pickerList = document.querySelector("[data-picker-list]");
const closePickerButtons = document.querySelectorAll("[data-close-picker]");
const confirmModal = document.querySelector("[data-confirm-modal]");
const confirmText = document.querySelector("[data-confirm-text]");
const confirmAcceptButton = document.querySelector("[data-confirm-accept]");
const closeConfirmButtons = document.querySelectorAll("[data-close-confirm]");
const themeToggleButton = document.querySelector("[data-theme-toggle]");
const focusToggleButton = document.querySelector("[data-focus-toggle]");
const sidebarResizer = document.querySelector("[data-sidebar-resizer]");

let state = loadState();
let pickerGroupId = null;
let confirmAction = null;
let savedSelection = null;
let lastRenderedGroupId = state.activeGroupId;
let editingGroupId = null;
let activeGroupFlyoutId = null;
let dragDepth = 0;
let activeEditorImage = null;
let sidebarResizeSession = null;

initialize();

function initialize() {
  document.execCommand("styleWithCSS", false, true);
  registerServiceWorker();
  startClock();
  state.theme = resolveInitialTheme(state.theme);
  applyTheme(state.theme);
  applyFocusMode(Boolean(state.focusMode));
  applyEditorPreferences();
  applySidebarWidth();
  syncEditorStatsDensity();

  if (!state.groups.length) {
    state.groups = [
      { id: "all", name: "All Entries", pinned: false },
      { id: "notes", name: "Notes", pinned: false },
      { id: "personal", name: "Personal", pinned: false },
    ];
  }

  if (!state.entries.length) {
    const firstEntry = createEntry();
    state.entries = [firstEntry];
    state.selectedId = firstEntry.id;
  } else if (!state.selectedId) {
    state.selectedId = state.entries[0].id;
  }

  persistState();
  render();
  bindEvents();
}

function registerServiceWorker() {
  if (!("serviceWorker" in navigator)) return;
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("./sw.js").catch(() => {
      // Keep the app usable even if service worker registration fails.
    });
  });
}

function startClock() {
  if (!currentTime) return;
  updateClock();
  const now = new Date();
  const delayUntilNextMinute =
    (60 - now.getSeconds()) * 1000 - now.getMilliseconds();

  window.setTimeout(() => {
    updateClock();
    window.setInterval(updateClock, 60 * 1000);
  }, delayUntilNextMinute);
}

function updateClock() {
  if (!currentTime) return;
  currentTime.textContent = new Intl.DateTimeFormat(undefined, {
    hour: "numeric",
    minute: "2-digit",
  }).format(new Date());
}

function bindEvents() {
  newEntryButton.addEventListener("click", () => {
    createAndSelectEntry();
  });

  openGroupFormButton.addEventListener("click", () => {
    openGroupForm();
  });

  importJsonButton.addEventListener("click", () => {
    importFileInput.click();
  });

  exportJsonButton.addEventListener("click", exportJsonState);

  importFileInput.addEventListener("change", async () => {
    const file = importFileInput.files?.[0];
    if (!file) return;

    try {
      const text = await file.text();
      const parsed = JSON.parse(text);
      state = normalizeImportedState(parsed);
      persistState();
      render();
      pulseStatus("Data imported");
    } catch {
      pulseStatus("Import failed");
    } finally {
      importFileInput.value = "";
    }
  });

  imageUploadInput?.addEventListener("change", async () => {
    const files = Array.from(imageUploadInput.files || []);
    if (!files.length) return;
    await insertFilesAsMedia(files);
    imageUploadInput.value = "";
  });

  cancelGroupButton.addEventListener("click", closeGroupForm);

  groupForm.addEventListener("submit", (event) => {
    event.preventDefault();
    const name = groupInput.value.trim().slice(0, 18);
    if (!name) return;

    if (editingGroupId) {
      const group = state.groups.find((item) => item.id === editingGroupId);
      if (!group || group.id === "all") return;
      group.name = name;
      persistState();
      render();
      closeGroupForm();
      pulseStatus("Group renamed");
      return;
    }

    state.groups.push({
      id: createId(),
      name,
      pinned: false,
    });
    state.activeGroupId = state.groups[state.groups.length - 1].id;

    persistState();
    render();
    closeGroupForm();
    pulseStatus("Group created");
  });

  searchInput.addEventListener("input", () => {
    state.searchTerm = searchInput.value.trim().toLowerCase();
    renderList();
  });

  entrySortTrigger?.addEventListener("click", () => {
    const isOpen = !entrySortMenu?.classList.contains("is-hidden");
    setEntrySortMenuOpen(!isOpen);
  });

  entrySortOptions.forEach((option) => {
    option.addEventListener("click", () => {
      state.entrySort = option.dataset.entrySortOption || "recent";
      persistState();
      renderEntrySort();
      renderList();
      setEntrySortMenuOpen(false);
      pulseStatus(`Sorted by ${getEntrySortLabel(state.entrySort)}`);
    });
  });

  titleInput.addEventListener("input", () => {
    const entry = getSelectedEntry();
    if (!entry) return;
    entry.title = titleInput.value;
    touchEntry(entry);
  });

  editor.addEventListener("input", () => {
    const entry = getSelectedEntry();
    if (!entry) return;
    entry.body = normalizeEditorMarkup(editor.innerHTML);
    touchEntry(entry);
  });

  editor.addEventListener("mouseup", syncControlsToSelection);
  editor.addEventListener("keyup", syncControlsToSelection);
  editor.addEventListener("focus", syncControlsToSelection);
  editor.addEventListener("click", (event) => {
    const image = event.target.closest(".editor-media img");
    if (image && editor.contains(image)) {
      event.preventDefault();
      selectEditorImage(image);
      return;
    }

    const checklistItem = event.target.closest(".editor-checklist li");
    if (checklistItem && editor.contains(checklistItem)) {
      const listRect = checklistItem.getBoundingClientRect();
      if (event.clientX <= listRect.left + 28) {
        event.preventDefault();
        checklistItem.classList.toggle("is-checked");
        handleFormattedContentChange();
        return;
      }
    }

    const link = event.target.closest("a[href]");
    if (!link || !editor.contains(link)) return;
    if (event.metaKey || event.ctrlKey) {
      event.preventDefault();
      window.open(link.href, "_blank", "noopener");
      return;
    }
    pulseStatus("Use Ctrl/Cmd-click to open links");
  });
  editor.addEventListener("paste", async (event) => {
    const files = Array.from(event.clipboardData?.files || []).filter((file) =>
      file.type.startsWith("image/"),
    );
    if (!files.length) return;
    event.preventDefault();
    await insertFilesAsMedia(files);
  });
  editor.addEventListener("dragenter", (event) => {
    if (!hasImageTransfer(event.dataTransfer)) return;
    event.preventDefault();
    dragDepth += 1;
    editor.classList.add("is-drag-over");
  });
  editor.addEventListener("dragover", (event) => {
    if (!hasImageTransfer(event.dataTransfer)) return;
    event.preventDefault();
    event.dataTransfer.dropEffect = "copy";
    editor.classList.add("is-drag-over");
  });
  editor.addEventListener("dragleave", (event) => {
    event.preventDefault();
    dragDepth = Math.max(0, dragDepth - 1);
    if (dragDepth === 0) {
      editor.classList.remove("is-drag-over");
    }
  });
  editor.addEventListener("drop", async (event) => {
    const files = Array.from(event.dataTransfer?.files || []).filter((file) =>
      file.type.startsWith("image/"),
    );
    if (!files.length) return;
    event.preventDefault();
    dragDepth = 0;
    editor.classList.remove("is-drag-over");
    placeCaretFromPoint(event.clientX, event.clientY);
    await insertFilesAsMedia(files);
  });
  editor.addEventListener("contextmenu", (event) => {
    if (!hasEditableSelection()) {
      hideEditorContextMenu();
      return;
    }

    event.preventDefault();
    cacheSelection();
    openEditorContextMenu(event.clientX, event.clientY);
  });
  document.addEventListener("selectionchange", () => {
    if (
      document.activeElement === editor ||
      editor.contains(document.activeElement)
    ) {
      cacheSelection();
      syncControlsToSelection();
    }
  });

  fontSelect.addEventListener("change", () => {
    state.fontKey = fontSelect.value;
    applyFontToSelection(fontSelect.value);
    persistState();
  });

  fontSizeSelect.addEventListener("change", () => {
    state.fontSize = fontSizeSelect.value;
    applyFontSizeToSelection(fontSizeSelect.value);
    persistState();
    applyEditorPreferences();
  });

  pageWidthSelect?.addEventListener("change", () => {
    state.pageWidth = pageWidthSelect.value;
    applyEditorPreferences();
    persistState();
    pulseStatus("Page width updated");
  });

  // Custom page-width dropdown
  if (pageWidthTrigger && pageWidthMenu) {
    pageWidthTrigger.addEventListener("click", (e) => {
      e.stopPropagation();
      const isOpen = !pageWidthMenu.classList.contains("is-hidden");
      pageWidthMenu.classList.toggle("is-hidden", isOpen);
      pageWidthTrigger.setAttribute("aria-expanded", String(!isOpen));
    });

    pageWidthOptions.forEach((btn) => {
      btn.addEventListener("click", () => {
        const val = btn.dataset.pageWidthOption;
        state.pageWidth = val;
        applyEditorPreferences();
        persistState();
        pulseStatus("Page width updated");
        pageWidthMenu.classList.add("is-hidden");
        pageWidthTrigger.setAttribute("aria-expanded", "false");
      });
    });

    document.addEventListener("click", (e) => {
      if (!pageWidthDropdown?.contains(e.target)) {
        pageWidthMenu.classList.add("is-hidden");
        pageWidthTrigger.setAttribute("aria-expanded", "false");
      }
    });

    document.addEventListener("keydown", (e) => {
      if (
        e.key === "Escape" &&
        !pageWidthMenu.classList.contains("is-hidden")
      ) {
        pageWidthMenu.classList.add("is-hidden");
        pageWidthTrigger.setAttribute("aria-expanded", "false");
      }
    });
  }

  fontSelect.addEventListener("wheel", (event) => {
    event.preventDefault();
    cycleSelectOption(fontSelect, event.deltaY);
    state.fontKey = fontSelect.value;
    applyFontToSelection(fontSelect.value);
    persistState();
  });

  fontSizeSelect.addEventListener("wheel", (event) => {
    event.preventDefault();
    cycleSelectOption(fontSizeSelect, event.deltaY);
    state.fontSize = fontSizeSelect.value;
    applyFontSizeToSelection(fontSizeSelect.value);
    persistState();
    applyEditorPreferences();
  });

  commandButtons.forEach((button) => {
    button.addEventListener("mousedown", (event) => {
      event.preventDefault();
    });

    button.addEventListener("click", () => {
      const command = button.dataset.command;
      const value = button.dataset.commandValue;
      restoreSelection();
      document.execCommand(command, false, value);
      editor.focus();
      handleFormattedContentChange();
    });
  });

  clearHighlightToolbarButton?.addEventListener("mousedown", (event) => {
    event.preventDefault();
  });

  clearHighlightToolbarButton?.addEventListener("click", () => {
    removeHighlightFromSelection();
  });

  linkToolbarButton?.addEventListener("mousedown", (event) => {
    event.preventDefault();
  });

  linkToolbarButton?.addEventListener("click", () => {
    openLinkEditor();
  });

  toggleChecklistButton?.addEventListener("mousedown", (event) => {
    event.preventDefault();
  });

  toggleChecklistButton?.addEventListener("click", () => {
    toggleChecklist();
  });

  imageUploadTrigger?.addEventListener("click", () => {
    imageUploadInput?.click();
  });

  imageUploadTrigger?.addEventListener("mousedown", (event) => {
    event.preventDefault();
  });

  imageSizeRange?.addEventListener("input", () => {
    resizeActiveImage(imageSizeRange.value);
  });

  sidebarResizer?.addEventListener("mousedown", (event) => {
    event.preventDefault();
    startSidebarResize(event.clientX);
  });

  sidebarResizer?.addEventListener("keydown", (event) => {
    const currentWidth = normalizeSidebarWidth(state.sidebarWidth);
    if (event.key === "ArrowLeft") {
      event.preventDefault();
      updateSidebarWidth(currentWidth - 16);
      return;
    }
    if (event.key === "ArrowRight") {
      event.preventDefault();
      updateSidebarWidth(currentWidth + 16);
    }
  });

  exportPdfToolbarButton?.addEventListener("click", () => {
    exportCurrentEntryAsPdf();
  });

  duplicateEntryButton?.addEventListener("click", () => {
    duplicateSelectedEntry();
  });

  highlightColorButtons.forEach((button) => {
    button.addEventListener("mousedown", (event) => {
      event.preventDefault();
    });
    button.addEventListener("click", () => {
      applyHighlightToSelection(button.dataset.highlightColor);
      hideEditorContextMenu();
    });
  });

  textColorButtons.forEach((button) => {
    button.addEventListener("mousedown", (event) => {
      event.preventDefault();
    });
    button.addEventListener("click", () => {
      applyTextColorToSelection(button.dataset.textColor);
      hideEditorContextMenu();
    });
  });

  clearHighlightButton?.addEventListener("mousedown", (event) => {
    event.preventDefault();
  });

  clearHighlightButton?.addEventListener("click", () => {
    removeHighlightFromSelection();
    hideEditorContextMenu();
  });

  clearTextColorButton?.addEventListener("mousedown", (event) => {
    event.preventDefault();
  });

  clearTextColorButton?.addEventListener("click", () => {
    removeTextColorFromSelection();
    hideEditorContextMenu();
  });

  document.addEventListener("mousedown", (event) => {
    if (
      imageTools?.contains(event.target) ||
      event.target.closest(".editor-media img")
    ) {
      return;
    }
    clearActiveImageSelection();
  });

  document.addEventListener("mousedown", (event) => {
    if (editorContextMenu.classList.contains("is-hidden")) return;
    if (editorContextMenu.contains(event.target)) return;
    hideEditorContextMenu();
  });

  document.addEventListener("mousedown", (event) => {
    if (
      event.target.closest(".group-menu-wrap") ||
      groupFlyout?.contains(event.target)
    ) {
      return;
    }
    closeGroupFlyout();
  });

  document.addEventListener("mousedown", (event) => {
    if (entrySortDropdown?.contains(event.target)) return;
    setEntrySortMenuOpen(false);
  });

  document.addEventListener("keydown", (event) => {
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "e") {
      event.preventDefault();
      createAndSelectEntry();
      return;
    }

    if (
      (event.ctrlKey || event.metaKey) &&
      event.shiftKey &&
      event.key.toLowerCase() === "d"
    ) {
      event.preventDefault();
      duplicateSelectedEntry();
      return;
    }

    if ((event.ctrlKey || event.metaKey) && event.shiftKey) {
      const lowerKey = event.key.toLowerCase();
      if (lowerKey === "7") {
        event.preventDefault();
        toggleChecklist();
        return;
      }
      if (lowerKey === "k") {
        event.preventDefault();
        openLinkEditor();
        return;
      }
      if (lowerKey === "1" || lowerKey === "2" || lowerKey === "3") {
        event.preventDefault();
        document.execCommand("formatBlock", false, `h${lowerKey}`);
        handleFormattedContentChange();
        return;
      }
    }

    if (event.key === "Escape") {
      hideEditorContextMenu();
      closeGroupFlyout();
      setEntrySortMenuOpen(false);
      closeEntryPicker();
      closeConfirmModal();
      editor.classList.remove("is-drag-over");
      dragDepth = 0;
      clearActiveImageSelection();
      stopSidebarResize();
    }
  });

  window.addEventListener("mousemove", (event) => {
    if (!sidebarResizeSession) return;
    updateSidebarWidth(
      sidebarResizeSession.startWidth +
        event.clientX -
        sidebarResizeSession.startX,
      false,
    );
  });

  window.addEventListener("mouseup", () => {
    if (!sidebarResizeSession) return;
    stopSidebarResize(true);
  });

  if (window.ResizeObserver && editorStats) {
    const resizeObserver = new ResizeObserver(() => {
      syncEditorStatsDensity();
    });
    resizeObserver.observe(editorStats);
  }

  groupFlyoutActions.forEach((button) => {
    button.addEventListener("click", () => {
      const group = state.groups.find(
        (item) => item.id === activeGroupFlyoutId,
      );
      if (!group || group.id === "all") return;

      closeGroupFlyout();

      if (button.dataset.groupFlyoutAction === "edit") {
        openGroupForm(group);
        return;
      }

      if (button.dataset.groupFlyoutAction === "export") {
        exportJsonBlob(
          {
            groups: [group],
            entries: state.entries.filter(
              (entry) => entry.groupId === group.id,
            ),
          },
          `${slugify(group.name)}-group.json`,
        );
        pulseStatus("Group exported");
      }
    });
  });

  themeToggleButton.addEventListener("click", () => {
    const previousTheme = state.theme === "dark" ? "dark" : "light";
    state.theme = state.theme === "dark" ? "light" : "dark";
    applyTheme(state.theme, previousTheme);
    persistState();
    pulseStatus(
      state.theme === "dark" ? "Dark mode enabled" : "Light mode enabled",
    );
  });

  focusToggleButton?.addEventListener("click", () => {
    state.focusMode = !state.focusMode;
    applyFocusMode(state.focusMode);
    renderFocusToggle();
    persistState();
    pulseStatus(state.focusMode ? "Focus mode enabled" : "Focus mode disabled");
  });

  window.addEventListener("scroll", hideEditorContextMenu, true);
  window.addEventListener("resize", hideEditorContextMenu);

  closePickerButtons.forEach((button) => {
    button.addEventListener("click", closeEntryPicker);
  });

  closeConfirmButtons.forEach((button) => {
    button.addEventListener("click", closeConfirmModal);
  });

  confirmAcceptButton.addEventListener("click", () => {
    if (confirmAction) confirmAction();
    closeConfirmModal();
  });
}

function render() {
  renderThemeToggle();
  renderFocusToggle();
  renderEntrySort();
  applyEditorPreferences();
  renderGroups();
  renderList();
  renderEditor();
  syncControlsToSelection();
}

function renderEntrySort() {
  if (!entrySortLabel) return;
  entrySortLabel.textContent = getEntrySortLabel(state.entrySort);
  entrySortOptions.forEach((option) => {
    const isActive = option.dataset.entrySortOption === state.entrySort;
    option.classList.toggle("is-active", isActive);
    option.setAttribute("aria-selected", String(isActive));
  });
}

function setEntrySortMenuOpen(isOpen) {
  if (!entrySortMenu || !entrySortTrigger) return;
  entrySortMenu.classList.toggle("is-hidden", !isOpen);
  entrySortTrigger.setAttribute("aria-expanded", String(isOpen));
}

function renderThemeToggle() {
  if (!themeToggleButton) return;
  const isDark = state.theme === "dark";
  themeToggleButton.setAttribute("aria-pressed", String(isDark));
  themeToggleButton.setAttribute(
    "aria-label",
    isDark ? "Switch to light mode" : "Switch to dark mode",
  );
  themeToggleButton.title = isDark ? "Light mode" : "Dark mode";
  syncHighlightButtons();
}

function renderFocusToggle() {
  if (!focusToggleButton) return;
  const isFocused = Boolean(state.focusMode);
  focusToggleButton.textContent = isFocused ? "Exit Focus" : "Focus";
  focusToggleButton.setAttribute("aria-pressed", String(isFocused));
  focusToggleButton.setAttribute(
    "aria-label",
    isFocused ? "Exit focus mode" : "Enter focus mode",
  );
  focusToggleButton.title = isFocused ? "Exit focus mode" : "Enter focus mode";
}

function renderGroups() {
  closeGroupFlyout();
  groupList.innerHTML = "";

  getVisibleGroups().forEach((group) => {
    const fragment = groupTemplate.content.cloneNode(true);
    const row = fragment.querySelector(".group-row");
    const actions = fragment.querySelector(".group-actions");
    const filterButton = fragment.querySelector(".group-item");
    const pinButton = fragment.querySelector(".group-pin-button");
    const pinIcon = fragment.querySelector(".fa-solid");
    const addButton = fragment.querySelector(".group-add-button");
    const menuButton = fragment.querySelector(".group-menu-button");
    const deleteButton = fragment.querySelector(".group-delete-button");
    const name = fragment.querySelector(".group-name");
    const count = fragment.querySelector(".group-count");

    name.textContent = group.name;
    count.textContent = String(countEntriesForGroup(group.id));
    const isActive = group.id === state.activeGroupId;
    filterButton.classList.toggle("is-active", isActive);
    row.classList.toggle("is-active", isActive);
    row.classList.toggle("is-pinned", Boolean(group.pinned));
    pinButton.classList.toggle("is-pinned", Boolean(group.pinned));
    pinIcon.classList.toggle("fa-thumbtack", !group.pinned);
    pinIcon.classList.toggle("fa-thumbtack-slash", Boolean(group.pinned));
    pinButton.setAttribute(
      "aria-label",
      group.pinned ? "Unpin group" : "Pin group",
    );
    pinButton.title = group.pinned ? "Unpin group" : "Pin group";

    if (group.id === "all") {
      actions.classList.add("is-hidden");
      row.classList.add("is-simple");
    }

    menuButton.addEventListener("click", (event) => {
      event.stopPropagation();
      const isOpen = activeGroupFlyoutId === group.id;
      closeGroupFlyout();
      if (!isOpen) openGroupFlyout(group, menuButton);
    });

    filterButton.addEventListener("click", () => {
      state.activeGroupId = group.id;
      const visibleEntries = getVisibleEntries();
      if (!visibleEntries.some((entry) => entry.id === state.selectedId)) {
        state.selectedId =
          visibleEntries[0]?.id || state.entries[0]?.id || null;
      }
      persistState();
      render();
    });

    pinButton.addEventListener("click", () => {
      toggleGroupPin(group.id);
    });

    addButton.addEventListener("click", () => {
      if (group.id === "all") return;
      openEntryPicker(group.id);
    });

    deleteButton.addEventListener("click", () => {
      closeGroupFlyout();
      openConfirmModal(
        `Delete "${group.name}"? Entries inside it will stay, but they will become ungrouped.`,
        () => {
          state.groups = state.groups.filter((item) => item.id !== group.id);
          state.entries = state.entries.map((entry) =>
            entry.groupId === group.id ? { ...entry, groupId: "" } : entry,
          );
          if (state.activeGroupId === group.id) {
            state.activeGroupId = "all";
          }
          persistState();
          render();
          pulseStatus("Group deleted");
        },
      );
    });

    groupList.appendChild(fragment);
  });
}

function renderList() {
  const shouldAnimateGroupSwitch = lastRenderedGroupId !== state.activeGroupId;
  entryList.innerHTML = "";
  const visibleEntries = getVisibleEntries();

  if (!visibleEntries.length) {
    const emptyState = document.createElement("div");
    emptyState.className = "empty-state";
    emptyState.innerHTML = `
      <div class="empty-state-eyebrow">No entries found</div>
      <h3 class="empty-state-title">${
        state.searchTerm
          ? "Nothing matches your search."
          : state.activeGroupId === "all"
            ? "Your journal is ready for its first page."
            : "This group is still empty."
      }</h3>
      <p class="empty-state-copy">${
        state.searchTerm
          ? "Try a broader keyword or clear the search field to see more entries."
          : state.activeGroupId === "all"
            ? "Create a new entry and start building a calmer, better organized writing space."
            : "Move an existing entry into this group or create a new one here."
      }</p>
    `;
    entryList.appendChild(emptyState);
    lastRenderedGroupId = state.activeGroupId;
    return;
  }

  visibleEntries.forEach((entry, index) => {
    const fragment = entryTemplate.content.cloneNode(true);
    const card = fragment.querySelector(".entry-item");
    const openButton = fragment.querySelector(".entry-open-button");
    const pinButton = fragment.querySelector(".entry-pin-button");
    const pinIcon = fragment.querySelector(".fa-solid");
    const exportButton = fragment.querySelector(".entry-export-button");
    const deleteButton = fragment.querySelector(".entry-delete-button");
    const title = fragment.querySelector(".entry-item-title");
    const meta = fragment.querySelector(".entry-item-meta");
    const date = fragment.querySelector(".entry-item-date");
    const group = fragment.querySelector(".entry-item-group");

    title.textContent = entry.title.trim() || "Untitled entry";
    const groupName =
      state.groups.find((groupItem) => groupItem.id === entry.groupId)?.name ??
      "";
    group.textContent = groupName;
    group.classList.toggle("is-hidden", !groupName);
    meta.classList.toggle("has-group", Boolean(groupName));
    date.textContent = formatCardDate(entry.updatedAt);
    card.classList.toggle("is-active", entry.id === state.selectedId);
    card.classList.toggle("is-pinned", Boolean(entry.pinned));
    if (shouldAnimateGroupSwitch) {
      card.classList.add("is-entering");
      card.style.setProperty("--entry-delay", `${Math.min(index * 28, 140)}ms`);
    }
    pinButton.classList.toggle("is-pinned", Boolean(entry.pinned));
    pinIcon.classList.toggle("fa-thumbtack", !entry.pinned);
    pinIcon.classList.toggle("fa-thumbtack-slash", Boolean(entry.pinned));
    pinButton.setAttribute(
      "aria-label",
      entry.pinned ? "Unpin entry" : "Pin entry",
    );
    pinButton.title = entry.pinned ? "Unpin entry" : "Pin entry";

    const openEntry = () => {
      state.selectedId = entry.id;
      renderEditor();
      renderList();
    };

    openButton.addEventListener("click", openEntry);

    card.addEventListener("click", (event) => {
      if (event.target.closest(".entry-item-actions")) return;
      openEntry();
    });

    pinButton.addEventListener("click", () => {
      toggleEntryPin(entry.id);
    });

    exportButton.addEventListener("click", () => {
      exportJsonBlob(
        {
          entries: [entry],
          groups: state.groups.filter(
            (groupItem) => groupItem.id === entry.groupId,
          ),
        },
        `${slugify(entry.title || "entry")}.json`,
      );
      pulseStatus("Entry exported");
    });

    deleteButton.addEventListener("click", () => {
      const label = entry.title.trim() || "Untitled entry";
      openConfirmModal(`Delete "${label}"? This cannot be undone.`, () => {
        deleteEntryById(entry.id);
      });
    });

    entryList.appendChild(fragment);
  });

  lastRenderedGroupId = state.activeGroupId;
}

function closeGroupFlyout() {
  activeGroupFlyoutId = null;
  if (groupFlyout) {
    groupFlyout.classList.add("is-hidden");
  }
  document.querySelectorAll(".group-menu-button").forEach((button) => {
    button.setAttribute("aria-expanded", "false");
  });
}

function openGroupFlyout(group, button) {
  if (!groupFlyout || !groupFlyoutTitle) return;

  activeGroupFlyoutId = group.id;
  groupFlyoutTitle.textContent = group.name;
  button.setAttribute("aria-expanded", "true");
  groupFlyout.classList.remove("is-hidden");

  const buttonRect = button.getBoundingClientRect();
  const flyoutRect = groupFlyout.getBoundingClientRect();
  const left = Math.min(
    window.innerWidth - flyoutRect.width - 16,
    buttonRect.right + 14,
  );
  const top = Math.min(
    window.innerHeight - flyoutRect.height - 16,
    Math.max(16, buttonRect.top - 10),
  );

  groupFlyout.style.left = `${left}px`;
  groupFlyout.style.top = `${top}px`;
}

function renderEditor() {
  const entry = getSelectedEntry();
  if (!entry) return;

  titleInput.value = entry.title;
  editor.innerHTML = normalizeEditorMarkup(entry.body || "");
  clearActiveImageSelection();
  applyEditorPreferences();
  renderEditorMeta();
  saveStatus.textContent = "Saved";
  updateCounts();
}

function openEntryPicker(groupId) {
  pickerGroupId = groupId;
  pickerList.innerHTML = "";

  state.entries.forEach((entry) => {
    const fragment = pickerEntryTemplate.content.cloneNode(true);
    const button = fragment.querySelector(".picker-entry-main");
    const actionButton = fragment.querySelector(".picker-entry-action");
    const title = fragment.querySelector(".picker-entry-title");
    const date = fragment.querySelector(".picker-entry-date");
    const words = fragment.querySelector(".picker-entry-words");
    const isInCurrentGroup = entry.groupId === pickerGroupId;

    title.textContent = entry.title.trim() || "Untitled entry";
    date.textContent = formatCardDate(entry.updatedAt);
    words.textContent = formatWordCount(countWordsFromHtml(entry.body));

    if (isInCurrentGroup && pickerGroupId !== "all") {
      actionButton.classList.remove("is-hidden");
    }

    button.addEventListener("click", () => {
      if (isInCurrentGroup) return;
      entry.groupId = pickerGroupId;
      state.selectedId = entry.id;
      persistState();
      render();
      closeEntryPicker();
      pulseStatus(`Moved to ${getGroupName(pickerGroupId)}`);
    });

    actionButton.addEventListener("click", () => {
      entry.groupId = "";
      if (state.activeGroupId === pickerGroupId) {
        state.activeGroupId = "all";
      }
      persistState();
      render();
      closeEntryPicker();
      pulseStatus("Entry removed from group");
    });

    pickerList.appendChild(fragment);
  });

  entryPicker.classList.remove("is-hidden");
}

function closeEntryPicker() {
  entryPicker.classList.add("is-hidden");
  pickerGroupId = null;
}

function openConfirmModal(message, onAccept) {
  confirmText.textContent = message;
  confirmAction = onAccept;
  confirmModal.classList.remove("is-hidden");
}

function closeConfirmModal() {
  confirmModal.classList.add("is-hidden");
  confirmAction = null;
}

function deleteEntryById(entryId) {
  if (state.entries.length === 1) {
    const onlyEntry = state.entries[0];
    onlyEntry.title = "";
    onlyEntry.body = "";
    onlyEntry.groupId = "";
    onlyEntry.updatedAt = new Date().toISOString();
    persistState();
    render();
    pulseStatus("Entry cleared");
    return;
  }

  state.entries = state.entries.filter((entry) => entry.id !== entryId);
  const visibleEntries = getVisibleEntries().filter(
    (entry) => entry.id !== entryId,
  );
  state.selectedId = visibleEntries[0]?.id || state.entries[0]?.id || null;
  persistState();
  render();
  pulseStatus("Entry deleted");
}

function getVisibleEntries() {
  return sortEntries(
    state.entries.filter((entry) => {
      const matchesGroup =
        state.activeGroupId === "all" || entry.groupId === state.activeGroupId;
      if (!matchesGroup) return false;

      if (!state.searchTerm) return true;
      const haystack = `${entry.title} ${stripHtml(entry.body)}`.toLowerCase();
      return haystack.includes(state.searchTerm);
    }),
  );
}

function getSelectedEntry() {
  return state.entries.find((entry) => entry.id === state.selectedId);
}

function countEntriesForGroup(groupId) {
  if (groupId === "all") return state.entries.length;
  return state.entries.filter((entry) => entry.groupId === groupId).length;
}

function getGroupName(groupId) {
  if (!groupId) return "";
  return state.groups.find((group) => group.id === groupId)?.name || "";
}

function createEntry(overrides = {}) {
  const now = new Date().toISOString();
  return {
    id: createId(),
    title: "",
    body: "",
    groupId: "",
    pinned: false,
    updatedAt: now,
    ...overrides,
  };
}

function getVisibleGroups() {
  const allGroup = state.groups.find((group) => group.id === "all");
  const otherGroups = sortPinnedItems(
    state.groups.filter((group) => group.id !== "all"),
  );
  return allGroup ? [allGroup, ...otherGroups] : otherGroups;
}

function sortPinnedItems(items) {
  return [...items].sort(
    (a, b) => Number(Boolean(b.pinned)) - Number(Boolean(a.pinned)),
  );
}

function sortEntries(items) {
  return [...items].sort((a, b) => {
    const pinnedDelta = Number(Boolean(b.pinned)) - Number(Boolean(a.pinned));
    if (pinnedDelta !== 0) return pinnedDelta;

    if (state.entrySort === "oldest") {
      return new Date(a.updatedAt).getTime() - new Date(b.updatedAt).getTime();
    }

    if (state.entrySort === "title") {
      return (a.title.trim() || "Untitled entry").localeCompare(
        b.title.trim() || "Untitled entry",
        undefined,
        { sensitivity: "base" },
      );
    }

    return new Date(b.updatedAt).getTime() - new Date(a.updatedAt).getTime();
  });
}

function createAndSelectEntry() {
  const entry = createEntry({
    groupId: state.activeGroupId === "all" ? "" : state.activeGroupId,
  });
  state.entries.unshift(entry);
  state.selectedId = entry.id;
  persistState();
  render();
  titleInput.focus();
  pulseStatus("New entry created");
}

function duplicateSelectedEntry() {
  const entry = getSelectedEntry();
  if (!entry) return;

  const duplicate = {
    ...entry,
    id: createId(),
    title: createDuplicateTitle(entry.title),
    pinned: false,
    updatedAt: new Date().toISOString(),
  };

  state.entries.unshift(duplicate);
  state.selectedId = duplicate.id;
  persistState();
  render();
  titleInput.focus();
  titleInput.select();
  pulseStatus("Entry duplicated");
}

function toggleEntryPin(entryId) {
  const entry = state.entries.find((item) => item.id === entryId);
  if (!entry) return;
  entry.pinned = !entry.pinned;
  persistState();
  renderList();
  pulseStatus(entry.pinned ? "Entry pinned" : "Entry unpinned");
}

function toggleGroupPin(groupId) {
  const group = state.groups.find((item) => item.id === groupId);
  if (!group || group.id === "all") return;
  group.pinned = !group.pinned;
  persistState();
  renderGroups();
  pulseStatus(group.pinned ? "Group pinned" : "Group unpinned");
}

function touchEntry(entry, status = "Saving...") {
  entry.updatedAt = new Date().toISOString();
  persistState();
  renderList();
  renderGroups();
  renderEditorMeta();
  updateCounts();
  pulseStatus(status);
}

function renderEditorMeta() {
  const entry = getSelectedEntry();
  if (!entry) return;
  const activeGroup =
    state.activeGroupId === "all"
      ? null
      : state.groups.find((group) => group.id === state.activeGroupId);
  activeGroupLabel.textContent = activeGroup?.name ?? "";
  activeGroupLabel.classList.toggle("is-hidden", !activeGroup);
  entryDate.textContent = formatLongDate(entry.updatedAt);
}

function applyEditorPreferences() {
  const fontKey = state.fontKey || "crimson";
  const fontSize = state.fontSize || "18px";
  const lineHeight = state.lineHeight || "1.82";
  const pageWidth = state.pageWidth || "860px";

  document.documentElement.style.setProperty(
    "--editor-font",
    `"${FONT_MAP[fontKey] || FONT_MAP.crimson}", sans-serif`,
  );
  document.documentElement.style.setProperty("--editor-size", fontSize);
  document.documentElement.style.setProperty(
    "--editor-line-height",
    lineHeight,
  );
  document.documentElement.style.setProperty("--page-width", pageWidth);

  if (fontSelect) fontSelect.value = fontKey;
  if (fontSizeSelect) fontSizeSelect.value = fontSize;
  if (pageWidthSelect) {
    pageWidthSelect.value = PAGE_WIDTH_OPTIONS.includes(pageWidth)
      ? pageWidth
      : "860px";
  }
  // Sync custom page-width dropdown
  if (pageWidthLabel) {
    const activeOpt = document.querySelector(
      `[data-page-width-option="${pageWidth}"]`,
    );
    if (activeOpt) {
      pageWidthLabel.textContent =
        activeOpt.querySelector(".page-width-option-label")?.textContent ||
        pageWidth;
    }
  }
  pageWidthOptions.forEach((btn) => {
    btn.classList.toggle(
      "is-active",
      btn.dataset.pageWidthOption === pageWidth,
    );
  });
}

function applySidebarWidth() {
  const width = normalizeSidebarWidth(state.sidebarWidth);
  document.documentElement.style.setProperty("--sidebar-width", `${width}px`);
  document.documentElement.classList.toggle(
    "sidebar-at-max",
    width >= normalizeSidebarWidth(MAX_SIDEBAR_WIDTH),
  );
  syncEditorStatsDensity();
}

function normalizeSidebarWidth(width) {
  const viewportLimitedMax = Math.max(
    MIN_SIDEBAR_WIDTH,
    window.innerWidth - 10 - MIN_EDITOR_PANEL_WIDTH,
  );
  return Math.min(
    Math.min(MAX_SIDEBAR_WIDTH, viewportLimitedMax),
    Math.max(MIN_SIDEBAR_WIDTH, parseInt(width, 10) || 320),
  );
}

function startSidebarResize(startX) {
  if (window.innerWidth <= 920) return;
  sidebarResizeSession = {
    startX,
    startWidth: normalizeSidebarWidth(state.sidebarWidth),
  };
  document.body.classList.add("is-resizing-sidebar");
}

function updateSidebarWidth(width, shouldPersist = true) {
  state.sidebarWidth = normalizeSidebarWidth(width);
  applySidebarWidth();
  if (shouldPersist) {
    persistState();
  }
}

function stopSidebarResize(shouldPersist = false) {
  if (!sidebarResizeSession) return;
  sidebarResizeSession = null;
  document.body.classList.remove("is-resizing-sidebar");
  if (shouldPersist) {
    persistState();
  }
}

function syncEditorStatsDensity() {
  if (!editorStats) return;
  const width = editorStats.getBoundingClientRect().width;
  let density = "roomy";

  if (width < 1040) density = "compact";
  if (width < 880) density = "tight";
  if (width < 720) density = "minimal";

  editorStats.dataset.density = density;
}

function updateCounts() {
  const text = stripHtml(normalizeEditorMarkup(editor.innerHTML))
    .replace(/\s+/g, " ")
    .trim();
  const lineSource = editor.innerText
    .replace(/\r/g, "")
    .replace(/\n$/, "")
    .trim();
  const lines = lineSource ? lineSource.split("\n").length : 1;
  const words = text ? text.split(" ").length : 0;
  const characters = text.length;
  lineCount.textContent = String(lines);
  wordCount.textContent = String(words);
  characterCount.textContent = String(characters);
}

function cacheSelection() {
  const selection = window.getSelection();
  if (!selection || selection.rangeCount === 0) return;
  const range = selection.getRangeAt(0);
  if (editor.contains(range.commonAncestorContainer)) {
    savedSelection = range.cloneRange();
  }
}

function restoreSelection() {
  if (!savedSelection) return;
  const selection = window.getSelection();
  selection.removeAllRanges();
  selection.addRange(savedSelection);
}

function applyFontToSelection(fontKey) {
  restoreSelection();
  const selection = window.getSelection();
  if (!selection || selection.rangeCount === 0) {
    syncControlsToSelection();
    return;
  }
  const range = selection.getRangeAt(0);
  if (!editor.contains(range.commonAncestorContainer)) {
    syncControlsToSelection();
    return;
  }

  if (range.collapsed) {
    insertStyledCaretMarker(range, { fontFamily: FONT_MAP[fontKey] });
    syncControlsToSelection();
    return;
  }

  wrapRangeWithStyle(range, { fontFamily: FONT_MAP[fontKey] });
  handleFormattedContentChange();
}

function applyFontSizeToSelection(size) {
  restoreSelection();
  const selection = window.getSelection();
  if (!selection || selection.rangeCount === 0) {
    syncControlsToSelection();
    return;
  }
  const range = selection.getRangeAt(0);
  if (!editor.contains(range.commonAncestorContainer)) {
    syncControlsToSelection();
    return;
  }

  if (range.collapsed) {
    insertStyledCaretMarker(range, { fontSize: size });
    syncControlsToSelection();
    return;
  }

  wrapRangeWithStyle(range, { fontSize: size });
  handleFormattedContentChange();
}

function openLinkEditor() {
  restoreSelection();
  const range = getEditorSelectionRange();
  if (!range) {
    pulseStatus("Place the caret or select text first");
    return;
  }

  const existingLink = getSelectedLink(range);
  const selectedText = window.getSelection()?.toString().trim() || "";
  const suggestedUrl =
    existingLink?.getAttribute("href") ||
    (URL_PATTERN.test(selectedText) ? normalizeUrl(selectedText) : "https://");
  const url = window.prompt(
    "Enter a link URL. Leave it blank to remove the current link.",
    suggestedUrl,
  );

  if (url === null) {
    editor.focus();
    return;
  }

  const trimmedUrl = url.trim();
  if (!trimmedUrl) {
    if (existingLink) {
      unwrapNode(existingLink);
      handleFormattedContentChange();
      pulseStatus("Link removed");
    }
    return;
  }

  const normalizedHref = normalizeUrl(trimmedUrl);
  if (existingLink) {
    existingLink.href = normalizedHref;
    existingLink.target = "_blank";
    existingLink.rel = "noopener noreferrer";
    handleFormattedContentChange();
    pulseStatus("Link updated");
    return;
  }

  if (!range.collapsed && selectedText) {
    document.execCommand("createLink", false, normalizedHref);
    const link = getSelectedLink(getEditorSelectionRange());
    if (link) {
      link.target = "_blank";
      link.rel = "noopener noreferrer";
    }
    handleFormattedContentChange();
    pulseStatus("Link added");
    return;
  }

  const label = window.prompt("Link text", trimmedUrl) ?? trimmedUrl;
  insertHtmlAtSelection(
    `<a href="${escapeHtml(normalizedHref)}" target="_blank" rel="noopener noreferrer">${escapeHtml(label.trim() || trimmedUrl)}</a>`,
  );
  handleFormattedContentChange();
  pulseStatus("Link inserted");
}

function toggleChecklist() {
  restoreSelection();
  const range = getEditorSelectionRange();
  if (!range) return;

  const currentChecklist = getClosestChecklist(range.startContainer);
  if (currentChecklist) {
    convertChecklistToParagraphs(currentChecklist);
    editor.focus();
    handleFormattedContentChange();
    pulseStatus("Checklist removed");
    return;
  }

  const blocks = getSelectedBlocks(range);
  const checklist = document.createElement("ul");
  checklist.className = "editor-checklist";

  if (!blocks.length) {
    const item = document.createElement("li");
    item.innerHTML = range.collapsed
      ? "Checklist item"
      : range.toString().trim();
    checklist.appendChild(item);
    range.deleteContents();
    range.insertNode(checklist);
  } else {
    blocks.forEach((block) => {
      const item = document.createElement("li");
      item.innerHTML = block.innerHTML.trim() || "Checklist item";
      checklist.appendChild(item);
    });
    const firstBlock = blocks[0];
    firstBlock.before(checklist);
    blocks.forEach((block) => block.remove());
  }

  placeCaretInside(checklist.querySelector("li"));
  handleFormattedContentChange();
  pulseStatus("Checklist added");
}

function wrapRangeWithStyle(range, styles, options = {}) {
  const contents = range.extractContents();
  const span = document.createElement("span");
  Object.assign(span.style, styles);
  span.appendChild(contents);
  range.insertNode(span);

  const selection = window.getSelection();
  const newRange = document.createRange();
  if (options.collapseToEnd) {
    newRange.setStartAfter(span);
    newRange.collapse(true);
  } else {
    newRange.selectNodeContents(span);
  }
  selection.removeAllRanges();
  selection.addRange(newRange);
  cacheSelection();
}

function insertStyledCaretMarker(range, styles) {
  const marker = document.createElement("span");
  marker.dataset.typingMarker = "true";
  Object.assign(marker.style, styles);
  marker.appendChild(document.createTextNode("\u200B"));
  range.insertNode(marker);

  const selection = window.getSelection();
  const newRange = document.createRange();
  newRange.setStart(marker.firstChild, 1);
  newRange.collapse(true);
  selection.removeAllRanges();
  selection.addRange(newRange);
  cacheSelection();
  editor.focus();
  handleFormattedContentChange();
}

function handleFormattedContentChange() {
  editor.focus();
  const entry = getSelectedEntry();
  if (!entry) return;
  entry.body = normalizeEditorMarkup(editor.innerHTML);
  touchEntry(entry, "Saved");
  syncControlsToSelection();
}

function applyHighlightToSelection(color) {
  restoreSelection();
  const selection = window.getSelection();
  if (!selection || selection.rangeCount === 0) return;
  const range = selection.getRangeAt(0);
  if (!editor.contains(range.commonAncestorContainer)) return;
  if (range.collapsed || !selection.toString().trim()) return;

  const contents = range.extractContents();
  const mark = document.createElement("mark");
  mark.className = "editor-highlight";
  mark.style.backgroundColor = color;
  mark.appendChild(contents);
  range.insertNode(mark);

  const marker = document.createTextNode("\u200B");
  mark.after(marker);

  const newRange = document.createRange();
  newRange.setStart(marker, 0);
  newRange.collapse(true);
  selection.removeAllRanges();
  selection.addRange(newRange);
  cacheSelection();

  editor.focus();
  handleFormattedContentChange();
}

function removeHighlightFromSelection() {
  restoreSelection();
  const selection = window.getSelection();
  if (!selection || selection.rangeCount === 0) return;
  const range = selection.getRangeAt(0);
  if (!editor.contains(range.commonAncestorContainer)) return;

  if (range.collapsed) {
    const currentNode =
      range.startContainer.nodeType === Node.ELEMENT_NODE
        ? range.startContainer
        : range.startContainer.parentElement;
    const activeHighlight =
      currentNode instanceof HTMLElement
        ? currentNode.closest("mark.editor-highlight")
        : null;

    if (!activeHighlight || !editor.contains(activeHighlight)) return;

    const marker = document.createTextNode("\u200B");
    activeHighlight.after(marker);

    const exitRange = document.createRange();
    exitRange.setStart(marker, 1);
    exitRange.collapse(true);
    selection.removeAllRanges();
    selection.addRange(exitRange);
    cacheSelection();
    editor.focus();
    return;
  }

  if (!selection.toString().trim()) return;

  const highlights = [];
  const walker = document.createTreeWalker(editor, NodeFilter.SHOW_ELEMENT, {
    acceptNode(node) {
      if (
        node instanceof HTMLElement &&
        node.matches("mark.editor-highlight") &&
        range.intersectsNode(node)
      ) {
        return NodeFilter.FILTER_ACCEPT;
      }
      return NodeFilter.FILTER_SKIP;
    },
  });

  while (walker.nextNode()) {
    highlights.push(walker.currentNode);
  }

  let node =
    range.commonAncestorContainer.nodeType === Node.ELEMENT_NODE
      ? range.commonAncestorContainer
      : range.commonAncestorContainer.parentElement;
  while (node && node !== editor) {
    if (node instanceof HTMLElement && node.matches("mark.editor-highlight")) {
      highlights.push(node);
    }
    node = node.parentElement;
  }

  [...new Set(highlights)].forEach(unwrapNode);
  editor.focus();
  handleFormattedContentChange();
}

function removeTextColorFromSelection() {
  restoreSelection();
  const selection = window.getSelection();
  if (!selection || selection.rangeCount === 0) return;
  const range = selection.getRangeAt(0);
  if (!editor.contains(range.commonAncestorContainer)) return;
  if (range.collapsed || !selection.toString().trim()) return;

  const coloredNodes = [];
  const walker = document.createTreeWalker(editor, NodeFilter.SHOW_ELEMENT, {
    acceptNode(node) {
      if (
        node instanceof HTMLElement &&
        node.style.color &&
        range.intersectsNode(node)
      ) {
        return NodeFilter.FILTER_ACCEPT;
      }
      return NodeFilter.FILTER_SKIP;
    },
  });

  while (walker.nextNode()) {
    coloredNodes.push(walker.currentNode);
  }

  let node =
    range.commonAncestorContainer.nodeType === Node.ELEMENT_NODE
      ? range.commonAncestorContainer
      : range.commonAncestorContainer.parentElement;
  while (node && node !== editor) {
    if (node instanceof HTMLElement && node.style.color) {
      coloredNodes.push(node);
    }
    node = node.parentElement;
  }

  [...new Set(coloredNodes)].forEach((element) => {
    element.style.removeProperty("color");
    cleanupStyledNode(element);
  });

  editor.focus();
  handleFormattedContentChange();
}

function applyTextColorToSelection(color) {
  restoreSelection();
  const selection = window.getSelection();
  if (!selection || selection.rangeCount === 0) return;
  const range = selection.getRangeAt(0);
  if (!editor.contains(range.commonAncestorContainer)) return;
  if (range.collapsed || !selection.toString().trim()) return;

  wrapRangeWithStyle(range, { color }, { collapseToEnd: true });

  editor.focus();
  handleFormattedContentChange();
}

function unwrapNode(node) {
  const parent = node.parentNode;
  if (!parent) return;
  while (node.firstChild) {
    parent.insertBefore(node.firstChild, node);
  }
  parent.removeChild(node);
}

function hasEditableSelection() {
  const selection = window.getSelection();
  if (!selection || selection.rangeCount === 0 || selection.isCollapsed)
    return false;
  const range = selection.getRangeAt(0);
  if (!editor.contains(range.commonAncestorContainer)) return false;
  return Boolean(selection.toString().trim());
}

function openEditorContextMenu(x, y) {
  const menuWidth = 170;
  const menuHeight = 68;
  const left = Math.min(x, window.innerWidth - menuWidth - 12);
  const top = Math.min(y, window.innerHeight - menuHeight - 12);

  editorContextMenu.style.left = `${Math.max(12, left)}px`;
  editorContextMenu.style.top = `${Math.max(12, top)}px`;
  editorContextMenu.classList.remove("is-hidden");
}

function hideEditorContextMenu() {
  editorContextMenu.classList.add("is-hidden");
}

function syncControlsToSelection() {
  const selection = window.getSelection();
  if (!selection || selection.rangeCount === 0) {
    resetControlState();
    return;
  }
  const range = selection.getRangeAt(0);
  if (!editor.contains(range.commonAncestorContainer)) {
    resetControlState();
    return;
  }

  cacheSelection();

  let node =
    range.startContainer.nodeType === Node.TEXT_NODE
      ? range.startContainer.parentElement
      : range.startContainer;

  let fontFamily = "";
  let fontSize = "";

  while (node && node !== editor) {
    const style = window.getComputedStyle(node);
    if (!fontFamily && style.fontFamily) {
      fontFamily = style.fontFamily;
    }
    if (!fontSize && style.fontSize) {
      fontSize = style.fontSize;
    }
    node = node.parentElement;
  }

  const resolvedFontKey = resolveFontKey(fontFamily);
  if (resolvedFontKey) {
    fontSelect.value = resolvedFontKey;
  }

  const normalizedSize = normalizeFontSize(fontSize);
  if (normalizedSize) {
    fontSizeSelect.value = normalizedSize;
  }
}

function resetControlState() {
  fontSelect.value = state.fontKey || "crimson";
  fontSizeSelect.value = state.fontSize || "18px";
  if (pageWidthSelect) pageWidthSelect.value = state.pageWidth || "860px";
  // Sync custom page-width dropdown
  const pw = state.pageWidth || "860px";
  pageWidthOptions.forEach((btn) => {
    btn.classList.toggle("is-active", btn.dataset.pageWidthOption === pw);
  });
  if (pageWidthLabel) {
    const activeOpt = document.querySelector(
      `[data-page-width-option="${pw}"]`,
    );
    if (activeOpt)
      pageWidthLabel.textContent =
        activeOpt.querySelector(".page-width-option-label")?.textContent || pw;
  }
}

function resolveFontKey(fontFamily) {
  const lower = fontFamily.toLowerCase();
  return (
    Object.entries(FONT_MAP).find(([, value]) =>
      lower.includes(value.toLowerCase()),
    )?.[0] || "crimson"
  );
}

function normalizeFontSize(fontSize) {
  const parsed = parseInt(fontSize, 10);
  if (!parsed) return "18px";
  const nearest = FONT_SIZE_OPTIONS.reduce((best, current) => {
    return Math.abs(parseInt(current, 10) - parsed) <
      Math.abs(parseInt(best, 10) - parsed)
      ? current
      : best;
  }, FONT_SIZE_OPTIONS[0]);
  return nearest;
}

function cleanupStyledNode(node) {
  if (!(node instanceof HTMLElement)) return;
  if (node.dataset.typingMarker === "true") return;
  if (node.tagName !== "SPAN") return;
  if (node.getAttribute("style")) return;

  const parent = node.parentNode;
  if (!parent) return;

  while (node.firstChild) {
    parent.insertBefore(node.firstChild, node);
  }
  parent.removeChild(node);
}

function getEditorSelectionRange() {
  const selection = window.getSelection();
  if (!selection || selection.rangeCount === 0) return null;
  const range = selection.getRangeAt(0);
  return editor.contains(range.commonAncestorContainer) ? range : null;
}

function getSelectedLink(range) {
  if (!range) return null;
  let node =
    range.startContainer.nodeType === Node.TEXT_NODE
      ? range.startContainer.parentElement
      : range.startContainer;

  while (node && node !== editor) {
    if (node instanceof HTMLAnchorElement) return node;
    node = node.parentElement;
  }

  return null;
}

function getClosestChecklist(node) {
  const element =
    node instanceof HTMLElement ? node : node.parentElement || null;
  return element?.closest(".editor-checklist") || null;
}

function getSelectedBlocks(range) {
  const blockSelectors = "p,h1,h2,h3,blockquote,li";
  const blocks = Array.from(editor.querySelectorAll(blockSelectors)).filter(
    (block) => range.intersectsNode(block),
  );

  if (blocks.length) return blocks;

  const currentBlock =
    (range.startContainer instanceof HTMLElement
      ? range.startContainer
      : range.startContainer.parentElement
    )?.closest(blockSelectors) || null;

  return currentBlock ? [currentBlock] : [];
}

function convertChecklistToParagraphs(checklist) {
  if (!checklist) return;
  const paragraphs = Array.from(checklist.querySelectorAll("li")).map(
    (item) => {
      const paragraph = document.createElement("p");
      paragraph.innerHTML = item.innerHTML.trim() || "<br>";
      return paragraph;
    },
  );

  checklist.replaceWith(...paragraphs);
}

function placeCaretInside(node) {
  if (!node) return;
  const target =
    node.lastChild && node.lastChild.nodeType === Node.TEXT_NODE
      ? node.lastChild
      : node;
  const range = document.createRange();
  range.selectNodeContents(node);
  range.collapse(false);
  const selection = window.getSelection();
  selection.removeAllRanges();
  selection.addRange(range);
  cacheSelection();
  editor.focus();
}

function insertHtmlAtSelection(html) {
  restoreSelection();
  document.execCommand("insertHTML", false, html);
}

function selectEditorImage(image) {
  if (!image) return;
  clearActiveImageSelection();
  activeEditorImage = image;
  image.classList.add("is-selected");

  const figure = image.closest(".editor-media");
  const width = parseInt(figure?.dataset.imageWidth || "100", 10);
  if (imageTools) imageTools.classList.remove("is-hidden");
  if (imageSizeRange) imageSizeRange.value = String(width || 100);
}

function clearActiveImageSelection() {
  if (activeEditorImage) {
    activeEditorImage.classList.remove("is-selected");
  }
  activeEditorImage = null;
  if (imageTools) imageTools.classList.add("is-hidden");
}

function resizeActiveImage(value) {
  if (!activeEditorImage) return;
  const figure = activeEditorImage.closest(".editor-media");
  if (!figure) return;

  const width = Math.min(100, Math.max(30, parseInt(value, 10) || 100));
  figure.dataset.imageWidth = String(width);
  figure.style.width = `${width}%`;
  handleFormattedContentChange();
}

async function insertFilesAsMedia(files) {
  const images = files.filter((file) => file.type.startsWith("image/"));
  if (!images.length) return;

  const htmlParts = [];
  for (const file of images) {
    const src = await readFileAsDataUrl(file);
    htmlParts.push(createImageHtml(src, file.name));
  }

  insertHtmlAtSelection(htmlParts.join(""));
  handleFormattedContentChange();
  pulseStatus(images.length === 1 ? "Image inserted" : "Images inserted");
}

function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function createImageHtml(src, name = "") {
  const alt = name
    .replace(/\.[^.]+$/, "")
    .replace(/[-_]+/g, " ")
    .trim();
  return `
    <figure class="editor-media" data-image-width="100" style="width: 100%">
      <img src="${escapeHtml(src)}" alt="${escapeHtml(alt)}" loading="lazy" />
    </figure>
    <p><br></p>
  `;
}

function hasImageTransfer(dataTransfer) {
  if (!dataTransfer) return false;
  return Array.from(dataTransfer.items || []).some((item) =>
    item.type.startsWith("image/"),
  );
}

function placeCaretFromPoint(x, y) {
  const range =
    document.caretRangeFromPoint?.(x, y) ||
    (() => {
      const position = document.caretPositionFromPoint?.(x, y);
      if (!position) return null;
      const nextRange = document.createRange();
      nextRange.setStart(position.offsetNode, position.offset);
      nextRange.collapse(true);
      return nextRange;
    })();

  if (!range || !editor.contains(range.startContainer)) return;
  const selection = window.getSelection();
  selection.removeAllRanges();
  selection.addRange(range);
  cacheSelection();
}

function cycleSelectOption(select, deltaY) {
  const options = Array.from(select.options);
  const currentIndex = options.findIndex(
    (option) => option.value === select.value,
  );
  const safeIndex = currentIndex === -1 ? 0 : currentIndex;
  const nextIndex =
    deltaY > 0
      ? Math.max(0, safeIndex - 1)
      : Math.min(options.length - 1, safeIndex + 1);
  select.value = options[nextIndex].value;
}

function exportJsonState() {
  exportJsonBlob(
    {
      entries: state.entries,
      groups: state.groups,
      selectedId: state.selectedId,
      activeGroupId: state.activeGroupId,
      entrySort: state.entrySort,
      theme: state.theme,
      fontKey: state.fontKey,
      fontSize: state.fontSize,
      lineHeight: state.lineHeight,
      pageWidth: state.pageWidth,
      focusMode: state.focusMode,
    },
    "pensera-data.json",
  );
  pulseStatus("Data exported");
}

function exportCurrentEntryAsPdf() {
  const entry = getSelectedEntry();
  if (!entry) return;

  const title = entry.title.trim() || "Untitled entry";
  const content = normalizeEditorMarkup(entry.body || "<p></p>");
  const printableHtml = `<!doctype html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>${escapeHtml(title)}</title>
    <link rel="preconnect" href="https://fonts.googleapis.com" />
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
    <link
      href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@400;500;600&family=Crimson+Text:wght@400;600&family=DM+Sans:wght@400;500;700&family=EB+Garamond:wght@400;500;600&family=Instrument+Serif:ital@0;1&family=Inter:wght@400;500;600&family=Libre+Baskerville:wght@400;700&family=Manrope:wght@400;500;700&family=Space+Grotesk:wght@400;500;700&display=swap"
      rel="stylesheet"
    />
    <style>
      :root {
        color-scheme: light;
        --text: #111111;
        --muted: #6f6f69;
        --line: #ddddda;
        --accent: #2d6cdf;
        --paper: #ffffff;
      }
      * { box-sizing: border-box; }
      html, body { margin: 0; background: #fff; color: var(--text); }
      body {
        font-family: "Inter", sans-serif;
        -webkit-print-color-adjust: exact;
        print-color-adjust: exact;
      }
      .sheet {
        width: min(820px, 100%);
        margin: 0 auto;
        padding: 48px 44px 56px;
      }
      .meta {
        margin: 0 0 16px;
        color: var(--muted);
        font-size: 12px;
        letter-spacing: 0.12em;
        text-transform: uppercase;
      }
      h1 {
        margin: 0 0 28px;
        font-family: "Instrument Serif", serif;
        font-size: 42px;
        line-height: 1.05;
        font-weight: 400;
        letter-spacing: -0.03em;
      }
      .content {
        color: var(--text);
      font-size: ${state.fontSize || "18px"};
        line-height: ${state.lineHeight || "1.82"};
      }
      .content p { margin: 0 0 1em; }
      .content h2,
      .content h3,
      .content h4 {
        margin: 1.4em 0 0.55em;
        line-height: 1.15;
      }
      .content h2 { font-size: 1.8em; }
      .content h3 { font-size: 1.4em; }
      .content h4 { font-size: 1.15em; letter-spacing: 0.02em; text-transform: uppercase; }
      .content .editor-checklist {
        margin: 0 0 1.1em;
        padding: 0;
        list-style: none;
      }
      .content .editor-checklist li {
        position: relative;
        margin: 0 0 0.7em;
        padding-left: 1.8em;
      }
      .content .editor-checklist li::before {
        content: "";
        position: absolute;
        top: 0.35em;
        left: 0;
        width: 0.9em;
        height: 0.9em;
        border: 1px solid var(--line);
        border-radius: 0.25em;
      }
      .content .editor-checklist li.is-checked {
        color: var(--muted);
        text-decoration: line-through;
      }
      .content .editor-checklist li.is-checked::before {
        background: var(--accent);
        border-color: var(--accent);
      }
      .content .editor-checklist li.is-checked::after {
        content: "";
        position: absolute;
        top: 0.54em;
        left: 0.28em;
        width: 0.34em;
        height: 0.18em;
        border-left: 2px solid #fff;
        border-bottom: 2px solid #fff;
        transform: rotate(-45deg);
      }
      .content .editor-media { margin: 1.4em 0; }
      .content .editor-media img {
        display: block;
        width: 100%;
        border-radius: 18px;
      }
      .content a {
        color: var(--accent);
        text-decoration: underline;
        text-underline-offset: 0.16em;
      }
      .content mark.editor-highlight {
        color: inherit;
      }
      @page {
        size: auto;
        margin: 16mm;
      }
      @media print {
        .sheet {
          width: auto;
          margin: 0;
          padding: 0;
        }
      }
    </style>
  </head>
  <body>
    <main class="sheet">
      <p class="meta">${escapeHtml(formatLongDate(entry.updatedAt))}</p>
      <h1>${escapeHtml(title)}</h1>
      <div class="content">${content}</div>
    </main>
  </body>
</html>`;

  const printFrame = document.createElement("iframe");
  printFrame.setAttribute("aria-hidden", "true");
  printFrame.style.position = "fixed";
  printFrame.style.right = "0";
  printFrame.style.bottom = "0";
  printFrame.style.width = "0";
  printFrame.style.height = "0";
  printFrame.style.border = "0";
  document.body.appendChild(printFrame);

  const frameWindow = printFrame.contentWindow;
  if (!frameWindow) {
    printFrame.remove();
    pulseStatus("PDF failed");
    return;
  }

  frameWindow.document.open();
  frameWindow.document.write(printableHtml);
  frameWindow.document.close();

  const cleanup = () => {
    window.setTimeout(() => {
      printFrame.remove();
    }, 300);
  };

  const triggerPrint = () => {
    frameWindow.focus();
    frameWindow.print();
    cleanup();
  };

  if (frameWindow.document.fonts?.ready) {
    frameWindow.document.fonts.ready.then(triggerPrint).catch(triggerPrint);
  } else {
    frameWindow.addEventListener("load", triggerPrint, { once: true });
  }

  pulseStatus("PDF ready");
}

function exportJsonBlob(data, filename) {
  const blob = new Blob([JSON.stringify(data, null, 2)], {
    type: "application/json",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

function escapeHtml(value) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function normalizeImportedState(parsed) {
  const groups = Array.isArray(parsed.groups) ? parsed.groups : [];
  const entries = Array.isArray(parsed.entries) ? parsed.entries : [];
  const normalizedGroups = groups.length
    ? [
        { id: "all", name: "All Entries" },
        ...groups.filter((group) => group.id !== "all"),
      ]
    : [
        { id: "all", name: "All Entries" },
        { id: "notes", name: "Notes" },
        { id: "personal", name: "Personal" },
      ];

  return {
    entries: entries.map((entry) => ({
      id: entry.id || createId(),
      title: entry.title || "",
      body: entry.body || "",
      groupId: entry.groupId || "",
      pinned: Boolean(entry.pinned),
      updatedAt: entry.updatedAt || new Date().toISOString(),
    })),
    groups: normalizedGroups.map((group) => ({
      ...group,
      pinned: group.id === "all" ? false : Boolean(group.pinned),
    })),
    selectedId: entries[0]?.id || null,
    activeGroupId: "all",
    searchTerm: "",
    entrySort: parsed.entrySort || "recent",
    fontKey: parsed.fontKey || "crimson",
    fontSize: parsed.fontSize || "18px",
    lineHeight: parsed.lineHeight || "1.82",
    pageWidth: parsed.pageWidth || "860px",
    sidebarWidth: normalizeSidebarWidth(parsed.sidebarWidth),
    theme: resolveInitialTheme(parsed.theme),
    focusMode: Boolean(parsed.focusMode),
  };
}

function closeGroupForm() {
  groupForm.reset();
  editingGroupId = null;
  groupSubmitButton.textContent = "Create";
  groupForm.classList.add("is-hidden");
}

function openGroupForm(group = null) {
  editingGroupId = group?.id || null;
  groupForm.classList.remove("is-hidden");
  groupInput.value = group?.name || "";
  groupSubmitButton.textContent = editingGroupId ? "Rename" : "Create";
  groupInput.focus();
  groupInput.select();
}

function loadState() {
  try {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (!saved) return { ...DEFAULT_STATE };
    const parsed = JSON.parse(saved);
    return {
      entries: (parsed.entries || []).map((entry) => ({
        ...entry,
        pinned: Boolean(entry.pinned),
      })),
      groups: (parsed.groups || []).map((group) => ({
        ...group,
        pinned: group.id === "all" ? false : Boolean(group.pinned),
      })),
      selectedId: parsed.selectedId || null,
      activeGroupId: parsed.activeGroupId || "all",
      searchTerm: "",
      entrySort: parsed.entrySort || "recent",
      fontKey: parsed.fontKey || "crimson",
      fontSize: parsed.fontSize || "18px",
      lineHeight: parsed.lineHeight || "1.82",
      pageWidth: parsed.pageWidth || "860px",
      sidebarWidth: normalizeSidebarWidth(parsed.sidebarWidth),
      theme: resolveInitialTheme(parsed.theme),
      focusMode: Boolean(parsed.focusMode),
    };
  } catch {
    return { ...DEFAULT_STATE };
  }
}

function persistState() {
  localStorage.setItem(
    STORAGE_KEY,
    JSON.stringify({
      entries: state.entries,
      groups: state.groups,
      selectedId: state.selectedId,
      activeGroupId: state.activeGroupId,
      entrySort: state.entrySort,
      fontKey: state.fontKey,
      fontSize: state.fontSize,
      lineHeight: state.lineHeight,
      pageWidth: state.pageWidth,
      sidebarWidth: state.sidebarWidth,
      theme: state.theme,
      focusMode: state.focusMode,
    }),
  );
}

function resolveInitialTheme(savedTheme) {
  if (savedTheme === "light" || savedTheme === "dark") {
    return savedTheme;
  }
  return window.matchMedia("(prefers-color-scheme: dark)").matches
    ? "dark"
    : "light";
}

function applyTheme(theme, previousTheme = null) {
  const resolvedTheme = theme === "dark" ? "dark" : "light";
  const sourceTheme =
    previousTheme ||
    document.documentElement.dataset.theme ||
    (resolvedTheme === "dark" ? "light" : "dark");
  remapHighlightColors(sourceTheme, resolvedTheme);
  remapTextColors(sourceTheme, resolvedTheme);
  document.documentElement.dataset.theme = resolvedTheme;
  document
    .querySelector(THEME_META_SELECTOR)
    ?.setAttribute("content", THEME_COLORS[resolvedTheme]);
  syncHighlightButtons();
  syncTextColorButtons();
}

function applyFocusMode(enabled) {
  document.documentElement.classList.toggle("focus-mode", Boolean(enabled));
}

function syncHighlightButtons() {
  const palette = HIGHLIGHT_COLORS[state.theme === "dark" ? "dark" : "light"];
  highlightColorButtons.forEach((button, index) => {
    const color = palette[index] || palette[0];
    button.dataset.highlightColor = color;
    button.style.setProperty("--highlight-color", color);
  });
}

function syncTextColorButtons() {
  const palette = TEXT_COLORS[state.theme === "dark" ? "dark" : "light"];
  textColorButtons.forEach((button, index) => {
    const color = palette[index] || palette[0];
    button.dataset.textColor = color;
    button.style.setProperty("--highlight-color", color);
  });
}

function remapHighlightColors(fromTheme, toTheme) {
  if (fromTheme === toTheme) return;
  const fromPalette = HIGHLIGHT_COLORS[fromTheme] || HIGHLIGHT_COLORS.light;
  const toPalette = HIGHLIGHT_COLORS[toTheme] || HIGHLIGHT_COLORS.light;

  editor.querySelectorAll("mark.editor-highlight").forEach((mark) => {
    const current = normalizeColor(mark.style.backgroundColor);
    const matchIndex = fromPalette.findIndex(
      (color) => normalizeColor(color) === current,
    );

    if (matchIndex !== -1) {
      mark.style.backgroundColor = toPalette[matchIndex];
    }
  });
}

function remapTextColors(fromTheme, toTheme) {
  if (fromTheme === toTheme) return;
  const fromPalette = TEXT_COLORS[fromTheme] || TEXT_COLORS.light;
  const toPalette = TEXT_COLORS[toTheme] || TEXT_COLORS.light;

  editor.querySelectorAll("[style*='color']").forEach((element) => {
    if (!(element instanceof HTMLElement)) return;
    if (element.matches("mark.editor-highlight")) return;
    const current = normalizeColor(element.style.color);
    const matchIndex = fromPalette.findIndex(
      (color) => normalizeColor(color) === current,
    );

    if (matchIndex !== -1) {
      element.style.color = toPalette[matchIndex];
      cleanupStyledNode(element);
    }
  });
}

function normalizeColor(color) {
  const probe = document.createElement("div");
  probe.style.color = color;
  document.body.appendChild(probe);
  const normalized = getComputedStyle(probe).color;
  probe.remove();
  return normalized;
}

function createPreview(html) {
  const text = stripHtml(normalizeEditorMarkup(html)).trim();
  return text ? text.slice(0, 90) : "Empty entry";
}

function countWordsFromHtml(html) {
  const text = stripHtml(normalizeEditorMarkup(html))
    .replace(/\s+/g, " ")
    .trim();
  return text ? text.split(" ").length : 0;
}

function normalizeEditorMarkup(html) {
  const cleaned = html
    .replace(/\u200B/g, "")
    .replace(/<span[^>]*data-typing-marker="true"[^>]*><\/span>/g, "");
  const container = document.createElement("div");
  container.innerHTML = linkifyHtml(cleaned);

  container.querySelectorAll("img").forEach((image) => {
    image.removeAttribute("width");
    image.removeAttribute("height");
  });

  container.querySelectorAll("figure").forEach((figure) => {
    if (!figure.classList.contains("editor-media")) {
      figure.classList.add("editor-media");
    }
    if (figure.classList.contains("editor-media")) {
      const width = Math.min(
        100,
        Math.max(30, parseInt(figure.dataset.imageWidth || "100", 10) || 100),
      );
      figure.dataset.imageWidth = String(width);
      figure.style.width = `${width}%`;
    }
  });

  container.querySelectorAll("ul").forEach((list) => {
    if (list.classList.contains("editor-checklist")) {
      Array.from(list.children).forEach((child) => {
        if (child.tagName !== "LI") {
          const item = document.createElement("li");
          item.innerHTML =
            child.innerHTML || child.textContent || "Checklist item";
          child.replaceWith(item);
        }
      });
    }
  });

  return container.innerHTML;
}

function linkifyHtml(html) {
  const container = document.createElement("div");
  container.innerHTML = html;
  const walker = document.createTreeWalker(container, NodeFilter.SHOW_TEXT, {
    acceptNode(node) {
      const parent = node.parentElement;
      if (!parent) return NodeFilter.FILTER_REJECT;
      if (parent.closest("a")) return NodeFilter.FILTER_REJECT;
      if (["SCRIPT", "STYLE"].includes(parent.tagName)) {
        return NodeFilter.FILTER_REJECT;
      }
      return URL_PATTERN.test(node.textContent)
        ? NodeFilter.FILTER_ACCEPT
        : NodeFilter.FILTER_REJECT;
    },
  });

  const textNodes = [];
  while (walker.nextNode()) {
    textNodes.push(walker.currentNode);
  }

  textNodes.forEach((node) => {
    const fragment = createLinkedFragment(node.textContent);
    node.parentNode?.replaceChild(fragment, node);
  });

  return container.innerHTML;
}

function createLinkedFragment(text) {
  const fragment = document.createDocumentFragment();
  let lastIndex = 0;

  for (const match of text.matchAll(URL_GLOBAL_PATTERN)) {
    const url = match[0];
    const offset = match.index ?? 0;
    if (offset > lastIndex) {
      fragment.appendChild(
        document.createTextNode(text.slice(lastIndex, offset)),
      );
    }
    const anchor = document.createElement("a");
    anchor.href = normalizeUrl(url);
    anchor.textContent = url;
    anchor.target = "_blank";
    anchor.rel = "noopener noreferrer";
    fragment.appendChild(anchor);
    lastIndex = offset + url.length;
  }

  if (lastIndex < text.length) {
    fragment.appendChild(document.createTextNode(text.slice(lastIndex)));
  }

  return fragment;
}

function normalizeUrl(url) {
  return url.startsWith("http://") || url.startsWith("https://")
    ? url
    : `https://${url}`;
}

function stripHtml(html) {
  const div = document.createElement("div");
  div.innerHTML = html;
  return div.textContent || div.innerText || "";
}

function formatDate(value) {
  return new Intl.DateTimeFormat("en", {
    month: "short",
    day: "numeric",
  }).format(new Date(value));
}

function formatCardDate(value) {
  return new Intl.DateTimeFormat("en", {
    month: "short",
    day: "numeric",
    year: "numeric",
  }).format(new Date(value));
}

function formatWordCount(count) {
  return `${count} ${count === 1 ? "word" : "words"}`;
}

function formatLongDate(value) {
  return new Intl.DateTimeFormat("en", {
    weekday: "short",
    month: "long",
    day: "numeric",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  }).format(new Date(value));
}

function pulseStatus(message) {
  saveStatus.textContent = message;
  window.clearTimeout(pulseStatus.timer);
  pulseStatus.timer = window.setTimeout(() => {
    saveStatus.textContent = "Saved";
  }, 900);
}

function createId() {
  if (window.crypto?.randomUUID) {
    return window.crypto.randomUUID();
  }
  return `entry-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function slugify(value) {
  return (
    value
      .toLowerCase()
      .trim()
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/^-|-$/g, "") || "export"
  );
}

function createDuplicateTitle(title) {
  const baseTitle = title.trim() || "Untitled entry";
  const copyMatch = baseTitle.match(/^(.*?)(?: Copy(?: (\d+))?)$/);
  if (!copyMatch) return `${baseTitle} Copy`;

  const sourceTitle = copyMatch[1].trim() || "Untitled entry";
  const copyNumber = Number(copyMatch[2] || "1") + 1;
  return `${sourceTitle} Copy ${copyNumber}`;
}

function getEntrySortLabel(sortKey) {
  if (sortKey === "oldest") return "Oldest";
  if (sortKey === "title") return "Title";
  return "Latest";
}

/* THE END OF THE JS FILE */
