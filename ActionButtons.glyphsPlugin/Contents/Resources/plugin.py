# encoding: utf-8

from __future__ import division, print_function, unicode_literals
import objc
import copy
from datetime import datetime
import json
import os
import plistlib
import platform
import re
import ssl
import subprocess
import threading
import time
import traceback
import urllib.request
import urllib.parse
import uuid
import webbrowser

from GlyphsApp import Glyphs, WINDOW_MENU
try:
    from GlyphsApp import FILTER_MENU
except Exception:
    FILTER_MENU = None
try:
    from GlyphsApp import SCRIPTS_MENU
except Exception:
    SCRIPTS_MENU = None

from GlyphsApp.plugins import GeneralPlugin
from AppKit import (
    NSApp, NSMenuItem, NSMenu, NSAlert, NSEvent,
    NSImage, NSButton, NSView,
    NSMutableAttributedString, NSAttributedString,
    NSMutableParagraphStyle, NSTextTab, NSParagraphStyleAttributeName,
    NSFontAttributeName, NSForegroundColorAttributeName,
    NSFont, NSColor,
    NSRightTextAlignment, NSLeftTextAlignment,
    NSLineBreakByTruncatingTail, NSLineBreakByWordWrapping,
    NSOnState, NSOffState,
    NSWindowCloseButton, NSImageOnly, NSImageScaleProportionallyDown,
    NSOpenPanel, NSSavePanel, NSModalResponseOK,
    NSViewMinXMargin,
)
from Foundation import NSURL, NSUserDefaults
from vanilla import Box, Button, CheckBox, EditText, FloatingWindow, Group, List, PopUpButton, ScrollView, SegmentedButton, TextBox, TextEditor


BUTTONS_PREF_KEY = "com.glyphs.ActionButtons.buttons"
VIEWMODE_PREF_KEY = "com.glyphs.ActionButtons.viewMode"
GRID_COLS_PREF_KEY = "com.glyphs.ActionButtons.gridColumns"
BUTTON_SIZE_PREF_KEY = "com.glyphs.ActionButtons.buttonSize"
COMPACT_HEIGHT_PREF_KEY = "com.glyphs.ActionButtons.compactHeight"
SHOW_CATEGORY_PREF_KEY = "com.glyphs.ActionButtons.showCategory"
SHOW_SHORTCUT_IN_BUTTON_PREF_KEY = "com.glyphs.ActionButtons.showShortcutInButton"
SHOW_TYPE_IN_BUTTON_PREF_KEY = "com.glyphs.ActionButtons.showTypeInButton"
RUN_ON_DOUBLE_CLICK_PREF_KEY = "com.glyphs.ActionButtons.runOnDoubleClick"
TELEMETRY_ENABLED_PREF_KEY = "com.glyphs.ActionButtons.telemetryEnabled"
TELEMETRY_PROMPTED_PREF_KEY = "com.glyphs.ActionButtons.telemetryPrompted"
TELEMETRY_INSTALL_ID_PREF_KEY = "com.glyphs.ActionButtons.telemetryInstallId"
RELEASE_VERSION_INFO_KEY = "ActionButtonsReleaseVersion"
RELEASE_NOTES_INFO_KEY = "ActionButtonsReleaseNotes"
TELEMETRY_ENDPOINT_INFO_KEY = "ActionButtonsTelemetryEndpoint"
TELEMETRY_APIKEY_INFO_KEY = "ActionButtonsTelemetryApiKey"
DEBUG_TITLEBAR_COG = False
LIST_MODE = 0
GRID_MODE = 1
GRID_COLS = 3  # legacy constant; runtime uses self.gridColumns
MAX_GRID_BUTTONS = 60

BUTTON_SIZE_HEIGHTS = {"small": 34, "medium": 46, "large": 52}
COMPACT_HEIGHT_ONE_HIDDEN = {"small": 26, "medium": 38, "large": 44}
COMPACT_HEIGHT_BOTH_HIDDEN = {"small": 24, "medium": 26, "large": 34}
BUTTON_SIZE_FONTS = {
    "small":  {"category": 9,  "name": 11, "shortcut": 9},
    "medium": {"category": 11, "name": 13, "shortcut": 11},
    "large":  {"category": 12, "name": 15, "shortcut": 12},
}
COLUMN_WINDOW_WIDTHS = {1: 224, 2: 400, 3: 520}
DEBUG_SHORTCUT_SYNC = False
DEBUG_BUTTON_COPY = False
INSTRUCTIONS_PSEUDO_SCRIPT_TARGET = "__actionbuttons__/instructions.py"
INSTRUCTIONS_PSEUDO_SCRIPT_TITLE = "Instructions"
COMPACT_BUTTON_DIALOG_HEIGHT = 250
FULL_BUTTON_DIALOG_HEIGHT = 490
IMPORT_EXPORT_FILE_EXTENSION = "actionbuttons"
IMPORT_EXPORT_FORMAT_NAME = "com.glyphs.ActionButtons.export"
IMPORT_EXPORT_FORMAT_VERSION = 1
IMPORT_EXPORT_MAX_ITEMS = 250
IMPORT_EXPORT_MAX_FILE_BYTES = 1024 * 1024  # 1 MiB

# Modifier key bit masks (stable across macOS versions)
_MOD_CONTROL  = 1 << 18   # NSControlKeyMask
_MOD_OPTION   = 1 << 19   # NSAlternateKeyMask
_MOD_SHIFT    = 1 << 17   # NSShiftKeyMask
_MOD_COMMAND  = 1 << 20   # NSCommandKeyMask
_MASK_KEY_DOWN = 1 << 10  # NSKeyDownMask / NSEventMaskKeyDown
_MOD_ANY = _MOD_CONTROL | _MOD_OPTION | _MOD_SHIFT | _MOD_COMMAND

_FUNCTION_KEY_NAMES = {
    122: "F1",
    120: "F2",
    99: "F3",
    118: "F4",
    96: "F5",
    97: "F6",
    98: "F7",
    100: "F8",
    101: "F9",
    109: "F10",
    103: "F11",
    111: "F12",
    105: "F13",
    107: "F14",
    113: "F15",
    106: "F16",
    64: "F17",
    79: "F18",
    80: "F19",
    90: "F20",
}

_FUNCTION_KEY_CHARS = {
    0xF704: "F1",
    0xF705: "F2",
    0xF706: "F3",
    0xF707: "F4",
    0xF708: "F5",
    0xF709: "F6",
    0xF70A: "F7",
    0xF70B: "F8",
    0xF70C: "F9",
    0xF70D: "F10",
    0xF70E: "F11",
    0xF70F: "F12",
    0xF710: "F13",
    0xF711: "F14",
    0xF712: "F15",
    0xF713: "F16",
    0xF714: "F17",
    0xF715: "F18",
    0xF716: "F19",
    0xF717: "F20",
}


class ActionButtons(GeneralPlugin):

    @objc.python_method
    def _remove_event_monitors(self):
        for attr_name in ("_eventMonitor", "_recordingMonitor"):
            monitor = getattr(self, attr_name, None)
            if monitor is None:
                continue
            try:
                NSEvent.removeMonitor_(monitor)
            except Exception:
                pass
            setattr(self, attr_name, None)

    def dealloc(self):
        self._remove_event_monitors()
        try:
            super(ActionButtons, self).dealloc()
        except Exception:
            pass

    def __del__(self):
        try:
            self._remove_event_monitors()
        except Exception:
            pass

    @objc.python_method
    def settings(self):
        self.name = Glyphs.localize({
            "en": "Action Buttons",
        })
        self.releaseVersion = self._bundle_release_version()

    @objc.python_method
    def start(self):
        # Defensive cleanup in case plugin lifecycle starts twice without full teardown.
        self._remove_event_monitors()
        Glyphs.registerDefault(BUTTONS_PREF_KEY, "[]")
        Glyphs.registerDefault(VIEWMODE_PREF_KEY, GRID_MODE)
        Glyphs.registerDefault(GRID_COLS_PREF_KEY, 1)
        Glyphs.registerDefault(BUTTON_SIZE_PREF_KEY, "medium")
        Glyphs.registerDefault(COMPACT_HEIGHT_PREF_KEY, False)
        Glyphs.registerDefault(SHOW_CATEGORY_PREF_KEY, True)
        Glyphs.registerDefault(SHOW_SHORTCUT_IN_BUTTON_PREF_KEY, True)
        Glyphs.registerDefault(SHOW_TYPE_IN_BUTTON_PREF_KEY, True)
        Glyphs.registerDefault(RUN_ON_DOUBLE_CLICK_PREF_KEY, False)
        Glyphs.registerDefault(TELEMETRY_ENABLED_PREF_KEY, False)
        Glyphs.registerDefault(TELEMETRY_PROMPTED_PREF_KEY, False)
        Glyphs.registerDefault(TELEMETRY_INSTALL_ID_PREF_KEY, "")
        self.items = self._load_items()
        self.viewMode = GRID_MODE
        self.gridColumns = int(Glyphs.defaults[GRID_COLS_PREF_KEY] or 1)
        self.buttonSize = str(Glyphs.defaults[BUTTON_SIZE_PREF_KEY] or "medium")
        self.compactHeightEnabled = bool(Glyphs.defaults[COMPACT_HEIGHT_PREF_KEY])
        self.showCategory = bool(Glyphs.defaults[SHOW_CATEGORY_PREF_KEY])
        self.showShortcutInButton = bool(Glyphs.defaults[SHOW_SHORTCUT_IN_BUTTON_PREF_KEY])
        self.showTypeInButton = bool(Glyphs.defaults[SHOW_TYPE_IN_BUTTON_PREF_KEY])
        self.runOnDoubleClick = bool(Glyphs.defaults[RUN_ON_DOUBLE_CLICK_PREF_KEY])
        self.selectionModeEnabled = False
        self.selectedItemIDs = set()
        self._suppressSelectionCheckboxCallback = False
        self.window = None
        self.addDialog = None
        self.editDialog = None
        self.actionStepDialog = None
        self.aboutDialog = None
        self.instructionsDialog = None
        self.exportDialog = None
        self.cachedFilters = []
        self.cachedScripts = []
        self.cachedScriptTitles = {}
        self.cachedScriptAbsolutePaths = {}
        self._lastScriptMenuPathTokens = None
        self.addDialogSelectedScript = None  # Track selected script from hierarchical menu
        self.addDialogActionSteps = []
        self.editDialogActionSteps = []
        self.editDialogSelectedScript = None
        self.actionStepSelectedScript = None
        self.actionStepDialogOwner = "add"
        self.lastStatusDetailsTitle = None
        self.lastStatusDetailsMessage = None
        self.shortcutCaptureOwner = None
        self._lastScriptSyncDiagnosticKey = None
        self._lastShortcutSyncDebugMessage = None
        self._seenShortcutSyncDebugKeys = set()
        self._telemetryQueue = []
        self._telemetryQueueLock = threading.Lock()
        self._telemetryFlushInProgress = False
        self._telemetrySessionId = str(uuid.uuid4())
        self._telemetryInstallId = self._telemetry_install_id()
        self._eventMonitor = NSEvent.addLocalMonitorForEventsMatchingMask_handler_(
            _MASK_KEY_DOWN,
            lambda event: self._handle_key_event(event),
        )
        self._recordingMonitor = NSEvent.addLocalMonitorForEventsMatchingMask_handler_(
            _MASK_KEY_DOWN,
            lambda event: self._record_shortcut_key_event(event),
        )

        menu_item = NSMenuItem.alloc().initWithTitle_action_keyEquivalent_(
            self.name,
            "showWindow:",
            "",
        )
        menu_item.setTarget_(self)
        Glyphs.menu[WINDOW_MENU].append(menu_item)
        self._telemetry_track("plugin_start", {
            "buttonCount": len(self.items),
            "viewMode": int(self.viewMode),
            "gridColumns": int(self.gridColumns),
        })

    def showWindow_(self, sender):
        self._ensure_items_for_window_open()
        if not self._is_main_window_valid():
            self.window = None
            self._build_main_window()

        self._refresh_available_targets()
        self._refresh_ui()
        self.window.open()
        self.window.makeKey()
        self._maybe_prompt_telemetry_consent()

    @objc.python_method
    def _ensure_items_for_window_open(self):
        # If the saved list is empty, restore the starter Instructions button on next open.
        if self.items:
            return
        self.items = self._seed_default_items()
        self._save_items()

    @objc.python_method
    def _is_main_window_valid(self):
        if self.window is None:
            return False
        try:
            return hasattr(self.window, "gridGroup")
        except Exception:
            return False

    @objc.python_method
    def _main_window_closed(self, sender):
        self._telemetry_flush_if_needed()
        self._remove_titlebar_buttons()
        self.window = None

    @objc.python_method
    def _main_window_resized(self, sender):
        try:
            self._layout_grid_buttons()
            self._layout_status_bar()
        except Exception:
            pass

    @objc.python_method
    def _build_main_window(self):
        init_w = COLUMN_WINDOW_WIDTHS.get(self.gridColumns, 280)
        self.window = FloatingWindow((init_w, 500), self.name, minSize=(220, 300))
        self.window.bind("close", self._main_window_closed)
        self.window.bind("resize", self._main_window_resized)
        try:
            self.window._window.setBecomesKeyOnlyIfNeeded_(False)
        except Exception:
            pass

        # ---- Grid area ----
        self.window.gridGroup = Group((0, 0, -0, -0))
        self.window.gridScroll = ScrollView(
                        (0, 0, -0, -64),
            self.window.gridGroup.getNSView(),
            hasHorizontalScroller=False,
            hasVerticalScroller=True,
            autohidesScrollers=True,
        )
        self.gridButtons = []
        self.gridSelectionChecks = []
        for i in range(MAX_GRID_BUTTONS):
            btn = Button((0, 0, 80, 52), "", callback=self._grid_button_clicked)
            try:
                # Rounded bezel styles tend to render at a fixed visual height.
                # Use a square style so large/medium/small frame heights are visible.
                btn.getNSButton().setBezelStyle_(2)
            except Exception:
                pass
            btn.itemIndex = i
            btn.show(False)
            setattr(self.window.gridGroup, "gridButton_%d" % i, btn)
            self.gridButtons.append(btn)

            chk = CheckBox((0, 0, 18, 20), "", callback=self._grid_selection_checkbox_toggled)
            chk.itemIndex = i
            chk.show(False)
            setattr(self.window.gridGroup, "gridCheck_%d" % i, chk)
            self.gridSelectionChecks.append(chk)

        # ---- Separator line between action row and status bar ----
        self.window.divider = Group((10, -28, -10, 1))
        divider_view = self.window.divider.getNSView()
        divider_view.setWantsLayer_(True)
        divider_layer = divider_view.layer()
        if divider_layer is not None:
            try:
                divider_color = NSColor.separatorColor().colorWithAlphaComponent_(0.45)
            except Exception:
                divider_color = NSColor.gridColor().colorWithAlphaComponent_(0.35)
            divider_layer.setBackgroundColor_(divider_color.CGColor())

        # ---- Action row ----
        # Left side: Select button (shown in normal mode), or icon buttons (shown in select mode)
        self.window.selectToggleButton = Button(
            (10, -60, 60, 26), "Select", callback=self._toggle_selection_mode
        )
        self.window.editIconButton = Button(
            (10, -61, 20, 32), "", callback=self._edit_selected_grid_item
        )
        self.window.duplicateIconButton = Button(
            (46, -61, 20, 32), "", callback=self._duplicate_selected_grid_item
        )
        self.window.deleteIconButton = Button(
            (82, -61, 20, 32), "", callback=self._delete_selected_grid_items
        )
        self.window.upIconButton = Button(
            (118, -61, 20, 32), "", callback=self._move_selected_grid_items_up
        )
        self.window.downIconButton = Button(
            (154, -61, 20, 32), "", callback=self._move_selected_grid_items_down
        )
        # Right side: + add button (always visible)
        self.window.addIconButton = Button(
            (-36, -61, 28, 28), "", callback=self._open_add_dialog
        )

        self._configure_symbol_button(self.window.editIconButton, "square.and.pencil", "Edit")
        self._configure_symbol_button(self.window.duplicateIconButton, "doc.on.doc", "Duplicate")
        self._configure_symbol_button(self.window.deleteIconButton, "trash", "Delete")
        self._configure_symbol_button(self.window.upIconButton, "arrow.up", "Up")
        self._configure_symbol_button(self.window.downIconButton, "arrow.down", "Down")
        self._configure_symbol_button(self.window.addIconButton, "plus", "+")
        self._set_control_tooltip(self.window.editIconButton.getNSButton(), "Edit selected button")
        self._set_control_tooltip(self.window.duplicateIconButton.getNSButton(), "Duplicate selected action button")
        self._set_control_tooltip(self.window.deleteIconButton.getNSButton(), "Delete selected buttons")
        self._set_control_tooltip(self.window.upIconButton.getNSButton(), "Move selected buttons up")
        self._set_control_tooltip(self.window.downIconButton.getNSButton(), "Move selected buttons down")
        self._set_control_tooltip(self.window.addIconButton.getNSButton(), "Add button")

        # ---- Status bar ----
        self.window.help = TextBox(
            (10, -22, -140, 16),
            self._window_help_text(),
            sizeStyle="small",
        )
        self.window.detailsButton = Button(
            (-120, -26, 110, 20), "View Details...", callback=self._show_last_status_details,
            sizeStyle="small",
        )
        self.window.detailsButton.show(False)
        self._layout_status_bar()

        # Hide select-mode icon buttons initially
        self.window.editIconButton.show(False)
        self.window.duplicateIconButton.show(False)
        self.window.deleteIconButton.show(False)
        self.window.upIconButton.show(False)
        self.window.downIconButton.show(False)

        # Inject utility and settings buttons into native title bar.
        self._install_titlebar_buttons()

        # Keep a hidden list view for compatibility with remaining list-dependent internals
        # until full retirement; it is never shown to the user.
        self.window.listView = List(
            (0, 0, 0, 0),
            [],
            columnDescriptions=[
                {"title": "", "key": "icon", "width": 28},
                {"title": "Name", "key": "name", "width": 160},
                {"title": "Type", "key": "type", "width": 60},
                {"title": "Shortcut", "key": "shortcut", "width": 90},
                {"title": "Action", "key": "target"},
            ],
            allowsEmptySelection=True,
            drawFocusRing=False,
            selectionCallback=self._list_selection_changed,
        )
        self.window.listView.show(False)

    @objc.python_method
    def _window_help_text(self):
        release_version = getattr(self, "releaseVersion", None) or self._bundle_release_version()
        if release_version:
            return "Action Buttons Beta %s" % release_version
        return ""

    # ------------------------------------------------------------------
    # Title-bar buttons
    # ------------------------------------------------------------------

    @objc.python_method
    def _set_control_tooltip(self, control, tooltip):
        if control is None or not tooltip:
            return
        try:
            control.setToolTip_(tooltip)
        except Exception:
            pass

    @objc.python_method
    def _install_titlebar_buttons(self):
        """Attach compact utility and settings buttons directly to the native title-bar row."""
        try:
            ns_win = self.window._window
            close_button = ns_win.standardWindowButton_(NSWindowCloseButton)
            if close_button is None:
                return
            titlebar_view = close_button.superview()
            if titlebar_view is None:
                return

            titlebar_frame = titlebar_view.frame()
            close_frame = close_button.frame()

            # Fit the control inside the actual titlebar height (often very small in Glyphs windows).
            row_height = max(12.0, float(titlebar_frame.size.height))
            button_side = max(14.0, min(18.0, row_height - 1.0))
            close_center_y = close_frame.origin.y + (close_frame.size.height / 2.0)
            button_y = close_center_y - (button_side / 2.0)
            max_y = max(0.0, row_height - button_side)
            button_y = max(0.0, min(button_y, max_y))

            self._remove_titlebar_buttons()

            right_margin = 6.0
            gap = 8.0
            current_x = max(0, titlebar_frame.size.width - button_side - right_margin)

            utility_btn = self._make_titlebar_symbol_button(
                button_side,
                "ellipsis.circle",
                "Menu",
                "utilityButtonClicked:",
                "Utility menu",
            )
            utility_btn.setFrameOrigin_((current_x, button_y))
            titlebar_view.addSubview_(utility_btn)
            self._utilityMenuButton = utility_btn

            current_x = max(0, current_x - button_side - gap)
            settings_btn = self._make_titlebar_symbol_button(
                button_side,
                "gearshape",
                "Settings",
                "cogButtonClicked:",
                "Settings",
            )
            settings_btn.setFrameOrigin_((current_x, button_y))
            titlebar_view.addSubview_(settings_btn)
            self._settingsCogButton = settings_btn

            if DEBUG_TITLEBAR_COG:
                self._apply_titlebar_debug_visuals(titlebar_view, close_button, settings_btn)
                print(
                    "ActionButtons titlebar debug:",
                    "titlebar=", titlebar_frame,
                    "close=", close_frame,
                    "gear=", settings_btn.frame(),
                    "flipped=", bool(getattr(titlebar_view, "isFlipped", lambda: False)()),
                )
        except Exception:
            # Fallback: no title-bar buttons; features remain reachable elsewhere.
            pass

    @objc.python_method
    def _remove_titlebar_buttons(self):
        for attr_name in ("_settingsCogButton", "_utilityMenuButton"):
            button = getattr(self, attr_name, None)
            if button is None:
                continue
            try:
                button.removeFromSuperview()
            except Exception:
                pass
            setattr(self, attr_name, None)

    @objc.python_method
    def _make_titlebar_symbol_button(self, button_side, symbol_name, fallback_title, action_name, tooltip):
        btn = NSButton.alloc().initWithFrame_(((0, 0), (button_side, button_side)))
        btn.setBezelStyle_(0)
        btn.setBordered_(False)
        btn.setTarget_(self)
        btn.setAction_(action_name)
        btn.setImagePosition_(NSImageOnly)
        btn.setImageScaling_(NSImageScaleProportionallyDown)
        btn.setAutoresizingMask_(NSViewMinXMargin)

        try:
            img = NSImage.imageWithSystemSymbolName_accessibilityDescription_(
                symbol_name, fallback_title
            )
            if img:
                btn.setImage_(img)
            else:
                btn.setTitle_(fallback_title)
        except Exception:
            btn.setTitle_(fallback_title)

        self._set_control_tooltip(btn, tooltip)
        return btn

    @objc.python_method
    def _apply_titlebar_debug_visuals(self, titlebar_view, close_button, cog_button):
        try:
            self._apply_debug_frame(titlebar_view, NSColor.systemBlueColor(), 0.03)
            self._apply_debug_frame(close_button, NSColor.systemRedColor(), 0.10)
            self._apply_debug_frame(cog_button, NSColor.systemGreenColor(), 0.10)
        except Exception:
            pass

    @objc.python_method
    def _apply_debug_frame(self, view, color, fill_alpha):
        if view is None:
            return
        view.setWantsLayer_(True)
        layer = view.layer()
        if layer is None:
            return
        layer.setBorderWidth_(1.0)
        layer.setCornerRadius_(4.0)
        layer.setBorderColor_(color.CGColor())
        layer.setBackgroundColor_(color.colorWithAlphaComponent_(fill_alpha).CGColor())

    @objc.python_method
    def _configure_symbol_button(self, vanilla_button, symbol_name, fallback_title, symbol_size=26.0):
        ns_button = vanilla_button.getNSButton()
        ns_button.setBezelStyle_(0)
        ns_button.setBordered_(False)
        ns_button.setImagePosition_(NSImageOnly)
        ns_button.setImageScaling_(NSImageScaleProportionallyDown)
        try:
            ns_button.setContentTintColor_(NSColor.labelColor())
        except Exception:
            pass

        try:
            img = NSImage.imageWithSystemSymbolName_accessibilityDescription_(
                symbol_name, fallback_title
            )
            if img:
                try:
                    img.setSize_((symbol_size, symbol_size))
                except Exception:
                    pass
                ns_button.setImage_(img)
                ns_button.setTitle_("")
                return
        except Exception:
            pass

        ns_button.setTitle_(fallback_title)

    def cogButtonClicked_(self, sender):
        self._open_settings_menu(sender)

    def utilityButtonClicked_(self, sender):
        self._open_utility_menu(sender)

    # ------------------------------------------------------------------
    # Settings flyout menu
    # ------------------------------------------------------------------

    @objc.python_method
    def _make_section_header(self, title):
        item = NSMenuItem.alloc().initWithTitle_action_keyEquivalent_("", None, "")
        item.setEnabled_(False)
        try:
            attrs = {
                NSFontAttributeName: NSFont.systemFontOfSize_(10),
                NSForegroundColorAttributeName: NSColor.secondaryLabelColor(),
            }
            attr_str = NSAttributedString.alloc().initWithString_attributes_(
                title.upper(), attrs
            )
            item.setAttributedTitle_(attr_str)
        except Exception:
            item.setTitle_(title.upper())
        return item

    @objc.python_method
    def _build_settings_menu(self):
        menu = NSMenu.alloc().init()
        menu.setAutoenablesItems_(False)

        def add(title, callback, tag=0, checked=False, enabled=True):
            mi = NSMenuItem.alloc().initWithTitle_action_keyEquivalent_(title, None, "")
            mi.setTarget_(self)
            mi.setEnabled_(bool(enabled))
            mi.setTag_(tag)
            mi.setState_(NSOnState if checked else NSOffState)
            # Store callback as represented object (plain Python callable safe to hold)
            mi.setRepresentedObject_(callback)
            mi.setAction_("settingsMenuItemClicked:")
            menu.addItem_(mi)

        # LAYOUT
        menu.addItem_(self._make_section_header("Layout"))
        add("1 Column",  self._settings_set_columns_1, tag=1, checked=(self.gridColumns == 1))
        add("2 Columns", self._settings_set_columns_2, tag=2, checked=(self.gridColumns == 2))
        add("3 Columns", self._settings_set_columns_3, tag=3, checked=(self.gridColumns == 3))
        menu.addItem_(NSMenuItem.separatorItem())

        # BUTTON SIZE
        menu.addItem_(self._make_section_header("Button Size"))
        add("Small",  self._settings_size_small,  checked=(self.buttonSize == "small"))
        add("Medium", self._settings_size_medium, checked=(self.buttonSize == "medium"))
        add("Large",  self._settings_size_large,  checked=(self.buttonSize == "large"))
        add(
            "Compact Height",
            self._settings_toggle_compact_height,
            checked=self.compactHeightEnabled,
            enabled=self._is_compact_height_selectable(),
        )
        menu.addItem_(NSMenuItem.separatorItem())

        # BUTTON CONTENT
        menu.addItem_(self._make_section_header("Button Content"))
        add("Show Category", self._settings_toggle_category, checked=self.showCategory)
        add("Show Shortcut", self._settings_toggle_shortcut, checked=self.showShortcutInButton)
        add("Show Type (S/F/A)", self._settings_toggle_type, checked=self.showTypeInButton)
        menu.addItem_(NSMenuItem.separatorItem())

        # RUN ON
        menu.addItem_(self._make_section_header("Run On"))
        add("Single Click", self._settings_run_single, checked=(not self.runOnDoubleClick))
        add("Double Click", self._settings_run_double, checked=self.runOnDoubleClick)
        menu.addItem_(NSMenuItem.separatorItem())

        # PRIVACY
        menu.addItem_(self._make_section_header("Privacy"))
        add("Share Anonymous Usage Data", self._settings_toggle_telemetry, checked=self._telemetry_enabled())
        menu.addItem_(NSMenuItem.separatorItem())

        return menu

    def settingsMenuItemClicked_(self, sender):
        callback = sender.representedObject()
        if callable(callback):
            callback(sender)

    @objc.python_method
    def _open_settings_menu(self, sender):
        menu = self._build_settings_menu()
        self._popup_menu_from_button(menu, sender, getattr(self, "_settingsCogButton", None))

    @objc.python_method
    def _build_utility_menu(self):
        menu = NSMenu.alloc().init()
        menu.setAutoenablesItems_(False)

        def add(title, callback):
            mi = NSMenuItem.alloc().initWithTitle_action_keyEquivalent_(title, None, "")
            mi.setTarget_(self)
            mi.setEnabled_(True)
            mi.setRepresentedObject_(callback)
            mi.setAction_("settingsMenuItemClicked:")
            menu.addItem_(mi)

        add("Export\u2026", self._open_export_dialog)
        add("Import\u2026", self._open_import_dialog)
        menu.addItem_(NSMenuItem.separatorItem())
        add("Instructions\u2026", self._open_instructions_dialog)
        add("About\u2026", self._open_about_dialog)
        menu.addItem_(NSMenuItem.separatorItem())
        add("Buy Me a Coffee", self._open_buy_me_a_coffee)
        return menu

    @objc.python_method
    def _open_utility_menu(self, sender):
        menu = self._build_utility_menu()
        self._popup_menu_from_button(menu, sender, getattr(self, "_utilityMenuButton", None))

    @objc.python_method
    def _popup_menu_from_button(self, menu, sender, fallback_button):
        try:
            ns_btn = sender or fallback_button
            if ns_btn is not None:
                menu.popUpMenuPositioningItem_atLocation_inView_(
                    None, (0, 26), ns_btn
                )
            else:
                event = NSApp.currentEvent()
                menu.popUpMenuPositioningItem_atLocation_inView_(None, event.locationInWindow(), None)
        except Exception:
            pass

    # Settings callbacks
    @objc.python_method
    def _settings_set_columns_1(self, sender): self._apply_column_setting(1)
    @objc.python_method
    def _settings_set_columns_2(self, sender): self._apply_column_setting(2)
    @objc.python_method
    def _settings_set_columns_3(self, sender): self._apply_column_setting(3)

    @objc.python_method
    def _apply_column_setting(self, cols):
        self.gridColumns = cols
        Glyphs.defaults[GRID_COLS_PREF_KEY] = cols
        self._adjust_window_for_columns()
        self._layout_grid_buttons()
        self._telemetry_track("setting_changed", {"setting": "gridColumns", "value": cols})

    @objc.python_method
    def _adjust_window_for_columns(self):
        new_w = COLUMN_WINDOW_WIDTHS.get(self.gridColumns, 280)
        try:
            x, y, _, h = self.window.getPosSize()
            self.window.setPosSize((x, y, new_w, h))
        except Exception:
            pass

    @objc.python_method
    def _settings_size_small(self, sender):
        self.buttonSize = "small"
        Glyphs.defaults[BUTTON_SIZE_PREF_KEY] = "small"
        self._layout_grid_buttons()
        self._telemetry_track("setting_changed", {"setting": "buttonSize", "value": "small"})

    @objc.python_method
    def _settings_size_medium(self, sender):
        self.buttonSize = "medium"
        Glyphs.defaults[BUTTON_SIZE_PREF_KEY] = "medium"
        self._layout_grid_buttons()
        self._telemetry_track("setting_changed", {"setting": "buttonSize", "value": "medium"})

    @objc.python_method
    def _settings_size_large(self, sender):
        self.buttonSize = "large"
        Glyphs.defaults[BUTTON_SIZE_PREF_KEY] = "large"
        self._layout_grid_buttons()
        self._telemetry_track("setting_changed", {"setting": "buttonSize", "value": "large"})

    @objc.python_method
    def _is_compact_height_selectable(self):
        return not (self.showCategory and self.showShortcutInButton)

    @objc.python_method
    def _settings_toggle_compact_height(self, sender):
        if not self._is_compact_height_selectable():
            return
        self.compactHeightEnabled = not self.compactHeightEnabled
        Glyphs.defaults[COMPACT_HEIGHT_PREF_KEY] = self.compactHeightEnabled
        self._layout_grid_buttons()
        self._telemetry_track("setting_changed", {"setting": "compactHeight", "value": self.compactHeightEnabled})

    @objc.python_method
    def _settings_toggle_category(self, sender):
        self.showCategory = not self.showCategory
        Glyphs.defaults[SHOW_CATEGORY_PREF_KEY] = self.showCategory
        self._layout_grid_buttons()
        self._telemetry_track("setting_changed", {"setting": "showCategory", "value": self.showCategory})

    @objc.python_method
    def _settings_toggle_shortcut(self, sender):
        self.showShortcutInButton = not self.showShortcutInButton
        Glyphs.defaults[SHOW_SHORTCUT_IN_BUTTON_PREF_KEY] = self.showShortcutInButton
        self._layout_grid_buttons()
        self._telemetry_track("setting_changed", {"setting": "showShortcutInButton", "value": self.showShortcutInButton})

    @objc.python_method
    def _settings_toggle_type(self, sender):
        self.showTypeInButton = not self.showTypeInButton
        Glyphs.defaults[SHOW_TYPE_IN_BUTTON_PREF_KEY] = self.showTypeInButton
        self._layout_grid_buttons()
        self._telemetry_track("setting_changed", {"setting": "showTypeInButton", "value": self.showTypeInButton})

    @objc.python_method
    def _settings_run_single(self, sender):
        self.runOnDoubleClick = False
        Glyphs.defaults[RUN_ON_DOUBLE_CLICK_PREF_KEY] = False
        self._telemetry_track("setting_changed", {"setting": "runOnDoubleClick", "value": False})

    @objc.python_method
    def _settings_run_double(self, sender):
        self.runOnDoubleClick = True
        Glyphs.defaults[RUN_ON_DOUBLE_CLICK_PREF_KEY] = True
        self._telemetry_track("setting_changed", {"setting": "runOnDoubleClick", "value": True})

    @objc.python_method
    def _settings_toggle_telemetry(self, sender):
        new_value = not self._telemetry_enabled()
        self._set_telemetry_enabled(new_value, source="settings")

    @objc.python_method
    def _bundle_release_version(self):
        info = self._bundle_info_dictionary()
        if info.get(RELEASE_VERSION_INFO_KEY):
            return info.get(RELEASE_VERSION_INFO_KEY)
        short_version = info.get("CFBundleShortVersionString", "")
        build_version = info.get("CFBundleVersion", "")
        if short_version and build_version:
            return "%s (%s)" % (short_version, build_version)
        return short_version or ""

    @objc.python_method
    def _bundle_info_dictionary(self):
        info_path = os.path.join(os.path.dirname(__file__), "..", "Info.plist")
        try:
            with open(info_path, "rb") as handle:
                info = plistlib.load(handle)
            if isinstance(info, dict):
                return info
        except Exception:
            pass
        return {}

    @objc.python_method
    def _telemetry_enabled(self):
        try:
            return bool(Glyphs.defaults[TELEMETRY_ENABLED_PREF_KEY])
        except Exception:
            return False

    @objc.python_method
    def _telemetry_install_id(self):
        try:
            install_id = str(Glyphs.defaults[TELEMETRY_INSTALL_ID_PREF_KEY] or "").strip()
        except Exception:
            install_id = ""
        if install_id:
            return install_id
        install_id = str(uuid.uuid4())
        Glyphs.defaults[TELEMETRY_INSTALL_ID_PREF_KEY] = install_id
        return install_id

    @objc.python_method
    def _telemetry_backend_config(self):
        info = self._bundle_info_dictionary()
        endpoint = str(info.get(TELEMETRY_ENDPOINT_INFO_KEY, "") or "").strip()
        endpoint = self._validated_telemetry_endpoint(endpoint)
        api_key = str(info.get(TELEMETRY_APIKEY_INFO_KEY, "") or "").strip()
        return endpoint, api_key

    @objc.python_method
    def _validated_telemetry_endpoint(self, endpoint):
        endpoint_text = str(endpoint or "").strip()
        if not endpoint_text:
            return ""
        try:
            parsed = urllib.parse.urlparse(endpoint_text)
        except Exception:
            return ""
        if parsed.scheme.lower() != "https":
            return ""
        if not parsed.netloc:
            return ""
        return endpoint_text

    @objc.python_method
    def _maybe_prompt_telemetry_consent(self):
        try:
            already_prompted = bool(Glyphs.defaults[TELEMETRY_PROMPTED_PREF_KEY])
        except Exception:
            already_prompted = False
        if already_prompted:
            return

        alert = NSAlert.alloc().init()
        alert.setMessageText_("Help Improve Action Buttons?")
        alert.setInformativeText_(
            "Share anonymous usage data to help improve stability and usability.\n\n"
            "Collected data includes button run success/failure, run duration, import/export counts, and settings changes. "
            "No script code, file paths, or glyph data are sent."
        )
        alert.addButtonWithTitle_("Share Anonymous Data")
        alert.addButtonWithTitle_("Not Now")
        self._set_alert_preferred_width(alert, 520)
        result = alert.runModal()
        self._set_telemetry_enabled(result == 1000, source="first_prompt")

    @objc.python_method
    def _set_telemetry_enabled(self, enabled, source="unknown"):
        enabled = bool(enabled)
        previous = self._telemetry_enabled()
        Glyphs.defaults[TELEMETRY_ENABLED_PREF_KEY] = enabled
        Glyphs.defaults[TELEMETRY_PROMPTED_PREF_KEY] = True

        if not enabled:
            with self._telemetryQueueLock:
                self._telemetryQueue = []
            if self.window is not None:
                self._set_status_message("Anonymous telemetry disabled.")
            return

        if self.window is not None:
            self._set_status_message("Anonymous telemetry enabled.")

        # Log this only when newly enabled to avoid duplicate events.
        if not previous:
            self._telemetry_track("telemetry_enabled", {"source": source})
            self._telemetry_flush_if_needed()

    @objc.python_method
    def _telemetry_context(self):
        try:
            glyphs_version = getattr(Glyphs, "versionString", None)
            if callable(glyphs_version):
                glyphs_version = glyphs_version()
            glyphs_version = str(glyphs_version or "")
        except Exception:
            glyphs_version = ""
        return {
            "installId": self._telemetryInstallId,
            "sessionId": self._telemetrySessionId,
            "pluginVersion": str(getattr(self, "releaseVersion", None) or self._bundle_release_version() or ""),
            "glyphsVersion": glyphs_version,
            "platform": "macOS",
            "platformVersion": str(platform.mac_ver()[0] or ""),
        }

    @objc.python_method
    def _telemetry_track(self, event_name, payload=None):
        if not self._telemetry_enabled():
            return

        payload = payload or {}
        event = {
            "eventName": str(event_name),
            "timestamp": int(time.time() * 1000),
            "context": self._telemetry_context(),
            "payload": payload,
        }
        with self._telemetryQueueLock:
            self._telemetryQueue.append(event)
            queue_size = len(self._telemetryQueue)
        if queue_size >= 5:
            self._telemetry_flush_async()

    @objc.python_method
    def _telemetry_flush_async(self):
        if not self._telemetry_enabled():
            return

        endpoint, api_key = self._telemetry_backend_config()
        if not endpoint:
            return

        with self._telemetryQueueLock:
            if self._telemetryFlushInProgress or not self._telemetryQueue:
                return
            self._telemetryFlushInProgress = True

        def worker():
            try:
                self._telemetry_flush_worker(endpoint, api_key)
            finally:
                with self._telemetryQueueLock:
                    self._telemetryFlushInProgress = False

        thread = threading.Thread(target=worker)
        thread.daemon = True
        thread.start()

    @objc.python_method
    def _telemetry_flush_worker(self, endpoint, api_key):
        with self._telemetryQueueLock:
            pending = list(self._telemetryQueue)
            self._telemetryQueue = []

        if not pending:
            return

        body = json.dumps({"events": pending}, default=str).encode("utf-8")
        req = urllib.request.Request(endpoint, data=body, method="POST")
        req.add_header("Content-Type", "application/json")
        if api_key:
            req.add_header("apikey", api_key)
            if self._telemetry_key_is_jwt(api_key):
                req.add_header("Authorization", "Bearer %s" % api_key)

        try:
            urlopen_kwargs = {"timeout": 4.0}
            ssl_context = self._telemetry_ssl_context()
            if ssl_context is not None:
                urlopen_kwargs["context"] = ssl_context

            with urllib.request.urlopen(req, **urlopen_kwargs) as response:
                _ = response.read()
        except Exception as error:
            error_text = str(error)
            if "CERTIFICATE_VERIFY_FAILED" in error_text:
                if self._telemetry_post_via_curl(endpoint, api_key, body):
                    return
            print("ActionButtons telemetry flush failed")
            traceback.print_exc()
            if "CERTIFICATE_VERIFY_FAILED" in error_text:
                print(
                    "ActionButtons telemetry TLS verification failed. "
                    "Attempted fallback to macOS curl trust store but it also failed."
                )
            # Requeue on transient failures and cap growth.
            with self._telemetryQueueLock:
                self._telemetryQueue = pending + self._telemetryQueue
                if len(self._telemetryQueue) > 200:
                    self._telemetryQueue = self._telemetryQueue[-200:]

    @objc.python_method
    def _telemetry_post_via_curl(self, endpoint, api_key, body):
        """Fallback sender using macOS curl trust store when Python SSL fails."""
        try:
            command = [
                "/usr/bin/curl",
                "-sS",
                "-o",
                "/dev/null",
                "-w",
                "%{http_code}",
                "-X",
                "POST",
                endpoint,
                "-H",
                "Content-Type: application/json",
                "--data-binary",
                "@-",
                "--max-time",
                "4",
            ]
            if api_key:
                command.extend(["-H", "apikey: %s" % api_key])
                if self._telemetry_key_is_jwt(api_key):
                    command.extend(["-H", "Authorization: Bearer %s" % api_key])

            result = subprocess.run(
                command,
                input=body,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                timeout=6,
            )
            if result.returncode != 0:
                try:
                    err_text = (result.stderr or b"").decode("utf-8", "ignore").strip()
                except Exception:
                    err_text = ""
                print("ActionButtons telemetry curl fallback failed: returncode=%s stderr=%s" % (result.returncode, err_text))
                return False

            http_code = (result.stdout or b"").decode("utf-8", "ignore").strip()
            if not http_code:
                print("ActionButtons telemetry curl fallback failed: empty HTTP status")
                return False
            if not http_code.startswith("2"):
                try:
                    err_text = (result.stderr or b"").decode("utf-8", "ignore").strip()
                except Exception:
                    err_text = ""
                print("ActionButtons telemetry curl fallback failed: http_status=%s stderr=%s" % (http_code, err_text))
                return False
            return True
        except Exception:
            return False

    @objc.python_method
    def _telemetry_key_is_jwt(self, value):
        key = str(value or "").strip()
        # Legacy anon/service_role keys are JWT-like. New publishable keys are opaque.
        return key.count(".") == 2

    @objc.python_method
    def _telemetry_ssl_context(self):
        """Return an SSL context that prefers certifi's CA bundle when available."""
        try:
            import certifi

            return ssl.create_default_context(cafile=certifi.where())
        except Exception:
            try:
                return ssl.create_default_context()
            except Exception:
                return None

    @objc.python_method
    def _telemetry_flush_if_needed(self):
        with self._telemetryQueueLock:
            has_pending = bool(self._telemetryQueue)
        if has_pending:
            self._telemetry_flush_async()

    @objc.python_method
    def _load_items(self):
        raw_value = Glyphs.defaults[BUTTONS_PREF_KEY]
        if not raw_value:
            return self._seed_default_items()

        try:
            parsed = json.loads(raw_value)
            if isinstance(parsed, list):
                normalized = []
                for index, item in enumerate(parsed):
                    normalized.append(self._normalize_item(item, index))
                sorted_items = self._sorted_items(normalized)
                if sorted_items:
                    return sorted_items
                return self._seed_default_items()
        except Exception:
            print("ActionButtons: could not parse saved buttons")
            traceback.print_exc()
        return self._seed_default_items()

    @objc.python_method
    def _seed_default_items(self):
        """Return a starter set with one Instructions example button."""
        return [self._normalize_item({
            "name": "Instructions",
            "type": "script",
            "target": INSTRUCTIONS_PSEUDO_SCRIPT_TARGET,
            "category": "Click Me",
            "shortcut": "\u2303\u2325\u2318I",  # ⌃⌥⌘I
            "_isInstructions": True,
        }, 0)]

    @objc.python_method
    def _save_items(self):
        for index, item in enumerate(self.items):
            item["orderIndex"] = index

        try:
            Glyphs.defaults[BUTTONS_PREF_KEY] = json.dumps(self.items)
        except Exception:
            print("ActionButtons: could not save button configuration")
            traceback.print_exc()

    @objc.python_method
    def _save_view_mode(self):
        Glyphs.defaults[VIEWMODE_PREF_KEY] = self.viewMode

    @objc.python_method
    def _exportable_item_entry(self, item):
        exported = {
            "name": item.get("name", ""),
            "type": item.get("type", "script"),
            "target": item.get("target", ""),
            "category": item.get("category", ""),
            "shortcut": item.get("shortcut", ""),
        }
        if item.get("type") == "action":
            exported["actions"] = [
                {
                    "actionType": action.get("actionType", "script"),
                    "target": action.get("target", ""),
                }
                for action in item.get("actions", [])
                if action.get("target")
            ]
            exported["continueOnError"] = bool(item.get("continueOnError", False))
        if item.get("_isInstructions"):
            exported["_isInstructions"] = True
        return exported

    @objc.python_method
    def _export_payload_for_items(self, items):
        return {
            "format": IMPORT_EXPORT_FORMAT_NAME,
            "version": IMPORT_EXPORT_FORMAT_VERSION,
            "exportedAt": datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"),
            "pluginVersion": getattr(self, "releaseVersion", None) or self._bundle_release_version(),
            "buttons": [self._exportable_item_entry(item) for item in items],
        }

    @objc.python_method
    def _parse_import_payload(self, payload):
        if isinstance(payload, list):
            return payload, None
        if not isinstance(payload, dict):
            return None, "The selected file does not contain a valid Action Buttons export payload."

        format_name = str(payload.get("format", "") or "").strip()
        if format_name and format_name != IMPORT_EXPORT_FORMAT_NAME:
            return None, "The selected file is not a supported Action Buttons export."

        version = payload.get("version")
        if version not in (None, IMPORT_EXPORT_FORMAT_VERSION, str(IMPORT_EXPORT_FORMAT_VERSION)):
            return None, "This export version is not supported by the current plugin build."

        buttons = payload.get("buttons")
        if buttons is None:
            buttons = payload.get("items")
        if not isinstance(buttons, list):
            return None, "The selected export file does not contain a valid buttons list."
        return buttons, None

    @objc.python_method
    def _validated_normalized_import_items(self, imported_entries):
        normalized_items = []
        for index, entry in enumerate(imported_entries):
            if not isinstance(entry, dict):
                return None, "Button %d is not a valid object." % (index + 1)

            raw_type = entry.get("type", "script")
            if bool(entry.get("_isInstructions", False)):
                raw_type = "script"
            if raw_type not in ("filter", "script", "action"):
                return None, "Button %d has an unsupported type: %s" % (index + 1, raw_type)

            normalized_item = self._normalize_item(entry, index)
            item_type = normalized_item.get("type", "script")
            if item_type in ("filter", "script") and not normalized_item.get("target"):
                return None, "Button %d is missing its target." % (index + 1)
            if item_type == "script":
                script_target = normalized_item.get("target", "")
                if script_target != INSTRUCTIONS_PSEUDO_SCRIPT_TARGET and not self._is_script_target_safe_path(script_target):
                    return None, "Button %d has an unsafe script path. Use a relative path inside your Glyphs Scripts folder." % (index + 1)
            if item_type == "action" and len(normalized_item.get("actions", [])) < 2:
                return None, "Action button %d must contain at least 2 valid steps." % (index + 1)
            if item_type == "action":
                for action in normalized_item.get("actions", []):
                    if action.get("actionType", "script") != "script":
                        continue
                    action_target = action.get("target", "")
                    if action_target != INSTRUCTIONS_PSEUDO_SCRIPT_TARGET and not self._is_script_target_safe_path(action_target):
                        return None, "Action button %d contains an unsafe script path. Use relative paths inside your Glyphs Scripts folder." % (index + 1)
            normalized_items.append(normalized_item)
        return normalized_items, None

    @objc.python_method
    def _normalized_duplicate_script_target(self, target):
        return str(target or "").replace("\\", "/").strip().lower()

    @objc.python_method
    def _duplicate_key_for_item(self, item):
        if not isinstance(item, dict):
            return None

        item_type = item.get("type", "script")
        if item_type == "script":
            target = self._normalized_duplicate_script_target(item.get("target", ""))
            if not target:
                return None
            return ("script", target)

        if item_type == "filter":
            target = str(item.get("target", "")).strip()
            if not target:
                return None
            return ("filter", target)

        if item_type == "action":
            steps = []
            for action in item.get("actions", []):
                action_type = action.get("actionType", "script")
                action_target = action.get("target", "")
                if action_type == "script":
                    normalized_target = self._normalized_duplicate_script_target(action_target)
                else:
                    normalized_target = str(action_target).strip()
                steps.append((action_type, normalized_target))
            if not steps:
                return None
            return ("action", bool(item.get("continueOnError", False)), tuple(steps))

        return None

    @objc.python_method
    def _item_report_name(self, item):
        if not isinstance(item, dict):
            return "Unnamed"
        name = str(item.get("name", "") or "").strip()
        if name:
            return name
        target = str(item.get("target", "") or "").strip()
        if target:
            return target
        return "Unnamed"

    @objc.python_method
    def _duplicate_pairs_display_text(self, duplicate_pairs):
        lines = []
        for pair in duplicate_pairs:
            existing_name = self._item_report_name(pair.get("existingItem"))
            imported_name = self._item_report_name(pair.get("importedItem"))
            lines.append("- %s > %s" % (existing_name, imported_name))
        return "\n".join(lines)

    @objc.python_method
    def _next_unique_import_name(self, name, reserved_names):
        clean_name = str(name or "").strip() or "Unnamed"
        if clean_name not in reserved_names:
            return clean_name

        match = re.match(r"^(.*?)(?:\s+(\d+))?$", clean_name)
        base_name = clean_name
        next_number = 2
        if match:
            base_name = (match.group(1) or "").strip() or clean_name
            if match.group(2):
                next_number = int(match.group(2)) + 1

        candidate = "%s %d" % (base_name, next_number)
        while candidate in reserved_names:
            next_number += 1
            candidate = "%s %d" % (base_name, next_number)
        return candidate

    @objc.python_method
    def _missing_targets_for_item(self, item):
        item_type = item.get("type", "script")
        missing = []

        if item_type == "filter":
            target = str(item.get("target", "") or "").strip()
            if target and target not in self.cachedFilters:
                missing.append("Filter: %s" % target)
            return missing

        if item_type == "script":
            target = item.get("target", "")
            if target != INSTRUCTIONS_PSEUDO_SCRIPT_TARGET and not self._script_absolute_path_for_target(target):
                missing.append("Script: %s" % target)
            return missing

        if item_type == "action":
            for action in item.get("actions", []):
                action_type = action.get("actionType", "script")
                target = action.get("target", "")
                if action_type == "filter":
                    if target and target not in self.cachedFilters:
                        missing.append("Filter: %s" % target)
                else:
                    if target != INSTRUCTIONS_PSEUDO_SCRIPT_TARGET and not self._script_absolute_path_for_target(target):
                        missing.append("Script: %s" % target)
        return missing

    @objc.python_method
    def _merge_import_items(self, imported_items, duplicate_mode="skip"):
        existing_by_key = {}
        for item in self.items:
            key = self._duplicate_key_for_item(item)
            if key is None or key in existing_by_key:
                continue
            existing_by_key[key] = item

        duplicate_pairs = []
        items_to_add = []
        for item in imported_items:
            key = self._duplicate_key_for_item(item)
            existing_item = existing_by_key.get(key)
            if existing_item is not None:
                duplicate_pairs.append({
                    "existingItem": existing_item,
                    "importedItem": item,
                })
                continue
            items_to_add.append(item)

        skipped_duplicates = 0
        if duplicate_pairs:
            if duplicate_mode == "copy":
                items_to_add.extend(pair.get("importedItem") for pair in duplicate_pairs)
            else:
                skipped_duplicates = len(duplicate_pairs)

        reserved_names = set()
        for item in self.items:
            name = str(item.get("name", "") or "").strip()
            if name:
                reserved_names.add(name)

        reserved_shortcuts = set()
        for item in self.items:
            shortcut = self._normalized_shortcut_value(item.get("shortcut", ""))
            if shortcut:
                reserved_shortcuts.add(shortcut)

        merged_items = []
        renamed_items = []
        cleared_shortcuts = []
        missing_targets = []

        for item in items_to_add:
            merged_item = copy.deepcopy(item)
            merged_item["id"] = str(uuid.uuid4())
            merged_item["orderIndex"] = len(self.items) + len(merged_items)

            original_name = str(merged_item.get("name", "") or "").strip() or "Unnamed"
            unique_name = self._next_unique_import_name(original_name, reserved_names)
            if unique_name != original_name:
                renamed_items.append({
                    "from": original_name,
                    "to": unique_name,
                })
            merged_item["name"] = unique_name
            reserved_names.add(unique_name)

            shortcut = self._normalized_shortcut_value(merged_item.get("shortcut", ""))
            if shortcut and shortcut in reserved_shortcuts:
                cleared_shortcuts.append({
                    "name": unique_name,
                    "shortcut": shortcut,
                })
                merged_item["shortcut"] = ""
            elif shortcut:
                reserved_shortcuts.add(shortcut)

            item_missing_targets = self._missing_targets_for_item(merged_item)
            if item_missing_targets:
                missing_targets.append({
                    "name": unique_name,
                    "targets": item_missing_targets,
                })

            merged_items.append(merged_item)

        return {
            "duplicatePairs": duplicate_pairs,
            "itemsToAdd": merged_items,
            "skippedDuplicates": skipped_duplicates,
            "renamedItems": renamed_items,
            "clearedShortcuts": cleared_shortcuts,
            "missingTargets": missing_targets,
        }

    @objc.python_method
    def _prompt_duplicate_import_choice(self, duplicate_pairs):
        if not duplicate_pairs:
            return "skip"

        alert = NSAlert.alloc().init()
        alert.setMessageText_("Duplicate Buttons Detected")
        alert.setInformativeText_(
            "The following buttons were detected as duplicates:\n\n%s\n\n"
            "The following choice applies to all detected duplicates."
            % self._duplicate_pairs_display_text(duplicate_pairs)
        )
        alert.addButtonWithTitle_("Skip")
        alert.addButtonWithTitle_("Import as Copy")
        self._set_alert_preferred_width(alert, 460)
        result = alert.runModal()
        if result == 1001:
            return "copy"
        return "skip"

    @objc.python_method
    def _set_alert_preferred_width(self, alert, min_width):
        if alert is None or not min_width:
            return
        try:
            # NSAlert often ignores direct frame-only resizing; accessory width is more reliable.
            spacer = NSView.alloc().initWithFrame_(((0, 0), (float(min_width), 1.0)))
            alert.setAccessoryView_(spacer)
        except Exception:
            pass
        try:
            window = alert.window()
            if window is None:
                return
            frame = window.frame()
            if frame.size.width >= float(min_width):
                return
            frame.size.width = float(min_width)
            try:
                window.setFrame_display_animate_(frame, True, False)
            except Exception:
                window.setFrame_display_(frame, True)
        except Exception:
            pass

    @objc.python_method
    def _show_error_alert(self, title, message, min_width=0):
        alert = NSAlert.alloc().init()
        alert.setMessageText_(title)
        alert.setInformativeText_(message)
        alert.addButtonWithTitle_("OK")
        self._set_alert_preferred_width(alert, min_width)
        alert.runModal()

    @objc.python_method
    def _import_export_limits_message(self, operation, item_limit_hit=False, size_limit_hit=False, item_count=0, byte_count=0):
        operation_text = str(operation or "transfer").strip().lower()
        action_word = "imported" if operation_text == "import" else "exported"
        lines = [
            "This %s file is too large for Action Buttons to process safely." % operation_text,
            "",
            "To keep the plugin responsive and stable, %s files are limited to:" % operation_text,
            "- %d buttons" % int(IMPORT_EXPORT_MAX_ITEMS),
            "- %d bytes (about %.1f MB)" % (int(IMPORT_EXPORT_MAX_FILE_BYTES), float(IMPORT_EXPORT_MAX_FILE_BYTES) / (1024.0 * 1024.0)),
        ]
        if item_limit_hit:
            lines.append("")
            lines.append("Detected buttons: %d" % int(item_count))
        if size_limit_hit:
            lines.append("")
            lines.append("Detected file size: %d bytes" % int(byte_count))
        lines.append("")
        lines.append("Please split your configuration into smaller files and try again.")
        lines.append("This helps avoid failed runs and keeps your data reliable.")
        return "\n".join(lines)

    @objc.python_method
    def _item_count_limit_error(self, operation, item_count):
        count = int(item_count or 0)
        if count <= int(IMPORT_EXPORT_MAX_ITEMS):
            return None
        return self._import_export_limits_message(
            operation,
            item_limit_hit=True,
            size_limit_hit=False,
            item_count=count,
            byte_count=0,
        )

    @objc.python_method
    def _byte_size_limit_error(self, operation, byte_count):
        size = int(byte_count or 0)
        if size <= int(IMPORT_EXPORT_MAX_FILE_BYTES):
            return None
        return self._import_export_limits_message(
            operation,
            item_limit_hit=False,
            size_limit_hit=True,
            item_count=0,
            byte_count=size,
        )

    @objc.python_method
    def _payload_size_bytes(self, payload):
        try:
            return len(json.dumps(payload, ensure_ascii=False, sort_keys=True).encode("utf-8"))
        except Exception:
            return 0

    @objc.python_method
    def _shortcut_updates_text(self, report):
        cleared_shortcuts = report.get("clearedShortcuts", [])
        if not cleared_shortcuts:
            return ""

        lines = ["Conflicting shortcuts were removed from the following imported buttons:"]
        if cleared_shortcuts:
            for item in cleared_shortcuts:
                lines.append("- %s (%s)" % (item.get("name"), item.get("shortcut")))

        return "\n".join(lines).rstrip()

    @objc.python_method
    def _missing_targets_text(self, report):
        missing_targets = report.get("missingTargets", []) if isinstance(report, dict) else []
        if not missing_targets:
            return ""

        lines = [
            "Some imported buttons reference scripts or filters that are not currently available on this Mac.",
            "The buttons were imported successfully, but they may not run until those targets are installed.",
            "",
        ]
        for item in missing_targets:
            item_name = str(item.get("name", "Unnamed") or "Unnamed")
            lines.append("- %s" % item_name)
            for target in item.get("targets", []):
                lines.append("  %s" % str(target))
        return "\n".join(lines).rstrip()

    @objc.python_method
    def _write_export_payload_to_path(self, path, payload):
        with open(path, "w", encoding="utf-8") as handle:
            json.dump(payload, handle, indent=2, ensure_ascii=False, sort_keys=True)
            handle.write("\n")

    @objc.python_method
    def _sorted_items(self, items):
        return sorted(items, key=lambda item: int(item.get("orderIndex", 0)))

    @objc.python_method
    def _normalize_item(self, item, fallback_order):
        name = item.get("name", "Unnamed")
        item_type = item.get("type", "script")
        target = item.get("target", "")
        is_instructions = bool(item.get("_isInstructions", False))
        if is_instructions:
            # Keep Instructions represented as a script in edit UI/list displays.
            item_type = "script"
            # Force legacy/accidental targets to canonical Instructions pseudo-script.
            target = INSTRUCTIONS_PSEUDO_SCRIPT_TARGET
        item_id = item.get("id") or str(uuid.uuid4())
        order_index = int(item.get("orderIndex", fallback_order))
        actions = item.get("actions", []) if item_type == "action" else []
        continue_on_error = bool(item.get("continueOnError", False)) if item_type == "action" else False

        normalized_actions = []
        if isinstance(actions, list):
            for action in actions:
                if not isinstance(action, dict):
                    continue
                action_type = action.get("actionType", "script")
                action_target = action.get("target", "")
                if action_type not in ("filter", "script"):
                    continue
                if not action_target:
                    continue
                normalized_actions.append({
                    "actionType": action_type,
                    "target": action_target,
                })

        return {
            "id": item_id,
            "name": name,
            "type": item_type,
            "target": target,
            "orderIndex": order_index,
            "actions": normalized_actions,
            "continueOnError": continue_on_error,
            "shortcut": item.get("shortcut", ""),
            "category": item.get("category", ""),
            "_isInstructions": is_instructions,
        }

    @objc.python_method
    def _refresh_available_targets(self):
        self.cachedFilters = self._list_filters()
        self.cachedScripts = self._list_scripts()

    @objc.python_method
    def _list_filters(self):
        names = []

        if FILTER_MENU is not None:
            try:
                filter_menu = Glyphs.menu[FILTER_MENU]
                self._collect_menu_titles(filter_menu, names)
            except Exception:
                pass

        # Fallback: scan all menus if filter menu constant is unavailable.
        if not names:
            try:
                for menu_key in Glyphs.menu:
                    self._collect_menu_titles(Glyphs.menu[menu_key], names)
            except Exception:
                pass

        filtered = []
        seen = set()
        for name in names:
            if not name:
                continue
            title = name.strip()
            if not title or title in seen:
                continue
            if title.startswith("-"):
                continue
            seen.add(title)
            filtered.append(title)

        return sorted(filtered)

    @objc.python_method
    def _collect_menu_titles(self, menu_or_item, names):
        if menu_or_item is None:
            return

        if hasattr(menu_or_item, "numberOfItems"):
            for i in range(menu_or_item.numberOfItems()):
                item = menu_or_item.itemAtIndex_(i)
                self._collect_menu_titles(item, names)
            return

        if isinstance(menu_or_item, (list, tuple)):
            for element in menu_or_item:
                self._collect_menu_titles(element, names)
            return

        if hasattr(menu_or_item, "submenu"):
            title = menu_or_item.title() or ""
            submenu = menu_or_item.submenu()
            if submenu and submenu.numberOfItems() > 0:
                self._collect_menu_titles(submenu, names)
            else:
                names.append(title)

    @objc.python_method
    def _list_scripts(self):
        possible_paths = self._script_root_paths()
        
        scripts_root = None
        for path in possible_paths:
            if os.path.isdir(path):
                scripts_root = path
                break
        
        if scripts_root is None:
            return []

        scripts = []
        script_titles = {}
        script_absolute_paths = {}
        excluded_dir_names = set(["fontTools", "vanilla", "__pycache__"])
        for entry_name in sorted(os.listdir(scripts_root)):
            if entry_name.startswith(".") or entry_name in excluded_dir_names:
                continue

            entry_path = os.path.join(scripts_root, entry_name)

            if os.path.isdir(entry_path):
                self._collect_script_paths_from_directory(
                    scripts,
                    entry_path,
                    entry_name,
                    excluded_dir_names,
                    script_titles,
                    script_absolute_paths,
                )
                continue

            resolved_alias_path = self._resolve_alias_path(entry_path)
            if resolved_alias_path and os.path.isdir(resolved_alias_path):
                self._collect_script_paths_from_directory(
                    scripts,
                    resolved_alias_path,
                    entry_name,
                    excluded_dir_names,
                    script_titles,
                    script_absolute_paths,
                )
                continue

            if entry_name.endswith(".py") and entry_name != "__init__.py":
                scripts.append(entry_name)
                script_absolute_paths[entry_name] = entry_path
                menu_title = self._extract_script_menu_title(entry_path)
                if menu_title:
                    script_titles[entry_name] = menu_title

        self.cachedScriptTitles = script_titles
        self.cachedScriptAbsolutePaths = script_absolute_paths
        print("ActionButtons: Found %d scripts" % len(scripts))
        return sorted(scripts)

    @objc.python_method
    def _collect_script_paths_from_directory(self, scripts, base_dir, display_root, excluded_dir_names, script_titles, script_absolute_paths):
        for root, dirs, files in os.walk(base_dir, followlinks=True):
            dirs[:] = [
                d for d in dirs
                if d not in excluded_dir_names and not d.startswith(".")
            ]

            relative_root = os.path.relpath(root, base_dir)
            if relative_root == ".":
                relative_root = ""

            for filename in files:
                if not filename.endswith(".py") or filename == "__init__.py" or filename.startswith("."):
                    continue

                if relative_root:
                    relative_path = os.path.join(display_root, relative_root, filename)
                else:
                    relative_path = os.path.join(display_root, filename)

                relative_script_path = relative_path.replace(os.sep, "/")
                absolute_script_path = os.path.join(root, filename)
                scripts.append(relative_script_path)
                script_absolute_paths[relative_script_path] = absolute_script_path
                menu_title = self._extract_script_menu_title(absolute_script_path)
                if menu_title:
                    script_titles[relative_script_path] = menu_title

    @objc.python_method
    def _extract_script_menu_title(self, script_path):
        if not script_path or not os.path.isfile(script_path):
            return None

        try:
            with open(script_path, "r", encoding="utf-8", errors="replace") as handle:
                for _ in range(120):
                    line = handle.readline()
                    if not line:
                        break
                    match = re.match(r"^\s*#\s*MenuTitle\s*:\s*(.+?)\s*$", line, re.IGNORECASE)
                    if match:
                        menu_title = match.group(1).strip()
                        if menu_title:
                            return menu_title
        except Exception:
            return None

        return None

    @objc.python_method
    def _script_absolute_path_for_target(self, script_relative_path):
        if not script_relative_path:
            return None
        if not self._is_script_target_safe_path(script_relative_path):
            return None

        cached_paths = getattr(self, "cachedScriptAbsolutePaths", {})
        cached = cached_paths.get(script_relative_path)
        if cached and os.path.isfile(cached):
            return cached

        possible_paths = self._script_root_paths()
        for scripts_root in possible_paths:
            candidate = os.path.join(scripts_root, script_relative_path)
            if os.path.isfile(candidate):
                return candidate
        return None

    @objc.python_method
    def _script_root_paths(self):
        # Try Glyphs 3 first, then fallback to Glyphs
        return [
            os.path.expanduser("~/Library/Application Support/Glyphs 3/Scripts"),
            os.path.expanduser("~/Library/Application Support/Glyphs/Scripts"),
        ]

    @objc.python_method
    def _is_script_target_safe_path(self, script_relative_path):
        target_text = str(script_relative_path or "").replace("\\", "/").strip()
        if not target_text:
            return False
        if os.path.isabs(target_text) or target_text.startswith("~"):
            return False

        parts = [part for part in target_text.split("/")]
        if not parts:
            return False
        for part in parts:
            if part in ("", ".", ".."):
                return False

        for scripts_root in self._script_root_paths():
            root_real = os.path.realpath(scripts_root)
            candidate_real = os.path.realpath(os.path.join(root_real, target_text))
            if candidate_real == root_real or candidate_real.startswith(root_real + os.sep):
                return True
        return False

    @objc.python_method
    def _script_menu_title_for_target(self, script_relative_path):
        if not script_relative_path:
            return None

        cached_titles = getattr(self, "cachedScriptTitles", {})
        cached_title = cached_titles.get(script_relative_path)
        if cached_title:
            return cached_title

        script_path = self._script_absolute_path_for_target(script_relative_path)
        if not script_path:
            return None
        return self._extract_script_menu_title(script_path)

    @objc.python_method
    def _script_display_title_for_target(self, script_relative_path):
        if script_relative_path == INSTRUCTIONS_PSEUDO_SCRIPT_TARGET:
            return INSTRUCTIONS_PSEUDO_SCRIPT_TITLE
        menu_title = self._script_menu_title_for_target(script_relative_path)
        if menu_title:
            return menu_title
        return self._pretty_script_display_token(os.path.basename(script_relative_path))

    @objc.python_method
    def _is_instructions_item(self, item):
        if not isinstance(item, dict):
            return False
        return bool(item.get("_isInstructions", False))

    @objc.python_method
    def _is_editing_instructions_item(self):
        if self.editDialog is None:
            return False
        index = getattr(self.editDialog, "currentIndex", None)
        if index is None or index < 0 or index >= len(self.items):
            return False
        return self._is_instructions_item(self.items[index])

    @objc.python_method
    def _resolve_alias_path(self, path):
        if not os.path.isfile(path):
            return None

        try:
            alias_url = NSURL.fileURLWithPath_(path)
            resolved_url, error = NSURL.URLByResolvingAliasFileAtURL_options_error_(alias_url, 0, None)
            if error is None and resolved_url is not None:
                resolved = resolved_url.path()
                if resolved:
                    return resolved
        except Exception:
            return None

        return None

    @objc.python_method
    def _icon_for_type(self, item_type):
        if item_type == "filter":
            return "F"
        if item_type == "action":
            return "A"
        return "S"

    @objc.python_method
    def _target_display(self, item):
        if item.get("type") == "action":
            action_count = len(item.get("actions", []))
            mode = "continue" if item.get("continueOnError", False) else "stop"
            return "%d steps (%s on error)" % (action_count, mode)
        if item.get("type") == "script":
            target = item.get("target", "")
            if not target:
                return ""
            if target == INSTRUCTIONS_PSEUDO_SCRIPT_TARGET:
                return INSTRUCTIONS_PSEUDO_SCRIPT_TITLE
            parts = str(target).split("/")
            pretty_parts = [self._pretty_script_display_token(part) for part in parts[:-1]]
            pretty_parts.append(self._script_display_title_for_target(target))
            return "/".join(pretty_parts)
        return item.get("target", "")

    @objc.python_method
    def _display_name(self, item):
        name = item.get("name", "")
        if item.get("type") != "script":
            return name

        target = item.get("target", "")
        if not target:
            return name

        raw_basename = os.path.basename(target)
        raw_without_ext = raw_basename[:-3] if raw_basename.endswith(".py") else raw_basename
        pretty_basename = self._pretty_script_display_token(raw_basename)
        metadata_title = self._script_display_title_for_target(target)
        # Keep explicit custom labels untouched. Only pretty-format legacy
        # default labels derived from the raw filename.
        if name in (raw_basename, raw_without_ext, pretty_basename):
            return metadata_title
        return name

    @objc.python_method
    def _refresh_ui(self):
        self.items = self._sorted_items(self.items)

        # Keep the hidden list view in sync (used by list-era internals still in use)
        rows = []
        for item in self.items:
            rows.append({
                "icon": self._icon_for_type(item["type"]),
                "name": self._display_name(item),
                "type": item["type"].capitalize(),
                "shortcut": item.get("shortcut", ""),
                "target": self._target_display(item),
            })
        self.window.listView.set(rows)
        self._update_selection_mode_controls()
        self._update_grid_action_icon_state()
        self._layout_grid_buttons()

    @objc.python_method
    def _layout_grid_buttons(self):
        _, _, window_width, window_height = self.window.getPosSize()
        # Grid scroll viewport fills from top to -56 (action row + status bar)
        grid_height = max(50, window_height - 56)
        viewport_width = window_width
        try:
            ns_scroll = self.window.gridScroll.getNSScrollView()
            if ns_scroll is not None:
                viewport_width = ns_scroll.contentSize().width
        except Exception:
            pass

        side_padding = 6
        inner_width = max(40.0, float(viewport_width - (side_padding * 2)))

        cols = max(1, self.gridColumns)
        h_gap = 12
        v_gap = 12
        cell_height = self._effective_grid_button_height()
        if self.compactHeightEnabled and not (self.showCategory and self.showShortcutInButton):
            v_gap = 8
        slot_width = float(inner_width - (cols - 1) * h_gap) / cols
        check_w = 18 if self.selectionModeEnabled else 0
        check_gap = 4 if self.selectionModeEnabled else 0
        cell_width = max(40, slot_width - check_w - check_gap)

        item_count = len(self.items)
        rows = int((item_count + cols - 1) / cols) if item_count else 1
        content_height = max(
            grid_height,
            int(8 + rows * cell_height + max(0, rows - 1) * v_gap + 8),
        )
        try:
            self.window.gridGroup.setPosSize((0, 0, int(viewport_width), int(content_height)))
        except Exception:
            pass

        for index, button in enumerate(self.gridButtons):
            check = self.gridSelectionChecks[index]
            if index < len(self.items):
                item = self.items[index]
                row = index // cols
                col = index % cols
                slot_x = side_padding + col * (slot_width + h_gap)
                x = slot_x + check_w + check_gap
                y = 8 + row * (cell_height + v_gap)

                button.setPosSize((int(x), int(y), int(cell_width), int(cell_height)))
                self._apply_button_attributed_title(button, item, cell_width)
                self._apply_button_tooltip(button, item)
                button.show(True)

                if self.selectionModeEnabled:
                    check_y = y + max(0, (cell_height - 20) / 2.0)
                    check.setPosSize((int(slot_x), int(check_y), 18, 20))
                    self._suppressSelectionCheckboxCallback = True
                    try:
                        check.set(1 if item.get("id") in self.selectedItemIDs else 0)
                    finally:
                        self._suppressSelectionCheckboxCallback = False
                    check.show(True)
                else:
                    check.show(False)
            else:
                button.show(False)
                check.show(False)

    @objc.python_method
    def _effective_grid_button_height(self):
        size_key = self.buttonSize if self.buttonSize in BUTTON_SIZE_HEIGHTS else "medium"
        base_height = BUTTON_SIZE_HEIGHTS.get(size_key, 46)
        if not self.compactHeightEnabled:
            return base_height
        if self.showCategory and self.showShortcutInButton:
            return base_height
        if (not self.showCategory) and (not self.showShortcutInButton):
            return COMPACT_HEIGHT_BOTH_HIDDEN.get(size_key, 26)
        return COMPACT_HEIGHT_ONE_HIDDEN.get(size_key, 38)

    @objc.python_method
    def _apply_button_attributed_title(self, button, item, cell_width):
        """Render button copy on multiple styled lines.

        Category and shortcut use smaller muted text. Name uses primary text.
        """
        type_label_map = {"filter": "F", "script": "S", "action": "A"}
        type_indicator = type_label_map.get(item.get("type", ""), "") if self.showTypeInButton else ""
        size_key = self.buttonSize if self.buttonSize in BUTTON_SIZE_FONTS else "medium"
        fspec = BUTTON_SIZE_FONTS[size_key]

        show_category_line = bool(self.showCategory)
        show_shortcut_line = bool(self.showShortcutInButton)
        category = (item.get("category", "") if show_category_line else "").upper()
        name_text = self._display_name(item)
        name_line = name_text
        if type_indicator:
            name_line = "%s\t%s" % (name_text, type_indicator)
        shortcut = item.get("shortcut", "") if show_shortcut_line else ""

        # Plain fallback title for environments that ignore attributed content.
        fallback_lines = []
        if show_category_line:
            fallback_lines.append(category or " ")
        fallback_lines.append(name_line)
        if show_shortcut_line:
            fallback_lines.append(shortcut or " ")
        fallback_title = "\n".join(fallback_lines)
        if not fallback_title:
            fallback_title = self._display_name(item)
        button.setTitle(fallback_title)

        try:
            ns_button = button._nsObject
            cell = ns_button.cell()
            if cell is not None:
                cell.setWraps_(True)
                cell.setUsesSingleLineMode_(False)
                cell.setScrollable_(False)
                cell.setLineBreakMode_(NSLineBreakByWordWrapping)
                try:
                    cell.setAlignment_(NSLeftTextAlignment)
                except Exception:
                    pass
                try:
                    cell.setTruncatesLastVisibleLine_(False)
                except Exception:
                    pass

            # Create base paragraph style
            paragraph = NSMutableParagraphStyle.alloc().init()
            paragraph.setLineBreakMode_(NSLineBreakByWordWrapping)
            paragraph.setAlignment_(NSLeftTextAlignment)
            if type_indicator:
                try:
                    right_tab = max(24.0, float(cell_width) - 20.0)
                    paragraph.setTabStops_([
                        NSTextTab.alloc().initWithType_location_(NSRightTextAlignment, right_tab)
                    ])
                except Exception:
                    pass

            # Create name paragraph with scaled spacing (25% of name font size)
            paragraph_name = NSMutableParagraphStyle.alloc().init()
            paragraph_name.setLineBreakMode_(NSLineBreakByWordWrapping)
            paragraph_name.setAlignment_(NSLeftTextAlignment)
            if self.compactHeightEnabled and not (self.showCategory and self.showShortcutInButton):
                spacing = 1.0
            else:
                spacing = max(1.0, fspec["name"] * 0.25)
            try:
                paragraph_name.setParagraphSpacingAfter_(spacing)
            except Exception:
                pass
            if type_indicator:
                try:
                    right_tab = max(24.0, float(cell_width) - 20.0)
                    paragraph_name.setTabStops_([
                        NSTextTab.alloc().initWithType_location_(NSRightTextAlignment, right_tab)
                    ])
                except Exception:
                    pass

            self._debug_button_copy(
                item,
                size_key,
                fspec,
                spacing,
                category,
                name_line,
                shortcut,
                show_category_line,
                show_shortcut_line,
            )

            attrs_name = {
                NSParagraphStyleAttributeName: paragraph_name,
                NSFontAttributeName: NSFont.systemFontOfSize_(fspec["name"]),
                NSForegroundColorAttributeName: NSColor.labelColor(),
            }
            attrs_muted = {
                NSParagraphStyleAttributeName: paragraph,
                NSFontAttributeName: NSFont.systemFontOfSize_(fspec["category"]),
                NSForegroundColorAttributeName: NSColor.secondaryLabelColor(),
            }
            attrs_shortcut = {
                NSParagraphStyleAttributeName: paragraph,
                NSFontAttributeName: NSFont.systemFontOfSize_(fspec["shortcut"]),
                NSForegroundColorAttributeName: NSColor.secondaryLabelColor(),
            }

            attributed = NSMutableAttributedString.alloc().init()
            if show_category_line:
                attributed.appendAttributedString_(
                    NSAttributedString.alloc().initWithString_attributes_((category or " ") + "\n", attrs_muted)
                )

            name_paragraph = name_line + ("\n" if show_shortcut_line else "")
            attributed.appendAttributedString_(
                NSAttributedString.alloc().initWithString_attributes_(name_paragraph, attrs_name)
            )

            if show_shortcut_line:
                attributed.appendAttributedString_(
                    NSAttributedString.alloc().initWithString_attributes_(shortcut or " ", attrs_shortcut)
                )

            ns_button.setAttributedTitle_(attributed)
        except Exception:
            pass

    @objc.python_method
    def _apply_button_tooltip(self, button, item):
        """Set a tooltip showing full type, action target, and shortcut."""
        try:
            type_full = item.get("type", "").capitalize()
            target = self._target_display(item)
            shortcut = item.get("shortcut", "")
            lines = ["Type:    %s" % type_full, "Action:  %s" % target]
            if shortcut:
                lines.append("Shortcut: %s" % shortcut)
            button._nsObject.setToolTip_("\n".join(lines))
        except Exception:
            pass

    @objc.python_method
    def _selected_index(self):
        selection = self.window.listView.getSelection()
        if not selection:
            return None
        return selection[0]

    @objc.python_method
    def _selected_item(self):
        index = self._selected_index()
        if index is None or index < 0 or index >= len(self.items):
            return None
        return self.items[index]

    @objc.python_method
    def _update_action_edit_button_state(self):
        # Legacy stub — state is now managed by _update_grid_action_icon_state.
        self._update_grid_action_icon_state()

    @objc.python_method
    def _layout_status_bar(self):
        if self.window is None:
            return
        details_visible = bool(self.window.detailsButton.isVisible()) if hasattr(self.window, "detailsButton") else False
        if details_visible:
            self.window.help.setPosSize((10, -22, -140, 16))
            self.window.detailsButton.setPosSize((-120, -26, 110, 20))
        else:
            self.window.help.setPosSize((10, -22, -10, 16))

    @objc.python_method
    def _set_status_message(self, message, details_title=None, details_message=None):
        self.window.help.set(message)
        self.lastStatusDetailsTitle = details_title
        self.lastStatusDetailsMessage = details_message

        show_details = bool(details_message)
        self.window.detailsButton.show(show_details)
        self.window.detailsButton.enable(show_details)
        self._layout_status_bar()

    @objc.python_method
    def _show_last_status_details(self, sender):
        if not self.lastStatusDetailsMessage:
            return
        alert = NSAlert.alloc().init()
        alert.setMessageText_(self.lastStatusDetailsTitle or "Details")
        alert.setInformativeText_(self.lastStatusDetailsMessage)
        alert.addButtonWithTitle_("OK")
        alert.runModal()

    @objc.python_method
    def _about_dialog_closed(self, sender):
        self.aboutDialog = None

    @objc.python_method
    def _instructions_dialog_closed(self, sender):
        self.instructionsDialog = None

    @objc.python_method
    def _export_dialog_closed(self, sender):
        self.exportDialog = None

    @objc.python_method
    def _close_about_dialog(self, sender):
        if self.aboutDialog is None:
            return
        try:
            self.aboutDialog.close()
        finally:
            self.aboutDialog = None

    @objc.python_method
    def _close_instructions_dialog(self, sender):
        if self.instructionsDialog is None:
            return
        try:
            self.instructionsDialog.close()
        finally:
            self.instructionsDialog = None

    @objc.python_method
    def _open_author_website(self, sender):
        try:
            webbrowser.open("https://www.neilcarding.co.uk")
        except Exception:
            pass

    @objc.python_method
    def _open_buy_me_a_coffee(self, sender):
        try:
            webbrowser.open("https://buymeacoffee.com/infom72e")
        except Exception:
            pass

    @objc.python_method
    def _export_dialog_rows(self):
        rows = []
        for item in self.items:
            rows.append({
                "name": self._display_name(item),
                "type": item.get("type", "").capitalize(),
                "target": self._target_display(item),
            })
        return rows

    @objc.python_method
    def _update_export_dialog_buttons(self):
        if self.exportDialog is None:
            return
        selection = []
        try:
            selection = self.exportDialog.itemList.getSelection()
        except Exception:
            selection = []
        has_selection = bool(selection)
        self.exportDialog.exportButton.enable(has_selection)
        self.exportDialog.selectAllButton.enable(bool(self.items))
        self.exportDialog.selectNoneButton.enable(bool(self.items))

    @objc.python_method
    def _open_export_dialog(self, sender):
        if not self.items:
            self._show_error_alert("Export Buttons", "There are no buttons to export yet.")
            return

        if self.exportDialog is None:
            self.exportDialog = FloatingWindow((640, 360), "Export Buttons", minSize=(520, 300))
            self.exportDialog.bind("close", self._export_dialog_closed)
            self.exportDialog.instructions = TextBox(
                (12, 12, -12, 20),
                "Select the buttons you want to export.",
            )
            self.exportDialog.itemList = List(
                (12, 40, -12, -54),
                [],
                columnDescriptions=[
                    {"title": "Name", "key": "name", "width": 180},
                    {"title": "Type", "key": "type", "width": 70},
                    {"title": "Action", "key": "target"},
                ],
                allowsEmptySelection=True,
                allowsMultipleSelection=True,
                drawFocusRing=False,
                selectionCallback=self._export_dialog_selection_changed,
            )
            self.exportDialog.selectAllButton = Button((12, -36, 90, 24), "Select All", callback=self._select_all_export_items)
            self.exportDialog.selectNoneButton = Button((108, -36, 96, 24), "Select None", callback=self._select_none_export_items)
            self.exportDialog.cancelButton = Button((-204, -36, 90, 24), "Cancel", callback=self._close_export_dialog)
            self.exportDialog.exportButton = Button((-106, -36, 90, 24), "Export", callback=self._confirm_export_dialog)

        self.exportDialog.itemList.set(self._export_dialog_rows())
        self.exportDialog.itemList.setSelection(list(range(len(self.items))))
        self._update_export_dialog_buttons()

        try:
            self.exportDialog.open()
        except ValueError:
            self.exportDialog = None
            self._open_export_dialog(sender)
            return
        self.exportDialog.makeKey()

    @objc.python_method
    def _close_export_dialog(self, sender):
        if self.exportDialog is None:
            return
        self.exportDialog.close()
        self.exportDialog = None

    @objc.python_method
    def _export_dialog_selection_changed(self, sender):
        self._update_export_dialog_buttons()

    @objc.python_method
    def _select_all_export_items(self, sender):
        if self.exportDialog is None:
            return
        self.exportDialog.itemList.setSelection(list(range(len(self.items))))
        self._update_export_dialog_buttons()

    @objc.python_method
    def _select_none_export_items(self, sender):
        if self.exportDialog is None:
            return
        self.exportDialog.itemList.setSelection([])
        self._update_export_dialog_buttons()

    @objc.python_method
    def _export_panel_default_path(self):
        return "ActionButtons Export.%s" % IMPORT_EXPORT_FILE_EXTENSION

    @objc.python_method
    def _confirm_export_dialog(self, sender):
        if self.exportDialog is None:
            return

        selected_indices = self.exportDialog.itemList.getSelection() or []
        if not selected_indices:
            self._show_error_alert("Export Buttons", "Select at least one button to export.")
            return

        items_to_export = [self.items[index] for index in selected_indices if 0 <= index < len(self.items)]
        if not items_to_export:
            self._show_error_alert("Export Buttons", "The selected buttons could not be prepared for export.")
            return

        count_error = self._item_count_limit_error("export", len(items_to_export))
        if count_error:
            self._show_error_alert("Export Limit Reached", count_error, min_width=560)
            return

        payload = self._export_payload_for_items(items_to_export)
        payload_bytes = self._payload_size_bytes(payload)
        size_error = self._byte_size_limit_error("export", payload_bytes)
        if size_error:
            self._show_error_alert("Export Limit Reached", size_error, min_width=560)
            return

        panel = NSSavePanel.savePanel()
        try:
            panel.setAllowedFileTypes_([IMPORT_EXPORT_FILE_EXTENSION])
        except Exception:
            pass
        try:
            panel.setNameFieldStringValue_(self._export_panel_default_path())
        except Exception:
            pass

        if panel.runModal() != NSModalResponseOK:
            return

        try:
            export_path = panel.URL().path()
        except Exception:
            export_path = None
        if not export_path:
            return
        if not export_path.lower().endswith("." + IMPORT_EXPORT_FILE_EXTENSION):
            export_path += "." + IMPORT_EXPORT_FILE_EXTENSION

        try:
            self._write_export_payload_to_path(export_path, payload)
        except Exception:
            self._show_error_alert("Export Failed", "The selected buttons could not be exported.")
            traceback.print_exc()
            return

        self._set_status_message("Exported %d button%s." % (len(items_to_export), "s" if len(items_to_export) != 1 else ""))
        self._telemetry_track("buttons_exported", {
            "count": len(items_to_export),
        })
        self._telemetry_flush_if_needed()
        self._close_export_dialog(sender)

    @objc.python_method
    def _open_import_dialog(self, sender):
        panel = NSOpenPanel.openPanel()
        try:
            panel.setAllowsMultipleSelection_(False)
        except Exception:
            pass
        try:
            panel.setAllowedFileTypes_([IMPORT_EXPORT_FILE_EXTENSION])
        except Exception:
            pass

        if panel.runModal() != NSModalResponseOK:
            return

        try:
            import_path = panel.URL().path()
        except Exception:
            import_path = None
        if not import_path:
            return

        self._import_buttons_from_path(import_path)

    @objc.python_method
    def _import_buttons_from_path(self, import_path):
        try:
            source_size = int(os.path.getsize(import_path) or 0)
        except Exception:
            source_size = 0
        size_error = self._byte_size_limit_error("import", source_size)
        if size_error:
            self._show_error_alert("Import Limit Reached", size_error, min_width=560)
            return

        try:
            with open(import_path, "r", encoding="utf-8") as handle:
                payload = json.load(handle)
        except Exception:
            self._show_error_alert("Import Failed", "The selected file could not be read as JSON.")
            traceback.print_exc()
            return

        imported_entries, error_message = self._parse_import_payload(payload)
        if error_message:
            self._show_error_alert("Import Failed", error_message)
            return

        normalized_items, error_message = self._validated_normalized_import_items(imported_entries)
        if error_message:
            self._show_error_alert("Import Failed", error_message)
            return
        if not normalized_items:
            self._show_error_alert("Import Failed", "The selected file did not contain any buttons to import.")
            return

        count_error = self._item_count_limit_error("import", len(normalized_items))
        if count_error:
            self._show_error_alert("Import Limit Reached", count_error, min_width=560)
            return

        self._refresh_available_targets()

        preview_report = self._merge_import_items(normalized_items, duplicate_mode="skip")
        duplicate_pairs = preview_report.get("duplicatePairs", [])
        duplicate_mode = "skip"
        if duplicate_pairs:
            duplicate_mode = self._prompt_duplicate_import_choice(duplicate_pairs)

        merge_report = self._merge_import_items(normalized_items, duplicate_mode=duplicate_mode)
        items_to_add = merge_report.get("itemsToAdd", [])
        if not items_to_add:
            if merge_report.get("skippedDuplicates"):
                self._set_status_message("Import skipped: all detected duplicates were skipped.")
            else:
                self._show_error_alert("Import Failed", "No buttons could be imported from the selected file.")
            return

        self.items.extend(items_to_add)
        self._save_items()
        self._refresh_ui()

        self._set_status_message("Imported %d button%s." % (len(items_to_add), "s" if len(items_to_add) != 1 else ""))
        skipped_duplicates_count = int(merge_report.get("skippedDuplicates", 0) or 0)
        self._telemetry_track("buttons_imported", {
            "count": len(items_to_add),
            "duplicateMode": duplicate_mode,
            "duplicatesDetected": len(duplicate_pairs),
            "skippedDuplicates": skipped_duplicates_count,
            "clearedShortcuts": len(merge_report.get("clearedShortcuts", [])),
        })
        self._telemetry_flush_if_needed()
        if merge_report.get("clearedShortcuts"):
            summary_text = self._shortcut_updates_text(merge_report)
            self._show_error_alert("Import Shortcut Updates", summary_text, min_width=460)
        if merge_report.get("missingTargets"):
            summary_text = self._missing_targets_text(merge_report)
            self._show_error_alert("Import Missing Targets", summary_text, min_width=560)

    @objc.python_method
    def _release_notes_text(self):
        info = self._bundle_info_dictionary()
        release_notes = info.get(RELEASE_NOTES_INFO_KEY, "")
        if release_notes:
            release_notes = release_notes.strip()
        if release_notes:
            return release_notes
        return "No release notes available in this build."

    @objc.python_method
    def _instructions_text(self):
        return (
            "Action Buttons lets you create one-click shortcuts for Glyphs filters, scripts, and multi-step actions.\n\n"
            "ADD A BUTTON\n"
            "Tap the + button to add a new Filter, Script, or Action button. "
            "Category and Shortcut are optional. The Add button enables when required inputs are set:\n"
            "- Filter: choose an action\n"
            "- Script: choose a script\n"
            "- Action: add at least 2 steps\n\n"
            "RUN A BUTTON\n"
            "Click any button to run it (or double-click if that option is set in Settings).\n\n"
            "SHORTCUTS\n"
            "Use Record to capture a shortcut, Clear to remove it, and Use Existing to import a shortcut from Glyphs when available.\n\n"
            "ACTION STEPS\n"
            "For Action buttons, use Add Step, Remove Step, Step Up, and Step Down above the steps list. "
            "Remove/Up/Down are enabled only when a step is selected (and Up/Down follow top/bottom limits). "
            "Friendly heads-up: not every script and filter combination works well together, "
            "so please test new combinations on a copy or unimportant file first.\n\n"
            "EDIT OR DELETE\n"
            "Tap Select, check the buttons you want to manage, then use the action-row icons "
            "(Edit, Duplicate, Delete, Move Up, Move Down).\n\n"
            "UTILITY MENU\n"
            "Use the utility icon in the title bar for Export, Import, Instructions, and About.\n\n"
            "SETTINGS\n"
            "Use the settings icon in the title bar to adjust layout, button size, content, and run behavior."
        )

    @objc.python_method
    def _open_instructions_dialog(self, sender):
        if self.instructionsDialog is not None:
            try:
                self.instructionsDialog.open()
                self.instructionsDialog.makeKey()
                return
            except Exception:
                self.instructionsDialog = None

        self.instructionsDialog = FloatingWindow((700, 580), "How to use Action Buttons", minSize=(640, 500))
        self.instructionsDialog.bind("close", self._instructions_dialog_closed)
        body_padding = 20
        self.instructionsDialog.title = TextBox(
            (body_padding, body_padding - 2, -body_padding, 20),
            "Instructions",
        )
        self.instructionsDialog.body = TextEditor(
            (body_padding, body_padding + 24, -body_padding, -56),
            self._instructions_text(),
            readOnly=True,
        )
        try:
            text_view = self.instructionsDialog.body.getNSTextView()
            if text_view is not None:
                text_view.setTextContainerInset_((10.0, 10.0))
        except Exception:
            pass
        self.instructionsDialog.closeButton = Button(
            (-body_padding - 90, -36, 90, 24),
            "Close",
            callback=self._close_instructions_dialog,
        )
        self.instructionsDialog.open()
        self.instructionsDialog.makeKey()

    @objc.python_method
    def _open_about_dialog(self, sender):
        if self.aboutDialog is not None:
            try:
                self.aboutDialog.open()
                self.aboutDialog.makeKey()
                return
            except Exception:
                self.aboutDialog = None

        info = self._bundle_info_dictionary()
        release_version = getattr(self, "releaseVersion", None) or self._bundle_release_version() or "Unknown"
        short_version = info.get("CFBundleShortVersionString", "")
        build_version = info.get("CFBundleVersion", "")
        version_line = "Release: %s" % release_version
        if short_version and build_version:
            version_line = "%s  |  Bundle: %s  Build: %s" % (version_line, short_version, build_version)

        self.aboutDialog = FloatingWindow((560, 450), "About Action Buttons", minSize=(500, 340))
        self.aboutDialog.bind("close", self._about_dialog_closed)
        self.aboutDialog.title = TextBox((12, 12, -12, 20), "Action Buttons")
        self.aboutDialog.version = TextBox((12, 36, -12, 20), version_line)
        self.aboutDialog.author = TextBox((12, 60, -12, 20), "Created by Neil Carding")
        self.aboutDialog.websiteLabel = TextBox((12, 84, 60, 20), "Website")
        self.aboutDialog.websiteButton = Button((74, 82, 220, 24), "www.neilcarding.co.uk", callback=self._open_author_website)
        self.aboutDialog.supportLabel = TextBox((12, 112, 120, 20), "Buy Me a Coffee")
        self.aboutDialog.supportButton = Button((132, 110, 220, 24), "buymeacoffee.com/infom72e", callback=self._open_buy_me_a_coffee)
        self.aboutDialog.disclaimer = TextBox(
            (12, 140, -12, 34),
            "Built for personal use and shared as-is. Please back up your files and test safely; you use this plugin at your own risk.",
        )
        self.aboutDialog.notesLabel = TextBox((12, 178, -12, 20), "Release Notes")
        self.aboutDialog.notes = TextBox((12, 200, -12, -50), self._release_notes_text())
        self.aboutDialog.closeButton = Button((-100, -34, 90, 24), "Close", callback=self._close_about_dialog)
        self.aboutDialog.open()
        self.aboutDialog.makeKey()

    @objc.python_method
    def _format_exception_reason(self):
        exc_type, exc_value, _ = traceback.sys.exc_info()
        if exc_value is not None:
            type_name = getattr(exc_type, "__name__", "Error") if exc_type else "Error"
            value_text = str(exc_value).strip()
            if value_text:
                return "%s: %s" % (type_name, value_text)
            return type_name

        formatted = traceback.format_exc().strip()
        if not formatted:
            return "Unknown error"

        lines = [line.strip() for line in formatted.splitlines() if line.strip()]
        if not lines:
            return "Unknown error"
        return lines[-1]

    @objc.python_method
    def _list_selection_changed(self, sender):
        # Hidden list used only to track selection index for edit dialog; no status message needed.
        self._update_action_edit_button_state()

    @objc.python_method
    def _view_mode_changed(self, sender):
        self.viewMode = sender.get()
        self._save_view_mode()
        self._refresh_ui()
        mode_name = "grid" if self.viewMode == GRID_MODE else "list"
        self._telemetry_track("setting_changed", {"setting": "viewMode", "value": mode_name})

    @objc.python_method
    def _grid_button_clicked(self, sender):
        index = getattr(sender, "itemIndex", None)
        if index is None or index >= len(self.items):
            return

        # In selection mode: toggle checkbox rather than run
        if self.selectionModeEnabled:
            item_id = self.items[index].get("id")
            if item_id in self.selectedItemIDs:
                self.selectedItemIDs.discard(item_id)
            else:
                self.selectedItemIDs.add(item_id)
            self._update_grid_action_icon_state()
            self._layout_grid_buttons()
            return

        # Respect run-on-double-click setting
        if self.runOnDoubleClick:
            try:
                if NSEvent.currentEvent().clickCount() < 2:
                    return
            except Exception:
                pass

        # Check if this is the Instructions seed button
        item = self.items[index]
        if item.get("_isInstructions"):
            self._open_instructions_dialog(None)
            return

        self._run_item(index)

    @objc.python_method
    def _grid_selection_checkbox_toggled(self, sender):
        if self._suppressSelectionCheckboxCallback:
            return
        if not self.selectionModeEnabled:
            return

        index = getattr(sender, "itemIndex", None)
        if index is None or index >= len(self.items):
            return

        item_id = self.items[index].get("id")
        if bool(sender.get()):
            self.selectedItemIDs.add(item_id)
        else:
            self.selectedItemIDs.discard(item_id)

        self._update_grid_action_icon_state()
        self._layout_grid_buttons()

    # ------------------------------------------------------------------
    # Selection mode
    # ------------------------------------------------------------------

    @objc.python_method
    def _toggle_selection_mode(self, sender):
        if not self.selectionModeEnabled and not self.items:
            return
        self.selectionModeEnabled = not self.selectionModeEnabled
        if not self.selectionModeEnabled:
            self.selectedItemIDs.clear()
        self._update_selection_mode_controls()
        self._layout_grid_buttons()

    @objc.python_method
    def _update_selection_mode_controls(self):
        in_sel = self.selectionModeEnabled
        has_items = len(self.items) > 0
        self.window.selectToggleButton.show(True)
        self.window.selectToggleButton.setTitle("Done" if in_sel else "Select")
        self.window.selectToggleButton.enable(in_sel or has_items)

        # Keep the Select/Done toggle visible so selection mode can always be exited.
        if in_sel:
            self.window.selectToggleButton.setPosSize((10, -60, 60, 26))
            self.window.editIconButton.setPosSize((78, -63, 20, 32))
            self.window.duplicateIconButton.setPosSize((103, -63, 20, 32))
            self.window.deleteIconButton.setPosSize((128, -63, 20, 32))
            self.window.upIconButton.setPosSize((150, -63, 20, 32))
            self.window.downIconButton.setPosSize((172, -63, 20, 32))

        self.window.editIconButton.show(in_sel)
        self.window.duplicateIconButton.show(in_sel)
        self.window.deleteIconButton.show(in_sel)
        self.window.upIconButton.show(in_sel)
        self.window.downIconButton.show(in_sel)
        if in_sel:
            self._update_grid_action_icon_state()

    @objc.python_method
    def _update_grid_action_icon_state(self):
        if self.window is None:
            return
        if not self.selectionModeEnabled:
            return
        n = len(self.selectedItemIDs)
        self.window.editIconButton.enable(n == 1)
        # Duplicate only enabled when exactly one Action-type button is selected
        if n == 1:
            dup_idx = self._single_selected_index_from_ids()
            can_dup = (dup_idx is not None and
                       self.items[dup_idx].get("type") == "action")
        else:
            can_dup = False
        self.window.duplicateIconButton.enable(can_dup)
        self.window.deleteIconButton.enable(n >= 1)
        # Move up/down: enabled when selection exists and there is room to move
        indices = self._selected_indices_from_ids()
        can_up = bool(indices and min(indices) > 0)
        can_down = bool(indices and max(indices) < len(self.items) - 1)
        self.window.upIconButton.enable(can_up)
        self.window.downIconButton.enable(can_down)

    @objc.python_method
    def _selected_indices_from_ids(self):
        """Return sorted list of item indices for currently checked items."""
        result = []
        for i, item in enumerate(self.items):
            if item.get("id") in self.selectedItemIDs:
                result.append(i)
        return sorted(result)

    @objc.python_method
    def _single_selected_index_from_ids(self):
        indices = self._selected_indices_from_ids()
        return indices[0] if len(indices) == 1 else None

    # ------------------------------------------------------------------
    # Grid action row callbacks
    # ------------------------------------------------------------------

    @objc.python_method
    def _edit_selected_grid_item(self, sender):
        index = self._single_selected_index_from_ids()
        if index is None:
            return
        # Sync hidden list selection so existing edit dialog reads the right item
        self.window.listView.setSelection([index])
        self._open_edit_dialog(sender)

    @objc.python_method
    def _duplicate_selected_grid_item(self, sender):
        index = self._single_selected_index_from_ids()
        if index is None:
            return
        item = self.items[index]
        if item.get("type") != "action":
            return
        new_item = copy.deepcopy(item)
        new_item["id"] = str(uuid.uuid4())
        new_item["name"] = new_item.get("name", "") + " COPY"
        new_item["shortcut"] = ""
        self.items.insert(index + 1, new_item)
        self._save_items()
        self.selectedItemIDs = {new_item["id"]}
        self._refresh_ui()

    @objc.python_method
    def _delete_selected_grid_items(self, sender):
        indices = self._selected_indices_from_ids()
        if not indices:
            return
        count = len(indices)
        alert = NSAlert.alloc().init()
        alert.setMessageText_("Delete %d button%s?" % (count, "s" if count > 1 else ""))
        alert.setInformativeText_(
            "This cannot be undone. The button%s will be permanently removed." % ("s" if count > 1 else "")
        )
        alert.addButtonWithTitle_("Delete")
        alert.addButtonWithTitle_("Cancel")
        result = alert.runModal()
        if result != 1000:  # 1000 = NSAlertFirstButtonReturn
            return
        # Delete in reverse order to preserve lower indices
        for i in sorted(indices, reverse=True):
            del self.items[i]
        self.selectedItemIDs.clear()
        self._save_items()
        self._refresh_ui()
        self._toggle_selection_mode(None)  # exit selection mode after delete

    @objc.python_method
    def _move_selected_grid_items_up(self, sender):
        self._move_selected_block(direction=-1)

    @objc.python_method
    def _move_selected_grid_items_down(self, sender):
        self._move_selected_block(direction=1)

    @objc.python_method
    def _move_selected_block(self, direction):
        """Move selected items as one block while preserving relative order."""
        indices = self._selected_indices_from_ids()
        if not indices:
            return
        if direction == -1 and min(indices) == 0:
            return
        if direction == 1 and max(indices) == len(self.items) - 1:
            return

        # Build new order
        items = list(self.items)
        if direction == -1:
            for i in indices:
                items[i - 1], items[i] = items[i], items[i - 1]
            new_indices = [i - 1 for i in indices]
        else:
            for i in reversed(indices):
                items[i + 1], items[i] = items[i], items[i + 1]
            new_indices = [i + 1 for i in indices]

        self.items = items
        self.selectedItemIDs = {self.items[i].get("id") for i in new_indices}
        self._save_items()
        self._refresh_ui()
        self._update_grid_action_icon_state()



    @objc.python_method
    def _run_selected(self, sender):
        index = self._selected_index()
        if index is None:
            self._set_status_message("Select a row first.")
            return
        self._run_item(index)

    @objc.python_method
    def _run_item(self, index):
        item = self.items[index]
        started_at = time.time()
        ok = False
        status_message = None
        status_details_title = None
        status_details_message = None

        if item["type"] == "filter":
            ok, error_reason = self._run_filter_with_error(item["target"])
            if not ok and error_reason:
                status_details_title = "Filter Error"
                status_details_message = error_reason
        elif item["type"] == "script":
            ok, error_reason = self._run_script_with_error(item["target"])
            if not ok and error_reason:
                status_details_title = "Script Error"
                status_details_message = error_reason
        elif item["type"] == "action":
            ok, status_message, status_details_title, status_details_message = self._run_action_button(item)
        else:
            status_message = "Unknown button type: %s" % item.get("type", "")

        if status_message:
            self._set_status_message(status_message, status_details_title, status_details_message)
        elif ok:
            self._set_status_message("Ran: %s" % item["name"])
        else:
            self._set_status_message("Could not run: %s" % item["name"])

        self._telemetry_track("button_run", {
            "buttonName": item.get("name", ""),
            "buttonType": item.get("type", ""),
            "success": bool(ok),
            "durationMs": int((time.time() - started_at) * 1000),
        })
        self._telemetry_flush_if_needed()

    @objc.python_method
    def _run_action_button(self, item):
        actions = item.get("actions", [])
        if len(actions) < 2:
            return False, "Action button requires at least 2 steps.", None, None

        continue_on_error = bool(item.get("continueOnError", False))
        success_count = 0
        total = len(actions)
        failure_lines = []

        for i, action in enumerate(actions):
            step_label = "%d/%d" % (i + 1, total)
            action_type = action.get("actionType", "script")
            action_target = action.get("target", "")

            self._set_status_message("[%s] Running %s: %s" % (step_label, action_type, action_target))

            if action_type == "filter":
                ok, error_reason = self._run_filter_with_error(action_target)
            else:
                ok, error_reason = self._run_script_with_error(action_target)

            if ok:
                success_count += 1
                continue

            failure_line = "Step %d out of %d failed (%s): %s" % (i + 1, total, action_type, action_target)
            if error_reason:
                failure_line += "\nReason: %s" % error_reason
            failure_lines.append(failure_line)

            if not continue_on_error:
                error_message = (
                    "The action sequence stopped at step %d out of %d while running %s: %s.\n"
                    "Completed %d out of %d steps successfully."
                ) % (i + 1, total, action_type, action_target, success_count, total)
                if error_reason:
                    error_message += "\n\nReason: %s" % error_reason
                self._show_action_error_dialog(item.get("name", "Action"), error_message)
                return False, error_message, None, None

        if success_count == total:
            return True, "Action '%s' finished successfully: %d out of %d steps completed." % (
                item.get("name", "Action"),
                success_count,
                total,
            ), None, None

        summary = "Action '%s' finished with warnings: %d out of %d steps completed successfully." % (
            item.get("name", "Action"),
            success_count,
            total,
        )
        details = "Some steps failed but execution continued because 'Continue on error' is enabled.\n\n" + "\n".join(failure_lines)
        return False, summary, "Action Button Warning", details

    @objc.python_method
    def _show_action_error_dialog(self, action_name, message):
        alert = NSAlert.alloc().init()
        alert.setMessageText_("Action Button Error")
        alert.setInformativeText_("%s\n\n%s" % (action_name, message))
        alert.addButtonWithTitle_("OK")
        alert.runModal()

    @objc.python_method
    def _run_filter_with_error(self, filter_name):
        print("ActionButtons: Running filter: '%s'" % filter_name)
        menu_item = self._find_menu_item_by_title(filter_name)
        if menu_item is None:
            print("ActionButtons: Could not find menu item for filter: '%s'" % filter_name)
            return False, "The filter could not be found in Glyphs' menu system."

        try:
            target = menu_item.target()
            action = menu_item.action()
            if action is not None:
                # Target can be None - responder chain will find the appropriate target
                NSApp.sendAction_to_from_(action, target, menu_item)
                return True, None
            else:
                print("ActionButtons: Menu item missing action")
                return False, "The filter menu item exists but does not expose a runnable action."
        except Exception:
            print("ActionButtons: Error running filter '%s'" % filter_name)
            traceback.print_exc()
            return False, self._format_exception_reason()

    @objc.python_method
    def _find_menu_item_by_title(self, title):
        if FILTER_MENU is not None:
            try:
                found = self._search_menu_for_title(Glyphs.menu[FILTER_MENU], title)
                if found is not None:
                    return found
            except Exception:
                pass

        try:
            for menu_key in Glyphs.menu:
                found = self._search_menu_for_title(Glyphs.menu[menu_key], title)
                if found is not None:
                    return found
        except Exception:
            pass

        return None

    @objc.python_method
    def _search_menu_for_title(self, menu_or_item, title):
        if menu_or_item is None:
            return None

        if hasattr(menu_or_item, "numberOfItems"):
            for i in range(menu_or_item.numberOfItems()):
                item = menu_or_item.itemAtIndex_(i)
                found = self._search_menu_for_title(item, title)
                if found is not None:
                    return found
            return None

        if isinstance(menu_or_item, (list, tuple)):
            for element in menu_or_item:
                found = self._search_menu_for_title(element, title)
                if found is not None:
                    return found
            return None

        if hasattr(menu_or_item, "submenu"):
            item_title = (menu_or_item.title() or "").strip()
            if item_title == title:
                return menu_or_item

            submenu = menu_or_item.submenu()
            if submenu and submenu.numberOfItems() > 0:
                return self._search_menu_for_title(submenu, title)

        return None

    @objc.python_method
    def _run_script_with_error(self, script_relative_path):
        script_path = self._script_absolute_path_for_target(script_relative_path)

        if script_path is None or not os.path.isfile(script_path):
            print("ActionButtons: Script not found: %s" % script_relative_path)
            return False, "The script file could not be found at the saved relative path."

        try:
            print("ActionButtons: Running script path: %s" % script_path)
            namespace = {
                "__file__": script_path,
                "__name__": "__main__",
                "Glyphs": Glyphs,
                "Font": getattr(Glyphs, "font", None),
            }
            with open(script_path, "r") as handle:
                script_source = handle.read()
            exec(compile(script_source, script_path, "exec"), namespace, namespace)
            return True, None
        except Exception:
            print("ActionButtons: error running script '%s'" % script_relative_path)
            traceback.print_exc()
            return False, self._format_exception_reason()

    @objc.python_method
    def _open_add_dialog(self, sender):
        self._refresh_available_targets()

        if self.addDialog is None:
            self.addDialog = FloatingWindow((520, 490), "Add Button")
            self.addDialog.bind("close", self._add_dialog_closed)
            self.addDialog.typeLabel = TextBox((12, 14, 70, 20), "Type")
            self.addDialog.typeMode = SegmentedButton(
                (80, 10, 200, 26),
                [{"title": "Filter"}, {"title": "Script"}, {"title": "Action"}],
                callback=self._add_dialog_type_changed,
            )

            self.addDialog.targetLabel = TextBox((12, 52, 70, 20), "Action")
            self.addDialog.targetPopup = PopUpButton((80, 48, -12, 26), [], callback=self._add_dialog_target_changed)
            self.addDialog.targetButton = Button((80, 48, -12, 26), "Choose Script...", callback=self._open_script_menu)

            self.addDialog.actionListLabel = TextBox((12, 52, 120, 20), "Action Steps")
            self.addDialog.actionList = List(
                (12, 106, -12, 176),
                [],
                columnDescriptions=[
                    {"title": "#", "key": "step", "width": 28},
                    {"title": "Type", "key": "type", "width": 70},
                    {"title": "Action", "key": "target"},
                ],
                allowsEmptySelection=True,
                drawFocusRing=False,
                selectionCallback=self._add_action_list_selection_changed,
            )
            self.addDialog.actionAddStep = Button((12, 76, 90, 24), "Add Step", callback=self._add_action_step)
            self.addDialog.actionRemoveStep = Button((108, 76, 110, 24), "Remove Step", callback=self._remove_action_step)
            self.addDialog.actionStepUp = Button((224, 76, 80, 24), "Step Up", callback=self._move_action_step_up)
            self.addDialog.actionStepDown = Button((310, 76, 100, 24), "Step Down", callback=self._move_action_step_down)
            self.addDialog.actionContinueOnError = CheckBox((12, 287, 230, 20), "Continue on error", value=False)

            self.addDialog.nameLabel = TextBox((12, 324, 70, 20), "Name")
            self.addDialog.nameEdit = EditText((80, 320, -12, 26), "")

            self.addDialog.categoryLabel = TextBox((12, 360, 70, 20), "Category")
            self.addDialog.categoryEdit = EditText((80, 356, -12, 26), "")
            try:
                self.addDialog.categoryEdit.getNSTextField().cell().setPlaceholderString_("Optional")
            except Exception:
                pass

            self.addDialog.shortcutLabel = TextBox((12, 396, 70, 20), "Shortcut")
            self.addDialog.shortcutEdit = EditText((80, 392, 150, 26), "")
            self.addDialog.shortcutRecordButton = Button((240, 393, 80, 24), "Record", callback=self._record_add_shortcut)
            self.addDialog.shortcutClearButton = Button((326, 393, 80, 24), "Clear", callback=self._clear_add_shortcut)
            self.addDialog.shortcutUseGlyphsButton = Button((412, 393, 96, 24), "Use Existing", callback=self._sync_add_shortcut_from_glyphs)
            self.addDialog.shortcutEdit.getNSTextField().setEditable_(False)
            try:
                self.addDialog.shortcutEdit.getNSTextField().cell().setPlaceholderString_("Optional")
            except Exception:
                pass
            self.addDialog.shortcutNote = TextBox(
                (80, 419, -12, 18),
                "Action Buttons' shortcuts are independent from Glyphs' shortcuts.",
                sizeStyle="small",
            )

            self.addDialog.cancelButton = Button((165, -36, 90, 24), "Cancel", callback=self._close_add_dialog)
            self.addDialog.addButton = Button((265, -36, 90, 24), "Add", callback=self._confirm_add_dialog)

        self.addDialog.typeMode.set(0)
        self.addDialogSelectedScript = None
        self.addDialogActionSteps = []
        self.addDialog.categoryEdit.set("")
        self.addDialog.shortcutEdit.set("")
        self._stop_shortcut_capture()
        self._populate_add_dialog_targets("filter")
        self._update_add_dialog_glyphs_shortcut_button()
        self._update_add_dialog_shortcut_clear_button()
        self._update_add_dialog_submit_button()
        try:
            self.addDialog.open()
        except ValueError:
            # Vanilla windows cannot be reopened after close; recreate the dialog.
            self.addDialog = None
            self._open_add_dialog(sender)
            return
        self.addDialog.makeKey()

    @objc.python_method
    def _close_add_dialog(self, sender):
        self._close_action_step_dialog(sender)
        self._stop_shortcut_capture()
        if self.addDialog is not None:
            self.addDialog.close()
            self.addDialog = None

    @objc.python_method
    def _add_dialog_closed(self, sender):
        self._stop_shortcut_capture()
        self.addDialog = None

    @objc.python_method
    def _add_dialog_type_changed(self, sender):
        mode = sender.get()
        if mode == 0:
            item_type = "filter"
        elif mode == 1:
            item_type = "script"
        else:
            item_type = "action"
        self._populate_add_dialog_targets(item_type)

    @objc.python_method
    def _add_dialog_target_changed(self, sender):
        selection = sender.getItem()
        if selection:
            self.addDialog.nameEdit.set(selection)
        self._update_add_dialog_glyphs_shortcut_button()
        self._update_add_dialog_submit_button()

    @objc.python_method
    def _add_dialog_has_required_inputs(self):
        if self.addDialog is None:
            return False
        mode = self.addDialog.typeMode.get()
        if mode == 0:
            target = self.addDialog.targetPopup.getItem()
            return bool(target and target != "(No items found)")
        if mode == 1:
            return bool(self.addDialogSelectedScript)
        return len(self.addDialogActionSteps) >= 2

    @objc.python_method
    def _update_add_dialog_submit_button(self):
        if self.addDialog is None or not hasattr(self.addDialog, "addButton"):
            return
        self.addDialog.addButton.enable(self._add_dialog_has_required_inputs())

    @objc.python_method
    def _set_dialog_height(self, dialog, height):
        if dialog is None:
            return
        try:
            x, y, w, _ = dialog.getPosSize()
            dialog.setPosSize((x, y, w, height))
        except Exception:
            pass

    @objc.python_method
    def _layout_add_edit_fields(self, dialog, item_type):
        if dialog is None:
            return

        compact = item_type in ("filter", "script")
        if compact:
            name_label_y = 88
            name_edit_y = 84
            category_label_y = 124
            category_edit_y = 120
            shortcut_label_y = 160
            shortcut_edit_y = 156
            shortcut_row_y = 157
            shortcut_note_y = 183
            self._set_dialog_height(dialog, COMPACT_BUTTON_DIALOG_HEIGHT)
        else:
            name_label_y = 324
            name_edit_y = 320
            category_label_y = 360
            category_edit_y = 356
            shortcut_label_y = 396
            shortcut_edit_y = 392
            shortcut_row_y = 393
            shortcut_note_y = 419
            self._set_dialog_height(dialog, FULL_BUTTON_DIALOG_HEIGHT)

        dialog.nameLabel.setPosSize((12, name_label_y, 70, 20))
        dialog.nameEdit.setPosSize((80, name_edit_y, -12, 26))
        dialog.categoryLabel.setPosSize((12, category_label_y, 70, 20))
        dialog.categoryEdit.setPosSize((80, category_edit_y, -12, 26))
        dialog.shortcutLabel.setPosSize((12, shortcut_label_y, 70, 20))
        dialog.shortcutEdit.setPosSize((80, shortcut_edit_y, 150, 26))
        dialog.shortcutRecordButton.setPosSize((240, shortcut_row_y, 80, 24))
        dialog.shortcutClearButton.setPosSize((326, shortcut_row_y, 80, 24))
        dialog.shortcutUseGlyphsButton.setPosSize((412, shortcut_row_y, 96, 24))
        dialog.shortcutNote.setPosSize((80, shortcut_note_y, -12, 18))

    @objc.python_method
    def _populate_add_dialog_targets(self, item_type):
        if item_type == "filter":
            # Show dropdown for filters (flat list)
            self.addDialog.targetPopup.show(True)
            self.addDialog.targetButton.show(False)
            self.addDialog.actionListLabel.show(False)
            self.addDialog.actionList.show(False)
            self.addDialog.actionAddStep.show(False)
            self.addDialog.actionRemoveStep.show(False)
            self.addDialog.actionStepUp.show(False)
            self.addDialog.actionStepDown.show(False)
            self.addDialog.actionContinueOnError.show(False)
            options = self.cachedFilters if self.cachedFilters else ["(No items found)"]
            self.addDialog.targetPopup.setItems(options)
            self.addDialog.targetPopup.set(0)
            self.addDialog.nameEdit.set(options[0] if options else "")
            self._layout_add_edit_fields(self.addDialog, item_type)
            self._update_add_dialog_shortcut_clear_button()
        else:
            if item_type == "script":
                # Show button for scripts (hierarchical menu)
                self.addDialog.targetPopup.show(False)
                self.addDialog.targetButton.show(True)
                self.addDialog.actionListLabel.show(False)
                self.addDialog.actionList.show(False)
                self.addDialog.actionAddStep.show(False)
                self.addDialog.actionRemoveStep.show(False)
                self.addDialog.actionStepUp.show(False)
                self.addDialog.actionStepDown.show(False)
                self.addDialog.actionContinueOnError.show(False)
                self.addDialogSelectedScript = None
                if self.cachedScripts:
                    self.addDialog.targetButton.setTitle("Choose Script...")
                else:
                    self.addDialog.targetButton.setTitle("(No scripts found)")
                self.addDialog.nameEdit.set("")
                self._layout_add_edit_fields(self.addDialog, item_type)
                self._update_add_dialog_shortcut_clear_button()
            else:
                # Show action builder controls
                self.addDialog.targetPopup.show(False)
                self.addDialog.targetButton.show(False)
                self.addDialog.actionListLabel.show(True)
                self.addDialog.actionList.show(True)
                self.addDialog.actionAddStep.show(True)
                self.addDialog.actionRemoveStep.show(True)
                self.addDialog.actionStepUp.show(True)
                self.addDialog.actionStepDown.show(True)
                self.addDialog.actionContinueOnError.show(True)
                self.addDialogActionSteps = []
                self.addDialog.actionContinueOnError.set(False)
                self.addDialog.nameEdit.set("")
                self._refresh_action_list()
                self._layout_add_edit_fields(self.addDialog, item_type)
                self._update_add_dialog_shortcut_clear_button()

            self._update_add_dialog_glyphs_shortcut_button()
        self._update_add_dialog_submit_button()

    @objc.python_method
    def _refresh_action_list(self):
        rows = []
        for i, action in enumerate(self.addDialogActionSteps):
            rows.append({
                "step": str(i + 1),
                "type": action.get("actionType", "").capitalize(),
                "target": action.get("target", ""),
            })
        self.addDialog.actionList.set(rows)
        self._update_action_step_buttons_for_owner("add")
        self._update_add_dialog_submit_button()

    @objc.python_method
    def _refresh_edit_action_list(self):
        rows = []
        for i, action in enumerate(self.editDialogActionSteps):
            rows.append({
                "step": str(i + 1),
                "type": action.get("actionType", "").capitalize(),
                "target": action.get("target", ""),
            })
        if self.editDialog is not None:
            self.editDialog.actionList.set(rows)
        self._update_action_step_buttons_for_owner("edit")

    @objc.python_method
    def _add_action_list_selection_changed(self, sender):
        self._update_action_step_buttons_for_owner("add")

    @objc.python_method
    def _edit_action_list_selection_changed(self, sender):
        self._update_action_step_buttons_for_owner("edit")

    @objc.python_method
    def _update_action_step_buttons_for_owner(self, owner):
        if owner == "edit":
            dialog = self.editDialog
            steps = self.editDialogActionSteps
        else:
            dialog = self.addDialog
            steps = self.addDialogActionSteps

        if dialog is None or not hasattr(dialog, "actionRemoveStep"):
            return

        index = self._selected_action_step_index_for_owner(owner)
        has_selection = index is not None and 0 <= index < len(steps)

        dialog.actionRemoveStep.enable(has_selection)
        dialog.actionStepUp.enable(has_selection and index > 0)
        dialog.actionStepDown.enable(has_selection and index < (len(steps) - 1))

    @objc.python_method
    def _refresh_action_list_for_owner(self):
        if self.actionStepDialogOwner == "edit":
            self._refresh_edit_action_list()
        else:
            self._refresh_action_list()

    @objc.python_method
    def _add_action_step(self, sender):
        self._open_action_step_dialog(sender, owner="add")

    @objc.python_method
    def _remove_action_step(self, sender):
        self._remove_action_step_for_owner("add")

    @objc.python_method
    def _move_action_step_up(self, sender):
        self._move_action_step_for_owner("add", -1)

    @objc.python_method
    def _move_action_step_down(self, sender):
        self._move_action_step_for_owner("add", 1)

    @objc.python_method
    def _selected_action_step_index_for_owner(self, owner):
        if owner == "edit":
            if self.editDialog is None:
                return None
            selection = self.editDialog.actionList.getSelection()
        else:
            if self.addDialog is None:
                return None
            selection = self.addDialog.actionList.getSelection()
        if not selection:
            return None
        return selection[0]

    @objc.python_method
    def _remove_action_step_for_owner(self, owner):
        steps = self.editDialogActionSteps if owner == "edit" else self.addDialogActionSteps
        dialog = self.editDialog if owner == "edit" else self.addDialog
        index = self._selected_action_step_index_for_owner(owner)
        if index is None or index < 0 or index >= len(steps):
            return
        del steps[index]
        self.actionStepDialogOwner = owner
        self._refresh_action_list_for_owner()
        if dialog is not None and len(steps) > 0:
            next_index = min(index, len(steps) - 1)
            dialog.actionList.setSelection([next_index])
        self._update_action_step_buttons_for_owner(owner)

    @objc.python_method
    def _move_action_step_for_owner(self, owner, delta):
        steps = self.editDialogActionSteps if owner == "edit" else self.addDialogActionSteps
        index = self._selected_action_step_index_for_owner(owner)
        if index is None:
            return
        new_index = index + delta
        if new_index < 0 or new_index >= len(steps):
            return
        steps[new_index], steps[index] = steps[index], steps[new_index]
        self.actionStepDialogOwner = owner
        self._refresh_action_list_for_owner()
        if owner == "edit" and self.editDialog is not None:
            self.editDialog.actionList.setSelection([new_index])
        elif owner == "add" and self.addDialog is not None:
            self.addDialog.actionList.setSelection([new_index])
        self._update_action_step_buttons_for_owner(owner)

    @objc.python_method
    def _open_action_step_dialog(self, sender, owner="add"):
        self.actionStepDialogOwner = owner
        if self.actionStepDialog is None:
            self.actionStepDialog = FloatingWindow((460, 180), "Add Action Step")
            self.actionStepDialog.typeLabel = TextBox((12, 14, 70, 20), "Type")
            self.actionStepDialog.typeMode = SegmentedButton(
                (80, 10, 200, 26),
                [{"title": "Filter"}, {"title": "Script"}],
                callback=self._action_step_type_changed,
            )
            self.actionStepDialog.targetLabel = TextBox((12, 52, 70, 20), "Action")
            self.actionStepDialog.targetPopup = PopUpButton((80, 48, -12, 26), [])
            self.actionStepDialog.targetButton = Button((80, 48, -12, 26), "Choose Script...", callback=self._open_action_step_script_menu)
            self.actionStepDialog.cancelButton = Button((250, -36, 90, 24), "Cancel", callback=self._close_action_step_dialog)
            self.actionStepDialog.addButton = Button((350, -36, 90, 24), "Add", callback=self._confirm_action_step_dialog)

        self.actionStepSelectedScript = None
        self.actionStepDialog.typeMode.set(0)
        self._populate_action_step_targets("filter")

        try:
            self.actionStepDialog.open()
        except ValueError:
            self.actionStepDialog = None
            self._open_action_step_dialog(sender)
            return
        self.actionStepDialog.makeKey()

    @objc.python_method
    def _close_action_step_dialog(self, sender):
        if self.actionStepDialog is not None:
            self.actionStepDialog.close()
            self.actionStepDialog = None

    @objc.python_method
    def _action_step_type_changed(self, sender):
        mode = sender.get()
        item_type = "filter" if mode == 0 else "script"
        self._populate_action_step_targets(item_type)

    @objc.python_method
    def _populate_action_step_targets(self, item_type):
        if item_type == "filter":
            self.actionStepDialog.targetPopup.show(True)
            self.actionStepDialog.targetButton.show(False)
            options = self.cachedFilters if self.cachedFilters else ["(No items found)"]
            self.actionStepDialog.targetPopup.setItems(options)
            self.actionStepDialog.targetPopup.set(0)
        else:
            self.actionStepDialog.targetPopup.show(False)
            self.actionStepDialog.targetButton.show(True)
            self.actionStepSelectedScript = None
            if self.cachedScripts:
                self.actionStepDialog.targetButton.setTitle("Choose Script...")
            else:
                self.actionStepDialog.targetButton.setTitle("(No scripts found)")

    @objc.python_method
    def _open_action_step_script_menu(self, sender):
        if not self.cachedScripts:
            return
        menu = NSMenu.alloc().init()
        self._build_script_menu(menu, self.cachedScripts, "", "actionStepScriptMenuItemSelected:")
        button_view = sender.getNSButton()
        menu.popUpMenuPositioningItem_atLocation_inView_(None, (0, button_view.frame().size.height), button_view)

    def actionStepScriptMenuItemSelected_(self, sender):
        script_path = sender.representedObject()
        if script_path:
            self.actionStepSelectedScript = script_path
            display_name = self._script_display_title_for_target(script_path)
            self.actionStepDialog.targetButton.setTitle(display_name)

    @objc.python_method
    def _confirm_action_step_dialog(self, sender):
        mode = self.actionStepDialog.typeMode.get()
        item_type = "filter" if mode == 0 else "script"

        if item_type == "filter":
            target = self.actionStepDialog.targetPopup.getItem()
            if not target or target == "(No items found)":
                return
        else:
            target = self.actionStepSelectedScript
            if not target:
                return

        if self.actionStepDialogOwner == "edit":
            target_steps = self.editDialogActionSteps
        else:
            target_steps = self.addDialogActionSteps

        target_steps.append({
            "actionType": item_type,
            "target": target,
        })
        self._refresh_action_list_for_owner()
        self._update_add_dialog_submit_button()
        self._close_action_step_dialog(sender)

    @objc.python_method
    def _open_script_menu(self, sender):
        if not self.cachedScripts:
            return
        
        # Build hierarchical menu from script paths
        menu = NSMenu.alloc().init()
        self._build_script_menu(menu, self.cachedScripts, "")
        
        # Display menu at button location
        button_view = sender.getNSButton()
        menu.popUpMenuPositioningItem_atLocation_inView_(None, (0, button_view.frame().size.height), button_view)

    @objc.python_method
    def _pretty_script_display_token(self, token):
        if token is None:
            return ""
        pretty = str(token)
        if pretty.endswith(".py"):
            pretty = pretty[:-3]
        pretty = pretty.replace("_", " ").replace("-", " ")
        return " ".join(pretty.split())

    @objc.python_method
    def _build_script_menu(self, menu, script_paths, current_prefix, selector_name="scriptMenuItemSelected:"):
        # Group scripts by their immediate folder
        folders = {}
        files_in_folder = []
        
        for script_path in script_paths:
            if current_prefix:
                if not script_path.startswith(current_prefix):
                    continue
                relative = script_path[len(current_prefix):].lstrip("/")
            else:
                relative = script_path
            
            parts = relative.split("/")
            if len(parts) == 1:
                # File at this level
                files_in_folder.append(script_path)
            else:
                # File in subfolder
                folder = parts[0]
                if folder not in folders:
                    folders[folder] = []
                folders[folder].append(script_path)
        
        # Add subfolders as submenus
        for folder_name in sorted(folders.keys()):
            submenu = NSMenu.alloc().init()
            menu_item = NSMenuItem.alloc().initWithTitle_action_keyEquivalent_(
                self._pretty_script_display_token(folder_name),
                None,
                "",
            )
            menu_item.setSubmenu_(submenu)
            menu.addItem_(menu_item)
            
            # Recursively build submenu
            prefix = (current_prefix + "/" + folder_name) if current_prefix else folder_name
            self._build_script_menu(submenu, folders[folder_name], prefix, selector_name)
        
        # Add files at this level
        for script_path in sorted(files_in_folder):
            display_name = self._script_display_title_for_target(script_path)
            menu_item = NSMenuItem.alloc().initWithTitle_action_keyEquivalent_(display_name, selector_name, "")
            menu_item.setTarget_(self)
            menu_item.setRepresentedObject_(script_path)
            menu.addItem_(menu_item)

    def scriptMenuItemSelected_(self, sender):
        script_path = sender.representedObject()
        if script_path:
            self.addDialogSelectedScript = script_path
            # Update button title and name field
            display_name = self._script_display_title_for_target(script_path)
            self.addDialog.targetButton.setTitle(display_name)
            self.addDialog.nameEdit.set(display_name)
            self._update_add_dialog_glyphs_shortcut_button()
            self._update_add_dialog_submit_button()

    @objc.python_method
    def _confirm_add_dialog(self, sender):
        mode = self.addDialog.typeMode.get()
        if mode == 0:
            item_type = "filter"
        elif mode == 1:
            item_type = "script"
        else:
            item_type = "action"
        
        if item_type == "filter":
            target = self.addDialog.targetPopup.getItem()
            if not target or target == "(No items found)":
                return
            actions = []
            continue_on_error = False
        elif item_type == "script":
            target = self.addDialogSelectedScript
            if not target:
                return
            actions = []
            continue_on_error = False
        else:
            if len(self.addDialogActionSteps) < 2:
                self._set_status_message("Action buttons need at least 2 steps.")
                return
            target = ""
            actions = list(self.addDialogActionSteps)
            continue_on_error = bool(self.addDialog.actionContinueOnError.get())
        
        category = self.addDialog.categoryEdit.get().strip()
        name = self.addDialog.nameEdit.get().strip()
        if not name:
            if item_type == "script":
                name = self._script_display_title_for_target(target)
            elif item_type == "action":
                name = "Action (%d steps)" % len(actions)
            else:
                name = target

        shortcut = self._normalized_shortcut_value(self.addDialog.shortcutEdit.get())
        if shortcut and not self._shortcut_has_required_modifier(shortcut):
            self._show_shortcut_validation_error()
            return
        conflict_name = self._find_shortcut_conflict_name(shortcut)
        if conflict_name:
            self._show_shortcut_conflict_alert(shortcut, conflict_name)
            return

        self.items.append({
            "id": str(uuid.uuid4()),
            "category": category,
            "name": name,
            "type": item_type,
            "target": target,
            "orderIndex": len(self.items),
            "actions": actions,
            "continueOnError": continue_on_error,
            "shortcut": shortcut,
        })

        self._save_items()
        self._refresh_ui()
        self._close_add_dialog(sender)

    @objc.python_method
    def _add_edit_action_step(self, sender):
        self._open_action_step_dialog(sender, owner="edit")

    @objc.python_method
    def _remove_edit_action_step(self, sender):
        self._remove_action_step_for_owner("edit")

    @objc.python_method
    def _move_edit_action_step_up(self, sender):
        self._move_action_step_for_owner("edit", -1)

    @objc.python_method
    def _move_edit_action_step_down(self, sender):
        self._move_action_step_for_owner("edit", 1)

    @objc.python_method
    def _open_edit_dialog(self, sender):
        self._refresh_available_targets()
        index = self._selected_index()
        item = self._selected_item()
        if index is None or item is None:
            self._set_status_message("Select a row to edit.")
            return

        if self.editDialog is None:
            self.editDialog = FloatingWindow((520, 490), "Edit Button")
            self.editDialog.bind("close", self._edit_dialog_closed)
            self.editDialog.typeLabel = TextBox((12, 14, 70, 20), "Type")
            self.editDialog.typeValue = TextBox((80, 14, -12, 20), "")
            self.editDialog.targetLabel = TextBox((12, 52, 70, 20), "Action")
            self.editDialog.targetPopup = PopUpButton((80, 48, -12, 26), [], callback=self._edit_dialog_target_changed)
            self.editDialog.targetButton = Button((80, 48, -12, 26), "Choose Script...", callback=self._open_edit_script_menu)

            self.editDialog.actionListLabel = TextBox((12, 52, 120, 20), "Action Steps")
            self.editDialog.actionList = List(
                (12, 106, -12, 176),
                [],
                columnDescriptions=[
                    {"title": "#", "key": "step", "width": 28},
                    {"title": "Type", "key": "type", "width": 70},
                    {"title": "Action", "key": "target"},
                ],
                allowsEmptySelection=True,
                drawFocusRing=False,
                selectionCallback=self._edit_action_list_selection_changed,
            )
            self.editDialog.actionAddStep = Button((12, 76, 90, 24), "Add Step", callback=self._add_edit_action_step)
            self.editDialog.actionRemoveStep = Button((108, 76, 110, 24), "Remove Step", callback=self._remove_edit_action_step)
            self.editDialog.actionStepUp = Button((224, 76, 80, 24), "Step Up", callback=self._move_edit_action_step_up)
            self.editDialog.actionStepDown = Button((310, 76, 100, 24), "Step Down", callback=self._move_edit_action_step_down)
            self.editDialog.actionContinueOnError = CheckBox((12, 287, 230, 20), "Continue on error", value=False)

            self.editDialog.nameLabel = TextBox((12, 324, 70, 20), "Name")
            self.editDialog.nameEdit = EditText((80, 320, -12, 26), "")

            self.editDialog.categoryLabel = TextBox((12, 360, 70, 20), "Category")
            self.editDialog.categoryEdit = EditText((80, 356, -12, 26), "")
            try:
                self.editDialog.categoryEdit.getNSTextField().cell().setPlaceholderString_("Optional")
            except Exception:
                pass

            self.editDialog.shortcutLabel = TextBox((12, 396, 70, 20), "Shortcut")
            self.editDialog.shortcutEdit = EditText((80, 392, 150, 26), "")
            self.editDialog.shortcutRecordButton = Button((240, 393, 80, 24), "Record", callback=self._record_edit_shortcut)
            self.editDialog.shortcutClearButton = Button((326, 393, 80, 24), "Clear", callback=self._clear_edit_shortcut)
            self.editDialog.shortcutUseGlyphsButton = Button((412, 393, 96, 24), "Use Existing", callback=self._sync_edit_shortcut_from_glyphs)
            self.editDialog.shortcutEdit.getNSTextField().setEditable_(False)
            try:
                self.editDialog.shortcutEdit.getNSTextField().cell().setPlaceholderString_("Optional")
            except Exception:
                pass
            self.editDialog.shortcutNote = TextBox(
                (80, 419, -12, 18),
                "Action Buttons' shortcuts are independent from Glyphs' shortcuts.",
                sizeStyle="small",
            )
            self.editDialog.cancelButton = Button((165, -36, 90, 24), "Cancel", callback=self._close_edit_dialog)
            self.editDialog.saveButton = Button((265, -36, 90, 24), "Save", callback=self._confirm_edit_dialog)

        self.editDialog.currentIndex = index
        self.editDialog.currentType = item.get("type", "script")
        self.editDialogSelectedScript = item.get("target", "") if item.get("type") == "script" else None
        self.editDialogActionSteps = [
            {"actionType": action.get("actionType", "script"), "target": action.get("target", "")}
            for action in item.get("actions", [])
            if action.get("target")
        ]
        self.editDialog.actionContinueOnError.set(bool(item.get("continueOnError", False)))
        self.editDialog.categoryEdit.set(item.get("category", ""))
        self.editDialog.nameEdit.set(item.get("name", ""))
        self.editDialog.shortcutEdit.set(item.get("shortcut", ""))
        self.editDialog.typeValue.set(item.get("type", "").capitalize())
        self._update_edit_dialog_shortcut_clear_button()

        self._populate_edit_dialog_targets(item)

        self.actionStepDialogOwner = "edit"
        self._refresh_edit_action_list()
        self._stop_shortcut_capture()
        self._update_edit_dialog_glyphs_shortcut_button()

        try:
            self.editDialog.open()
        except ValueError:
            self.editDialog = None
            self._open_edit_dialog(sender)
            return
        self.editDialog.makeKey()

    @objc.python_method
    def _populate_edit_dialog_targets(self, item):
        item_type = item.get("type", "script")

        self.editDialog.targetLabel.show(item_type in ("filter", "script"))
        self.editDialog.targetPopup.show(item_type == "filter")
        self.editDialog.targetButton.show(item_type == "script")

        self.editDialog.actionListLabel.show(item_type == "action")
        self.editDialog.actionList.show(item_type == "action")
        self.editDialog.actionAddStep.show(item_type == "action")
        self.editDialog.actionRemoveStep.show(item_type == "action")
        self.editDialog.actionStepUp.show(item_type == "action")
        self.editDialog.actionStepDown.show(item_type == "action")
        self.editDialog.actionContinueOnError.show(item_type == "action")

        if item_type == "filter":
            options = self.cachedFilters if self.cachedFilters else ["(No items found)"]
            self.editDialog.targetPopup.setItems(options)
            current_target = item.get("target", "")
            if current_target in options:
                self.editDialog.targetPopup.set(options.index(current_target))
            else:
                self.editDialog.targetPopup.set(0)
        elif item_type == "script":
            current_target = item.get("target", "")
            if current_target:
                display_name = self._script_display_title_for_target(current_target)
                self.editDialog.targetButton.setTitle(display_name)
            elif self.cachedScripts or self._is_instructions_item(item):
                self.editDialog.targetButton.setTitle("Choose Script...")
            else:
                self.editDialog.targetButton.setTitle("(No scripts found)")

        self._layout_add_edit_fields(self.editDialog, item_type)
        self._update_edit_dialog_shortcut_clear_button()
        self._update_edit_dialog_glyphs_shortcut_button()

    @objc.python_method
    def _edit_dialog_target_changed(self, sender):
        self._update_edit_dialog_glyphs_shortcut_button()

    @objc.python_method
    def _open_edit_script_menu(self, sender):
        is_instructions_item = self._is_editing_instructions_item()
        if not self.cachedScripts and not is_instructions_item:
            return
        menu = NSMenu.alloc().init()
        if is_instructions_item:
            menu_item = NSMenuItem.alloc().initWithTitle_action_keyEquivalent_(
                INSTRUCTIONS_PSEUDO_SCRIPT_TITLE, "editScriptMenuItemSelected:", ""
            )
            menu_item.setTarget_(self)
            menu_item.setRepresentedObject_(INSTRUCTIONS_PSEUDO_SCRIPT_TARGET)
            menu.addItem_(menu_item)
            if self.cachedScripts:
                menu.addItem_(NSMenuItem.separatorItem())
        self._build_script_menu(menu, self.cachedScripts, "", "editScriptMenuItemSelected:")
        button_view = sender.getNSButton()
        menu.popUpMenuPositioningItem_atLocation_inView_(None, (0, button_view.frame().size.height), button_view)

    def editScriptMenuItemSelected_(self, sender):
        script_path = sender.representedObject()
        if script_path:
            self.editDialogSelectedScript = script_path
            display_name = self._script_display_title_for_target(script_path)
            self.editDialog.targetButton.setTitle(display_name)
            self._update_edit_dialog_glyphs_shortcut_button()

    @objc.python_method
    def _iter_menu_items(self, menu_or_item):
        if menu_or_item is None:
            return

        if hasattr(menu_or_item, "numberOfItems"):
            for i in range(menu_or_item.numberOfItems()):
                item = menu_or_item.itemAtIndex_(i)
                for nested in self._iter_menu_items(item):
                    yield nested
            return

        if isinstance(menu_or_item, (list, tuple)):
            for element in menu_or_item:
                for nested in self._iter_menu_items(element):
                    yield nested
            return

        if hasattr(menu_or_item, "title"):
            yield menu_or_item
            submenu = None
            if hasattr(menu_or_item, "submenu"):
                submenu = menu_or_item.submenu()
            if submenu and submenu.numberOfItems() > 0:
                for nested in self._iter_menu_items(submenu):
                    yield nested

    @objc.python_method
    def _iter_menu_items_with_path(self, menu_or_item, path_tokens=None):
        if path_tokens is None:
            path_tokens = []

        if menu_or_item is None:
            return

        if hasattr(menu_or_item, "numberOfItems"):
            for i in range(menu_or_item.numberOfItems()):
                item = menu_or_item.itemAtIndex_(i)
                for nested in self._iter_menu_items_with_path(item, path_tokens):
                    yield nested
            return

        if isinstance(menu_or_item, (list, tuple)):
            for element in menu_or_item:
                for nested in self._iter_menu_items_with_path(element, path_tokens):
                    yield nested
            return

        if hasattr(menu_or_item, "title"):
            try:
                title = (menu_or_item.title() or "").strip()
            except Exception:
                title = ""

            next_tokens = list(path_tokens)
            if title:
                next_tokens.append(title)

            yield menu_or_item, next_tokens

            submenu = None
            if hasattr(menu_or_item, "submenu"):
                submenu = menu_or_item.submenu()
            if submenu and submenu.numberOfItems() > 0:
                for nested in self._iter_menu_items_with_path(submenu, next_tokens):
                    yield nested

    @objc.python_method
    def _common_suffix_length(self, left_tokens, right_tokens):
        length = 0
        li = len(left_tokens) - 1
        ri = len(right_tokens) - 1
        while li >= 0 and ri >= 0:
            if left_tokens[li] != right_tokens[ri]:
                break
            length += 1
            li -= 1
            ri -= 1
        return length

    @objc.python_method
    def _normalized_script_match_token(self, token):
        if token is None:
            return ""
        normalized = str(token).strip().lower()
        if normalized.endswith(".py"):
            normalized = normalized[:-3]
        normalized = normalized.replace("…", " ")
        normalized = normalized.replace("_", " ").replace("-", " ")
        normalized = " ".join(normalized.split())
        return normalized

    @objc.python_method
    def _compact_script_match_token(self, token):
        normalized = self._normalized_script_match_token(token)
        return "".join(char for char in normalized if char.isalnum())

    @objc.python_method
    def _log_script_sync_diagnostic_once(self, script_target, reason):
        if not script_target or not reason:
            return
        if not hasattr(self, "_lastScriptSyncDiagnosticKey"):
            self._lastScriptSyncDiagnosticKey = None
        key = "%s|%s" % (script_target, reason)
        if key == self._lastScriptSyncDiagnosticKey:
            return
        self._lastScriptSyncDiagnosticKey = key
        print("ActionButtons: Script shortcut sync skipped for '%s' (%s)." % (script_target, reason))

    @objc.python_method
    def _debug_shortcut_sync(self, message, key=None):
        if not DEBUG_SHORTCUT_SYNC:
            return
        if not message:
            return
        if not hasattr(self, "_seenShortcutSyncDebugKeys"):
            self._seenShortcutSyncDebugKeys = set()
        if key:
            if key in self._seenShortcutSyncDebugKeys:
                return
            self._seenShortcutSyncDebugKeys.add(key)
        if not hasattr(self, "_lastShortcutSyncDebugMessage"):
            self._lastShortcutSyncDebugMessage = None
        if self._lastShortcutSyncDebugMessage == message:
            return
        self._lastShortcutSyncDebugMessage = message
        print("ActionButtons Sync: %s" % message)

    @objc.python_method
    def _debug_button_copy(self, item, size_key, fspec, spacing, category, name_line, shortcut, show_category_line, show_shortcut_line):
        if not DEBUG_BUTTON_COPY:
            return
        item_id = item.get("id", "")
        debug_key = "%s|%s|%s|%s|%s" % (
            item_id,
            size_key,
            fspec.get("category"),
            fspec.get("name"),
            fspec.get("shortcut"),
        )
        if not hasattr(self, "_seenButtonCopyDebugKeys"):
            self._seenButtonCopyDebugKeys = set()
        if debug_key in self._seenButtonCopyDebugKeys:
            return
        self._seenButtonCopyDebugKeys.add(debug_key)
        print(
            "ActionButtons Copy: id=%s size=%s fonts=(cat:%s name:%s short:%s) spacing=%.2f category=%r showCategory=%s name=%r shortcut=%r showShortcut=%s" % (
                item_id,
                size_key,
                fspec.get("category"),
                fspec.get("name"),
                fspec.get("shortcut"),
                spacing,
                category,
                bool(show_category_line),
                name_line,
                shortcut,
                bool(show_shortcut_line),
            )
        )

    @objc.python_method
    def _parse_user_key_equivalent(self, value):
        if not value:
            return None

        text = str(value)
        flags = 0
        index = 0
        while index < len(text):
            marker = text[index]
            if marker == "^":
                flags |= _MOD_CONTROL
            elif marker == "~":
                flags |= _MOD_OPTION
            elif marker == "$":
                flags |= _MOD_SHIFT
            elif marker == "@":
                flags |= _MOD_COMMAND
            else:
                break
            index += 1

        key = text[index:]
        if not key:
            return None

        special_chars = {
            " ": "Space",
            "\t": "Tab",
            "\r": "Return",
            "\n": "Return",
            "\x7f": "Delete",
            "\x08": "Delete",
        }

        if len(key) == 1:
            codepoint = ord(key)
            display_key = _FUNCTION_KEY_CHARS.get(codepoint)
            if display_key is None:
                display_key = special_chars.get(key, key.upper())
        else:
            display_key = key.upper()

        if not display_key:
            return None

        prefix = ""
        if flags & _MOD_CONTROL:
            prefix += "⌃"
        if flags & _MOD_OPTION:
            prefix += "⌥"
        if flags & _MOD_SHIFT:
            prefix += "⇧"
        if flags & _MOD_COMMAND:
            prefix += "⌘"
        return prefix + display_key

    @objc.python_method
    def _shortcut_from_user_defaults_for_menu_path(self, path_tokens):
        if not path_tokens:
            self._debug_shortcut_sync("NSUserKeyEquivalents fallback skipped (no path tokens)")
            return None

        try:
            defaults = NSUserDefaults.standardUserDefaults()
            key_equivalents = defaults.dictionaryForKey_("NSUserKeyEquivalents")
        except Exception:
            self._debug_shortcut_sync("NSUserKeyEquivalents fallback failed (defaults unavailable)")
            return None
        if not key_equivalents:
            self._debug_shortcut_sync("NSUserKeyEquivalents fallback found no custom key equivalents")
            return None

        leaf = str(path_tokens[-1]).strip()
        if not leaf:
            self._debug_shortcut_sync("NSUserKeyEquivalents fallback skipped (empty leaf title)")
            return None

        candidates = []
        for i in range(len(path_tokens)):
            suffix = [str(token).strip() for token in path_tokens[i:] if str(token).strip()]
            if not suffix:
                continue
            candidates.append("->".join(suffix))
            candidates.append(suffix[-1])

        # De-duplicate while preserving order.
        seen = set()
        unique_candidates = []
        for candidate in candidates:
            if candidate in seen:
                continue
            seen.add(candidate)
            unique_candidates.append(candidate)

        for candidate in unique_candidates:
            if candidate in key_equivalents:
                shortcut = self._parse_user_key_equivalent(key_equivalents[candidate])
                if shortcut:
                    self._debug_shortcut_sync(
                        "NSUserKeyEquivalents matched '%s' as '%s' for path %s"
                        % (candidate, shortcut, " -> ".join(path_tokens))
                    )
                    return shortcut

        self._debug_shortcut_sync(
            "NSUserKeyEquivalents no match for path %s; tried: %s"
            % (" -> ".join(path_tokens), ", ".join(unique_candidates))
        )
        return None

    @objc.python_method
    def _shortcut_from_user_defaults_for_script_target(self, script_target):
        if not script_target:
            return None

        try:
            defaults = NSUserDefaults.standardUserDefaults()
            key_equivalents = defaults.dictionaryForKey_("NSUserKeyEquivalents")
        except Exception:
            self._debug_shortcut_sync("Script-target fallback failed (defaults unavailable)")
            return None
        if not key_equivalents:
            self._debug_shortcut_sync("Script-target fallback found no NSUserKeyEquivalents entries")
            return None

        normalized_target = str(script_target).replace("\\", "/")
        target_parts = [part for part in normalized_target.split("/") if part]
        if not target_parts:
            return None

        leaf_filename = target_parts[-1]
        leaf_basename = leaf_filename[:-3] if leaf_filename.lower().endswith(".py") else leaf_filename
        menu_title = self._script_menu_title_for_target(script_target)
        leaf_titles = [self._pretty_script_display_token(leaf_basename)]
        if menu_title:
            leaf_titles.append(menu_title)

        folder_parts = [self._pretty_script_display_token(part) for part in target_parts[:-1]]
        candidates = []

        # Candidate keys with folder path suffixes and title variants.
        for title in leaf_titles:
            if not title:
                continue
            full_parts = folder_parts + [title]
            for i in range(len(full_parts)):
                suffix = [part for part in full_parts[i:] if part]
                if not suffix:
                    continue
                joined = "->".join(suffix)
                candidates.append(joined)
                candidates.append(suffix[-1])

        # Some apps may include explicit top-level menu token.
        for title in leaf_titles:
            if not title:
                continue
            full_parts = ["Scripts"] + folder_parts + [title]
            for i in range(len(full_parts)):
                suffix = [part for part in full_parts[i:] if part]
                if not suffix:
                    continue
                candidates.append("->".join(suffix))

        seen = set()
        unique_candidates = []
        for candidate in candidates:
            candidate = str(candidate).strip()
            if not candidate or candidate in seen:
                continue
            seen.add(candidate)
            unique_candidates.append(candidate)

        for candidate in unique_candidates:
            if candidate in key_equivalents:
                shortcut = self._parse_user_key_equivalent(key_equivalents[candidate])
                if shortcut:
                    self._debug_shortcut_sync(
                        "Script-target fallback matched '%s' as '%s' for '%s'"
                        % (candidate, shortcut, script_target),
                        key="script_target_match:%s" % script_target,
                    )
                    return shortcut

        self._debug_shortcut_sync(
            "Script-target fallback no match for '%s'; tried: %s"
            % (script_target, ", ".join(unique_candidates)),
            key="script_target_nomatch:%s" % script_target,
        )
        return None

    @objc.python_method
    def _menu_item_shortcut_string(self, menu_item, path_tokens=None):
        try:
            key = menu_item.keyEquivalent() or ""
            mask = menu_item.keyEquivalentModifierMask()
        except Exception:
            self._debug_shortcut_sync("menu_item_shortcut_string failed reading keyEquivalent")
            return None

        if not key:
            self._debug_shortcut_sync(
                "Menu item has empty keyEquivalent; using NSUserKeyEquivalents fallback for %s"
                % (" -> ".join(path_tokens) if path_tokens else "(unknown path)")
            )
            return self._shortcut_from_user_defaults_for_menu_path(path_tokens)

        special_chars = {
            " ": "Space",
            "\t": "Tab",
            "\r": "Return",
            "\n": "Return",
            "\x7f": "Delete",
            "\x08": "Delete",
        }

        display_key = None
        if len(key) == 1:
            codepoint = ord(key)
            display_key = _FUNCTION_KEY_CHARS.get(codepoint)
            if display_key is None:
                display_key = special_chars.get(key, key.upper())
        elif key:
            display_key = key.upper()

        if not display_key:
            self._debug_shortcut_sync(
                "Menu item key '%s' did not map directly; using NSUserKeyEquivalents fallback for %s"
                % (key, " -> ".join(path_tokens) if path_tokens else "(unknown path)")
            )
            return self._shortcut_from_user_defaults_for_menu_path(path_tokens)

        prefix = ""
        if mask & _MOD_CONTROL:
            prefix += "⌃"
        if mask & _MOD_OPTION:
            prefix += "⌥"
        if mask & _MOD_SHIFT:
            prefix += "⇧"
        if mask & _MOD_COMMAND:
            prefix += "⌘"
        shortcut = prefix + display_key
        self._debug_shortcut_sync(
            "Menu item direct keyEquivalent resolved as '%s' (path: %s)"
            % (shortcut, " -> ".join(path_tokens) if path_tokens else "(unknown path)")
        )
        return shortcut

    @objc.python_method
    def _filter_menu_item_for_name(self, filter_name):
        if not filter_name:
            return None

        menus_to_check = []
        if FILTER_MENU is not None:
            try:
                menus_to_check.append(Glyphs.menu[FILTER_MENU])
            except Exception:
                pass

        if not menus_to_check:
            try:
                for menu_key in Glyphs.menu:
                    menus_to_check.append(Glyphs.menu[menu_key])
            except Exception:
                return None

        for menu in menus_to_check:
            for menu_item in self._iter_menu_items(menu):
                try:
                    title = (menu_item.title() or "").strip()
                except Exception:
                    continue
                if title == filter_name:
                    return menu_item
        return None

    @objc.python_method
    def _script_menu_item_for_target(self, script_target):
        if not script_target:
            return None
        self._lastScriptMenuPathTokens = None

        normalized_target = script_target.replace("\\", "/")
        basename = os.path.basename(normalized_target)
        if basename.endswith(".py"):
            basename = basename[:-3]
        if not basename:
            return None

        target_tokens = []
        for part in normalized_target.split("/"):
            part = part.strip()
            if not part:
                continue
            if part.endswith(".py"):
                part = part[:-3]
            token = self._normalized_script_match_token(part)
            if token:
                target_tokens.append(token)
        if not target_tokens:
            return None

        basename_token = self._normalized_script_match_token(basename)
        menu_title = self._script_menu_title_for_target(script_target)
        menu_title_token = self._normalized_script_match_token(menu_title) if menu_title else ""
        accepted_title_tokens = set([basename_token])
        if menu_title_token:
            accepted_title_tokens.add(menu_title_token)
        accepted_compact_tokens = set()
        for token in accepted_title_tokens:
            compact = self._compact_script_match_token(token)
            if compact:
                accepted_compact_tokens.add(compact)
        self._debug_shortcut_sync(
            "Script sync lookup target='%s', basename='%s', menuTitle='%s'"
            % (script_target, basename_token, menu_title_token or ""),
            key="script_lookup:%s" % script_target,
        )

        matches = []
        menus_to_check = []
        if SCRIPTS_MENU is not None:
            try:
                menus_to_check.append(Glyphs.menu[SCRIPTS_MENU])
            except Exception:
                pass

        if not menus_to_check:
            try:
                for menu_key in Glyphs.menu:
                    menus_to_check.append(Glyphs.menu[menu_key])
            except Exception:
                return None

        scanned_count = 0
        for menu in menus_to_check:
            try:
                iterator = self._iter_menu_items_with_path(menu)
                for menu_item, path_tokens in iterator:
                    scanned_count += 1
                    if not path_tokens:
                        continue
                    title = path_tokens[-1]
                    title_token = self._normalized_script_match_token(title)
                    title_compact = self._compact_script_match_token(title)
                    if (
                        title_token not in accepted_title_tokens
                        and title_compact not in accepted_compact_tokens
                    ):
                        continue
                    normalized_path = []
                    for token in path_tokens:
                        normalized_token = self._normalized_script_match_token(token)
                        if normalized_token:
                            normalized_path.append(normalized_token)
                    if not normalized_path:
                        continue
                    metadata_title_match = bool(menu_title_token and title_token == menu_title_token)
                    raw_path_tokens = [str(token).strip() for token in path_tokens if str(token).strip()]
                    matches.append((menu_item, normalized_path, raw_path_tokens, metadata_title_match))
            except Exception as e:
                self._debug_shortcut_sync("Menu scan exception while matching script '%s': %s" % (script_target, e))
                continue

        self._debug_shortcut_sync(
            "Script sync scanned %d menu items for target '%s'"
            % (scanned_count, script_target),
            key="script_scanned:%s" % script_target,
        )

        if not matches:
            self._debug_shortcut_sync("No menu matches found for script target '%s'" % script_target)
            return None
        if len(matches) == 1:
            self._lastScriptSyncDiagnosticKey = None
            self._lastScriptMenuPathTokens = matches[0][2]
            self._debug_shortcut_sync(
                "Unique menu match for '%s': %s"
                % (script_target, " -> ".join(matches[0][2])),
                key="script_unique:%s" % script_target,
            )
            return matches[0][0]

        # Best-effort disambiguation: prefer unique candidates whose menu path
        # most closely matches the saved script path. If confidence is low, skip.
        ranked = []
        target_folder_tokens = set(target_tokens[:-1])
        for menu_item, menu_tokens, raw_path_tokens, metadata_title_match in matches:
            suffix_len = self._common_suffix_length(menu_tokens, target_tokens)
            folder_hits = 0
            for token in menu_tokens[:-1]:
                if token in target_folder_tokens:
                    folder_hits += 1
            metadata_score = 1 if metadata_title_match else 0
            ranked.append((metadata_score, suffix_len, folder_hits, len(menu_tokens), menu_item, raw_path_tokens))

        ranked.sort(key=lambda row: (row[0], row[1], row[2], row[3]), reverse=True)
        best = ranked[0]
        best_key = (best[0], best[1], best[2], best[3])
        tied_count = 0
        for candidate in ranked:
            if (candidate[0], candidate[1], candidate[2], candidate[3]) == best_key:
                tied_count += 1
        if tied_count > 1:
            self._log_script_sync_diagnostic_once(script_target, "ambiguous menu matches")
            self._debug_shortcut_sync("Ambiguous top-ranked menu matches for '%s'" % script_target)
            return None

        # Require matching signal beyond basename-only equality.
        if best[1] <= 1 and best[2] == 0:
            self._log_script_sync_diagnostic_once(script_target, "low confidence")
            self._debug_shortcut_sync("Low-confidence script menu match for '%s'" % script_target)
            return None

        self._lastScriptSyncDiagnosticKey = None
        self._lastScriptMenuPathTokens = best[5]
        self._debug_shortcut_sync(
            "Selected menu match for '%s': %s"
            % (script_target, " -> ".join(best[5])),
            key="script_selected:%s" % script_target,
        )
        return best[4]

    @objc.python_method
    def _glyphs_shortcut_for_target(self, item_type, target):
        if item_type == "filter":
            menu_item = self._filter_menu_item_for_name(target)
            if menu_item is None:
                self._debug_shortcut_sync("Filter sync lookup failed for '%s'" % target)
                return None
            return self._menu_item_shortcut_string(menu_item)

        if item_type == "script":
            menu_item = self._script_menu_item_for_target(target)
            if menu_item is None:
                fallback_shortcut = self._shortcut_from_user_defaults_for_script_target(target)
                if fallback_shortcut:
                    self._debug_shortcut_sync(
                        "Script sync resolved '%s' via script-target fallback -> '%s'"
                        % (target, fallback_shortcut),
                        key="script_resolved_target_fallback:%s" % target,
                    )
                    return fallback_shortcut
                self._debug_shortcut_sync("Script sync lookup failed for '%s'" % target)
                return None
            shortcut = self._menu_item_shortcut_string(menu_item, path_tokens=self._lastScriptMenuPathTokens)
            self._debug_shortcut_sync(
                "Script sync resolved '%s' -> '%s'"
                % (target, shortcut or "<none>"),
                key="script_resolved_menu:%s" % target,
            )
            return shortcut

        return None

    @objc.python_method
    def _add_dialog_shortcut_sync_candidate(self):
        if self.addDialog is None:
            return None
        mode = self.addDialog.typeMode.get()
        if mode == 0:
            target = self.addDialog.targetPopup.getItem()
            if not target or target == "(No items found)":
                return None
            return self._glyphs_shortcut_for_target("filter", target)
        if mode == 1 and self.addDialogSelectedScript:
            return self._glyphs_shortcut_for_target("script", self.addDialogSelectedScript)
        return None

    @objc.python_method
    def _edit_dialog_shortcut_sync_candidate(self):
        if self.editDialog is None:
            return None
        item_type = getattr(self.editDialog, "currentType", "script")
        if item_type == "filter":
            target = self.editDialog.targetPopup.getItem()
            if not target or target == "(No items found)":
                return None
            return self._glyphs_shortcut_for_target("filter", target)
        if item_type == "script" and self.editDialogSelectedScript:
            return self._glyphs_shortcut_for_target("script", self.editDialogSelectedScript)
        return None

    @objc.python_method
    def _update_add_dialog_glyphs_shortcut_button(self):
        if self.addDialog is None or not hasattr(self.addDialog, "shortcutUseGlyphsButton"):
            return
        shortcut = self._add_dialog_shortcut_sync_candidate()
        self.addDialog.shortcutUseGlyphsButton.enable(bool(shortcut))

    @objc.python_method
    def _update_add_dialog_shortcut_clear_button(self):
        if self.addDialog is None or not hasattr(self.addDialog, "shortcutClearButton"):
            return
        value = self._normalized_shortcut_value(self.addDialog.shortcutEdit.get())
        self.addDialog.shortcutClearButton.enable(bool(value))

    @objc.python_method
    def _update_edit_dialog_glyphs_shortcut_button(self):
        if self.editDialog is None or not hasattr(self.editDialog, "shortcutUseGlyphsButton"):
            return
        shortcut = self._edit_dialog_shortcut_sync_candidate()
        self.editDialog.shortcutUseGlyphsButton.enable(bool(shortcut))

    @objc.python_method
    def _update_edit_dialog_shortcut_clear_button(self):
        if self.editDialog is None or not hasattr(self.editDialog, "shortcutClearButton"):
            return
        value = self._normalized_shortcut_value(self.editDialog.shortcutEdit.get())
        self.editDialog.shortcutClearButton.enable(bool(value))

    @objc.python_method
    def _sync_add_shortcut_from_glyphs(self, sender):
        if self.addDialog is None:
            return
        shortcut = self._add_dialog_shortcut_sync_candidate()
        if shortcut:
            self.addDialog.shortcutEdit.set(shortcut)
        self._update_add_dialog_shortcut_clear_button()
        self._update_add_dialog_glyphs_shortcut_button()

    @objc.python_method
    def _sync_edit_shortcut_from_glyphs(self, sender):
        if self.editDialog is None:
            return
        shortcut = self._edit_dialog_shortcut_sync_candidate()
        if shortcut:
            self.editDialog.shortcutEdit.set(shortcut)
        self._update_edit_dialog_shortcut_clear_button()
        self._update_edit_dialog_glyphs_shortcut_button()

    @objc.python_method
    def _close_edit_dialog(self, sender):
        self._close_action_step_dialog(sender)
        self._stop_shortcut_capture()
        if self.editDialog is not None:
            self.editDialog.close()
            self.editDialog = None

    @objc.python_method
    def _edit_dialog_closed(self, sender):
        self._stop_shortcut_capture()
        self.editDialog = None

    @objc.python_method
    def _confirm_edit_dialog(self, sender):
        index = getattr(self.editDialog, "currentIndex", None)
        if index is None or index < 0 or index >= len(self.items):
            self._close_edit_dialog(sender)
            return

        item_type = getattr(self.editDialog, "currentType", "script")
        if item_type == "filter":
            target = self.editDialog.targetPopup.getItem()
            if not target or target == "(No items found)":
                return
            actions = []
            continue_on_error = False
        elif item_type == "script":
            target = self.editDialogSelectedScript
            if not target:
                return
            actions = []
            continue_on_error = False
        else:
            if len(self.editDialogActionSteps) < 2:
                self._set_status_message("Action buttons need at least 2 steps.")
                return
            target = ""
            actions = list(self.editDialogActionSteps)
            continue_on_error = bool(self.editDialog.actionContinueOnError.get())

        new_category = self.editDialog.categoryEdit.get().strip()
        new_name = self.editDialog.nameEdit.get().strip()
        if not new_name:
            if item_type == "script":
                new_name = self._script_display_title_for_target(target)
            elif item_type == "action":
                new_name = "Action (%d steps)" % len(actions)
            else:
                new_name = target

        shortcut = self._normalized_shortcut_value(self.editDialog.shortcutEdit.get())
        if shortcut and not self._shortcut_has_required_modifier(shortcut):
            self._show_shortcut_validation_error()
            return
        conflict_name = self._find_shortcut_conflict_name(shortcut, ignore_index=index)
        if conflict_name:
            self._show_shortcut_conflict_alert(shortcut, conflict_name)
            return

        item = self.items[index]
        item["category"] = new_category
        item["name"] = new_name
        item["type"] = item_type
        item["target"] = target
        item["actions"] = actions
        item["continueOnError"] = continue_on_error
        item["shortcut"] = shortcut

        self._save_items()
        self._refresh_ui()
        self.window.listView.setSelection([index])
        self._close_edit_dialog(sender)

    @objc.python_method
    def _format_shortcut_from_event(self, event):
        """Return a shortcut string like '⌘⇧A' from an NSEvent, or None if not usable."""
        try:
            flags = event.modifierFlags()
            chars = event.charactersIgnoringModifiers()
            key_code = event.keyCode()
        except Exception:
            return None

        special_chars = {
            " ": "Space",
            "\t": "Tab",
            "\r": "Return",
            "\n": "Return",
            "\x7f": "Delete",
            "\x08": "Delete",
        }

        display_key = _FUNCTION_KEY_NAMES.get(key_code)
        if display_key is None:
            # Non-function keys still require at least one modifier to avoid accidental triggers.
            if not (flags & _MOD_ANY):
                return None
            if not chars:
                return None
            char = chars[0]
            display_key = special_chars.get(char, char.upper())
        if not display_key:
            return None
        # Standard macOS modifier order: ⌃⌥⇧⌘
        prefix = ""
        if flags & _MOD_CONTROL:
            prefix += "⌃"
        if flags & _MOD_OPTION:
            prefix += "⌥"
        if flags & _MOD_SHIFT:
            prefix += "⇧"
        if flags & _MOD_COMMAND:
            prefix += "⌘"
        return prefix + display_key

    @objc.python_method
    def _handle_key_event(self, event):
        """NSEvent local monitor handler — fires assigned shortcuts when this window is key."""
        if self.shortcutCaptureOwner is not None:
            return event
        if not self.items or not self._is_main_window_valid():
            return event
        try:
            if NSApp.keyWindow() != self.window._window:
                return event
        except Exception:
            return event
        shortcut_str = self._format_shortcut_from_event(event)
        if not shortcut_str:
            return event
        for idx, item in enumerate(self.items):
            if item.get("shortcut", "") == shortcut_str:
                self._run_item(idx)
                return None  # consume — prevent the event reaching anything else
        return event

    @objc.python_method
    def _normalized_shortcut_value(self, value):
        if value is None:
            return ""
        return str(value).strip()

    @objc.python_method
    def _shortcut_has_required_modifier(self, shortcut):
        if shortcut in _FUNCTION_KEY_NAMES.values():
            return True
        return any(symbol in shortcut for symbol in ("⌘", "⌃", "⌥", "⇧"))

    @objc.python_method
    def _find_shortcut_conflict_name(self, shortcut, ignore_index=None):
        if not shortcut:
            return None
        for i, item in enumerate(self.items):
            if ignore_index is not None and i == ignore_index:
                continue
            if item.get("shortcut", "") == shortcut:
                return item.get("name", "another button")
        return None

    @objc.python_method
    def _show_shortcut_validation_error(self):
        alert = NSAlert.alloc().init()
        alert.setMessageText_("Invalid Shortcut")
        alert.setInformativeText_(
            "Use either a function key (F1-F20) or include at least one modifier: "
            "⌘ (Command), ⌃ (Control), ⌥ (Option), or ⇧ (Shift)."
        )
        alert.addButtonWithTitle_("OK")
        alert.runModal()

    @objc.python_method
    def _show_shortcut_conflict_alert(self, shortcut, conflict_name):
        alert = NSAlert.alloc().init()
        alert.setMessageText_("Shortcut Conflict")
        alert.setInformativeText_(
            "The shortcut %s is already assigned to '%s'.\n"
            "Use a different combination or clear the existing assignment." % (shortcut, conflict_name)
        )
        alert.addButtonWithTitle_("OK")
        alert.runModal()

    @objc.python_method
    def _record_add_shortcut(self, sender):
        if self.addDialog is None:
            return
        self.shortcutCaptureOwner = "add"
        self.addDialog.shortcutRecordButton.setTitle("Recording...")

    @objc.python_method
    def _clear_add_shortcut(self, sender):
        if self.addDialog is None:
            return
        self.addDialog.shortcutEdit.set("")
        self._update_add_dialog_shortcut_clear_button()
        self._stop_shortcut_capture()

    @objc.python_method
    def _record_edit_shortcut(self, sender):
        if self.editDialog is None:
            return
        self.shortcutCaptureOwner = "edit"
        self.editDialog.shortcutRecordButton.setTitle("Recording...")

    @objc.python_method
    def _clear_edit_shortcut(self, sender):
        if self.editDialog is None:
            return
        self.editDialog.shortcutEdit.set("")
        self._update_edit_dialog_shortcut_clear_button()
        self._stop_shortcut_capture()

    @objc.python_method
    def _stop_shortcut_capture(self):
        if self.addDialog is not None and hasattr(self.addDialog, "shortcutRecordButton"):
            self.addDialog.shortcutRecordButton.setTitle("Record")
        if self.editDialog is not None and hasattr(self.editDialog, "shortcutRecordButton"):
            self.editDialog.shortcutRecordButton.setTitle("Record")
        self.shortcutCaptureOwner = None

    @objc.python_method
    def _record_shortcut_key_event(self, event):
        if self.shortcutCaptureOwner not in ("add", "edit"):
            return event

        dialog = self.addDialog if self.shortcutCaptureOwner == "add" else self.editDialog
        if dialog is None:
            self._stop_shortcut_capture()
            return event

        try:
            if NSApp.keyWindow() != dialog._window:
                return event
        except Exception:
            return event

        try:
            if event.keyCode() == 53:  # Escape — cancel
                self._stop_shortcut_capture()
                return None
        except Exception:
            return event

        shortcut_str = self._format_shortcut_from_event(event)
        if shortcut_str:
            dialog.shortcutEdit.set(shortcut_str)
            if self.shortcutCaptureOwner == "add":
                self._update_add_dialog_shortcut_clear_button()
            else:
                self._update_edit_dialog_shortcut_clear_button()
            self._stop_shortcut_capture()
            return None

        return None
