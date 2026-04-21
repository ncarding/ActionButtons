# Action Buttons

Action Buttons is a Glyphs 3 plugin that lets you save frequently used filters, scripts, and multi-step action sequences as reusable buttons.

Current beta version: `0.2.0-beta.1`

Public README refresh: 2026-04-21.

## Disclaimer

This plugin was initially built for personal use and is shared as-is. It has not yet been fully tested.

Please back up your files and test on copies first. If you choose to use it, you do so at your own risk.


## Features

- Three button types:
  - `Filter`
  - `Script`
  - `Action`
- Two views:
  - `List`
  - `Grid`
- Adjustable grid sizing with optional `Compact Height` mode
- Unified `Edit...` flow for all button types
- Shortcut recording inside Add/Edit dialogs
- `Use Existing` shortcut import for supported targets
- Ordered multi-step Action buttons with optional `Continue on error`
- Dedicated utility menu for import/export, instructions, and about
- Versioned `.actionbuttons` export/import format with duplicate detection
- Explicit opt-in anonymous telemetry for run outcomes and duration
- Persistent button order and view mode

Add/Edit dialog behavior highlights:
- `Category` and `Shortcut` are optional fields
- `Add` is enabled only when required inputs are valid:
  - `Filter`: a valid filter is selected
  - `Script`: a script is selected
  - `Action`: at least 2 steps are configured
- In `Action` mode, step controls are shown above the step list
- `Remove Step` / `Step Up` / `Step Down` are disabled until a step is selected
- `Step Up` and `Step Down` are disabled at the top/bottom list boundaries

## Installation

Current manual install (temporary until package manager onboarding):

1. Copy `ActionButtons.glyphsPlugin` to:

```text
~/Library/Application Support/Glyphs 3/Plugins/
```

2. Quit Glyphs completely.
3. Reopen Glyphs.
4. Open the plugin from the `Window` menu as `Action Buttons`.

Package-manager install instructions will be added once the plugin is listed in Glyphs packages.

## Main Window

The main window includes:
- `Select` and selection controls
- `+` Add button
- Grid buttons for quick execution
- Settings menu in the title bar (`⚙`)
- Utility menu in the title bar for `Export…`, `Import…`, `Instructions…`, and `About…`

Icon-only controls show tooltips, including the title-bar buttons and selection-mode action buttons.

Button sizing options:
- `Small`, `Medium`, `Large`
- `Compact Height` (optional): reduces tile height only when `Show Category` and/or `Show Shortcut` is hidden
- `Compact Height` cannot be selected while both `Show Category` and `Show Shortcut` are enabled
- If previously enabled, `Compact Height` state is remembered and reapplied automatically when either line is hidden again

The settings menu contains layout, button size, button content, run behavior, and a privacy telemetry toggle.
`Instructions…` and `About…` live in the separate utility menu.

List view columns include:
- Type icon (`F`, `S`, `A`)
- Name
- Type
- Shortcut
- Target

## Button Types

### Filter

Stores a Glyphs filter menu item by title.

### Script

Stores a relative script path from your Glyphs Scripts folder.

Display names prefer script metadata `# MenuTitle:` when available and fall back to a prettified filename.

### Action

An Action button is a sequence of at least two steps.

Each step can be:
- a `Filter`
- a `Script`

You can reorder steps and optionally enable `Continue on error`.

Friendly heads-up: not every script and filter combination works well together, so please test new combinations on a copy or unimportant file first.

## Shortcuts

Shortcuts are local to Action Buttons and are stored on each button.

Rules:
- `F1` to `F20` can be used on their own
- all other keys require at least one modifier
- supported modifiers:
  - `⌘`
  - `⌃`
  - `⌥`
  - `⇧`

Action Buttons prevents duplicate shortcuts inside the plugin.

Important:
- Shortcuts only trigger while the Action Buttons window is key
- Clicking in the content area focuses the window

## Glyphs Shortcut Import

The Add/Edit dialog includes a `Use Existing` button next to the shortcut field.

Behavior:
- Always visible
- Disabled when no Glyphs shortcut can be determined for the current target
- One-time import only
- No write-back to Glyphs settings

Notes:
- Filter sync is based on the matching Glyphs menu item
- Script sync uses best-effort matching against menu titles and macOS key-equivalent storage
- Some script setups may remain ambiguous and therefore unsyncable by design

The dialog also shows this reminder:

`Action Buttons' shortcuts are independent from Glyphs' shortcuts.`

## Script Discovery

Scripts are discovered from:
- `~/Library/Application Support/Glyphs 3/Scripts`
- fallback: `~/Library/Application Support/Glyphs/Scripts`

Aliased folders are supported.

Displayed script names:
- prefer `# MenuTitle:` metadata when present
- otherwise use a prettified filename/path token

## Import and Export

Use the utility menu in the title bar to access `Export…` and `Import…`.

Export behavior:
- `Export…` opens a dedicated selection dialog instead of reusing Select mode
- You can export any subset of current buttons
- Export files use JSON with the custom `.actionbuttons` extension
- Exported data includes button definitions only, not cog-menu preferences such as grid columns, button size, compact height, or run-on-double-click

Import behavior:
- `Import…` validates the selected `.actionbuttons` file before changing saved buttons
- Script buttons are matched by saved script path relative to the Glyphs Scripts root
- Filter buttons are matched by filter target
- Action buttons are matched by ordered step signature plus `Continue on error`
- If duplicates are detected, one global alert lists every duplicate pair as `Existing name > Imported name`
- You can choose `Skip` for the whole duplicate set or `Import as Copy` for the whole duplicate set
- Imported buttons always get fresh internal IDs
- If an imported button name collides with an existing name, the imported name is suffixed as `Name 2`, `Name 3`, and so on until unique
- If an imported shortcut conflicts with an existing Action Buttons shortcut, the imported shortcut is cleared and reported
- Missing scripts or filters do not block import, but they are reported after import completes

## Telemetry

Telemetry is disabled by default and uses explicit opt-in.

- First launch prompt offers `Share Anonymous Data` or `Not Now`
- You can change telemetry at any time from the settings menu: `Share Anonymous Usage Data`
- Captured events focus on plugin usage only: run success/failure, run duration, import/export counts, and settings changes
- No script source, file contents, glyph data, or local file paths are sent

## Troubleshooting

### The plugin does not appear in Glyphs

1. Restart Glyphs fully
2. Confirm the bundle exists in:

```text
~/Library/Application Support/Glyphs 3/Plugins/ActionButtons.glyphsPlugin
```

3. If missing, copy `ActionButtons.glyphsPlugin` into that Plugins folder and restart Glyphs.

### A script button does not run

- Confirm the script still exists in the Scripts folder
- Confirm the saved relative path still matches the file location
- Check Macro Panel for Python errors from the script itself

### `Use Existing` is disabled for a script

- The script may have no configured Glyphs shortcut
- The menu title may be ambiguous
- The underlying shortcut may not be exposed directly by the menu item

The plugin uses a fallback lookup for many script shortcuts, but not every setup can be resolved safely.

## Versioning

This plugin is still in beta.

Versioning approach:
- Public beta tags use semantic versioning with a beta suffix, for example `0.2.0-beta.1`
- Bundle marketing version in `Info.plist` stays numeric for compatibility
- Bundle build number increments on each packaged milestone
- Human-readable beta tag is stored in the project `VERSION` file

When making the next beta update:
1. Increment `VERSION`.
2. Update `CFBundleShortVersionString` if the numeric release changes.
3. Increment `CFBundleVersion`.
4. Add a new entry to `CHANGELOG.md`.

## Development

Internal development and operations notes are intentionally kept in local-only documentation.