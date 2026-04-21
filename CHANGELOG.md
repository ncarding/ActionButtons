# Changelog

This project uses semantic versioning during beta with `0.x.y-beta.n` release tags.

## Unreleased

- Added documentation disclaimer: plugin is shared as-is, built for personal use, and should be tested on copies/unimportant files first.

- Added empty-state safeguards:
	- `Select` is disabled when no buttons remain.
	- Reopening after delete-all reseeds the default `Instructions` button.
- Updated Instructions seed/edit behavior:
	- Instructions is represented as a script-style pseudo target.
	- Instructions script chooser shows a dedicated top-level `Instructions` entry.
- Refined Add/Edit dialog copy and controls:
	- Shortcut import label changed to `Use Existing`.
	- Shortcut independence note updated to possessive copy.
	- `Category` and `Shortcut` placeholders now show `Optional`.
	- Shortcut `Clear` button is disabled when no value is set.
	- In `Action` mode, step controls are positioned above the steps list.
	- Step control buttons now enable/disable by selection and top/bottom boundaries.
	- `Add` is disabled until required inputs are valid by type.
- Updated About window:
	- Added author line: `Neil Carding`.
	- Added website link: `www.neilcarding.co.uk`.

## 0.2.0-beta.1

- Established the first explicit beta version for Action Buttons.
- Added unified Add/Edit dialogs, shortcut recording, and Glyphs shortcut import.
- Added script metadata-title support and stronger script shortcut sync fallback logic.
- Added automated pytest coverage for shortcut and script-sync logic.
- Refreshed project documentation for the current UI and release workflow.