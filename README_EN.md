# EngineeringChangeMarker — Word Macro Template for Marking Document Revisions

A Word macro template (.dotm) that automatically marks and highlights edited content in a document when it is submitted. It leaves a clear "change trail" for every version, making proofreading and version comparison straightforward.

## Features

- **Automatic change marking**: scans the active document and marks locations where content has been modified
- **Highlighting of changes**: edited text is highlighted so reviewers can spot revisions at a glance
- **Proofreading & comparison support**: works together with the Review pane to compare before/after content
- **Built for engineering change scenarios**: suitable for drawing notes, technical documents, and design files that require strict version control

## Installation

The recommended approach is to install the template as a **global Word add-in**. The macro template is loaded globally by Word, so the macros become available in every document — no need to import them into each file. Install once, use everywhere.

### Method 1 (recommended): install as a global add-in

1. Download `EngineeringChangeMarker.dotm` from this repository.
2. Copy the file into Word's STARTUP folder:
   - Press `Win + R`, type `%APPDATA%\Microsoft\Word\STARTUP`, and press Enter.
   - Paste `EngineeringChangeMarker.dotm` into that folder.
   - If the folder does not exist, create it manually.
3. Close and reopen Word completely so the template is loaded automatically.
4. Verify the template is loaded: **File → Options → Add-ins** → select **Word Add-ins** in the "Manage" drop-down at the bottom → click **Go**. The template should appear in the list.
5. If the macros are disabled: right-click the `.dotm` file → **Properties** → check **Unblock**, then reopen Word.

Once installed, the macros are available in all Word documents (including ordinary .docx files) — no per-document setup is needed.

### Method 2: use the user templates folder

Copy the `.dotm` file to Word's templates folder (`%APPDATA%\Microsoft\Templates`) and create new documents from it via **File → New**. Documents created this way carry the macros, but this only affects new documents, not existing ones.

### Method 3: temporary single-document use

Double-click `EngineeringChangeMarker.dotm` to open and use it directly, or import the module via the VBA editor (`Alt + F11`). This only affects the current document and is fine for a quick trial, but not recommended as a long-term setup.

## Usage

1. Open the Word document you want to mark.
2. Press `Alt + F8` to open the Macro dialog, select the macro, and click **Run**.
3. After the macro finishes, the modified parts of the document are highlighted automatically.
4. To compare versions, use the Review tab to inspect tracked changes and comments.

> Note: The exact macro name depends on the actual code. Check it before running.

## Notes

- Save a copy of the document before running the macro, in case the result cannot be undone.
- With the global add-in setup, the macro acts on the active document; in a macro-enabled document (.docm) the behavior is identical to an embedded macro.
- This macro is written in VBA and targets Microsoft Word for Windows.

## Compatibility

- Supported: Microsoft Word 2016 / 2019 / Microsoft 365 (Windows)
- Note: VBA support in Word for Mac is limited; compatibility issues may occur.

## License

All rights reserved by default. If you want to allow others to use it freely, add a LICENSE file (e.g., MIT).
