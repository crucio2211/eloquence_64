# Eloquence for NVDA (Forked)

This is a fork of [fastfinge/eloquence_64](https://github.com/fastfinge/eloquence_64), the Eloquence synthesizer addon for NVDA with 64-bit support.

This fork adds several improvements on top of the original, focused on making pronunciation dictionaries easier to manage and making the synthesizer more responsive.

---

## What is new in this fork

### Pronunciation Dictionary Editor

A built-in dictionary editor is now accessible directly from NVDA Settings under the Eloquence category. Click the **Edit Pronunciation Dictionary** button to open it.

**Features:**

- **File selector** — lists all .dic files found in the eloquence folder so you can switch between them without leaving the editor
- **Real-time search** — filters entries as you type with no lag, even with 65,000 or more entries, using a virtual list that only renders visible rows
- **Add, Edit, and Remove entries** — manage words and their pronunciations directly
- **Import** — import entries from another dictionary file, with a destination chooser so you can send entries to any existing .dic file or create a new one, regardless of which file is currently open
- **Export** — export the current dictionary to a file
- **Auto-reload on save** — after saving or importing, a confirmation dialog appears. Once you dismiss it, the synthesizer reloads automatically so changes take effect immediately without restarting NVDA
- **Two pronunciation formats supported:**
  - Phoneme format: `` `[.1hE.0lo] ``
  - Spelled-out format: `heh loe` (used for abbreviations and acronyms)

### Word Variation Generator

When you select an entry in the dictionary editor, an **Add Variations** button appears. This opens a variation generator dialog that helps you quickly add related word forms from a root entry.

- Auto-generates common prefixed and suffixed forms of the selected root word
- All suggestions are unchecked by default so you choose only what you want
- Each variation shows its word and phoneme side by side
- Phoneme of each variation can be reviewed and edited individually before adding
- Custom variations can be added through a separate dialog window
- Space bar or click toggles the checkbox for each variation

### IPC Responsiveness Improvement

The synthesizer now uses fire-and-forget IPC calls for commands that do not require a response, such as addText, insertIndex, and temporary prosody changes. This significantly reduces latency during navigation and typing, making the addon feel as responsive as other synthesizers.

### Audio Device Switching Fix

Fixed a bug where switching audio output devices (for example from Bluetooth headphones to laptop speakers) would not take effect until NVDA was restarted. The synthesizer now properly reinitializes when the audio device changes.

---

## Installation

1. Download the latest `.nvda-addon` file from the [Releases](../../releases) page
2. Open it with NVDA to install
3. Restart NVDA when prompted

---

## Reporting bugs

If you find a bug, please [open an Issue](../../issues) and include:
- A description of what happened
- Steps to reproduce
- Your NVDA version
- Your Windows version
- The NVDA log file if possible (`NVDA menu > Tools > View log`)

## Contributing

Pull requests are welcome. If you want to propose a change, please open an Issue first to discuss it so we can make sure it fits the direction of the fork.

---

## Credits

- Original addon by [fastfinge](https://github.com/fastfinge) and contributors
- Improvements in this fork by crucio2211
