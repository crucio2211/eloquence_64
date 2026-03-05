# Copyright (C) 2009-2019 eloquence fans
# synthDrivers/eci.py
# todo: possibly add to this
import gui
import wx
import ctypes
import winsound
import shutil  # Added for Copy Helper tool

try:
	from speech import (
		IndexCommand,
		CharacterModeCommand,
		LangChangeCommand,
		BreakCommand,
		PitchCommand,
		RateCommand,
		VolumeCommand,
		PhonemeCommand,
	)
except ImportError:
	from speech.commands import (
		IndexCommand,
		CharacterModeCommand,
		LangChangeCommand,
		BreakCommand,
		PitchCommand,
		RateCommand,
		VolumeCommand,
		PhonemeCommand,
	)

try:
	from driverHandler import NumericDriverSetting, BooleanDriverSetting, DriverSetting
except ImportError:
	from autoSettingsUtils.driverSetting import (
		BooleanDriverSetting,
		DriverSetting,
		NumericDriverSetting,
	)

try:
	from autoSettingsUtils.utils import StringParameterInfo
except ImportError:

	class StringParameterInfo:
		def __init__(self, value, label):
			self.value = value
			self.label = label


punctuation = ",.?:;)(?!"
punctuation = [x for x in punctuation]
from ctypes import *
import ctypes.wintypes
import synthDriverHandler
import os
import config
import re
import logging
import globalVars
from synthDriverHandler import (
	SynthDriver,
	synthIndexReached,
	synthDoneSpeaking,
)
from . import _eloquence
from . import _text_preprocessing
from collections import OrderedDict
import unicodedata
import addonHandler

addonHandler.initTranslation()

log = logging.getLogger(__name__)


minRate = 40
maxRate = 150
pause_re = re.compile(r"([a-zA-Z0-9]|\s)([,.:;?!)])(\2*?)(\s|[\\/]|$|$)")
time_re = re.compile(r"(\d):(\d+):(\d+)")
VOICE_BCP47 = {
	"enu": "en-US",
	"eng": "en-GB",
	"esp": "es-ES",
	"esm": "es-419",
	"ptb": "pt-BR",
	"fra": "fr-FR",
	"frc": "fr-CA",
	"deu": "de-DE",
	"ita": "it-IT",
	"fin": "fi-FI",
	"chs": "zh-CN",  # Simplified Chinese
	"jpn": "ja-JP",  # Japanese
	"kor": "ko-KR",  # Korean
}

VOICE_CODE_TO_ID = {code: str(info[0]) for code, info in _eloquence.langs.items()}
VOICE_ID_TO_BCP47 = {
	voice_id: VOICE_BCP47.get(code) for code, voice_id in VOICE_CODE_TO_ID.items() if VOICE_BCP47.get(code)
}
LANGUAGE_TO_VOICE_ID = {
	lang.lower(): VOICE_CODE_TO_ID[code] for code, lang in VOICE_BCP47.items() if code in VOICE_CODE_TO_ID
}
PRIMARY_LANGUAGE_TO_VOICE_IDS = {}
for code, lang in VOICE_BCP47.items():
	voice_id = VOICE_CODE_TO_ID.get(code)
	if not voice_id:
		continue
	primary = lang.split("-", 1)[0].lower()
	PRIMARY_LANGUAGE_TO_VOICE_IDS.setdefault(primary, []).append(voice_id)

variants = {
	1: "Reed",
	2: "Shelley",
	3: "Bobby",
	4: "Rocko",
	5: "Glen",
	6: "Sandy",
	7: "Grandma",
	8: "Grandpa",
}


class _VirtualList(wx.ListCtrl):
	"""wx.ListCtrl subclass that supports virtual mode via OnGetItemText override."""
	def __init__(self, parent, **kwargs):
		super().__init__(parent, style=wx.LC_REPORT | wx.LC_SINGLE_SEL | wx.LC_VIRTUAL | wx.BORDER_SUNKEN, **kwargs)
		self._data_source = None  # callable(row, col) -> str

	def set_data_source(self, fn):
		self._data_source = fn

	def OnGetItemText(self, item, col):
		if self._data_source:
			return self._data_source(item, col)
		return ""


class DictionaryEditorDialog(wx.Dialog):
	"""Dialog for editing Eloquence pronunciation dictionary files.
	Uses a virtual ListCtrl so 65k+ entries never hang the UI.
	"""

	def __init__(self, parent):
		super().__init__(
			parent,
			title=_("Eloquence Pronunciation Dictionary Editor"),
			style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
		)
		self._dic_dir = os.path.join(os.path.dirname(__file__), "eloquence")
		self._entries = []        # full list of [word, phonetic]
		self._filtered = []       # current view (after search)
		self._current_file = None
		self._modified = False
		self._build_ui()
		self._populate_file_choice()

	# ------------------------------------------------------------------
	def _build_ui(self):
		main_sizer = wx.BoxSizer(wx.VERTICAL)
		inner = wx.BoxSizer(wx.VERTICAL)

		# File selector
		fs = wx.BoxSizer(wx.HORIZONTAL)
		fs.Add(wx.StaticText(self, label=_("Dictionary file:")), 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
		self._file_choice = wx.Choice(self, choices=[])
		self._file_choice.Bind(wx.EVT_CHOICE, self._on_file_change)
		fs.Add(self._file_choice, 1, wx.EXPAND)
		inner.Add(fs, 0, wx.EXPAND | wx.BOTTOM, 8)

		# Search
		ss = wx.BoxSizer(wx.HORIZONTAL)
		ss.Add(wx.StaticText(self, label=_("Search:")), 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
		self._search_ctrl = wx.SearchCtrl(self, style=wx.TE_PROCESS_ENTER)
		self._search_ctrl.ShowCancelButton(True)
		self._search_ctrl.Bind(wx.EVT_TEXT, self._on_search)
		self._search_ctrl.Bind(wx.EVT_SEARCHCTRL_CANCEL_BTN, self._on_search_cancel)
		ss.Add(self._search_ctrl, 1, wx.EXPAND)
		inner.Add(ss, 0, wx.EXPAND | wx.BOTTOM, 8)

		# Virtual ListCtrl — only renders visible rows, no lag with 65k entries
		self._list = _VirtualList(self)
		self._list.InsertColumn(0, _("Word"), width=200)
		self._list.InsertColumn(1, _("Pronunciation"), width=340)
		self._list.set_data_source(self.OnGetItemText)
		self._list.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self._on_edit)
		self._list.Bind(wx.EVT_LIST_ITEM_SELECTED, self._on_sel_changed)
		self._list.Bind(wx.EVT_LIST_ITEM_DESELECTED, self._on_sel_changed)
		inner.Add(self._list, 1, wx.EXPAND | wx.BOTTOM, 6)

		# Entry count
		self._count_label = wx.StaticText(self, label="")
		inner.Add(self._count_label, 0, wx.BOTTOM, 4)

		# Action buttons
		bs = wx.BoxSizer(wx.HORIZONTAL)
		self._add_btn    = wx.Button(self, label=_("&Add"))
		self._var_btn    = wx.Button(self, label=_("Add &Variations..."))
		self._edit_btn   = wx.Button(self, label=_("&Edit"))
		self._remove_btn = wx.Button(self, label=_("&Remove"))
		self._import_btn = wx.Button(self, label=_("&Import..."))
		self._export_btn = wx.Button(self, label=_("E&xport..."))
		self._save_btn   = wx.Button(self, label=_("&Save"))

		self._edit_btn.Disable()
		self._var_btn.Disable()
		self._remove_btn.Disable()

		self._add_btn.Bind(wx.EVT_BUTTON, self._on_add)
		self._var_btn.Bind(wx.EVT_BUTTON, self._on_variations)
		self._edit_btn.Bind(wx.EVT_BUTTON, self._on_edit)
		self._remove_btn.Bind(wx.EVT_BUTTON, self._on_remove)
		self._import_btn.Bind(wx.EVT_BUTTON, self._on_import)
		self._export_btn.Bind(wx.EVT_BUTTON, self._on_export)
		self._save_btn.Bind(wx.EVT_BUTTON, self._on_save)

		for btn in (self._add_btn, self._var_btn, self._edit_btn,
					self._remove_btn, self._import_btn, self._export_btn, self._save_btn):
			bs.Add(btn, 0, wx.RIGHT, 4)
		inner.Add(bs, 0, wx.BOTTOM, 8)

		main_sizer.Add(inner, 1, wx.EXPAND | wx.ALL, 10)
		close_btn = wx.Button(self, wx.ID_CLOSE, label=_("Close"))
		close_btn.Bind(wx.EVT_BUTTON, self._on_close)
		main_sizer.Add(close_btn, 0, wx.ALIGN_RIGHT | wx.RIGHT | wx.BOTTOM, 10)

		self.SetSizer(main_sizer)
		self.SetSize((660, 520))
		self.Centre()

	def OnGetItemText(self, item, col):
		if item < 0 or item >= len(self._filtered):
			return ""
		return self._filtered[item][col]

	# ------------------------------------------------------------------
	def _populate_file_choice(self):
		choices, self._dic_paths = [], []
		try:
			for fname in sorted(os.listdir(self._dic_dir)):
				if fname.lower().endswith(".dic"):
					choices.append(fname)
					self._dic_paths.append(os.path.join(self._dic_dir, fname))
		except Exception as e:
			log.error(f"Failed to list dic files: {e}")
		self._file_choice.SetItems(choices)
		if choices:
			self._file_choice.SetSelection(0)
			self._load_file(self._dic_paths[0])

	def _load_file(self, path):
		try:
			self._current_file = path
			self._entries = []
			with open(path, "r", encoding="iso-8859-1", errors="replace") as f:
				for line in f:
					line = line.rstrip("\r\n")
					if line and "\t" in line:
						parts = line.split("\t", 1)
						if len(parts) == 2:
							self._entries.append([parts[0], parts[1]])
			self._modified = False
			self._apply_filter()
		except Exception as e:
			log.error(f"Failed to load {path}: {e}")
			wx.MessageBox(_("Failed to load file: {e}").format(e=str(e)), _("Error"), wx.OK | wx.ICON_ERROR)

	def _apply_filter(self, search=""):
		"""Rebuild _filtered and update virtual list item count."""
		q = search.lower().strip()
		if q:
			self._filtered = [e for e in self._entries if q in e[0].lower() or q in e[1].lower()]
		else:
			self._filtered = list(self._entries)
		self._list.SetItemCount(len(self._filtered))
		self._list.Refresh()
		total = len(self._entries)
		shown = len(self._filtered)
		if q:
			self._count_label.SetLabel(_("{shown} of {total} entries").format(shown=shown, total=total))
		else:
			self._count_label.SetLabel(_("{total} entries").format(total=total))
		self._on_sel_changed()

	def _get_selected_filtered_idx(self):
		return self._list.GetFirstSelected()

	def _get_entry_idx_from_filtered(self, filtered_idx):
		"""Map a filtered row index back to the master _entries index."""
		if filtered_idx < 0 or filtered_idx >= len(self._filtered):
			return -1
		entry = self._filtered[filtered_idx]
		try:
			return self._entries.index(entry)
		except ValueError:
			return -1

	# ------------------------------------------------------------------
	def _on_file_change(self, evt):
		if self._modified and not self._confirm_unsaved():
			return
		sel = self._file_choice.GetSelection()
		if 0 <= sel < len(self._dic_paths):
			self._load_file(self._dic_paths[sel])

	def _on_search(self, evt):
		self._apply_filter(self._search_ctrl.GetValue())

	def _on_search_cancel(self, evt):
		self._search_ctrl.SetValue("")
		self._apply_filter()

	def _on_sel_changed(self, evt=None):
		has = self._list.GetFirstSelected() != -1
		self._edit_btn.Enable(has)
		self._var_btn.Enable(has)
		self._remove_btn.Enable(has)

	# ------------------------------------------------------------------
	def _on_add(self, evt):
		dlg = DictionaryEntryDialog(self, _("Add Entry"))
		if dlg.ShowModal() == wx.ID_OK:
			word, phonetic = dlg.get_values()
			if word:
				self._entries.append([word, phonetic])
				self._modified = True
				self._apply_filter(self._search_ctrl.GetValue())
		dlg.Destroy()

	def _on_edit(self, evt=None):
		fi = self._get_selected_filtered_idx()
		if fi == -1:
			return
		ei = self._get_entry_idx_from_filtered(fi)
		if ei == -1:
			return
		word, phonetic = self._entries[ei]
		dlg = DictionaryEntryDialog(self, _("Edit Entry"), word, phonetic)
		if dlg.ShowModal() == wx.ID_OK:
			nw, np = dlg.get_values()
			if nw:
				self._entries[ei] = [nw, np]
				self._modified = True
				self._apply_filter(self._search_ctrl.GetValue())
		dlg.Destroy()

	def _on_remove(self, evt):
		fi = self._get_selected_filtered_idx()
		if fi == -1:
			return
		ei = self._get_entry_idx_from_filtered(fi)
		if ei == -1:
			return
		word = self._entries[ei][0]
		dlg = wx.MessageDialog(
			self,
			_("Remove entry for '{word}'?").format(word=word),
			_("Confirm Remove"),
			wx.YES_NO | wx.ICON_QUESTION,
		)
		if dlg.ShowModal() == wx.ID_YES:
			del self._entries[ei]
			self._modified = True
			self._apply_filter(self._search_ctrl.GetValue())
		dlg.Destroy()

	def _on_variations(self, evt):
		fi = self._get_selected_filtered_idx()
		if fi == -1:
			return
		ei = self._get_entry_idx_from_filtered(fi)
		if ei == -1:
			return
		word, phonetic = self._entries[ei]
		existing_words = {e[0].lower() for e in self._entries}
		dlg = VariationGeneratorDialog(self, word, phonetic, existing_words)
		if dlg.ShowModal() == wx.ID_OK:
			new_entries = dlg.get_entries()
			if new_entries:
				self._entries.extend(new_entries)
				self._modified = True
				self._apply_filter(self._search_ctrl.GetValue())
				wx.MessageBox(
					_("Added {n} variation(s).").format(n=len(new_entries)),
					_("Done"),
					wx.OK | wx.ICON_INFORMATION,
				)
		dlg.Destroy()

	def _on_import(self, evt):
		# Step 1: Pick the source file to import from
		src_dlg = wx.FileDialog(
			self, _("Import Dictionary File"),
			wildcard=_("Dictionary files (*.dic)|*.dic|Text files (*.txt)|*.txt|All files (*.*)|*.*"),
			style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST,
		)
		if src_dlg.ShowModal() != wx.ID_OK:
			src_dlg.Destroy()
			return
		src_path = src_dlg.GetPath()
		src_dlg.Destroy()

		# Read source entries first
		try:
			imported = []
			with open(src_path, "r", encoding="iso-8859-1", errors="replace") as f:
				for line in f:
					line = line.rstrip("\r\n")
					if line and "\t" in line:
						parts = line.split("\t", 1)
						if len(parts) == 2:
							imported.append([parts[0], parts[1]])
		except Exception as e:
			wx.MessageBox(_("Failed to read file: {e}").format(e=str(e)), _("Error"), wx.OK | wx.ICON_ERROR)
			return

		if not imported:
			wx.MessageBox(_("No valid entries found in the selected file."), _("Import"), wx.OK | wx.ICON_INFORMATION)
			return

		# Step 2: Ask where to put the imported entries
		existing_files = [os.path.basename(p) for p in self._dic_paths]
		choices = existing_files + [_("Create new .dic file...")]
		dest_dlg = wx.SingleChoiceDialog(
			self,
			_("Found {n} entries. Where do you want to import them?").format(n=len(imported)),
			_("Import Destination"),
			choices,
		)
		# Pre-select current file
		if self._current_file:
			cur_name = os.path.basename(self._current_file)
			if cur_name in existing_files:
				dest_dlg.SetSelection(existing_files.index(cur_name))
		if dest_dlg.ShowModal() != wx.ID_OK:
			dest_dlg.Destroy()
			return
		sel = dest_dlg.GetSelection()
		dest_dlg.Destroy()

		# Step 3: Determine destination path
		if sel == len(existing_files):
			# Create new file
			new_dlg = wx.FileDialog(
				self, _("Create New Dictionary File"),
				defaultDir=self._dic_dir,
				wildcard=_("Dictionary files (*.dic)|*.dic"),
				style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT,
			)
			if new_dlg.ShowModal() != wx.ID_OK:
				new_dlg.Destroy()
				return
			dest_path = new_dlg.GetPath()
			if not dest_path.lower().endswith(".dic"):
				dest_path += ".dic"
			new_dlg.Destroy()
			dest_entries = []
			dest_existing = set()
		else:
			dest_path = self._dic_paths[sel]
			# Load destination if it's not the current file
			if dest_path == self._current_file:
				dest_entries = self._entries
				dest_existing = {e[0].lower() for e in self._entries}
			else:
				dest_entries = []
				dest_existing = set()
				try:
					with open(dest_path, "r", encoding="iso-8859-1", errors="replace") as f:
						for line in f:
							line = line.rstrip("\r\n")
							if line and "\t" in line:
								parts = line.split("\t", 1)
								if len(parts) == 2:
									dest_entries.append([parts[0], parts[1]])
									dest_existing.add(parts[0].lower())
				except Exception as e:
					wx.MessageBox(_("Failed to read destination: {e}").format(e=str(e)), _("Error"), wx.OK | wx.ICON_ERROR)
					return

		# Step 4: Merge and save
		added = 0
		for entry in imported:
			if entry[0].lower() not in dest_existing:
				dest_entries.append(entry)
				dest_existing.add(entry[0].lower())
				added += 1

		try:
			with open(dest_path, "w", encoding="iso-8859-1", errors="replace", newline="\r\n") as f:
				for word, phonetic in dest_entries:
					f.write(f"{word}\t{phonetic}\n")
		except Exception as e:
			wx.MessageBox(_("Failed to save: {e}").format(e=str(e)), _("Error"), wx.OK | wx.ICON_ERROR)
			return

		# If we wrote to the current file, refresh the view
		if dest_path == self._current_file:
			self._entries = dest_entries
			self._modified = False
			self._apply_filter(self._search_ctrl.GetValue())
		elif dest_path not in self._dic_paths:
			# New file — add to dropdown and switch to it
			self._dic_paths.append(dest_path)
			self._file_choice.Append(os.path.basename(dest_path))

		wx.MessageBox(
			_("Imported {n} new entries into {f}. Press OK to apply changes.").format(n=added, f=os.path.basename(dest_path)),
			_("Import Complete"), wx.OK | wx.ICON_INFORMATION,
		)
		# User already pressed OK — reinit now, speech is idle
		try:
			synthDriverHandler.setSynth(synthDriverHandler.getSynth().name)
		except Exception as e:
			log.error(f"Failed to reload synth after import: {e}", exc_info=True)


	def _on_export(self, evt):
		dlg = wx.FileDialog(
			self, _("Export Dictionary File"),
			wildcard=_("Dictionary files (*.dic)|*.dic|Text files (*.txt)|*.txt"),
			style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT,
		)
		if dlg.ShowModal() == wx.ID_OK:
			path = dlg.GetPath()
			try:
				with open(path, "w", encoding="iso-8859-1", errors="replace", newline="\r\n") as f:
					for word, phonetic in self._entries:
						f.write(f"{word}\t{phonetic}\n")
				wx.MessageBox(_("Exported {n} entries.").format(n=len(self._entries)), _("Export Complete"), wx.OK | wx.ICON_INFORMATION)
			except Exception as e:
				wx.MessageBox(_("Export failed: {e}").format(e=str(e)), _("Error"), wx.OK | wx.ICON_ERROR)
		dlg.Destroy()

	def _save_current(self):
		if not self._current_file:
			return
		try:
			with open(self._current_file, "w", encoding="iso-8859-1", errors="replace", newline="\r\n") as f:
				for word, phonetic in self._entries:
					f.write(f"{word}\t{phonetic}\n")
			self._modified = False
			wx.MessageBox(_("Dictionary saved. Press OK to apply changes."), _("Saved"), wx.OK | wx.ICON_INFORMATION)
			# User already pressed OK — reinit now, speech is idle
			try:
				synthDriverHandler.setSynth(synthDriverHandler.getSynth().name)
			except Exception as e:
				log.error(f"Failed to reload synth after dictionary save: {e}", exc_info=True)
		except Exception as e:
			wx.MessageBox(_("Failed to save: {e}").format(e=str(e)), _("Error"), wx.OK | wx.ICON_ERROR)

	def _on_save(self, evt):
		self._save_current()

	def _confirm_unsaved(self):
		"""Ask user what to do with unsaved changes. Returns True to proceed, False to cancel."""
		dlg = wx.MessageDialog(
			self,
			_("You have unsaved changes. Save before continuing?"),
			_("Unsaved Changes"),
			wx.YES_NO | wx.CANCEL | wx.ICON_QUESTION,
		)
		result = dlg.ShowModal()
		dlg.Destroy()
		if result == wx.ID_YES:
			self._save_current()
			return True
		elif result == wx.ID_NO:
			return True
		return False  # Cancel

	def _on_close(self, evt):
		if self._modified and not self._confirm_unsaved():
			return
		self.EndModal(wx.ID_OK)


class VariationGeneratorDialog(wx.Dialog):
	"""Generate word variations from a root word + phoneme.
	Uses a two-column ListCtrl: Word | Phoneme.
	Space bar or click toggles the checkbox (state image).
	Double-click or Edit button opens phoneme editor.
	"""

	PREFIXES = [
		("i-",    "I"),
		("in-",   "In"),
		("nag-",  "nAg"),
		("mag-",  "mAg"),
		("nang-", "nAG"),
		("mang-", "mAG"),
		("pag-",  "pAg"),
		("na-",   "nA"),
		("ma-",   "mA"),
		("ka-",   "kA"),
		("sang-", "sAG"),
	]
	SUFFIXES = [
		("-an",   "An"),
		("-in",   "In"),
		("-ng",   "G"),
		("-han",  "hAn"),
		("-hin",  "hIn"),
	]
	INFIXES = [
		("um",    "um"),
		("in",    "in"),
	]

	def __init__(self, parent, root_word, root_phoneme, existing_words):
		super().__init__(
			parent,
			title=_("Generate Variations for: {word}").format(word=root_word),
			style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
		)
		self._root_word = root_word
		self._root_phoneme = root_phoneme
		self._existing = existing_words
		self._variations = []  # [word, phoneme, checked]
		self._build_ui()
		self._generate_suggestions()

	def _build_ui(self):
		sizer = wx.BoxSizer(wx.VERTICAL)

		info = wx.StaticText(
			self,
			label=_("Root: {word}   Phoneme: {ph}\n"
					"Space/click = toggle. Double-click or Edit = change phoneme.").format(
				word=self._root_word, ph=self._root_phoneme
			)
		)
		sizer.Add(info, 0, wx.ALL, 10)

		# ListCtrl with native checkbox support via EnableCheckBoxes()
		self._list = wx.ListCtrl(self, style=wx.LC_REPORT | wx.LC_SINGLE_SEL | wx.BORDER_SUNKEN)
		self._list.EnableCheckBoxes(True)
		self._list.InsertColumn(0, _("Word"), width=200)
		self._list.InsertColumn(1, _("Phoneme"), width=300)
		self._list.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self._on_dbl_click)
		self._list.Bind(wx.EVT_KEY_DOWN, self._on_key_down)
		sizer.Add(self._list, 1, wx.EXPAND | wx.LEFT | wx.RIGHT, 10)

		# Buttons row
		btn_row = wx.BoxSizer(wx.HORIZONTAL)
		edit_btn = wx.Button(self, label=_("&Edit Phoneme"))
		all_btn  = wx.Button(self, label=_("Select &All"))
		none_btn = wx.Button(self, label=_("Select &None"))
		edit_btn.Bind(wx.EVT_BUTTON, self._on_edit)
		all_btn.Bind(wx.EVT_BUTTON,  lambda e: self._set_all(True))
		none_btn.Bind(wx.EVT_BUTTON, lambda e: self._set_all(False))
		for b in (edit_btn, all_btn, none_btn):
			btn_row.Add(b, 0, wx.RIGHT, 6)
		sizer.Add(btn_row, 0, wx.LEFT | wx.TOP | wx.BOTTOM, 10)

		# Add custom variation button — opens a separate dialog
		add_custom_btn = wx.Button(self, label=_("+ Add Custom Variation..."))
		add_custom_btn.Bind(wx.EVT_BUTTON, self._on_add_custom)
		sizer.Add(add_custom_btn, 0, wx.LEFT | wx.BOTTOM, 10)

		# OK / Cancel
		std = wx.StdDialogButtonSizer()
		ok_btn = wx.Button(self, wx.ID_OK, label=_("Add Checked"))
		ok_btn.SetDefault()
		std.AddButton(ok_btn)
		std.AddButton(wx.Button(self, wx.ID_CANCEL))
		std.Realize()
		sizer.Add(std, 0, wx.ALIGN_RIGHT | wx.ALL, 10)

		self.SetSizer(sizer)
		self.SetSize((580, 520))
		self.Centre()

	# ------------------------------------------------------------------
	@staticmethod
	def _strip_phoneme(ph):
		"""Remove surrounding `[ and ] from phoneme string, return inner content."""
		ph = ph.strip()
		if ph.startswith("`[") and ph.endswith("]"):
			return ph[2:-1]
		return ph

	@staticmethod
	def _strip_phoneme(ph):
		ph = ph.strip()
		if ph.startswith("`[") and ph.endswith("]"):
			return ph[2:-1]
		return ph

	@staticmethod
	def _wrap_phoneme(inner):
		return f"`[{inner}]"

	def _combine(self, prefix_inner, root_ph, suffix_inner=""):
		root_inner = self._strip_phoneme(root_ph)
		return self._wrap_phoneme(prefix_inner + root_inner + suffix_inner)

	def _generate_suggestions(self):
		word, phoneme = self._root_word, self._root_phoneme
		for label, ph in self.PREFIXES:
			nw = label.replace("-", "") + word
			if nw.lower() not in self._existing:
				self._variations.append([nw, self._combine(ph, phoneme), False])
		for label, ph in self.SUFFIXES:
			nw = word + label.replace("-", "")
			if nw.lower() not in self._existing:
				self._variations.append([nw, self._combine("", phoneme, ph), False])
		for ph_infix in self.INFIXES:
			nw = word[0] + ph_infix[0] + word[1:]
			if nw.lower() not in self._existing:
				self._variations.append([nw, self._combine(ph_infix[1], phoneme), False])
		self._rebuild()

	def _rebuild(self):
		self._list.DeleteAllItems()
		for i, (word, phoneme, checked) in enumerate(self._variations):
			idx = self._list.InsertItem(i, word)
			self._list.SetItem(idx, 1, phoneme)
			self._list.CheckItem(idx, checked)

	def _toggle(self, idx):
		if idx < 0 or idx >= len(self._variations):
			return
		new_state = not self._list.IsItemChecked(idx)
		self._variations[idx][2] = new_state
		self._list.CheckItem(idx, new_state)

	def _on_key_down(self, evt):
		if evt.GetKeyCode() == wx.WXK_SPACE:
			idx = self._list.GetFirstSelected()
			if idx != wx.NOT_FOUND:
				self._toggle(idx)
		else:
			evt.Skip()

	def _on_dbl_click(self, evt):
		self._open_editor(evt.GetIndex())

	def _on_edit(self, evt):
		idx = self._list.GetFirstSelected()
		if idx == wx.NOT_FOUND:
			wx.MessageBox(_("Select a variation first."), _("Edit Phoneme"), wx.OK | wx.ICON_INFORMATION)
			return
		self._open_editor(idx)

	def _open_editor(self, idx):
		if idx < 0 or idx >= len(self._variations):
			return
		word, phoneme, checked = self._variations[idx]
		dlg = DictionaryEntryDialog(self, _("Edit Variation"), word, phoneme)
		if dlg.ShowModal() == wx.ID_OK:
			nw, np = dlg.get_values()
			if nw:
				self._variations[idx] = [nw, np, checked]
				self._list.SetItem(idx, 0, nw)
				self._list.SetItem(idx, 1, np)
		dlg.Destroy()

	def _on_add_custom(self, evt):
		dlg = DictionaryEntryDialog(
			self,
			_("Add Custom Variation"),
			phonetic=self._root_phoneme,
		)
		while dlg.ShowModal() == wx.ID_OK:
			word, phoneme = dlg.get_values()
			if not word:
				wx.MessageBox(_("Please enter a word."), _("Error"), wx.OK | wx.ICON_WARNING)
				continue
			self._variations.append([word, phoneme, True])
			idx = self._list.InsertItem(len(self._variations) - 1, word)
			self._list.SetItem(idx, 1, phoneme)
			self._list.CheckItem(idx, True)
			# Reset for next entry
			dlg._word_ctrl.SetValue("")
			dlg._phonetic_ctrl.SetValue(self._root_phoneme)
			dlg._word_ctrl.SetFocus()
			# Ask if user wants to add another
			more = wx.MessageDialog(
				self,
				_("Variation '{word}' added. Add another?").format(word=word),
				_("Added"),
				wx.YES_NO | wx.ICON_QUESTION,
			)
			if more.ShowModal() != wx.ID_YES:
				more.Destroy()
				break
			more.Destroy()
		dlg.Destroy()

	def _set_all(self, state):
		for i in range(len(self._variations)):
			self._variations[i][2] = state
			self._list.CheckItem(i, state)

	def get_entries(self):
		return [[w, ph] for i, (w, ph, _) in enumerate(self._variations)
				if self._list.IsItemChecked(i)]


class DictionaryEntryDialog(wx.Dialog):
	"""Small dialog for adding or editing a single dictionary entry."""

	def __init__(self, parent, title, word="", phonetic=""):
		super().__init__(parent, title=title, style=wx.DEFAULT_DIALOG_STYLE)
		sizer = wx.BoxSizer(wx.VERTICAL)

		word_sizer = wx.BoxSizer(wx.HORIZONTAL)
		word_label = wx.StaticText(self, label=_("Word:"))
		self._word_ctrl = wx.TextCtrl(self, value=word)
		word_sizer.Add(word_label, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
		word_sizer.Add(self._word_ctrl, 1, wx.EXPAND)
		sizer.Add(word_sizer, 0, wx.EXPAND | wx.ALL, 10)

		phonetic_sizer = wx.BoxSizer(wx.HORIZONTAL)
		phonetic_label = wx.StaticText(self, label=_("Pronunciation:"))
		self._phonetic_ctrl = wx.TextCtrl(self, value=phonetic)
		phonetic_sizer.Add(phonetic_label, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
		phonetic_sizer.Add(self._phonetic_ctrl, 1, wx.EXPAND)
		sizer.Add(phonetic_sizer, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

		hint = wx.StaticText(self, label=_(
			"Phoneme format: `[.1hE.0lo]   or   Spelled-out: heh loe"
		))
		hint.SetForegroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_GRAYTEXT))
		sizer.Add(hint, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

		btn_sizer = wx.StdDialogButtonSizer()
		ok_btn = wx.Button(self, wx.ID_OK)
		ok_btn.SetDefault()
		cancel_btn = wx.Button(self, wx.ID_CANCEL)
		btn_sizer.AddButton(ok_btn)
		btn_sizer.AddButton(cancel_btn)
		btn_sizer.Realize()
		sizer.Add(btn_sizer, 0, wx.ALIGN_RIGHT | wx.ALL, 10)

		self.SetSizer(sizer)
		self.Fit()
		self.Centre()
		self._word_ctrl.SetFocus()

	def get_values(self):
		return self._word_ctrl.GetValue().strip(), self._phonetic_ctrl.GetValue().strip()


class EloquenceSettingsPanel(gui.settingsDialogs.SettingsPanel):
	# Translators: Name of the category for this add-on in the settings dialog
	title = _("Eloquence")

	def makeSettings(self, settings):
		try:
			sHelper = gui.guiHelper.BoxSizerHelper(self, sizer=settings)

			self.dictionarySources = {
				"https://github.com/mohamed00/AltIBMTTSDictionaries": "Alternative IBM TTS Dictionaries",
				"https://github.com/eigencrow/IBMTTSDictionaries": "IBM TTS Dictionaries",
			}

			self.dictionaryChoice = sHelper.addLabeledControl(
				# Translators: Label of a combobox in the Eloquence category of the settings dialog
				_("Dictionary:"), wx.Choice, choices=list(self.dictionarySources.values())
			)
			self.dictionaryChoice.SetStringSelection(
				config.conf.get("eloquence", {}).get("dictionary_name", "Alternative IBM TTS Dictionaries")
			)

			self.updateButton = sHelper.addItem(wx.Button(self, label=_("Check for updates")))
			self.Bind(wx.EVT_BUTTON, self.onUpdate, self.updateButton)
			# When NVDA is running in secure mode, one should not be able to save any setting to disk.
			if globalVars.appArgs.secure:
				self.updateButton.Disable()

			# Tool to automate copying eloquence_host32.exe for 64-bit NVDA secure screens
			self.copyHelperButton = sHelper.addItem(
				# Translators: Label of a button in the Eloquence category of the settings dialog
				wx.Button(self, label=_("Copy Helper to System Config (for Logon Screen)"))
			)
			self.Bind(wx.EVT_BUTTON, self.onCopyHelper, self.copyHelperButton)
			# Copying helper from secure mode is not allowed, following "Use NVDA during sign-in" button behaviour in
			# NVDA's General settings category.
			if globalVars.appArgs.secure:
				self.copyHelperButton.Disable()
				

			# NEW: Auto-update addon button
			# Translators: Label of a button in the Eloquence category of the settings dialog
			self.addonUpdateButton = sHelper.addItem(wx.Button(self, label=_("Check for Add-on Updates")))
			self.Bind(wx.EVT_BUTTON, self.onCheckAddonUpdate, self.addonUpdateButton)
			# Add-on updates are not allowed in secure mode.
			if globalVars.appArgs.secure:
				self.addonUpdateButton.Disable()

			# Dictionary editor button
			self.dictEditorButton = sHelper.addItem(wx.Button(self, label=_("Edit Pronunciation Dictionary...")))
			self.Bind(wx.EVT_BUTTON, self.onEditDictionary, self.dictEditorButton)
			if globalVars.appArgs.secure:
				self.dictEditorButton.Disable()
		except Exception as e:
			log.error(f"Error creating Eloquence settings panel: {e}")
			# Panel creation failed, but don't crash - synth will still work

	def onEditDictionary(self, evt):
		"""Open the pronunciation dictionary editor dialog."""
		try:
			dlg = DictionaryEditorDialog(self)
			dlg.ShowModal()
			dlg.Destroy()
		except Exception as e:
			log.error(f"Error opening dictionary editor: {e}", exc_info=True)
			wx.MessageBox(
				_("Failed to open dictionary editor: {e}").format(e=str(e)),
				_("Error"),
				wx.OK | wx.ICON_ERROR,
			)

	def onCopyHelper(self, evt):
		"""Copies eloquence_host32.exe with UAC elevation support and definitive feedback."""
		source_file = os.path.normpath(os.path.join(os.path.dirname(__file__), "eloquence_host32.exe"))
		prog_files = os.environ.get("ProgramFiles", "C:\\Program Files")
		target_addon_dir = os.path.normpath(
			os.path.join(prog_files, "NVDA", "systemConfig", "addons", "Eloquence")
		)

		# Security check: Ensure the target addon directory exists in systemConfig
		if not os.path.isdir(target_addon_dir):
			wx.MessageBox(
				_(
					# Translators: Text of a message dialog when copying the helper to system config
					"Eloquence folder not found in systemConfig.\n\nPlease go to NVDA Settings > General and click 'Use currently saved settings during sign-in' first to initialize folders."
				),
				# Translators: Title of a message dialog when copying the helper to system config
				_("Folder Missing"),
				wx.OK | wx.ICON_WARNING,
			)
			return

		dest_dir = os.path.normpath(os.path.join(target_addon_dir, "synthDrivers"))
		dest_file = os.path.normpath(os.path.join(dest_dir, "eloquence_host32.exe"))

		if not os.path.exists(source_file):
			wx.MessageBox(
				# Translators: Text of a message dialog when copying the helper to system config
				_("Source file not found at:\n{source_file}").format(source_file=source_file),
				# Translators: Title of a message dialog when copying the helper to system config
				_("Error"),
				wx.OK | wx.ICON_ERROR,
			)
			return

		# Prepare elevated command: ensure subdirectory exists and copy the helper
		cmd_params = f'/c mkdir "{dest_dir}" 2>nul & copy /y "{source_file}" "{dest_file}"'

		try:
			# Triggering UAC Elevation using ShellExecuteW's "runas" verb
			ret = ctypes.windll.shell32.ShellExecuteW(None, "runas", "cmd.exe", cmd_params, None, 0)

			if ret > 32:
				# Play Windows Asterisk sound for confirmation of successful launch
				winsound.MessageBeep(winsound.MB_ICONASTERISK)
				wx.MessageBox(
					_(
						# Translators: text of a message dialog when copying the helper to system config
						"Successfully copied eloquence_host32.exe to systemConfig!\n\nEloquence should now load normally on logon screen, start-up, and other secure screens."
					),
					# Translators: Title of a message dialog when copying the helper to system config
					_("Success"),
					wx.OK | wx.ICON_INFORMATION,
				)
			elif ret == 5:
				# SE_ERR_ACCESSDENIED: Elevation prompt was declined
				wx.MessageBox(
					# Translators: Text of a message dialog when copying the helper to system config
					_("Copy process was cancelled or permission was denied by the user."),
					# Translators: Title of a message dialog when copying the helper to system config
					_("Cancelled"),
					wx.OK | wx.ICON_ERROR,
				)
			else:
				wx.MessageBox(
					# Translators: Text of a message dialog when copying the helper to system config
					_("An error occurred while attempting to copy the file. (Error Code: {ret})").format(ret=ret),
					# Translators: Title of a message dialog when copying the helper to system config
					_("Error"),
					wx.OK | wx.ICON_ERROR,
				)
		except Exception as e:
			wx.MessageBox(
				# Translators: Text of a message dialog when copying the helper to system config
				_("An unexpected error occurred: {e}").format(e=str(e)),
				# Translators: Title of a message dialog when copying the helper to system config
				_("Error"),
				wx.OK | wx.ICON_ERROR,
			)

	def onCheckAddonUpdate(self, evt):
		"""Check for and apply addon updates from GitHub"""
		import sys
		import os

		# Import the update manager
		addon_dir = os.path.abspath(os.path.dirname(__file__))
		update_manager_path = os.path.join(addon_dir, "_eloquence_updater.py")

		# Check if updater exists
		if not os.path.exists(update_manager_path):
			wx.MessageBox(
				# Translators: Text of a message dialog when updating the add-on
				_("Update manager not found. Please reinstall the add-on."),
				# Translators: Title of a message dialog when updating the add-on
				_("Error"),
				wx.OK | wx.ICON_ERROR,
			)
			return

		# Import update manager
		sys.path.insert(0, addon_dir)
		try:
			from _eloquence_updater import EloquenceUpdateManager, show_update_dialog
		except ImportError as e:
			wx.MessageBox(
				# Translators: Text of a message dialog when updating the add-on
				_("Failed to load update manager: {e}").format(e=e),
				# Translators: Title of a message dialog when updating the add-on
				_("Error"),
				wx.OK | wx.ICON_ERROR,
			)
			return
		finally:
			if addon_dir in sys.path:
				sys.path.remove(addon_dir)

		# Create progress dialog
		progress = wx.ProgressDialog(
			# Translators: Title of a progress dialog when updating the add-on
			_("Checking for Updates"),
			# Translators: Message of a progress dialog when updating the add-on
			_("Connecting to GitHub..."),
			maximum=100,
			parent=self,
			style=wx.PD_APP_MODAL | wx.PD_AUTO_HIDE,
		)

		try:
			# Initialize update manager
			manager = EloquenceUpdateManager(addon_dir)

			# Check for updates
			# Translators: Message of a progress dialog when updating the add-on
			progress.Update(10, _("Checking for updates..."))
			(
				has_update,
				latest_version,
				download_url,
				changelog,
			) = manager.check_for_updates()

			if not has_update:
				# Translators: Message of a progress dialog when updating the add-on
				progress.Update(100, _("No updates available"))
				progress.Destroy()
				wx.MessageBox(
					# Translators: Text of a message dialog when updating the add-on
					_("You are using the latest version!"),
					# Translators: Title of a message dialog when updating the add-on
					_("Up to Date"),
					wx.OK | wx.ICON_INFORMATION,
				)
				return

			# Show changelog
			# Translators: Text of a message dialog when updating the add-on
			progress.Update(20, _("Update available!"))
			progress.Destroy()

			changelog_dialog = wx.MessageDialog(
				self,
				_(
					# Translators: Text of a message dialog when updating the add-on
					"New version available: {latest_version}\n\n"
					"Current version: {currVersion}\n\n"
					"Changelog:\n{changelog}\n\n"
					"Would you like to download and review the update?"
				).format(
					latest_version=latest_version,
					currVersion=manager.CURRENT_VERSION,
					changelog=changelog[:500],
				),
				# Translators: Title of a message dialog when updating the add-on
				_("Update Available"),
				wx.YES_NO | wx.ICON_INFORMATION,
			)

			if changelog_dialog.ShowModal() != wx.ID_YES:
				return

			# Download update
			progress = wx.ProgressDialog(		
				# Translators: Text of a progress dialog when updating the add-on
				_("Downloading Update"),
				# Translators: Title of a progress dialog when updating the add-on
				_("Downloading..."),
				maximum=100,
				parent=self,
				style=wx.PD_APP_MODAL | wx.PD_CAN_ABORT,
			)

			def download_progress(percent, message):
				cont, skip = progress.Update(percent, message)
				return cont

			zip_path = manager.download_update(download_url, download_progress)

			# Extract update
			# Translators: Text of a progress dialog when updating the add-on
			progress.Update(0, _("Extracting update..."))
			manager.extract_update(zip_path, download_progress)

			# Analyze changes
			# Translators: Text of a progress dialog when updating the add-on
			progress.Update(0, _("Analyzing changes..."))
			changes = manager.analyze_changes(download_progress)

			progress.Destroy()

			# Show update dialog with detailed changes
			apply_update, decisions = show_update_dialog(self, changes, latest_version)

			if not apply_update:
				manager.cleanup()
				wx.MessageBox(
					# Translators: Text of a message dialog when updating the add-on
					_("Update cancelled."),
					# Translators: Text of a message dialog when updating the add-on
					_("Cancelled"),
					wx.OK | wx.ICON_INFORMATION,
				)
				return

			# Apply update with progress
			progress = wx.ProgressDialog(
				# Translators: Text of a progress dialog when updating the add-on
				_("Applying Update"),
				# Translators: Text of a progress dialog when updating the add-on
				_("Please wait..."),
				maximum=100,
				parent=self,
				style=wx.PD_APP_MODAL | wx.PD_AUTO_HIDE,
			)

			def merge_progress(percent, message):
				progress.Update(percent, message)

			manager.smart_merge(changes, decisions, merge_progress)

			progress.Destroy()

			# Success!
			wx.MessageBox(
				_(
					# Translators: Text of a message dialog when updating the add-on
					"Update to {latest_version} applied successfully!\n\n"
					"Please restart NVDA for changes to take effect."
				).format(latest_version=latest_version),
				# Translators: Title of a message dialog when updating the add-on
				_("Update Successful"),
				wx.OK | wx.ICON_INFORMATION,
			)

			# Cleanup
			manager.cleanup()

		except Exception as e:
			progress.Destroy()
			log.error(f"Update failed: {e}")
			wx.MessageBox(
				_(
					# Translators: Text of a message dialog when updating the add-on
					"Update failed: {e}\n\n"
					"Your addon has not been modified."
				).format(e=str(e)),
				# Translators: Title of a message dialog when updating the add-on
				_("Update Failed"),
				wx.OK | wx.ICON_ERROR,
			)

	def onSave(self):
		if "eloquence" not in config.conf:
			config.conf["eloquence"] = {}
		selection = self.dictionaryChoice.GetStringSelection()
		for url, name in self.dictionarySources.items():
			if name == selection:
				config.conf["eloquence"]["dictionary_name"] = name
				config.conf["eloquence"]["dictionary_url"] = url
				break

	def onUpdate(self, evt):
		import urllib.request
		import zipfile
		import os

		self.onSave()
		dictionary_url = config.conf.get("eloquence", {}).get("dictionary_url")
		if not dictionary_url:
			wx.MessageBox(
				# Translators: Text of a message dialog when updating a dictionary
				_("Please select a dictionary first."),
				# Translators: Title of a message dialog when updating a dictionary
				_("Error"),
				wx.OK | wx.ICON_ERROR,
			)
			return

		try:
			# Add /archive/master.zip to the end of the URL to download the master branch
			zip_url = dictionary_url + "/archive/master.zip"
			zip_path, _unused = urllib.request.urlretrieve(zip_url)

			addon_dir = os.path.abspath(os.path.dirname(__file__))
			dest_folder = os.path.join(addon_dir, "eloquence")

			if not os.path.exists(dest_folder):
				os.makedirs(dest_folder)

			with zipfile.ZipFile(zip_path, "r") as zip_ref:
				zip_ref.extractall(addon_dir)
				zip_contents = zip_ref.namelist()
				extracted_root_name = zip_contents[0].split("/")[0]
				extracted_folder_path = os.path.join(addon_dir, extracted_root_name)

			updates_count = 0

			# --- HELPER: Ensure CP1252 compatibility ---
			def clean_key_text(text):
				try:
					text.encode("cp1252")
					return text
				except UnicodeEncodeError:
					# If not CP1252, fallback to stripping accents
					return "".join(
						c for c in unicodedata.normalize("NFD", text) if unicodedata.category(c) != "Mn"
					)

			# --- HELPER: Extract Key/Word only (Cleaned) ---
			def get_key(line):
				parts = line.strip().split(None, 1)
				if parts:
					raw_key = parts[0].lower()
					return clean_key_text(raw_key)  # Return CP1252-safe key
				return None

			# --- HELPER: Normalize Format (Space to Tab + Clean ALL text) ---
			def normalize_entry_format(line):
				line = line.strip()
				if " [" in line and "\t[" not in line:
					parts = line.split(" [", 1)
					if len(parts) == 2:
						word_part = parts[0].strip()
						pronunciation_part = parts[1]
						# Clean BOTH the word and pronunciation to ensure CP1252 compatibility
						clean_word = clean_key_text(word_part)
						clean_pronunciation = clean_key_text(pronunciation_part)
						return f"{clean_word}\t[{clean_pronunciation}"

				# Even if it's already tabbed, ensure ALL text is CP1252-safe
				if "\t[" in line:
					parts = line.split("\t[", 1)
					if len(parts) == 2:
						word_part = parts[0].strip()
						pronunciation_part = parts[1]
						# Clean BOTH parts
						clean_word = clean_key_text(word_part)
						clean_pronunciation = clean_key_text(pronunciation_part)
						return f"{clean_word}\t[{clean_pronunciation}"

				# If no bracket format, just clean the whole line
				return clean_key_text(line)

			# --- MAIN LOGIC ---
			if os.path.exists(extracted_folder_path):
				candidates = []
				for root, dirs, files in os.walk(extracted_folder_path):
					for f in files:
						if f.lower().endswith(".dic"):
							full_path = os.path.join(root, f)
							candidates.append((full_path, f))

				processed_filenames = set()
				encodings_to_try = ["utf-8", "cp1252", "iso-8859-1", "cp437"]

				for source_path, filename in candidates:
					dest_path = os.path.join(dest_folder, filename)

					# Auto-create new dictionary files with CP1252-safe content
					if not os.path.exists(dest_path):
						try:
							# Read source file with encoding detection
							source_lines = []
							read_success = False
							for enc in encodings_to_try:
								try:
									with open(source_path, "r", encoding=enc) as f:
										source_lines = f.readlines()
										read_success = True
										break
								except UnicodeDecodeError:
									continue

							if not read_success:
								with open(source_path, "r", encoding="iso-8859-1", errors="replace") as f:
									source_lines = f.readlines()

							# Process and strip accents from all lines
							processed_lines = []
							for line in source_lines:
								normalized_line = normalize_entry_format(line)
								if normalized_line.strip():  # Skip empty lines
									processed_lines.append(normalized_line)

							# Write as CP1252
							with open(dest_path, "w", encoding="cp1252") as f:
								for line in processed_lines:
									f.write(f"{line}\n")

							updates_count += len(processed_lines)
							log.info(
								f"Created new dictionary file: {filename} ({len(processed_lines)} entries, CP1252-safe)"
							)
						except Exception as e:
							log.error(f"Failed to create new dictionary {filename}: {e}")
						continue

					if filename.lower() in processed_filenames:
						continue
					processed_filenames.add(filename.lower())

					lines_to_append = []
					try:
						# 1. READ LOCAL: Extract CLEAN KEYS
						existing_keys = set()

						def load_local_keys(f_handle):
							for line in f_handle:
								key = get_key(line)
								if key:
									existing_keys.add(key)

						try:
							# Try CP1252 first as it is the standard for dictionaries
							with open(dest_path, "r", encoding="cp1252") as f:
								load_local_keys(f)
						except UnicodeDecodeError:
							# Fallback if it was previously written in a different encoding
							try:
								with open(dest_path, "r", encoding="utf-8") as f:
									load_local_keys(f)
							except UnicodeDecodeError:
								with open(dest_path, "r", encoding="mbcs", errors="ignore") as f:
									load_local_keys(f)

						# 2. READ SOURCE WITH AUTO-DETECT
						source_lines = []
						read_success = False
						for enc in encodings_to_try:
							try:
								with open(source_path, "r", encoding=enc) as f:
									source_lines = f.readlines()
									read_success = True
									break
							except UnicodeDecodeError:
								continue

						if not read_success:
							with open(source_path, "r", encoding="iso-8859-1", errors="replace") as f:
								source_lines = f.readlines()

						# 3. FILTER, CLEAN, & FORMAT
						for line in source_lines:
							# This cleans the visual word and normalizes spaces while preserving CP1252 accents
							normalized_line = normalize_entry_format(line)

							# Extract the clean key for comparison
							key = get_key(normalized_line)

							if not key:
								continue

							# Check duplicates using the key
							if key not in existing_keys:
								lines_to_append.append(normalized_line)
								existing_keys.add(key)

						# 4. WRITE UPDATES (Strictly CP1252)
						if lines_to_append:
							with open(dest_path, "a", encoding="cp1252") as f:
								f.write("\n")
								for item in lines_to_append:
									f.write(f"{item}\n")
							updates_count += len(lines_to_append)

					except Exception as e:
						log.error(f"Failed to merge dictionary {filename}: {e}")

				shutil.rmtree(extracted_folder_path)

			os.remove(zip_path)

			if updates_count > 0:
				# Count how many were new files vs updated entries
				new_files = sum(1 for f in os.listdir(dest_folder) if f.lower().endswith(".dic"))
				wx.MessageBox(
					_(
						# Translators: Text of a message dialog when updating a dictionary
						"Dictionary update successful!\n\n"
						"• Total updates: {updates_count}\n"
						"• Dictionary files: {new_files}\n\n"
						"Note: CP1252 encoding enforced; some accents may have been stripped for compatibility."
					).format(updates_count=updates_count, new_files=new_files),
					# Translators: Title of a message dialog when updating a dictionary
					_("Success"),
					wx.OK | wx.ICON_INFORMATION,
				)
			else:
				wx.MessageBox(
					# Translators: Text of a message dialog when updating a dictionary
					_("No new updates found. Your dictionaries are already up to date."),
					# Translators: Title of a message dialog when updating a dictionary
					_("Eloquence"),
					wx.OK | wx.ICON_INFORMATION,
				)

		except Exception as e:
			wx.MessageBox(
				# Translators: Text of a message dialog when updating a dictionary
				_("An error occurred while updating the dictionary: {e}").format(e=e),
				# Translators: Title of a message dialog when updating a dictionary
				_("Error"),
				wx.OK | wx.ICON_ERROR,
			)
		pass


class SynthDriver(synthDriverHandler.SynthDriver):
	settingsPanel = EloquenceSettingsPanel
	supportedSettings = (
		SynthDriver.VoiceSetting(),
		SynthDriver.VariantSetting(),
		SynthDriver.RateSetting(),
		SynthDriver.PitchSetting(),
		SynthDriver.InflectionSetting(),
		SynthDriver.VolumeSetting(),
		# Translators: A synth setting available in speech settings dialog
		NumericDriverSetting("hsz", _("Hea&d size")),
		# Translators: A synth setting available in speech settings dialog
		NumericDriverSetting("rgh", _("Rou&ghness")),
		# Translators: A synth setting available in speech settings dialog
		NumericDriverSetting("bth", _("Breathi&ness")),
		BooleanDriverSetting(
			# Translators: A synth setting available in speech settings dialog
			"backquoteVoiceTags", _("Enable backquote voice &tags"), True
		),
		# Translators: A synth setting available in speech settings dialog
		BooleanDriverSetting("ABRDICT", _("Enable &abbreviation dictionary"), False),
		# Translators: A synth setting available in speech settings dialog
		BooleanDriverSetting("phrasePrediction", _("Enable phras&e prediction"), False),
		# Translators: A synth setting available in speech settings dialog
		DriverSetting("pauseMode", _("&Pauses"), defaultVal="0"),
	)
	supportedCommands = {
		IndexCommand,
		CharacterModeCommand,
		LangChangeCommand,
		BreakCommand,
		PitchCommand,
		RateCommand,
		VolumeCommand,
		PhonemeCommand,
	}
	supportedNotifications = {synthIndexReached, synthDoneSpeaking}
	PROSODY_ATTRS = {
		PitchCommand: _eloquence.pitch,
		VolumeCommand: _eloquence.vlm,
		RateCommand: _eloquence.rate,
	}

	description = "ETI-Eloquence"
	name = "eloquence"

	# Initialize _pause_mode at class level to prevent issues with setting restoration
	_pause_mode = 0

	@classmethod
	def check(cls):
		try:
			log.info("Eloquence: Running check() to verify synth is available")
			result = _eloquence.eciCheck()
			log.info(f"Eloquence: check() returned {result}")
			return result
		except Exception as e:
			log.error(f"Eloquence: check() failed with error: {e}", exc_info=True)
			return False

	def __init__(self):
		# Safe settings panel registration - won't crash if API changes in different NVDA versions
		try:
			if hasattr(gui.settingsDialogs, "NVDASettingsDialog"):
				if hasattr(gui.settingsDialogs.NVDASettingsDialog, "categoryClasses"):
					if EloquenceSettingsPanel not in gui.settingsDialogs.NVDASettingsDialog.categoryClasses:
						gui.settingsDialogs.NVDASettingsDialog.categoryClasses.append(EloquenceSettingsPanel)
		except Exception as e:
			log.warning(f"Could not register Eloquence settings panel: {e}")
			# Continue initialization - synth will work without settings panel

		try:
			log.info("Eloquence: Starting initialization")
			_eloquence.initialize(self._onIndexReached)
			log.info("Eloquence: _eloquence.initialize completed successfully")
		except Exception as e:
			log.error(f"Eloquence: Failed to initialize _eloquence module: {e}", exc_info=True)
			raise

		try:
			voice_param = _eloquence.params.get(9)
			if voice_param is None:
				configured_voice = config.conf.get("speech", {}).get("eci", {}).get("voice", "enu")
				voice_info = _eloquence.langs.get(configured_voice) or _eloquence.langs.get("enu")
				voice_param = voice_info[0] if voice_info else 65536
			self._update_voice_state(voice_param, update_default=True)
			# Initialize _rate first before setting the rate property
			self._rate = self._percentToParam(50, minRate, maxRate)
			self.rate = 50
			self.variant = "1"
			self._pause_mode = 0
			log.info("Eloquence: Initialization completed successfully")
		except Exception as e:
			log.error(f"Eloquence: Failed during voice/parameter setup: {e}", exc_info=True)
			raise

	def terminate(self):
		# Shut down the eloquence backend so the WavePlayer is properly closed.
		# This ensures the next initialize_audio() creates a fresh player for
		# the currently configured audio device.
		try:
			_eloquence.terminate()
		except Exception as e:
			log.error(f"Eloquence: error terminating backend: {e}", exc_info=True)

		# Safe settings panel removal - won't crash if it was never registered
		try:
			if hasattr(gui.settingsDialogs, "NVDASettingsDialog"):
				if hasattr(gui.settingsDialogs.NVDASettingsDialog, "categoryClasses"):
					gui.settingsDialogs.NVDASettingsDialog.categoryClasses.remove(EloquenceSettingsPanel)
		except (ValueError, AttributeError) as e:
			log.debug(f"Settings panel already removed or never registered: {e}")
		except Exception as e:
			log.warning(f"Error removing Eloquence settings panel: {e}")

		super(SynthDriver, self).terminate()

	def combine_adjacent_strings(self, lst):
		result = []
		current_string = ""
		for item in lst:
			if isinstance(item, str):
				current_string += item
			else:
				if current_string:
					result.append(current_string)
					current_string = ""
				result.append(item)
		if current_string:
			result.append(current_string)
		return result

	def speak(self, speechSequence):
		last = None
		outlist = []
		pending_indexes = []
		queued_speech = False

		# Reset prosody to baseline at the start of each utterance to prevent
		# state leaks from previous speech sequences (issue #59).
		for pr in (_eloquence.rate, _eloquence.pitch, _eloquence.vlm):
			outlist.append((_eloquence.cmdProsody, (pr, 1, 0)))

		# IBMTTS Logic: Combine strings before processing regex
		speechSequence = self.combine_adjacent_strings(speechSequence)

		for item in speechSequence:
			if isinstance(item, str):
				s = str(item)
				s = self.xspeakText(s)
				outlist.append((_eloquence.speak, (s,)))
				last = s
				queued_speech = True
			elif isinstance(item, IndexCommand):
				pending_indexes.append(item.index)
				outlist.append((_eloquence.index, (item.index,)))
			elif isinstance(item, BreakCommand):
				# Eloquence doesn't respect delay time in milliseconds.
				# Therefor we need to adjust waiting time depending on curernt speech rate
				# The following table of adjustments has been measured empirically
				# Then we do linear approximation
				coefficients = {
					10: 1,
					43: 2,
					60: 3,
					75: 4,
					85: 5,
				}
				ck = sorted(coefficients.keys())
				if self.rate <= ck[0]:
					factor = coefficients[ck[0]]
				elif self.rate >= ck[-1]:
					factor = coefficients[ck[-1]]
				elif self.rate in ck:
					factor = coefficients[ck[0]]
				else:
					li = [index for index, r in enumerate(ck) if r < self.rate][-1]
					ri = li + 1
					ra = ck[li]
					rb = ck[ri]
					factor = 1.0 * coefficients[ra] + (coefficients[rb] - coefficients[ra]) * (
						self.rate - ra
					) / (rb - ra)
				pFactor = factor * item.time
				pFactor = int(pFactor)
				outlist.append((_eloquence.speak, (f"`p{pFactor}.",)))
				queued_speech = True
			elif isinstance(item, LangChangeCommand):
				voice_id = self._resolve_voice_for_language(item.lang)
				if voice_id is None:
					log.debug("No Eloquence voice mapped for language '%s'", item.lang)
					continue
				voice_str = str(voice_id)
				if voice_str == self.curvoice:
					if item.lang is None:
						self._languageOverrideActive = False
					continue
				try:
					queued_voice = int(voice_id)
				except (TypeError, ValueError):
					log.debug(
						"Skipping language change for '%s': invalid voice id %r",
						item.lang,
						voice_id,
					)
					continue
				outlist.append((_eloquence.set_voice, (queued_voice,)))
				self._update_voice_state(queued_voice, update_default=item.lang is None)
			elif type(item) in self.PROSODY_ATTRS:
				pr = self.PROSODY_ATTRS[type(item)]
				# Use the raw _offset/_multiplier values directly, NOT the
				# computed properties.  NVDA guarantees that only one of them
				# is specified (they are mutually exclusive).  The computed
				# .multiplier property already folds offset into a ratio
				# using the *current* defaultValue, so passing both would
				# double-count the change.  Raw values are stable constants
				# that do not depend on defaultValue and are safe to apply
				# later in the worker thread against the live base pitch.
				raw_offset = getattr(item, "_offset", 0)
				raw_multiplier = getattr(item, "_multiplier", 1)
				outlist.append(
					(
						_eloquence.cmdProsody,
						(pr, raw_multiplier, raw_offset),
					)
				)
		if not queued_speech:
			# No speech queued. Ensure any state changes apply and emit indexes immediately
			# so sayAll can advance even when there's nothing to speak.
			for func, args in outlist:
				if func is _eloquence.index:
					continue
				try:
					func(*args)
				except Exception:
					log.exception("Synthesis command failed")
			for index in pending_indexes:
				synthIndexReached.notify(synth=self, index=index)
			synthDoneSpeaking.notify(synth=self)
			return

		# Trailing Pause Logic from IBMTTS:
		if last is not None and last.rstrip()[-1] not in punctuation:
			# Mode 0 uses p0 for legacy speed performance
			# Mode 1 and 2 use p1 for standard modern speed
			p_val = "0" if self._pause_mode == 0 else "1"
			outlist.append((_eloquence.speak, (f"`p{p_val} ",)))

		outlist.append((_eloquence.index, (0xFFFF,)))
		outlist.append((_eloquence.synth, ()))
		seq = _eloquence._client._sequence
		_eloquence.synth_queue.put((outlist, seq))
		_eloquence.process()

	def xspeakText(self, text, should_pause=False):
		text = _text_preprocessing.preprocess(text, _eloquence.params[9])
		if not self._backquoteVoiceTags:
			text = text.replace("`", " ")
		text = "`vv%d %s" % (
			self.getVParam(_eloquence.vlm),
			text,
		)  # no embedded commands

		# IBMTTS Regex Injection Logic for dynamic pausing:
		if self._pause_mode == 0:
			# Mode 0 (Do not shorten) maps punctuation to p0 for legacy snappy performance.
			text = pause_re.sub(r"\1 `p0\2\3\4", text)
		elif self._pause_mode == 2:
			# Mode 2 (Shorten all pauses) maps punctuation to p1 for consistent modern shortening.
			text = pause_re.sub(r"\1 `p1\2\3\4", text)

		text = time_re.sub(r"\1:\2 \3", text)
		if self._ABRDICT:
			text = "`da1 " + text
		else:
			text = "`da0 " + text
		if self._phrasePrediction:
			text = "`pp1 " + text
		else:
			text = "`pp0 " + text
		# if two strings are sent separately, pause between them. This might fix some of the audio issues we're having.
		if should_pause:
			p_val = "0" if self._pause_mode == 0 else "1"
			text = text + f" `p{p_val}."
		return text
		#  _eloquence.speak(text, index)

	# def cancel(self):
	#  self.dll.eciStop(self.handle)

	def pause(self, switch):
		_eloquence.pause(switch)
		#  self.dll.eciPause(self.handle,switch)

	# Pause Mode Definitions:
	# 0: Injects p0 at all punctuation for Legacy Speed.
	# 1: Standard timing with a p1 pause at the end of speech blocks only.
	# 2: Injects p1 at all punctuation for consistent Modern Shortening.
	_pauseModes = {
		# Translators: One of the mode listed in the pause combobox synth setting available in speech settings dialog
		"0": StringParameterInfo("0", _("Do not shorten")),
		# Translators: One of the mode listed in the pause combobox synth setting available in speech settings dialog
		"1": StringParameterInfo("1", _("Shorten at end only")),
		# Translators: One of the mode listed in the pause combobox synth setting available in speech settings dialog
		"2": StringParameterInfo("2", _("Shorten all pauses")),
	}

	def _get_availablePausemodes(self):
		return self._pauseModes

	def _set_pauseMode(self, val):
		self._pause_mode = int(val)

	def _get_pauseMode(self):
		return str(self._pause_mode)

	_backquoteVoiceTags = False
	_ABRDICT = False
	_phrasePrediction = False

	def _get_backquoteVoiceTags(self):
		return self._backquoteVoiceTags

	def _set_backquoteVoiceTags(self, enable):
		if enable == self._backquoteVoiceTags:
			return
		self._backquoteVoiceTags = enable

	def _get_ABRDICT(self):
		return self._ABRDICT

	def _set_ABRDICT(self, enable):
		if enable == self._ABRDICT:
			return
		self._ABRDICT = enable

	def _get_phrasePrediction(self):
		return self._phrasePrediction

	def _set_phrasePrediction(self, enable):
		if enable == self._phrasePrediction:
			return
		self._phrasePrediction = enable

	def _get_rate(self):
		return self._paramToPercent(self.getVParam(_eloquence.rate), minRate, maxRate)

	def _set_rate(self, vl):
		self._rate = self._percentToParam(vl, minRate, maxRate)
		self.setVParam(_eloquence.rate, self._percentToParam(vl, minRate, maxRate))

	def _get_pitch(self):
		return self.getVParam(_eloquence.pitch)

	def _set_pitch(self, vl):
		self.setVParam(_eloquence.pitch, vl)

	def _get_volume(self):
		return self.getVParam(_eloquence.vlm)

	def _set_volume(self, vl):
		self.setVParam(_eloquence.vlm, int(vl))

	def _set_inflection(self, vl):
		vl = int(vl)
		self.setVParam(_eloquence.fluctuation, vl)

	def _get_inflection(self):
		return self.getVParam(_eloquence.fluctuation)

	def _set_hsz(self, vl):
		vl = int(vl)
		self.setVParam(_eloquence.hsz, vl)

	def _get_hsz(self):
		return self.getVParam(_eloquence.hsz)

	def _set_rgh(self, vl):
		vl = int(vl)
		self.setVParam(_eloquence.rgh, vl)

	def _get_rgh(self):
		return self.getVParam(_eloquence.rgh)

	def _set_bth(self, vl):
		vl = int(vl)
		self.setVParam(_eloquence.bth, vl)

	def _get_bth(self):
		return self.getVParam(_eloquence.bth)

	def _getAvailableVoices(self):
		o = OrderedDict()
		for name in os.listdir(_eloquence.eciPath[:-8]):
			if not name.lower().endswith(".syn"):
				continue
			voice_code = name.lower()[:-4]
			info = _eloquence.langs[voice_code]
			language = VOICE_BCP47.get(voice_code)
			o[str(info[0])] = synthDriverHandler.VoiceInfo(str(info[0]), info[1], language)
		return o

	def _get_voice(self):
		return str(_eloquence.params[9])

	def _set_voice(self, vl):
		_eloquence.set_voice(vl)
		self._update_voice_state(vl, update_default=True)

	def _update_voice_state(self, voice_id, update_default):
		voice_str = str(voice_id)
		try:
			_eloquence.params[9] = int(voice_str)
		except (TypeError, ValueError):
			log.debug("Unable to coerce Eloquence voice id '%s' to int", voice_id)
		if update_default or not getattr(self, "_defaultVoice", None):
			self._defaultVoice = voice_str
		self.curvoice = voice_str
		current_default = getattr(self, "_defaultVoice", None)
		self._languageOverrideActive = (
			(not update_default) and current_default is not None and voice_str != current_default
		)

	def _resolve_voice_for_language(self, language):
		if not language:
			return getattr(self, "_defaultVoice", None)
		normalized = language.lower().replace("_", "-")
		voice_id = LANGUAGE_TO_VOICE_ID.get(normalized)
		if voice_id:
			return voice_id
		primary, _, region = normalized.partition("-")
		default_voice = getattr(self, "_defaultVoice", None)
		default_lang = VOICE_ID_TO_BCP47.get(default_voice) if default_voice else None
		if default_lang:
			default_primary, _, default_region = default_lang.lower().partition("-")
			if default_primary == primary and (not region or default_region == region):
				return default_voice
		candidates = PRIMARY_LANGUAGE_TO_VOICE_IDS.get(primary, [])
		if not candidates:
			return None
		if region:
			for candidate in candidates:
				candidate_tag = VOICE_ID_TO_BCP47.get(candidate)
				if not candidate_tag:
					continue
				cand_primary, _, cand_region = candidate_tag.lower().partition("-")
				if cand_primary == primary and cand_region == region:
					return candidate
			if primary == "es":
				for candidate in candidates:
					candidate_tag = VOICE_ID_TO_BCP47.get(candidate)
					if candidate_tag and candidate_tag.lower().endswith("-419"):
						return candidate
		if default_lang and default_lang.lower().partition("-")[0] == primary:
			return default_voice
		return candidates[0]

	def getVParam(self, pr):
		return _eloquence.getVParam(pr)

	def setVParam(self, pr, vl):
		_eloquence.setVParam(pr, vl)

	def _get_lastIndex(self):
		# fix?
		return _eloquence.lastindex

	def cancel(self):
		_eloquence.stop()

	def _getAvailableVariants(self):
		global variants
		return OrderedDict(
			(str(id), synthDriverHandler.VoiceInfo(str(id), name)) for id, name in variants.items()
		)

	def _set_variant(self, v):
		global variants
		self._variant = v if int(v) in variants else "1"
		_eloquence.setVariant(int(v))
		self.setVParam(_eloquence.rate, self._rate)
		#  if 'eloquence' in config.conf['speech']:
		#   config.conf['speech']['eloquence']['pitch'] = self.pitch

	def _get_variant(self):
		return self._variant

	def _onIndexReached(self, index):
		if index is not None:
			synthIndexReached.notify(synth=self, index=index)
		else:
			synthDoneSpeaking.notify(synth=self)
