"""
import_dialog.py

xlwings-callable script to open a simple Tk GUI listing Excel files
in the directory of the calling workbook (Control Panel.xlsm).

Usage from Excel (VBA):
    Sub ImportButton()
        RunPython "import src.excel_tools.import_dialog as imp; imp.import_dialog()"
    End Sub

Behavior:
- When called from Excel via xlwings `Book.caller()` it determines the workbook's folder.
- Shows a Tk window with the list of Excel files (*.xls, *.xlsx, *.xlsm) in that folder.
- User selects a file and clicks "Import".
- The script writes the chosen file path to a named range `import_file` if it exists in the calling workbook; otherwise writes it to `Sheet1!A1`.
- Works in standalone mode for testing (it will try to attach to an open workbook named 'Control Panel.xlsm' or ask the user to pick a workbook file).

Notes:
- Requires `xlwings` installed and configured for your Excel -> RunPython hookup.
- The UI is intentionally minimal (Tk) so it runs reliably from Excel on Windows.
"""

import os
import sys
import glob
import tkinter as tk
from tkinter import messagebox, ttk

import xlwings as xw


def _get_caller_workbook():
    """Return attached workbook (xw.Book) using Book.caller() when possible, otherwise try to find an open workbook named 'Control Panel.xlsm' or open it from the current folder."""
    try:
        wb = xw.Book.caller()
        return wb
    except Exception:
        # fallback: attach to an already-open workbook named like 'Control Panel' (case-insensitive)
        app = xw.apps.active if xw.apps.count > 0 else xw.App(visible=True)
        try:
            wb = [b for b in app.books if b.name.lower().startswith("control panel")][0]
            return wb
        except Exception:
            # As final fallback, ask user to pick the workbook from the folder where this script sits
            return None


def _write_selection_to_workbook(wb, selected_path):
    """Write chosen path back to workbook. Prefer named range 'import_file', otherwise write to 'Sheet1'!A1."""
    try:
        # prefer named range 'import_file'
        nm = None
        try:
            nm = wb.names["import_file"]
        except Exception:
            nm = None

        if nm is not None and nm.refers_to:
            # write to the named range location
            try:
                rng = wb.names["import_file"].refers_to_range
                rng.value = selected_path
                return True
            except Exception:
                # fallback to setting via name.address
                try:
                    wb.names["import_file"].value = selected_path
                    return True
                except Exception:
                    pass

        # fallback: write to Sheet1!A1
        try:
            sh = wb.sheets[0]
            sh.range("A1").value = selected_path
            return True
        except Exception:
            # last resort: use active sheet
            wb.app.api.ActiveSheet.Range("A1").Value = selected_path
            return True
    except Exception as e:
        print("Error writing selection to workbook:", e)
        return False


def import_dialog():
    """Main entry point called from Excel via RunPython or tested standalone.

    Shows a small Tk window listing Excel files in the calling workbook's folder and
    writes the selected path back to the calling workbook.
    """
    wb = _get_caller_workbook()

    # Determine folder to scan for Excel files
    if wb is not None:
        folder = os.path.dirname(wb.fullname)
    else:
        # try current working directory (project root) as fallback
        folder = os.getcwd()

    # gather Excel files
    patterns = ["*.xlsm", "*.xlsx", "*.xls"]
    files = []
    for p in patterns:
        files.extend(glob.glob(os.path.join(folder, p)))
    files = sorted(files, key=lambda p: os.path.basename(p).lower())

    # If none found, try parent folder
    if not files:
        parent = os.path.dirname(folder)
        for p in patterns:
            files.extend(glob.glob(os.path.join(parent, p)))
        files = sorted(files, key=lambda p: os.path.basename(p).lower())

    # If still none found, try a recursive search within the folder (covers e.g. data/raw)
    if not files:
        for p in patterns:
            files.extend(glob.glob(os.path.join(folder, '**', p), recursive=True))
        # remove duplicates and sort
        files = sorted(set(files), key=lambda p: os.path.basename(p).lower())

    if not files:
        msg = f"No Excel files found in:\n{folder}\n\nPlease place files in the same folder as the Control Panel workbook."
        try:
            if wb is not None:
                xw.apps.active.api.MsgBox(msg)
            else:
                tk.messagebox.showinfo("No files", msg)
        except Exception:
            print(msg)
        return

    # Build GUI
    root = tk.Tk()
    root.title("Import Excel - Select file")
    root.geometry("640x360")

    frm = ttk.Frame(root, padding=10)
    frm.pack(fill=tk.BOTH, expand=True)

    lbl = ttk.Label(frm, text=f"Files in: {folder}")
    lbl.pack(anchor=tk.W)

    # tree/listbox with scrollbar
    tree_frame = ttk.Frame(frm)
    tree_frame.pack(fill=tk.BOTH, expand=True)

    scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
    listbox = tk.Listbox(tree_frame, yscrollcommand=scrollbar.set, selectmode=tk.SINGLE)
    scrollbar.config(command=listbox.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    for f in files:
        listbox.insert(tk.END, os.path.basename(f))

    # path preview
    preview_var = tk.StringVar(value="")
    preview_lbl = ttk.Label(frm, textvariable=preview_var)
    preview_lbl.pack(fill=tk.X)

    def on_select(event=None):
        sel = listbox.curselection()
        if not sel:
            preview_var.set("")
            return
        idx = sel[0]
        preview_var.set(files[idx])

    listbox.bind("<<ListboxSelect>>", on_select)

    def do_import():
        sel = listbox.curselection()
        if not sel:
            messagebox.showwarning("No selection", "Please select a file to import.")
            return
        idx = sel[0]
        chosen = files[idx]
        if wb is None:
            # try to attach again to write back; if still None we will open file chooser
            wb = _get_caller_workbook()
            if wb is None:
                messagebox.showinfo("Running standalone", f"Selected: {chosen}\n(Standalone mode - cannot write back to Excel automatically)")
                root.destroy()
                return
        ok = _write_selection_to_workbook(wb, chosen)
        if ok:
            messagebox.showinfo("Imported", f"Selected file written to workbook: {chosen}")
        else:
            messagebox.showerror("Error", "Could not write selection to workbook.")
        root.destroy()

    btn_frame = ttk.Frame(frm)
    btn_frame.pack(fill=tk.X)

    btn_import = ttk.Button(btn_frame, text="Import", command=do_import)
    btn_import.pack(side=tk.RIGHT, padx=6, pady=6)

    btn_cancel = ttk.Button(btn_frame, text="Cancel", command=root.destroy)
    btn_cancel.pack(side=tk.RIGHT, pady=6)

    # double-click = import
    def on_double(event):
        on_select()
        do_import()

    listbox.bind("<Double-Button-1>", on_double)

    root.mainloop()


# Allow running the script directly for testing. When run directly it will attempt to
# find an open workbook named like 'Control Panel' or will run in standalone mode.
if __name__ == "__main__":
    try:
        app = xw.apps.active if xw.apps.count > 0 else xw.App(visible=True)
        # attempt to attach to a workbook named Control Panel*.xlsm
        try:
            wb = [b for b in app.books if b.name.lower().startswith("control panel")][0]
            wb.set_mock_caller()
        except Exception:
            # nothing to attach, user can open the Control Panel workbook manually
            pass
    except Exception:
        pass

    import_dialog()
