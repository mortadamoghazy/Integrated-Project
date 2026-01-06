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
from tkinter import filedialog
from typing import List, Tuple, Dict

import xlwings as xw

# Ensure project root is on sys.path
project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

from src.excel_tools.mapping_utils import (
    ensure_custom_sheet,
    get_source_rows_display,
    save_mapping_row,
    get_month_sheets,
    DEFAULT_MONTH_SHEET_REGEX,
    CUSTOM_SHEET
)
from tests.test_monthly_mapping_engine import (
    DEST_HEADERS,
    load_saved_mappings,
    create_destination_sheet,
    read_employee_ids_from_source,
    DEFAULT_ID_ROW,
    DEFAULT_CODE_COL,
    DEFAULT_LABEL_COL,
    detect_employee_blocks_on_source,
    build_row_lookup,
    sum_rows_for_block,
)
from tests.mapped_data_loader import load_all_mapped_months


def remove_mapping_from_custommap(wb: xw.Book, label: str) -> None:
    """Remove a mapping row from CustomMap sheet."""
    from src.core.normalization import _norm_label
    
    try:
        sh = wb.sheets[CUSTOM_SHEET]
    except Exception:
        return  # CustomMap doesn't exist
    
    used = sh.range("A1").current_region
    if used.rows.count <= 1:
        return  # No data rows
    
    data = used.value
    target_norm = _norm_label(label)
    
    # Find and delete the row
    for idx, row in enumerate(data[1:], start=2):
        if row and row[0] and _norm_label(row[0]) == target_norm:
            sh.range(idx, 1).api.EntireRow.Delete()
            wb.save()
            print(f"Removed mapping for '{label}' from CustomMap")
            return


def get_all_destination_labels(wb: xw.Book, base_headers: List[str]) -> List[str]:
    """Get all destination labels including custom ones from CustomMap."""
    all_labels = list(base_headers)  # Start with base headers
    
    try:
        sh = wb.sheets[CUSTOM_SHEET]
        used = sh.range("A1").current_region
        if used.rows.count > 1:
            data = used.value
            for row in data[1:]:
                if row and row[0]:
                    label = str(row[0]).strip()
                    if label and label not in all_labels:
                        all_labels.append(label)
                        print(f"Loaded custom label from CustomMap: '{label}'")
    except Exception:
        pass  # CustomMap doesn't exist or error reading
    
    return all_labels


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
    print("DEBUG: import_dialog() called")
    wb = _get_caller_workbook()
    print(f"DEBUG: wb = {wb}")

    # Determine folder to scan for Excel files - always use data/raw
    project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
    folder = os.path.join(project_root, 'data', 'raw')
    print(f"DEBUG: Using data/raw folder: {folder}")
    
    # Create folder if it doesn't exist
    os.makedirs(folder, exist_ok=True)

    # gather Excel files from data/raw
    patterns = ["*.xlsm", "*.xlsx", "*.xls"]
    files = []
    for p in patterns:
        files.extend(glob.glob(os.path.join(folder, p)))
    files = sorted(files, key=lambda p: os.path.basename(p).lower())
    print(f"DEBUG: Found {len(files)} files in {folder}")

    if not files:
        msg = f"No Excel files found in:\n{folder}\n\nPlease place your payroll Excel files in the data/raw folder."
        print(f"DEBUG: {msg}")
        try:
            if wb is not None:
                xw.apps.active.api.MsgBox(msg)
            else:
                tk.messagebox.showinfo("No files", msg)
        except Exception:
            print(msg)
        return

    print(f"DEBUG: Building GUI with {len(files)} files")
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
        # Use a local reference for the workbook so we don't shadow outer `wb`
        target_wb = wb
        if target_wb is None:
            # try to attach again
            target_wb = _get_caller_workbook()

        if target_wb is None:
            # Automatically use data/raw/data_sorted.xlsx - open if exists, create if not
            try:
                app = xw.apps.active if xw.apps.count > 0 else xw.App(visible=True)
                project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
                target_path = os.path.join(project_root, 'data', 'raw', 'data_sorted.xlsx')
                
                # Check if data_sorted.xlsx already exists
                if os.path.exists(target_path):
                    # Open existing file
                    target_wb = app.books.open(target_path)
                    print(f"Opened existing target workbook: {target_path}")
                else:
                    # Create new workbook and save to target path
                    target_wb = app.books.add()
                    target_wb.save(target_path)
                    print(f"Created new target workbook at: {target_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Could not open/create target workbook: {e}")
                root.destroy()
                return

        ok_write = _write_selection_to_workbook(target_wb, chosen)
        if not ok_write:
            # still allow mapping, but warn
            messagebox.showwarning("Warning", "Selected file could not be written to workbook cell, proceeding to mapping GUI.")

        # Close this dialog first before opening mapping GUI
        root.destroy()
        
        # Launch advanced mapping GUI to map source columns to Control Panel targets
        try:
            launch_mapping_gui(target_wb, chosen)
        except Exception as e:
            messagebox.showerror("Mapping error", f"An error occurred while running mapping GUI: {e}")

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


def ensure_destination_columns(sh_dest: xw.Sheet, required_labels: List[str]) -> None:
    """Ensure all required destination columns exist in the sheet, adding new ones if needed."""
    from src.core.normalization import _norm_label
    
    # Get current headers
    last_col = sh_dest.range(1, sh_dest.cells.columns.count).end('left').column
    current_headers = sh_dest.range((1, 1), (1, last_col)).value
    if not isinstance(current_headers, list):
        current_headers = [current_headers] if current_headers else []
    
    current_headers_norm = {_norm_label(h): i+1 for i, h in enumerate(current_headers) if h}
    
    # Add missing labels
    next_col = len(current_headers) + 1
    for label in required_labels:
        if _norm_label(label) not in current_headers_norm:
            # Add new column
            sh_dest.range((1, next_col)).value = label
            sh_dest.range((1, next_col)).api.Font.Bold = True
            sh_dest.range((1, next_col)).color = (220, 230, 241)
            sh_dest.range((1, next_col)).column_width = 18
            print(f"    Added new column: '{label}' at position {next_col}")
            current_headers_norm[_norm_label(label)] = next_col
            next_col += 1


def get_dest_header_map_from_sheet(sh_dest: xw.Sheet) -> Dict[str, int]:
    """Get normalized header -> column index map from actual sheet headers (supports custom columns)."""
    from src.core.normalization import _norm_label
    
    last_col = sh_dest.range(1, sh_dest.cells.columns.count).end('left').column
    headers = sh_dest.range((1, 1), (1, last_col)).value
    if not isinstance(headers, list):
        headers = [headers] if headers else []
    
    return {_norm_label(h): i+1 for i, h in enumerate(headers) if h}


def apply_mappings_one_month_custom(
    sh_src: xw.Sheet,
    sh_dest: xw.Sheet,
    rules: List,
    *,
    id_row: int,
    code_col: int,
    label_col: int,
) -> None:
    """Apply mappings with dynamic header detection (supports custom columns)."""
    from src.core.normalization import _norm_label, _norm_emp_id
    
    # Use dynamic header map instead of hardcoded DEST_HEADERS
    dest_header_map = get_dest_header_map_from_sheet(sh_dest)

    raw_ids, id_width = read_employee_ids_from_source(sh_src, id_row=id_row)
    
    # Fill employee IDs
    if raw_ids:
        sh_dest.range((2, 1), (1 + len(raw_ids), 1)).value = [[rid] for rid in raw_ids]
    
    emp_to_dest_row = {}
    for idx, rid in enumerate(raw_ids, start=2):
        emp_norm = _norm_emp_id(rid, id_width)
        emp_to_dest_row[emp_norm] = idx
    
    emp_blocks = detect_employee_blocks_on_source(sh_src, id_row=id_row, id_width=id_width)
    code_to_rows, labelnorm_to_rows = build_row_lookup(sh_src, code_col=code_col, label_col=label_col)

    for rule in rules:
        dest_col = dest_header_map.get(rule.dest_label_norm)
        if not dest_col or dest_col == 1:
            continue  # skip employee_id col or unknown headers

        for emp_norm, block in emp_blocks.items():
            dest_row = emp_to_dest_row.get(emp_norm)
            if not dest_row:
                continue

            start_col, end_col = block["start_col"], block["end_col"]
            total = None

            for key, op, item_type in rule.keys_ops:
                if item_type == "code":
                    rows = code_to_rows.get(key, [])
                else:
                    rows = labelnorm_to_rows.get(_norm_label(key), [])

                if not rows:
                    continue

                part = sum_rows_for_block(sh_src, rows, start_col, end_col)

                if total is None:
                    total = 0.0 if op in ["+", "-"] else 1.0

                if op == "+":
                    total += part
                elif op == "-":
                    total -= part
                elif op == "*":
                    total *= part
                elif op == "/" and part != 0:
                    total /= part

            if total is not None:
                cell = sh_dest.range((dest_row, dest_col))
                cell.value = total
                cell.color = (204, 255, 204)
                cell.api.Font.Bold = True


def remove_label_from_sheets(
    target_wb: xw.Book,
    label: str,
    month_sheets: List[str]
) -> None:
    """Remove a label column from all destination sheets."""
    from src.core.normalization import _norm_label
    
    label_norm = _norm_label(label)
    
    for sheet_name in month_sheets:
        try:
            sh = target_wb.sheets[sheet_name]
            # Find the column
            last_col = sh.range(1, sh.cells.columns.count).end('left').column
            headers = sh.range((1, 1), (1, last_col)).value
            if not isinstance(headers, list):
                headers = [headers] if headers else []
            
            for idx, h in enumerate(headers, start=1):
                if h and _norm_label(h) == label_norm:
                    # Delete the column
                    sh.range((1, idx)).api.EntireColumn.Delete()
                    print(f"  Removed column '{label}' from sheet '{sheet_name}'")
                    break
        except Exception as e:
            print(f"  Warning: Could not remove column from '{sheet_name}': {e}")


def clear_label_from_sheets(
    target_wb: xw.Book,
    label: str,
    month_sheets: List[str]
) -> None:
    """Clear data from a label column in all destination sheets (keeps column)."""
    from src.core.normalization import _norm_label
    
    label_norm = _norm_label(label)
    
    for sheet_name in month_sheets:
        try:
            sh = target_wb.sheets[sheet_name]
            # Find the column
            last_col = sh.range(1, sh.cells.columns.count).end('left').column
            headers = sh.range((1, 1), (1, last_col)).value
            if not isinstance(headers, list):
                headers = [headers] if headers else []
            
            for idx, h in enumerate(headers, start=1):
                if h and _norm_label(h) == label_norm:
                    # Clear data (keep header)
                    last_row = sh.range((sh.cells.rows.count, idx)).end('up').row
                    if last_row > 1:
                        sh.range((2, idx), (last_row, idx)).value = ""
                        sh.range((2, idx), (last_row, idx)).color = None
                    print(f"  Cleared data from column '{label}' in sheet '{sheet_name}'")
                    break
        except Exception as e:
            print(f"  Warning: Could not clear column in '{sheet_name}': {e}")


def apply_mappings_from_source(
    source_wb: xw.Book,
    target_wb: xw.Book,
    month_regex: str = DEFAULT_MONTH_SHEET_REGEX,
    label_col: int = 2,
    id_row: int = DEFAULT_ID_ROW,
    code_col: int = DEFAULT_CODE_COL,
    modified_labels_only: List[str] = None,
) -> None:
    """Apply saved mappings from source workbook to create mapped sheets in target workbook.
    
    Args:
        modified_labels_only: If provided, only apply mappings for these destination labels.
                            This optimizes performance by not rewriting unchanged mappings.
    """
    # Load mapping rules
    all_rules = load_saved_mappings(source_wb)
    if not all_rules:
        messagebox.showinfo("No Mappings", 
                          "No mapping rules found. The mapping definitions will be used when you run this again.")
        return
    
    # Filter rules if only specific labels were modified
    if modified_labels_only:
        from src.core.normalization import _norm_label
        modified_norms = {_norm_label(lbl) for lbl in modified_labels_only}
        rules = [r for r in all_rules if r.dest_label_norm in modified_norms]
        if not rules:
            print("No rules match the modified labels, skipping mapping application.")
            return
        print(f"Applying only {len(rules)} modified mapping(s): {modified_labels_only}")
    else:
        rules = all_rules
        print(f"Applying all {len(rules)} mappings")
    
    # Get month sheets
    month_sheets = get_month_sheets(source_wb, month_regex=month_regex)
    if not month_sheets:
        messagebox.showwarning("No Month Sheets", 
                             f"No month sheets found matching regex: {month_regex}")
        return
    
    # Collect all destination labels from rules to ensure columns exist
    all_dest_labels = [r.dest_label_raw for r in rules]
    
    print(f"Applying mappings for {len(month_sheets)} month sheets: {month_sheets}")
    
    # Process each month sheet
    for src_name in month_sheets:
        sh_src = source_wb.sheets[src_name]
        
        # If only updating specific labels and sheet exists, update in place
        existing_names = [s.name for s in target_wb.sheets]
        if modified_labels_only and src_name in existing_names:
            # Update existing sheet (partial update)
            sh_dest = target_wb.sheets[src_name]
            print(f"  Updating '{src_name}' (modified labels only)")
            # Ensure new columns exist if custom labels were added
            ensure_destination_columns(sh_dest, all_dest_labels)
        else:
            # Create new sheet or full recreation if not in partial mode
            if src_name in existing_names:
                target_wb.sheets[src_name].delete()
            sh_dest = create_destination_sheet(target_wb, base_name=src_name, suffix="")
            print(f"  Creating '{src_name}' (full mapping)")
            # Add any custom labels that aren't in DEST_HEADERS
            ensure_destination_columns(sh_dest, all_dest_labels)
        
        apply_mappings_one_month_custom(
            sh_src=sh_src,
            sh_dest=sh_dest,
            rules=rules,
            id_row=id_row,
            code_col=code_col,
            label_col=label_col,
        )
    
    # Save target workbook to data/raw/data_sorted.xlsx
    try:
        # Determine save path
        project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
        save_path = os.path.join(project_root, 'data', 'raw', 'data_sorted.xlsx')
        
        # Save to specific location
        target_wb.save(save_path)
        target_wb_path = save_path
        print(f"✅ Mappings applied successfully to {len(month_sheets)} sheets")
        print(f"✅ Saved to: {save_path}")
        
        # Automatically convert to CSV for analysis
        try:
            print("\n📊 Converting mapped data to CSV...")
            csv_output_path = os.path.join(project_root, 'outputs', 'payroll_mapped_long.csv')
            df = load_all_mapped_months(target_wb_path, csv_output_path)
            messagebox.showinfo("Success", 
                              f"Mappings applied successfully!\n\n"
                              f"✅ Created {len(month_sheets)} mapped sheets\n"
                              f"✅ Saved: data/raw/data_sorted.xlsx\n"
                              f"✅ Generated CSV: outputs/payroll_mapped_long.csv\n\n"
                              f"Ready for plotting & forecasting!")
        except Exception as csv_error:
            print(f"⚠️ Warning: Could not generate CSV: {csv_error}")
            import traceback
            traceback.print_exc()
            messagebox.showinfo("Partial Success", 
                              f"Mappings applied successfully to {len(month_sheets)} sheets!\n\n"
                              f"✅ Saved: data/raw/data_sorted.xlsx\n"
                              f"⚠️ CSV conversion failed: {csv_error}\n\n"
                              f"You can run it manually: python tests/run_mapped_data_loader.py")
    except Exception as e:
        print(f"Warning: Could not save target workbook: {e}")
        import traceback
        traceback.print_exc()
        messagebox.showwarning("Save Warning", 
                             f"Mappings applied but could not auto-save: {e}\n\n"
                             f"Please save the workbook manually.")


def launch_mapping_gui(
    target_wb: xw.Book,
    source_path: str,
    month_regex: str = DEFAULT_MONTH_SHEET_REGEX,
    label_col: int = 2,
):
    """Launch advanced mapping GUI for source Excel file to define mappings."""
    # Open source workbook
    app = target_wb.app
    source_wb = None
    try:
        # Check if source is already open
        for b in app.books:
            if b.fullname.lower() == source_path.lower():
                source_wb = b
                break
        if source_wb is None:
            source_wb = app.books.open(source_path)
    except Exception as e:
        messagebox.showerror("Error", f"Could not open source workbook: {e}")
        return

    ensure_custom_sheet(source_wb)

    # Get template headers excluding employee_id (auto-mapped)
    base_headers = [h for h in DEST_HEADERS if h.lower() != "employee_id"]
    
    # Load all destination labels including custom ones from CustomMap
    template_headers = get_all_destination_labels(source_wb, base_headers)

    # Get month sheets from source workbook
    month_sheets = get_month_sheets(source_wb, month_regex=month_regex)
    if not month_sheets:
        messagebox.showerror("Error", f"No month sheets found matching regex: {month_regex}")
        return

    # Build GUI with improved styling
    root = tk.Tk()
    root.title("Payroll Mapping Configuration")
    root.geometry("800x650")
    root.resizable(True, True)
    root.configure(bg="#f0f0f0")
    
    # Configure style
    style = ttk.Style()
    style.theme_use('clam')
    style.configure('Title.TLabel', font=('Segoe UI', 11, 'bold'), foreground='#2c3e50')
    style.configure('TLabel', font=('Segoe UI', 9), foreground='#34495e')
    style.configure('TCombobox', font=('Segoe UI', 9))
    style.configure('TButton', font=('Segoe UI', 9), padding=6)
    
    # Main frame
    main_frame = ttk.Frame(root, padding=20)
    main_frame.grid(sticky="nsew", padx=10, pady=10)
    root.grid_rowconfigure(0, weight=1)
    root.grid_columnconfigure(0, weight=1)
    
    # Header
    header = ttk.Label(main_frame, text="Payroll Data Mapping Configuration", 
                      style='Title.TLabel')
    header.grid(row=0, column=0, columnspan=2, pady=(0, 20), sticky="w")
    
    # Separator
    ttk.Separator(main_frame, orient='horizontal').grid(row=1, column=0, columnspan=2, 
                                                         sticky="ew", pady=(0, 15))
    
    frame = main_frame

    ttk.Label(frame, text="1) Source Month Sheet:", font=('Segoe UI', 9, 'bold')).grid(row=2, column=0, sticky="w", pady=(5,2))
    var_src_sheet = tk.StringVar(value=month_sheets[0])
    cb_src_sheet = ttk.Combobox(frame, textvariable=var_src_sheet, state="readonly", values=month_sheets, width=30)
    cb_src_sheet.grid(row=3, column=0, sticky="ew", pady=(0, 15))

    ttk.Label(frame, text="2) Destination Column:", font=('Segoe UI', 9, 'bold')).grid(row=4, column=0, sticky="w", pady=(5,2))
    
    # Container for destination column selection and add button
    dest_frame = ttk.Frame(frame)
    dest_frame.grid(row=5, column=0, sticky="ew", pady=(0, 15))
    dest_frame.columnconfigure(0, weight=1)
    
    var_dest = tk.StringVar()
    cb_dest = ttk.Combobox(dest_frame, textvariable=var_dest, state="readonly", values=template_headers, width=50)
    cb_dest.grid(row=0, column=0, sticky="ew", padx=(0, 5))
    
    def add_new_label():
        """Add a new custom destination label."""
        # Create a popup dialog
        dialog = tk.Toplevel(root)
        dialog.title("Add New Destination Label")
        dialog.geometry("400x150")
        dialog.transient(root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="Enter new destination label name:", font=('Segoe UI', 9)).pack(pady=(20, 5))
        new_label_var = tk.StringVar()
        entry = ttk.Entry(dialog, textvariable=new_label_var, width=40)
        entry.pack(pady=5)
        entry.focus()
        
        def confirm():
            new_label = new_label_var.get().strip()
            if not new_label:
                messagebox.showwarning("Invalid", "Label name cannot be empty.", parent=dialog)
                return
            if new_label in template_headers:
                messagebox.showwarning("Duplicate", f"Label '{new_label}' already exists.", parent=dialog)
                return
            
            # Add to template headers
            template_headers.append(new_label)
            cb_dest['values'] = template_headers
            cb_dest.set(new_label)  # Auto-select the new label
            messagebox.showinfo("Success", f"Added new label: '{new_label}'", parent=dialog)
            dialog.destroy()
        
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=15)
        ttk.Button(btn_frame, text="Add", command=confirm).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Cancel", command=dialog.destroy).pack(side='left', padx=5)
        
        # Bind Enter key
        entry.bind('<Return>', lambda e: confirm())
    
    def remove_label():
        """Remove a destination label completely (deletes column from all sheets)."""
        label = var_dest.get()
        if not label:
            messagebox.showwarning("No Selection", "Please select a label to remove.")
            return
        
        if label.lower() == "employee_id":
            messagebox.showerror("Cannot Remove", "Cannot remove employee_id column.")
            return
        
        confirm = messagebox.askyesno(
            "Confirm Removal",
            f"Remove '{label}' completely?\n\n"
            f"This will:\n"
            f"• Delete the column from all destination sheets\n"
            f"• Delete its mapping from CustomMap\n\n"
            f"This action cannot be undone!"
        )
        
        if not confirm:
            return
        
        # Remove from template headers
        if label in template_headers:
            template_headers.remove(label)
            cb_dest['values'] = template_headers
            cb_dest.set('')
        
        # Remove from CustomMap
        remove_mapping_from_custommap(source_wb, label)
        
        # Track as modified for deletion from sheets
        if label not in modified_labels:
            modified_labels.append(label)
        
        messagebox.showinfo("Removed", f"Label '{label}' will be removed from all sheets when GUI closes.")
    
    def clear_label_mapping():
        """Clear the mapping for a label (keeps column but removes mapping data)."""
        label = var_dest.get()
        if not label:
            messagebox.showwarning("No Selection", "Please select a label to clear.")
            return
        
        if label.lower() == "employee_id":
            messagebox.showerror("Cannot Clear", "Cannot clear employee_id column.")
            return
        
        confirm = messagebox.askyesno(
            "Confirm Clear",
            f"Clear mapping for '{label}'?\n\n"
            f"This will:\n"
            f"• Keep the column in destination sheets\n"
            f"• Remove its mapping from CustomMap\n"
            f"• Clear all data in this column\n\n"
            f"The column will remain empty until you define a new mapping."
        )
        
        if not confirm:
            return
        
        # Remove from CustomMap
        remove_mapping_from_custommap(source_wb, label)
        
        # Track as modified for clearing from sheets
        if label not in modified_labels:
            modified_labels.append(label)
        
        messagebox.showinfo("Cleared", f"Mapping for '{label}' will be cleared when GUI closes.")
    
    # Button row with add, remove, and clear
    btn_row = ttk.Frame(dest_frame)
    btn_row.grid(row=0, column=1, sticky="e")
    
    ttk.Button(btn_row, text="➕ New", command=add_new_label, 
              style='Action.TButton', width=8).pack(side='left', padx=2)
    ttk.Button(btn_row, text="🗑 Remove", command=remove_label, 
              style='Action.TButton', width=10).pack(side='left', padx=2)
    ttk.Button(btn_row, text="⚠ Clear", command=clear_label_mapping, 
              style='Action.TButton', width=8).pack(side='left', padx=2)

    ttk.Label(frame, text="3) Operator:", font=('Segoe UI', 9, 'bold')).grid(row=6, column=0, sticky="w", pady=(5,2))
    var_op = tk.StringVar(value="+")
    cb_op = ttk.Combobox(frame, textvariable=var_op, state="readonly", values=["+", "-", "*", "/"], width=10)
    cb_op.grid(row=7, column=0, sticky="w", pady=(0, 15))

    ttk.Label(frame, text="4) Build Formula - Add by Code:", font=('Segoe UI', 9, 'bold')).grid(row=8, column=0, sticky="w", pady=(5,2))
    var_code = tk.StringVar()
    cb_code = ttk.Combobox(frame, textvariable=var_code, state="readonly", width=35)
    cb_code.grid(row=9, column=0, sticky="ew")

    ttk.Label(frame, text="Or Add by Label:", font=('Segoe UI', 9)).grid(row=10, column=0, sticky="w", pady=(10,2))
    var_label = tk.StringVar()
    cb_label = ttk.Combobox(frame, textvariable=var_label, state="readonly", width=70)
    cb_label.grid(row=11, column=0, sticky="ew")

    items: List[Tuple[str, str, str]] = []  # (key, op, type)
    modified_labels: List[str] = []  # Track which destination labels were modified this session
    
    # Current formula display
    formula_frame = tk.Frame(frame, bg='#ecf0f1', relief='solid', borderwidth=1)
    formula_frame.grid(row=12, column=0, sticky="ew", pady=(15, 10), padx=2)
    ttk.Label(formula_frame, text="Current Formula:", font=('Segoe UI', 8, 'bold'), 
             background='#ecf0f1').pack(anchor='w', padx=8, pady=(5,2))
    lbl_selected = ttk.Label(formula_frame, text="(none)", font=('Consolas', 9), 
                            background='#ecf0f1', foreground='#27ae60')
    lbl_selected.pack(anchor='w', padx=8, pady=(0, 5))

    display_to_raw: Dict[str, str] = {}

    def refresh_source_lists():
        sh = source_wb.sheets[var_src_sheet.get()]
        labels_display, map_disp_to_raw, codes_sorted = get_source_rows_display(sh, label_col=label_col)
        nonlocal display_to_raw
        display_to_raw = map_disp_to_raw
        cb_code["values"] = codes_sorted
        cb_label["values"] = labels_display

    def add_code():
        c = var_code.get()
        if not c:
            return
        op = var_op.get()
        entry = (c, op, "code")
        if entry not in items:
            items.append(entry)
        lbl_selected["text"] = "Selected items: " + ", ".join([f"{op}{key}" for key, op, _ in items])

    def add_label():
        disp = var_label.get()
        if not disp:
            return
        raw = display_to_raw.get(disp, disp)
        op = var_op.get()
        entry = (raw, op, "label")
        if entry not in items:
            items.append(entry)
        lbl_selected["text"] = "Selected items: " + ", ".join([f"{op}{key}" for key, op, _ in items])

    def clear_items():
        items.clear()
        lbl_selected["text"] = "Selected items: (none)"

    def save_rule():
        dest = var_dest.get()
        if not dest:
            messagebox.showerror("Error", "Select a destination column from template.")
            return
        if not items:
            messagebox.showerror("Error", "Add at least one code or label item.")
            return

        save_mapping_row(source_wb, dest_label_raw=dest, keys_ops=items)
        source_wb.save()
        
        # Track this label as modified
        if dest not in modified_labels:
            modified_labels.append(dest)
        
        messagebox.showinfo("Saved", f"Saved mapping for '{dest}' in CustomMap sheet.")
        clear_items()

    # Button styling
    style.configure('Action.TButton', font=('Segoe UI', 9), padding=8)
    style.configure('Primary.TButton', font=('Segoe UI', 9, 'bold'), padding=8)
    
    # Buttons layout
    btn_frame = ttk.Frame(frame)
    btn_frame.grid(row=13, column=0, sticky="ew", pady=(15, 0))
    
    ttk.Button(btn_frame, text="🔄 Refresh Lists", command=refresh_source_lists, 
              style='Action.TButton').pack(side='left', padx=(0, 5))
    ttk.Button(btn_frame, text="➕ Add Code", command=add_code, 
              style='Action.TButton').pack(side='left', padx=5)
    ttk.Button(btn_frame, text="➕ Add Label", command=add_label, 
              style='Action.TButton').pack(side='left', padx=5)
    ttk.Button(btn_frame, text="🗑 Clear", command=clear_items, 
              style='Action.TButton').pack(side='left', padx=5)
    
    # Bottom action buttons
    bottom_frame = ttk.Frame(frame)
    bottom_frame.grid(row=14, column=0, sticky="ew", pady=(20, 0))
    
    ttk.Button(bottom_frame, text="💾 Save Mapping Rule", command=save_rule, 
              style='Primary.TButton').pack(side='left', padx=(0, 10))
    ttk.Button(bottom_frame, text="❌ Close", command=root.destroy, 
              style='Action.TButton').pack(side='right')

    refresh_source_lists()
    root.mainloop()
    
    # After GUI closes, automatically run mapping engine to apply only the modified mappings
    if modified_labels:
        try:
            # Check which labels need to be removed or cleared vs updated
            labels_to_remove = []
            labels_to_clear = []
            labels_to_update = []
            
            # Load current mappings to see what still exists
            current_rules = load_saved_mappings(source_wb)
            current_mapped_labels = {r.dest_label_raw for r in current_rules}
            
            for label in modified_labels:
                if label not in template_headers:
                    # Label was removed from template list
                    labels_to_remove.append(label)
                elif label not in current_mapped_labels:
                    # Label exists but has no mapping (cleared)
                    labels_to_clear.append(label)
                else:
                    # Label has a mapping (normal update)
                    labels_to_update.append(label)
            
            # Get month sheets
            month_sheets = get_month_sheets(source_wb, month_regex)
            existing_sheet_names = [s.name for s in target_wb.sheets]
            actual_sheets = [s for s in month_sheets if s in existing_sheet_names]
            
            # Handle removals
            if labels_to_remove:
                print(f"Removing {len(labels_to_remove)} label(s): {labels_to_remove}")
                for label in labels_to_remove:
                    remove_label_from_sheets(target_wb, label, actual_sheets)
            
            # Handle clears
            if labels_to_clear:
                print(f"Clearing {len(labels_to_clear)} label(s): {labels_to_clear}")
                for label in labels_to_clear:
                    clear_label_from_sheets(target_wb, label, actual_sheets)
            
            # Handle updates
            if labels_to_update:
                apply_mappings_from_source(source_wb, target_wb, month_regex, label_col, 
                                         modified_labels_only=labels_to_update)
            
            # Save workbook and generate CSV
            if labels_to_remove or labels_to_clear or labels_to_update:
                try:
                    target_wb.save()
                    target_wb_path = target_wb.fullname
                    
                    # Automatically convert to CSV for analysis
                    print("\n📊 Converting mapped data to CSV...")
                    csv_output = "outputs/payroll_mapped_long.csv"
                    df = load_all_mapped_months(target_wb_path, csv_output)
                    print(f"✅ CSV generated: {csv_output}")
                    
                except Exception as csv_error:
                    print(f"⚠️ Warning: Could not generate CSV: {csv_error}")
                    import traceback
                    traceback.print_exc()
                
        except Exception as e:
            messagebox.showerror("Mapping Engine Error", 
                               f"Error applying changes: {e}\n\nChanges were saved but not applied. "
                               f"You can run the mapping engine manually later.")
    else:
        # Even if no labels were modified, try to generate CSV if workbook exists
        try:
            if hasattr(target_wb, 'fullname') and target_wb.fullname:
                print("No mappings modified this session, but generating CSV from existing data...")
                csv_output = "outputs/payroll_mapped_long.csv"
                df = load_all_mapped_months(target_wb.fullname, csv_output)
                print(f"✅ CSV generated: {csv_output}")
        except Exception as e:
            print(f"Could not generate CSV: {e}")
            pass


# Allow running the script directly for testing. When run directly it will attempt to
# find an open workbook named like 'Control Panel' or will run in standalone mode.
if __name__ == "__main__":
    print("DEBUG: Script started")
    try:
        print("DEBUG: Trying to connect to Excel...")
        app = xw.apps.active if xw.apps.count > 0 else xw.App(visible=True)
        print(f"DEBUG: Excel app created/connected: {app}")
        # attempt to attach to a workbook named Control Panel*.xlsm
        try:
            wb = [b for b in app.books if b.name.lower().startswith("control panel")][0]
            wb.set_mock_caller()
            print(f"DEBUG: Attached to workbook: {wb.name}")
        except Exception as e:
            # nothing to attach, user can open the Control Panel workbook manually
            print(f"DEBUG: Could not attach to Control Panel workbook: {e}")
            pass
    except Exception as e:
        print(f"DEBUG: Error in Excel connection: {e}")
        pass

    print("DEBUG: Calling import_dialog()")
    import_dialog()
