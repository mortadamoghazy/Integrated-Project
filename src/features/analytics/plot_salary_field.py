import matplotlib.pyplot as plt
import numpy as np
import xlwings as xw

from src.core.normalization import _norm_label
from src.core.config import TGT_SHEET


def plot_field(wb: xw.Book, field_label: str):
    sh = wb.sheets[TGT_SHEET]

    # Read headers
    headers = sh.range("A1").expand("right").value
    headers_norm = [_norm_label(h) for h in headers]

    target_norm = _norm_label(field_label)
    if target_norm not in headers_norm:
        raise ValueError(f"Field '{field_label}' not found in Sheet1")

    col_index = headers_norm.index(target_norm) + 1  # 0-based -> Excel column

    # Read employee IDs (col A)
    emp_ids = sh.range("A2").expand("down").value
    if not isinstance(emp_ids, list):
        emp_ids = [emp_ids]
    emp_ids = [str(e) for e in emp_ids]

    # Read field values
    values = sh.range((2, col_index + 1)).expand("down").value
    if not isinstance(values, list):
        values = [values]

    # Convert non-numeric to 0
    values = [v if isinstance(v, (int, float)) else 0 for v in values]

    # Compute average
    average = np.mean(values)

    # Compute differences
    diffs = [v - average for v in values]

    # Colors: above avg = blue, below = red
    colors = ["blue" if v >= average else "red" for v in values]

    # Plot
    plt.figure(figsize=(12, 6))
    plt.bar(emp_ids, values, color=colors)

    # Plot average line
    plt.axhline(average, color="green", linestyle="--", label=f"Avg = {average:.2f}")

    # Annotate each bar with +diff or -diff
    for x, y, d in zip(emp_ids, values, diffs):
        sign = "+" if d >= 0 else ""
        plt.text(x, y, f"{sign}{d:.2f}", ha="center", va="bottom", fontsize=8)

    plt.xlabel("Employee ID")
    plt.ylabel(field_label)
    plt.title(f"{field_label} per employee")

    plt.legend()
    plt.tight_layout()
    plt.show()
