"""
Visualization of All Regression Models

Fixes:
- Employee ID mismatch (int vs padded str)
- Empty plot for employee 00001
- Automatic detection of available employees
- Consistent ID formatting
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os

OUTPUTS = "outputs/"


# ------------------------------------------------------------
# Helper: Normalize employee IDs to zero-padded strings
# ------------------------------------------------------------
def normalize_emp_id(df):
    df["employee_id"] = (
        df["employee_id"]
        .astype(str)
        .str.replace(".0", "", regex=False)   # In case CSV had floats
        .str.zfill(5)
    )
    return df


# ------------------------------------------------------------
# Load all CSV results (metrics)
# ------------------------------------------------------------
def load_metrics():
    files = {
        "Trend + Seasonality (Fourier K1)": "results_fourier_k1_emp_metrics.csv",
        "AR(1)": "results_ar1_emp_metrics.csv",
        "Pooled FE": "results_pooled_fe_emp_metrics.csv",
        "Ridge": "results_regularized_panel_ridge_emp_metrics.csv",
        "Lasso": "results_regularized_panel_lasso_emp_metrics.csv",
    }

    metrics = {}
    for model, filename in files.items():
        path = os.path.join(OUTPUTS, filename)
        df = pd.read_csv(path)
        df = normalize_emp_id(df)
        metrics[model] = df

    return metrics


# ------------------------------------------------------------
# Load all prediction files for time-series plotting
# ------------------------------------------------------------
def load_predictions():
    files = {
        "Fourier": "results_fourier_k1_predictions.csv",
        "AR1": "results_ar1_predictions.csv",
        "PooledFE": "results_pooled_fe_predictions.csv",
        "RidgeLasso": "results_regularized_panel_predictions.csv",
    }

    preds = {}
    for model, filename in files.items():
        path = os.path.join(OUTPUTS, filename)
        df = pd.read_csv(path)
        df = normalize_emp_id(df)
        df["month"] = pd.to_datetime(df["month"])
        preds[model] = df

    return preds


# ------------------------------------------------------------
# Global Comparison Plots
# ------------------------------------------------------------
def plot_global_comparison(metrics):
    summary = []

    for model, df in metrics.items():
        summary.append([
            model,
            df["rmse"].mean(),
            df["mae"].mean(),
            df["mape"].mean(),
            df["r2"].mean(),
        ])

    summary_df = pd.DataFrame(summary, columns=["Model", "RMSE", "MAE", "MAPE", "R2"])
    summary_df = summary_df.set_index("Model")

    # Global Error Comparison
    summary_df[["RMSE", "MAE", "MAPE"]].plot(kind="bar", figsize=(12,6))
    plt.title("Global Error Comparison Across Models")
    plt.ylabel("Error Value")
    plt.grid(axis="y", linestyle="--", alpha=0.6)
    plt.tight_layout()
    plt.show()

    # Global R² Comparison
    summary_df["R2"].plot(kind="bar", figsize=(10,5), color="green")
    plt.title("Global R² Comparison Across Models")
    plt.ylabel("R² Value")
    plt.grid(axis="y", linestyle="--", alpha=0.6)
    plt.tight_layout()
    plt.show()


# ------------------------------------------------------------
# Per-Employee Comparison Plot
# ------------------------------------------------------------
def plot_per_employee_rmse(metrics):
    all_emp_ids = sorted(metrics["AR(1)"]["employee_id"].unique())
    rmse_df = pd.DataFrame(index=all_emp_ids)

    for model, df in metrics.items():
        rmse_df[model] = df.set_index("employee_id")["rmse"]

    rmse_df.plot(figsize=(14,7))
    plt.title("Per-Employee RMSE Comparison")
    plt.xlabel("Employee ID")
    plt.ylabel("RMSE")
    plt.grid(True, linestyle="--", alpha=0.5)
    plt.tight_layout()
    plt.show()


# ------------------------------------------------------------
# True vs Predicted for One Employee
# ------------------------------------------------------------
def plot_example_employee(preds, emp_id="00001"):
    # Detect available employee IDs
    available_ids = preds["Fourier"]["employee_id"].unique()

    if emp_id not in available_ids:
        print(f"WARNING: employee_id={emp_id} not found. Using first available: {available_ids[0]}")
        emp_id = available_ids[0]

    f = preds["Fourier"][preds["Fourier"]["employee_id"] == emp_id]
    a = preds["AR1"][preds["AR1"]["employee_id"] == emp_id]
    e = preds["PooledFE"][preds["PooledFE"]["employee_id"] == emp_id]
    r = preds["RidgeLasso"][preds["RidgeLasso"]["employee_id"] == emp_id]

    plt.figure(figsize=(12,6))

    plt.plot(f["month"], f["y_true"], label="True Salary", linewidth=3, color="black")
    plt.plot(f["month"], f["y_pred"], label="Fourier (K1)", linestyle="--")
    plt.plot(a["month"], a["y_pred"], label="AR(1)", linestyle="--")
    plt.plot(e["month"], e["y_pred"], label="Pooled FE", linestyle="--")
    plt.plot(r["month"], r["y_pred_ridge"], label="Ridge", linestyle="--")
    plt.plot(r["month"], r["y_pred_lasso"], label="Lasso", linestyle="--")

    plt.title(f"True vs Predicted Salary for Employee {emp_id}")
    plt.xlabel("Month")
    plt.ylabel("Salaire Brut (EUR)")
    plt.grid(True, linestyle="--", alpha=0.5)
    plt.legend()
    plt.tight_layout()
    plt.show()


# ------------------------------------------------------------
# MAIN
# ------------------------------------------------------------
if __name__ == "__main__":
    metrics = load_metrics()
    preds = load_predictions()

    print("Plotting global comparison...")
    plot_global_comparison(metrics)

    print("Plotting per-employee RMSE comparison...")
    plot_per_employee_rmse(metrics)

    print("Plotting example employee forecast...")
    plot_example_employee(preds, emp_id="00001")

    print("All visualizations completed.")
