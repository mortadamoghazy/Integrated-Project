"""
Per-employee autoregressive linear regression: y_t ~ t + y_{t-1} + y_{t-2}
Target: salaire_brut

Evaluation:
- Per-employee metrics: RMSE, MAE, MAPE, R^2
- Global metrics over all employees
- Learning process visualization:
  * True vs predicted over time
  * Residuals over time
  * Expanding-window learning curve
"""

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_squared_error, mean_absolute_error, r2_score


# =============================
# Configuration
# =============================
TARGET_LABEL = "salaire_brut"
CSV_PATH = "outputs/payroll_long.csv"
TEST_MONTHS = 6  # last months used as test
LAGS = 2  # number of autoregressive lags


# =============================
# Metrics
# =============================
def mape(y_true, y_pred):
    y_true = np.array(y_true)
    y_pred = np.array(y_pred)
    mask = y_true != 0
    if mask.sum() == 0:
        return np.nan
    return np.mean(np.abs((y_true[mask] - y_pred[mask]) / y_true[mask])) * 100.0


# =============================
# Data loading
# =============================
def load_and_prepare(csv_path: str) -> pd.DataFrame:
    df = pd.read_csv(csv_path)
    df["month"] = pd.to_datetime(df["month"], format="%Y-%m")
    df = df.sort_values(["employee_id", "month"]).reset_index(drop=True)
    return df


# =============================
# AR(LAGS) + trend per employee
# =============================
def run_ar1(df: pd.DataFrame):
    employees = df["employee_id"].unique()

    all_y_true = []
    all_y_pred = []
    metrics_per_employee = []
    preds_list = []

    for emp in employees:
        sub = df[df["employee_id"] == emp].copy()
        sub = sub.sort_values("month")

        if len(sub) <= TEST_MONTHS + LAGS + 3:
            continue

        # Lagged variables
        for lag in range(1, LAGS + 1):
            sub[f"y_lag{lag}"] = sub[TARGET_LABEL].shift(lag)
        sub = sub.dropna().copy()

        if len(sub) <= TEST_MONTHS + 3:
            continue

        sub["t"] = np.arange(len(sub))

        lag_cols = [f"y_lag{lag}" for lag in range(1, LAGS + 1)]
        X = sub[["t"] + lag_cols]
        y = sub[TARGET_LABEL].values

        split = len(sub) - TEST_MONTHS
        X_train, X_test = X.iloc[:split], X.iloc[split:]
        y_train, y_test = y[:split], y[split:]

        model = LinearRegression()
        model.fit(X_train, y_train)
        y_pred = model.predict(X_test)

        all_y_true.extend(y_test.tolist())
        all_y_pred.extend(y_pred.tolist())

        metrics_per_employee.append(
            {
                "employee_id": emp,
                "rmse": np.sqrt(mean_squared_error(y_test, y_pred)),
                "mae": mean_absolute_error(y_test, y_pred),
                "mape": mape(y_test, y_pred),
                "r2": r2_score(y_test, y_pred),
                "n_test": len(y_test),
            }
        )

        pred_df = sub.iloc[split:][["employee_id", "month"]].copy()
        pred_df["y_true"] = y_test
        pred_df["y_pred"] = y_pred
        preds_list.append(pred_df)

    metrics_df = pd.DataFrame(metrics_per_employee)
    preds_df = pd.concat(preds_list, ignore_index=True)

    global_metrics = {
        "rmse": np.sqrt(mean_squared_error(all_y_true, all_y_pred)),
        "mae": mean_absolute_error(all_y_true, all_y_pred),
        "mape": mape(all_y_true, all_y_pred),
        "r2": r2_score(all_y_true, all_y_pred),
    }

    return metrics_df, preds_df, global_metrics


# =============================
# Learning process visualization
# =============================
def plot_learning_process(preds_df: pd.DataFrame, employee_id=None):
    if employee_id is not None:
        data = preds_df[preds_df["employee_id"] == employee_id].copy()
        title_suffix = f"Employee {employee_id}"
    else:
        data = preds_df.groupby("month")[["y_true", "y_pred"]].mean().reset_index()
        title_suffix = "All Employees (Mean)"

    data = data.sort_values("month")

    # --- True vs predicted ---
    plt.figure()
    plt.plot(data["month"], data["y_true"], label="True")
    plt.plot(data["month"], data["y_pred"], label="Predicted")
    plt.xlabel("Month")
    plt.ylabel("Salaire brut")
    plt.title(f"AR(1) Learning Process – {title_suffix}")
    plt.legend()
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()

    # --- Residuals ---
    residuals = data["y_true"] - data["y_pred"]

    plt.figure()
    plt.plot(data["month"], residuals)
    plt.axhline(0)
    plt.xlabel("Month")
    plt.ylabel("Residual (True − Predicted)")
    plt.title(f"Residuals Over Time – {title_suffix}")
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()


# =============================
# Expanding-window learning curve
# =============================
def plot_expanding_learning_curve(df: pd.DataFrame, employee_id):
    sub = df[df["employee_id"] == employee_id].copy()
    sub = sub.sort_values("month")

    sub["y_lag1"] = sub[TARGET_LABEL].shift(1)
    sub = sub.dropna().copy()
    sub["t"] = np.arange(len(sub))

    X = sub[["t", "y_lag1"]].values
    y = sub[TARGET_LABEL].values

    rmses = []
    train_sizes = []

    for i in range(6, len(sub) - 1):
        model = LinearRegression()
        model.fit(X[:i], y[:i])
        y_pred = model.predict(X[i:i+1])
        rmse = np.sqrt(mean_squared_error(y[i:i+1], y_pred))
        rmses.append(rmse)
        train_sizes.append(i)

    plt.figure()
    plt.plot(train_sizes, rmses)
    plt.xlabel("Training size (months)")
    plt.ylabel("RMSE")
    plt.title(f"Expanding Window Learning Curve – Employee {employee_id}")
    plt.tight_layout()
    plt.show()


# =============================
# Main
# =============================
if __name__ == "__main__":
    df = load_and_prepare(CSV_PATH)
    metrics_df, preds_df, global_metrics = run_ar1(df)

    print("=== Global metrics (AR(2) + Trend) ===")
    for k, v in global_metrics.items():
        print(f"{k.upper()}: {v:.4f}")

    print("\n=== Mean per-employee metrics ===")
    print(metrics_df[["rmse", "mae", "mape", "r2"]].mean())

    metrics_df.to_csv("results_ar2_emp_metrics.csv", index=False)
    preds_df.to_csv("results_ar2_predictions.csv", index=False)

    # Visual diagnostics
    plot_learning_process(preds_df, employee_id=None)

    if not metrics_df.empty:
        example_emp = metrics_df.iloc[0]["employee_id"]
        plot_learning_process(preds_df, employee_id=example_emp)
        plot_expanding_learning_curve(df, employee_id=example_emp)
