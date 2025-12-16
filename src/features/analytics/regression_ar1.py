"""
Per-employee autoregressive linear regression: y_t ~ t + y_{t-1}
Target: salaire_brut

Evaluation:
- Per-employee metrics: RMSE, MAE, MAPE, R^2
- Global metrics over all employees
"""

import numpy as np
import pandas as pd
from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_squared_error, mean_absolute_error, r2_score

TARGET_LABEL = "salaire_brut"
CSV_PATH = "outputs/payroll_long.csv"
TEST_MONTHS = 6  # last 6 effective observations (after lag) as test


def mape(y_true, y_pred):
    y_true = np.array(y_true)
    y_pred = np.array(y_pred)
    mask = y_true != 0
    if mask.sum() == 0:
        return np.nan
    return np.mean(np.abs((y_true[mask] - y_pred[mask]) / y_true[mask])) * 100.0


def load_and_prepare(csv_path: str) -> pd.DataFrame:
    df = pd.read_csv(csv_path)
    df["month"] = pd.to_datetime(df["month"], format="%Y-%m")
    df = df.sort_values(["employee_id", "month"]).reset_index(drop=True)
    return df


def run_ar1(df: pd.DataFrame):
    employees = df["employee_id"].unique()
    all_y_true = []
    all_y_pred = []
    metrics_per_employee = []
    preds_list = []

    for emp in employees:
        sub = df[df["employee_id"] == emp].copy()
        sub = sub.sort_values("month")
        n = len(sub)
        if n <= TEST_MONTHS + 3:
            continue

        # Lagged feature
        sub["y_lag1"] = sub[TARGET_LABEL].shift(1)
        sub = sub.dropna(subset=["y_lag1"]).copy()
        n_eff = len(sub)
        if n_eff <= TEST_MONTHS + 3:
            continue

        sub["t"] = np.arange(n_eff)

        X = sub[["t", "y_lag1"]]
        y = sub[TARGET_LABEL].values

        train_idx = n_eff - TEST_MONTHS
        X_train, X_test = X.iloc[:train_idx], X.iloc[train_idx:]
        y_train, y_test = y[:train_idx], y[train_idx:]

        model = LinearRegression()
        model.fit(X_train, y_train)
        y_pred = model.predict(X_test)

        all_y_true.extend(y_test.tolist())
        all_y_pred.extend(y_pred.tolist())

        emp_rmse = np.sqrt(mean_squared_error(y_test, y_pred))
        emp_mae = mean_absolute_error(y_test, y_pred)
        emp_mape = mape(y_test, y_pred)
        emp_r2 = r2_score(y_test, y_pred)

        metrics_per_employee.append(
            {
                "employee_id": emp,
                "rmse": emp_rmse,
                "mae": emp_mae,
                "mape": emp_mape,
                "r2": emp_r2,
                "n_test": len(y_test),
            }
        )

        pred_df = sub.iloc[train_idx:][["employee_id", "month"]].copy()
        pred_df["y_true"] = y_test
        pred_df["y_pred"] = y_pred
        preds_list.append(pred_df)

    all_y_true = np.array(all_y_true)
    all_y_pred = np.array(all_y_pred)

    global_rmse = np.sqrt(mean_squared_error(all_y_true, all_y_pred))
    global_mae = mean_absolute_error(all_y_true, all_y_pred)
    global_mape = mape(all_y_true, all_y_pred)
    global_r2 = r2_score(all_y_true, all_y_pred)

    metrics_df = pd.DataFrame(metrics_per_employee)
    preds_df = pd.concat(preds_list, ignore_index=True)

    global_metrics = {
        "rmse": global_rmse,
        "mae": global_mae,
        "mape": global_mape,
        "r2": global_r2,
    }

    return metrics_df, preds_df, global_metrics


if __name__ == "__main__":
    df = load_and_prepare(CSV_PATH)
    metrics_df, preds_df, global_metrics = run_ar1(df)

    print("=== Global metrics (AR(1) + Trend) on test period ===")
    for k, v in global_metrics.items():
        print(f"{k.upper()}: {v:.4f}")

    print("\n=== Per-employee metrics summary (mean over employees) ===")
    print(metrics_df[["rmse", "mae", "mape", "r2"]].mean())

    metrics_df.to_csv("results_ar1_emp_metrics.csv", index=False)
    preds_df.to_csv("results_ar1_predictions.csv", index=False)
    print("\nSaved: results_ar1_emp_metrics.csv")
    print("Saved: results_ar1_predictions.csv")
