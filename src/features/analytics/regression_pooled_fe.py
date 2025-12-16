"""
Pooled panel regression with employee fixed effects and month dummies:
y_{i,t} ~ t + employee_dummies + month_dummies
Target: salaire_brut

Evaluation:
- Global metrics: RMSE, MAE, MAPE, R^2 on the last TEST_MONTHS months
- Per-employee metrics on the test period
"""

import numpy as np
import pandas as pd
from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_squared_error, mean_absolute_error, r2_score

TARGET_LABEL = "salaire_brut"
CSV_PATH = "outputs/payroll_long.csv"
TEST_MONTHS = 6  # last 6 calendar months as test period (panel-wide)


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
    df = df.sort_values(["month", "employee_id"]).reset_index(drop=True)
    return df


def build_features(df: pd.DataFrame):
    # panel time index from global months
    unique_months = np.sort(df["month"].unique())
    month_to_t = {m: i for i, m in enumerate(unique_months)}
    df["t"] = df["month"].map(month_to_t)

    df["month_num"] = df["month"].dt.month

    emp_dummies = pd.get_dummies(df["employee_id"], prefix="emp", drop_first=True)
    month_dummies = pd.get_dummies(df["month_num"], prefix="m", drop_first=True)

    X = pd.concat([df[["t"]], emp_dummies, month_dummies], axis=1)
    y = df[TARGET_LABEL].values

    return X, y, unique_months


def run_pooled_fe(df: pd.DataFrame):
    X, y, unique_months = build_features(df)

    test_months = unique_months[-TEST_MONTHS:]
    train_months = unique_months[:-TEST_MONTHS]

    train_mask = df["month"].isin(train_months)
    test_mask = df["month"].isin(test_months)

    X_train, X_test = X[train_mask], X[test_mask]
    y_train, y_test = y[train_mask], y[test_mask]

    model = LinearRegression()
    model.fit(X_train, y_train)
    y_pred = model.predict(X_test)

    # Global metrics
    rmse = np.sqrt(mean_squared_error(y_test, y_pred))
    mae = mean_absolute_error(y_test, y_pred)
    mape_val = mape(y_test, y_pred)
    r2 = r2_score(y_test, y_pred)
    global_metrics = {
        "rmse": rmse,
        "mae": mae,
        "mape": mape_val,
        "r2": r2,
    }

    preds_df = df.loc[test_mask, ["employee_id", "month"]].copy()
    preds_df["y_true"] = y_test
    preds_df["y_pred"] = y_pred

    # Per-employee metrics on test months
    metrics_per_employee = []
    for emp, sub in preds_df.groupby("employee_id"):
        yt = sub["y_true"].values
        yp = sub["y_pred"].values
        if len(yt) < 2:
            continue
        emp_rmse = np.sqrt(mean_squared_error(yt, yp))
        emp_mae = mean_absolute_error(yt, yp)
        emp_mape = mape(yt, yp)
        emp_r2 = r2_score(yt, yp)
        metrics_per_employee.append(
            {
                "employee_id": emp,
                "rmse": emp_rmse,
                "mae": emp_mae,
                "mape": emp_mape,
                "r2": emp_r2,
                "n_test": len(yt),
            }
        )

    metrics_df = pd.DataFrame(metrics_per_employee)
    return preds_df, metrics_df, global_metrics


if __name__ == "__main__":
    df = load_and_prepare(CSV_PATH)
    preds_df, metrics_df, global_metrics = run_pooled_fe(df)

    print("=== Global metrics (Pooled FE) on test period ===")
    for k, v in global_metrics.items():
        print(f"{k.upper()}: {v:.4f}")

    print("\n=== Per-employee metrics summary (mean over employees) ===")
    print(metrics_df[["rmse", "mae", "mape", "r2"]].mean())

    preds_df.to_csv("results_pooled_fe_predictions.csv", index=False)
    metrics_df.to_csv("results_pooled_fe_emp_metrics.csv", index=False)
    print("\nSaved: results_pooled_fe_predictions.csv")
    print("Saved: results_pooled_fe_emp_metrics.csv")
