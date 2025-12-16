"""
TUNED Trend + Seasonality Models (Multiple Seasonal Specifications)
Target: salaire_brut

We compare 4 variants:
1. Monthly dummies (12-month seasonality)
2. Quarterly dummies (4-season quarter effects)
3. Fourier 1 harmonic (sin/cos k=1)
4. Fourier 2 harmonics (sin/cos k=1,2)

This version is FIXED to:
- Save prediction CSVs for each model
- Save metrics CSVs for each model
- Use consistent naming with visualize_all_models.py
"""

import numpy as np
import pandas as pd
from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_squared_error, mean_absolute_error, r2_score

TARGET_LABEL = "salaire_brut"
CSV_PATH = "outputs/payroll_long.csv"
TEST_MONTHS = 6


def mape(y_true, y_pred):
    y_true = np.array(y_true)
    y_pred = np.array(y_pred)
    mask = y_true != 0
    if mask.sum() == 0:
        return np.nan
    return np.mean(np.abs((y_true[mask] - y_pred[mask]) / y_true[mask])) * 100


def load_data(path):
    df = pd.read_csv(path)
    df["month"] = pd.to_datetime(df["month"], format="%Y-%m")
    return df.sort_values(["employee_id", "month"]).reset_index(drop=True)


# ---------------------------------------------------------------------
# FEATURE BUILDERS
# ---------------------------------------------------------------------

def seasonal_monthly_dummies(sub):
    dummies = pd.get_dummies(sub["month"].dt.month, prefix="m", drop_first=True)
    dummies.index = sub.index
    return dummies


def seasonal_quarter_dummies(sub):
    q = ((sub["month"].dt.month - 1) // 3) + 1
    dummies = pd.get_dummies(q, prefix="q", drop_first=True)
    dummies.index = sub.index
    return dummies


def seasonal_fourier(sub, K):
    t = np.arange(len(sub))
    feats = {}
    for k in range(1, K + 1):
        feats[f"sin{k}"] = np.sin(2 * np.pi * k * t / 12)
        feats[f"cos{k}"] = np.cos(2 * np.pi * k * t / 12)
    df = pd.DataFrame(feats, index=sub.index)
    return df


# ---------------------------------------------------------------------
# GENERIC MODEL RUNNER
# ---------------------------------------------------------------------

def run_model(df, model_name, seasonal_fn):
    employees = df["employee_id"].unique()

    all_y_true = []
    all_y_pred = []
    per_emp_metrics = []
    all_predictions = []   # <-- Needed for plotting

    for emp in employees:
        sub = df[df["employee_id"] == emp].copy()
        sub = sub.sort_values("month")
        n = len(sub)
        if n <= TEST_MONTHS + 3:
            continue

        sub["t"] = np.arange(n)

        seasonal_features = seasonal_fn(sub)
        X = pd.concat([sub[["t"]], seasonal_features], axis=1)
        y = sub[TARGET_LABEL].values

        # Drop bad rows
        y_series = pd.Series(y, index=X.index)
        mask_valid = X.notnull().all(axis=1) & y_series.notnull()
        X = X.loc[mask_valid]
        y = y_series.loc[mask_valid].values
        sub = sub.loc[mask_valid]

        n = len(X)
        if n <= TEST_MONTHS + 3:
            continue

        # Split
        train_idx = n - TEST_MONTHS
        X_train, X_test = X.iloc[:train_idx], X.iloc[train_idx:]
        y_train, y_test = y[:train_idx], y[train_idx:]
        sub_test = sub.iloc[train_idx:]

        # Fit
        model = LinearRegression()
        model.fit(X_train, y_train)
        y_pred = model.predict(X_test)

        # Store predictions for plotting
        pred_df = sub_test[["employee_id", "month"]].copy()
        pred_df["y_true"] = y_test
        pred_df["y_pred"] = y_pred
        all_predictions.append(pred_df)

        # Extend global lists
        all_y_true.extend(y_test.tolist())
        all_y_pred.extend(y_pred.tolist())

        # Employee metrics
        per_emp_metrics.append({
            "employee_id": emp,
            "rmse": np.sqrt(mean_squared_error(y_test, y_pred)),
            "mae": mean_absolute_error(y_test, y_pred),
            "mape": mape(y_test, y_pred),
            "r2": r2_score(y_test, y_pred),
            "n_test": len(y_test)
        })

    # Global metrics
    all_y_true = np.array(all_y_true)
    all_y_pred = np.array(all_y_pred)

    global_metrics = {
        "rmse": np.sqrt(mean_squared_error(all_y_true, all_y_pred)),
        "mae": mean_absolute_error(all_y_true, all_y_pred),
        "mape": mape(all_y_true, all_y_pred),
        "r2": r2_score(all_y_true, all_y_pred)
    }

    return pd.DataFrame(per_emp_metrics), pd.concat(all_predictions, ignore_index=True), global_metrics


# ---------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------

if __name__ == "__main__":
    df = load_data(CSV_PATH)

    models = {
        "monthly_dummies": seasonal_monthly_dummies,
        "quarterly_dummies": seasonal_quarter_dummies,
        "fourier_k1": lambda sub: seasonal_fourier(sub, K=1),
        "fourier_k2": lambda sub: seasonal_fourier(sub, K=2),
    }

    results = {}

    for name, fn in models.items():
        print(f"\n=== Running model: {name} ===")
        metrics_df, preds_df, global_metrics = run_model(df, name, fn)

        # Save results — MATCH plotting script
        metrics_df.to_csv(f"results_{name}_emp_metrics.csv", index=False)
        preds_df.to_csv(f"results_{name}_predictions.csv", index=False)

        print(f"Global Metrics for {name}:")
        for k, v in global_metrics.items():
            print(f"  {k.upper()}: {v:.4f}")

        print("Per-employee mean metrics:")
        print(metrics_df[["rmse", "mae", "mape", "r2"]].mean())

        results[name] = global_metrics

    # Summary
    print("\n\n=== MODEL COMPARISON SUMMARY ===")
    for name, met in results.items():
        print(f"{name}: RMSE={met['rmse']:.2f}, MAPE={met['mape']:.2f}, R2={met['r2']:.4f}")

    best = min(results, key=lambda x: results[x]["rmse"])
    print(f"\n>>> BEST MODEL BASED ON GLOBAL RMSE: {best}")
