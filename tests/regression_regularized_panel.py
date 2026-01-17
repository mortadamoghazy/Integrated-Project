"""
Regularized pooled panel regression (Ridge & Lasso)
Same features as pooled FE:
y_{i,t} ~ t + employee_dummies + month_dummies
Target: salaire_brut

Evaluation:
- Global metrics (RMSE, MAE, MAPE, R^2) for Ridge and Lasso
- Per-employee metrics on the test period for both models
"""

import numpy as np
import pandas as pd
from sklearn.linear_model import Ridge, Lasso
from sklearn.metrics import mean_squared_error, mean_absolute_error, r2_score

TARGET_LABEL = "salaire_brut"
CSV_PATH = "outputs/payroll_long.csv"
TEST_MONTHS = 6     # last 6 months = outer test
VAL_MONTHS = 3      # last 3 months of train = inner validation
ALPHAS = [0.1, 1.0, 10.0, 100.0]


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
    unique_months = np.sort(df["month"].unique())
    month_to_t = {m: i for i, m in enumerate(unique_months)}
    df["t"] = df["month"].map(month_to_t)

    df["month_num"] = df["month"].dt.month

    emp_dummies = pd.get_dummies(df["employee_id"], prefix="emp", drop_first=True)
    month_dummies = pd.get_dummies(df["month_num"], prefix="m", drop_first=True)

    X = pd.concat([df[["t"]], emp_dummies, month_dummies], axis=1)
    y = df[TARGET_LABEL].values

    return X, y, unique_months


def select_alpha(model_class, X_train, y_train, df_train, train_months):
    unique_train_months = np.sort(train_months)
    if len(unique_train_months) <= VAL_MONTHS + 1:
        # not enough months for validation split
        return ALPHAS[1]

    val_months = unique_train_months[-VAL_MONTHS:]
    inner_train_months = unique_train_months[:-VAL_MONTHS]

    inner_train_mask = df_train["month"].isin(inner_train_months)
    val_mask = df_train["month"].isin(val_months)

    X_inner_train = X_train[inner_train_mask]
    y_inner_train = y_train[inner_train_mask]
    X_val = X_train[val_mask]
    y_val = y_train[val_mask]

    best_alpha = None
    best_rmse = np.inf

    for a in ALPHAS:
        model = model_class(alpha=a)
        model.fit(X_inner_train, y_inner_train)
        y_val_pred = model.predict(X_val)
        rmse = np.sqrt(mean_squared_error(y_val, y_val_pred))
        if rmse < best_rmse:
            best_rmse = rmse
            best_alpha = a

    return best_alpha


def run_regularized(df: pd.DataFrame):
    X, y, unique_months = build_features(df)

    test_months = unique_months[-TEST_MONTHS:]
    train_months = unique_months[:-TEST_MONTHS]

    train_mask = df["month"].isin(train_months)
    test_mask = df["month"].isin(test_months)

    X_train, X_test = X[train_mask], X[test_mask]
    y_train, y_test = y[train_mask], y[test_mask]
    df_train = df.loc[train_mask].copy()
    df_test = df.loc[test_mask].copy()

    # --- Ridge ---
    best_alpha_ridge = select_alpha(Ridge, X_train, y_train, df_train, train_months)
    ridge_model = Ridge(alpha=best_alpha_ridge)
    ridge_model.fit(X_train, y_train)
    y_pred_ridge = ridge_model.predict(X_test)

    ridge_metrics_global = {
        "rmse": np.sqrt(mean_squared_error(y_test, y_pred_ridge)),
        "mae": mean_absolute_error(y_test, y_pred_ridge),
        "mape": mape(y_test, y_pred_ridge),
        "r2": r2_score(y_test, y_pred_ridge),
        "alpha": best_alpha_ridge,
    }

    # --- Lasso ---
    best_alpha_lasso = select_alpha(Lasso, X_train, y_train, df_train, train_months)
    lasso_model = Lasso(alpha=best_alpha_lasso, max_iter=10000)
    lasso_model.fit(X_train, y_train)
    y_pred_lasso = lasso_model.predict(X_test)

    lasso_metrics_global = {
        "rmse": np.sqrt(mean_squared_error(y_test, y_pred_lasso)),
        "mae": mean_absolute_error(y_test, y_pred_lasso),
        "mape": mape(y_test, y_pred_lasso),
        "r2": r2_score(y_test, y_pred_lasso),
        "alpha": best_alpha_lasso,
    }

    preds_df = df_test[["employee_id", "month"]].copy()
    preds_df["y_true"] = y_test
    preds_df["y_pred_ridge"] = y_pred_ridge
    preds_df["y_pred_lasso"] = y_pred_lasso

    # Per-employee metrics for Ridge and Lasso
    ridge_emp_metrics = []
    lasso_emp_metrics = []

    for emp, sub in preds_df.groupby("employee_id"):
        yt = sub["y_true"].values
        yr = sub["y_pred_ridge"].values
        yl = sub["y_pred_lasso"].values
        if len(yt) < 2:
            continue

        # Ridge per-employee
        ridge_emp_metrics.append(
            {
                "employee_id": emp,
                "rmse": np.sqrt(mean_squared_error(yt, yr)),
                "mae": mean_absolute_error(yt, yr),
                "mape": mape(yt, yr),
                "r2": r2_score(yt, yr),
                "n_test": len(yt),
            }
        )

        # Lasso per-employee
        lasso_emp_metrics.append(
            {
                "employee_id": emp,
                "rmse": np.sqrt(mean_squared_error(yt, yl)),
                "mae": mean_absolute_error(yt, yl),
                "mape": mape(yt, yl),
                "r2": r2_score(yt, yl),
                "n_test": len(yt),
            }
        )

    ridge_emp_df = pd.DataFrame(ridge_emp_metrics)
    lasso_emp_df = pd.DataFrame(lasso_emp_metrics)

    return preds_df, ridge_emp_df, lasso_emp_df, ridge_metrics_global, lasso_metrics_global


if __name__ == "__main__":
    df = load_and_prepare(CSV_PATH)
    preds_df, ridge_emp_df, lasso_emp_df, ridge_global, lasso_global = run_regularized(df)

    print("=== Ridge (regularized panel) global metrics on test period ===")
    for k, v in ridge_global.items():
        if k == "alpha":
            print(f"{k.upper()}: {v}")
        else:
            print(f"{k.upper()}: {v:.4f}")

    print("\n=== Lasso (regularized panel) global metrics on test period ===")
    for k, v in lasso_global.items():
        if k == "alpha":
            print(f"{k.upper()}: {v}")
        else:
            print(f"{k.upper()}: {v:.4f}")

    print("\n=== Ridge per-employee metrics summary (mean over employees) ===")
    print(ridge_emp_df[["rmse", "mae", "mape", "r2"]].mean())

    print("\n=== Lasso per-employee metrics summary (mean over employees) ===")
    print(lasso_emp_df[["rmse", "mae", "mape", "r2"]].mean())

    preds_df.to_csv("results_regularized_panel_predictions.csv", index=False)
    ridge_emp_df.to_csv("results_regularized_panel_ridge_emp_metrics.csv", index=False)
    lasso_emp_df.to_csv("results_regularized_panel_lasso_emp_metrics.csv", index=False)

    print("\nSaved: results_regularized_panel_predictions.csv")
    print("Saved: results_regularized_panel_ridge_emp_metrics.csv")
    print("Saved: results_regularized_panel_lasso_emp_metrics.csv")
