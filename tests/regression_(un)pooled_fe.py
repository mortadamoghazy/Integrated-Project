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
from sklearn.linear_model import LinearRegression, Ridge, Lasso, RidgeCV, LassoCV
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


def run_pooled_ridge(df: pd.DataFrame, alphas=None):
    """Pooled model with Ridge regularization using cross-validation"""
    if alphas is None:
        alphas = [0.1, 1.0, 10.0, 100.0, 1000.0]
    
    X, y, unique_months = build_features(df)

    test_months = unique_months[-TEST_MONTHS:]
    train_months = unique_months[:-TEST_MONTHS]

    train_mask = df["month"].isin(train_months)
    test_mask = df["month"].isin(test_months)

    X_train, X_test = X[train_mask], X[test_mask]
    y_train, y_test = y[train_mask], y[test_mask]

    # Use RidgeCV for cross-validated alpha selection
    model = RidgeCV(alphas=alphas, cv=5)
    model.fit(X_train, y_train)
    y_pred = model.predict(X_test)
    
    best_alpha = model.alpha_

    # Global metrics
    global_metrics = {
        "rmse": np.sqrt(mean_squared_error(y_test, y_pred)),
        "mae": mean_absolute_error(y_test, y_pred),
        "mape": mape(y_test, y_pred),
        "r2": r2_score(y_test, y_pred),
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
        metrics_per_employee.append(
            {
                "employee_id": emp,
                "rmse": np.sqrt(mean_squared_error(yt, yp)),
                "mae": mean_absolute_error(yt, yp),
                "mape": mape(yt, yp),
                "r2": r2_score(yt, yp),
                "n_test": len(yt),
            }
        )

    metrics_df = pd.DataFrame(metrics_per_employee)
    return preds_df, metrics_df, global_metrics


def run_pooled_lasso(df: pd.DataFrame, alphas=None):
    """Pooled model with Lasso regularization using cross-validation"""
    if alphas is None:
        alphas = [0.1, 1.0, 10.0, 100.0, 1000.0]
    
    X, y, unique_months = build_features(df)

    test_months = unique_months[-TEST_MONTHS:]
    train_months = unique_months[:-TEST_MONTHS]

    train_mask = df["month"].isin(train_months)
    test_mask = df["month"].isin(test_months)

    X_train, X_test = X[train_mask], X[test_mask]
    y_train, y_test = y[train_mask], y[test_mask]

    # Use LassoCV for cross-validated alpha selection
    model = LassoCV(alphas=alphas, cv=5, max_iter=5000)
    model.fit(X_train, y_train)
    y_pred = model.predict(X_test)
    
    best_alpha = model.alpha_

    # Global metrics
    global_metrics = {
        "rmse": np.sqrt(mean_squared_error(y_test, y_pred)),
        "mae": mean_absolute_error(y_test, y_pred),
        "mape": mape(y_test, y_pred),
        "r2": r2_score(y_test, y_pred),
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
        metrics_per_employee.append(
            {
                "employee_id": emp,
                "rmse": np.sqrt(mean_squared_error(yt, yp)),
                "mae": mean_absolute_error(yt, yp),
                "mape": mape(yt, yp),
                "r2": r2_score(yt, yp),
                "n_test": len(yt),
            }
        )

    metrics_df = pd.DataFrame(metrics_per_employee)
    return preds_df, metrics_df, global_metrics


def run_unpooled(df: pd.DataFrame):
    all_y_true = []
    all_y_pred = []
    all_predictions = []
    per_emp_metrics = []

    # Global calendar time (same as pooled model)
    unique_months = np.sort(df["month"].unique())
    month_to_t = {m: i for i, m in enumerate(unique_months)}
    df["t"] = df["month"].map(month_to_t)
    df["month_num"] = df["month"].dt.month

    test_months = unique_months[-TEST_MONTHS:]
    train_months = unique_months[:-TEST_MONTHS]

    for emp, sub in df.groupby("employee_id"):
        sub = sub.sort_values("month")

        train_mask = sub["month"].isin(train_months)
        test_mask = sub["month"].isin(test_months)

        if train_mask.sum() < 6 or test_mask.sum() < 2:
            continue

        # Month dummies (employee-specific)
        month_dummies = pd.get_dummies(sub["month_num"], prefix="m", drop_first=True)
        X = pd.concat([sub[["t"]], month_dummies], axis=1)
        y = sub[TARGET_LABEL].values

        X_train = X.loc[train_mask]
        X_test = X.loc[test_mask]
        y_train = y[train_mask]
        y_test = y[test_mask]

        model = LinearRegression()
        model.fit(X_train, y_train)
        y_pred = model.predict(X_test)

        # Store predictions
        pred_df = sub.loc[test_mask, ["employee_id", "month"]].copy()
        pred_df["y_true"] = y_test
        pred_df["y_pred"] = y_pred
        all_predictions.append(pred_df)

        # Collect global values
        all_y_true.extend(y_test.tolist())
        all_y_pred.extend(y_pred.tolist())

        # Per-employee metrics
        per_emp_metrics.append({
            "employee_id": emp,
            "rmse": np.sqrt(mean_squared_error(y_test, y_pred)),
            "mae": mean_absolute_error(y_test, y_pred),
            "mape": mape(y_test, y_pred),
            "r2": r2_score(y_test, y_pred),
            "n_test": len(y_test),
        })

    # Global metrics
    all_y_true = np.array(all_y_true)
    all_y_pred = np.array(all_y_pred)

    global_metrics = {
        "rmse": np.sqrt(mean_squared_error(all_y_true, all_y_pred)),
        "mae": mean_absolute_error(all_y_true, all_y_pred),
        "mape": mape(all_y_true, all_y_pred),
        "r2": r2_score(all_y_true, all_y_pred),
    }

    preds_df = pd.concat(all_predictions, ignore_index=True)
    metrics_df = pd.DataFrame(per_emp_metrics)

    return preds_df, metrics_df, global_metrics


def run_unpooled_ridge(df: pd.DataFrame, alphas=None):
    """Unpooled model with Ridge regularization per employee"""
    if alphas is None:
        alphas = [0.1, 1.0, 10.0, 100.0, 1000.0]
    
    all_y_true = []
    all_y_pred = []
    all_predictions = []
    per_emp_metrics = []

    unique_months = np.sort(df["month"].unique())
    month_to_t = {m: i for i, m in enumerate(unique_months)}
    df["t"] = df["month"].map(month_to_t)
    df["month_num"] = df["month"].dt.month

    test_months = unique_months[-TEST_MONTHS:]
    train_months = unique_months[:-TEST_MONTHS]

    for emp, sub in df.groupby("employee_id"):
        sub = sub.sort_values("month")

        train_mask = sub["month"].isin(train_months)
        test_mask = sub["month"].isin(test_months)

        if train_mask.sum() < 6 or test_mask.sum() < 2:
            continue

        month_dummies = pd.get_dummies(sub["month_num"], prefix="m", drop_first=True)
        X = pd.concat([sub[["t"]], month_dummies], axis=1)
        y = sub[TARGET_LABEL].values

        X_train = X.loc[train_mask]
        X_test = X.loc[test_mask]
        y_train = y[train_mask]
        y_test = y[test_mask]

        # Use RidgeCV for cross-validated alpha selection
        model = RidgeCV(alphas=alphas, cv=min(5, len(X_train)))
        model.fit(X_train, y_train)
        y_pred = model.predict(X_test)

        pred_df = sub.loc[test_mask, ["employee_id", "month"]].copy()
        pred_df["y_true"] = y_test
        pred_df["y_pred"] = y_pred
        all_predictions.append(pred_df)

        all_y_true.extend(y_test.tolist())
        all_y_pred.extend(y_pred.tolist())

        per_emp_metrics.append({
            "employee_id": emp,
            "rmse": np.sqrt(mean_squared_error(y_test, y_pred)),
            "mae": mean_absolute_error(y_test, y_pred),
            "mape": mape(y_test, y_pred),
            "r2": r2_score(y_test, y_pred),
            "n_test": len(y_test),
        })

    all_y_true = np.array(all_y_true)
    all_y_pred = np.array(all_y_pred)

    global_metrics = {
        "rmse": np.sqrt(mean_squared_error(all_y_true, all_y_pred)),
        "mae": mean_absolute_error(all_y_true, all_y_pred),
        "mape": mape(all_y_true, all_y_pred),
        "r2": r2_score(all_y_true, all_y_pred),
    }

    preds_df = pd.concat(all_predictions, ignore_index=True)
    metrics_df = pd.DataFrame(per_emp_metrics)

    return preds_df, metrics_df, global_metrics


def run_unpooled_lasso(df: pd.DataFrame, alphas=None):
    """Unpooled model with Lasso regularization per employee"""
    if alphas is None:
        alphas = [0.1, 1.0, 10.0, 100.0, 1000.0]
    
    all_y_true = []
    all_y_pred = []
    all_predictions = []
    per_emp_metrics = []

    unique_months = np.sort(df["month"].unique())
    month_to_t = {m: i for i, m in enumerate(unique_months)}
    df["t"] = df["month"].map(month_to_t)
    df["month_num"] = df["month"].dt.month

    test_months = unique_months[-TEST_MONTHS:]
    train_months = unique_months[:-TEST_MONTHS]

    for emp, sub in df.groupby("employee_id"):
        sub = sub.sort_values("month")

        train_mask = sub["month"].isin(train_months)
        test_mask = sub["month"].isin(test_months)

        if train_mask.sum() < 6 or test_mask.sum() < 2:
            continue

        month_dummies = pd.get_dummies(sub["month_num"], prefix="m", drop_first=True)
        X = pd.concat([sub[["t"]], month_dummies], axis=1)
        y = sub[TARGET_LABEL].values

        X_train = X.loc[train_mask]
        X_test = X.loc[test_mask]
        y_train = y[train_mask]
        y_test = y[test_mask]

        # Use LassoCV for cross-validated alpha selection
        model = LassoCV(alphas=alphas, cv=min(5, len(X_train)), max_iter=5000)
        model.fit(X_train, y_train)
        y_pred = model.predict(X_test)

        pred_df = sub.loc[test_mask, ["employee_id", "month"]].copy()
        pred_df["y_true"] = y_test
        pred_df["y_pred"] = y_pred
        all_predictions.append(pred_df)

        all_y_true.extend(y_test.tolist())
        all_y_pred.extend(y_pred.tolist())

        per_emp_metrics.append({
            "employee_id": emp,
            "rmse": np.sqrt(mean_squared_error(y_test, y_pred)),
            "mae": mean_absolute_error(y_test, y_pred),
            "mape": mape(y_test, y_pred),
            "r2": r2_score(y_test, y_pred),
            "n_test": len(y_test),
        })

    all_y_true = np.array(all_y_true)
    all_y_pred = np.array(all_y_pred)

    global_metrics = {
        "rmse": np.sqrt(mean_squared_error(all_y_true, all_y_pred)),
        "mae": mean_absolute_error(all_y_true, all_y_pred),
        "mape": mape(all_y_true, all_y_pred),
        "r2": r2_score(all_y_true, all_y_pred),
    }

    preds_df = pd.concat(all_predictions, ignore_index=True)
    metrics_df = pd.DataFrame(per_emp_metrics)

    return preds_df, metrics_df, global_metrics


# ===============================
# MODEL COMPARISON: POOLED vs UNPOOLED, with and without REGULARIZATION
# ===============================

if __name__ == "__main__":
    df = load_and_prepare(CSV_PATH)

    results = []

    # 1. Pooled FE (no regularization)
    print("\n" + "="*60)
    print("1. POOLED FE (No Regularization)")
    print("="*60)
    preds, metrics, global_m = run_pooled_fe(df)
    print("\n=== Global metrics ===")
    for k, v in global_m.items():
        print(f"{k.upper()}: {v:.4f}")
    print("\n=== Per-employee metrics summary ===")
    print(metrics[["rmse", "mae", "mape", "r2"]].mean())
    preds.to_csv("results_pooled_fe_predictions.csv", index=False)
    metrics.to_csv("results_pooled_fe_emp_metrics.csv", index=False)
    results.append(("Pooled FE", global_m))

    # 2. Pooled Ridge
    print("\n" + "="*60)
    print("2. POOLED RIDGE (CV-optimized alpha)")
    print("="*60)
    preds, metrics, global_m = run_pooled_ridge(df)
    print("\n=== Global metrics ===")
    for k, v in global_m.items():
        print(f"{k.upper()}: {v:.4f}")
    print("\n=== Per-employee metrics summary ===")
    print(metrics[["rmse", "mae", "mape", "r2"]].mean())
    preds.to_csv("results_pooled_ridge_predictions.csv", index=False)
    metrics.to_csv("results_pooled_ridge_emp_metrics.csv", index=False)
    results.append(("Pooled Ridge", global_m))

    # 3. Pooled Lasso
    print("\n" + "="*60)
    print("3. POOLED LASSO (CV-optimized alpha)")
    print("="*60)
    preds, metrics, global_m = run_pooled_lasso(df)
    print("\n=== Global metrics ===")
    for k, v in global_m.items():
        print(f"{k.upper()}: {v:.4f}")
    print("\n=== Per-employee metrics summary ===")
    print(metrics[["rmse", "mae", "mape", "r2"]].mean())
    preds.to_csv("results_pooled_lasso_predictions.csv", index=False)
    metrics.to_csv("results_pooled_lasso_emp_metrics.csv", index=False)
    results.append(("Pooled Lasso", global_m))

    # 4. Unpooled
    print("\n" + "="*60)
    print("4. UNPOOLED (No Regularization)")
    print("="*60)
    preds, metrics, global_m = run_unpooled(df)
    print("\n=== Global metrics ===")
    for k, v in global_m.items():
        print(f"{k.upper()}: {v:.4f}")
    print("\n=== Per-employee metrics summary ===")
    print(metrics[["rmse", "mae", "mape", "r2"]].mean())
    preds.to_csv("results_unpooled_predictions.csv", index=False)
    metrics.to_csv("results_unpooled_emp_metrics.csv", index=False)
    results.append(("Unpooled", global_m))

    # 5. Unpooled Ridge
    print("\n" + "="*60)
    print("5. UNPOOLED RIDGE (CV-optimized alpha per employee)")
    print("="*60)
    preds, metrics, global_m = run_unpooled_ridge(df)
    print("\n=== Global metrics ===")
    for k, v in global_m.items():
        print(f"{k.upper()}: {v:.4f}")
    print("\n=== Per-employee metrics summary ===")
    print(metrics[["rmse", "mae", "mape", "r2"]].mean())
    preds.to_csv("results_unpooled_ridge_predictions.csv", index=False)
    metrics.to_csv("results_unpooled_ridge_emp_metrics.csv", index=False)
    results.append(("Unpooled Ridge", global_m))

    # 6. Unpooled Lasso
    print("\n" + "="*60)
    print("6. UNPOOLED LASSO (CV-optimized alpha per employee)")
    print("="*60)
    preds, metrics, global_m = run_unpooled_lasso(df)
    print("\n=== Global metrics ===")
    for k, v in global_m.items():
        print(f"{k.upper()}: {v:.4f}")
    print("\n=== Per-employee metrics summary ===")
    print(metrics[["rmse", "mae", "mape", "r2"]].mean())
    preds.to_csv("results_unpooled_lasso_predictions.csv", index=False)
    metrics.to_csv("results_unpooled_lasso_emp_metrics.csv", index=False)
    results.append(("Unpooled Lasso", global_m))

    # Summary comparison
    print("\n" + "="*60)
    print("SUMMARY COMPARISON")
    print("="*60)
    print(f"{'Model':<20} {'RMSE':<12} {'MAE':<12} {'MAPE':<12} {'R²':<12}")
    print("-" * 68)
    for name, metrics in results:
        print(f"{name:<20} {metrics['rmse']:<12.4f} {metrics['mae']:<12.4f} "
              f"{metrics['mape']:<12.4f} {metrics['r2']:<12.4f}")

    print("\n✅ All results saved to CSV files.")

