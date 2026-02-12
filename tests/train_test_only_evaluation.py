"""
Train/Test Split ONLY Forecasting Evaluation
=============================================
No validation set. Hyperparameter selection via cross-validation on training set.

Models:
- AR(1), AR(2): No hyperparameters
- Ridge, Lasso: CV-based alpha selection on training set from [0.1, 1.0, 10.0, 100.0]
- Pooled FE: CV-based alpha selection on training set from [0, 0.1, 1.0, 10.0, 100.0]
- Unpooled: CV-based alpha selection on training set from [0, 0.1, 1.0, 10.0, 100.0]

Evaluation Methods:
1. Train/Test Split (one-step-ahead)
2. Recursive Multi-Step Forecasting
3. Time Series Cross-Validation
"""

import sys
from pathlib import Path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pandas as pd
import numpy as np
import warnings
warnings.filterwarnings('ignore')

from sklearn.linear_model import LinearRegression, Ridge, Lasso
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import mean_squared_error, mean_absolute_error, r2_score
from sklearn.model_selection import TimeSeriesSplit

# FIX_ME: Configuration - Update these paths if your data files are located elsewhere
# TARGET_LABEL: The column name in your dataset containing the salary/target variable
# CSV_PATH: Path to the processed payroll data in long format (employee-month observations)
# TEST_MONTHS: Number of months to reserve for testing (typically 6 months)
TARGET_LABEL = "salaire_brut"
CSV_PATH = "outputs/payroll_long.csv"
TEST_MONTHS = 6


def nrmse(y_true, y_pred):
    """Normalized RMSE."""
    rmse = np.sqrt(mean_squared_error(y_true, y_pred))
    mean_true = np.mean(y_true)
    return rmse / mean_true if mean_true != 0 else np.inf


def compute_metrics(y_true, y_pred):
    """Compute all metrics."""
    return {
        'rmse': np.sqrt(mean_squared_error(y_true, y_pred)),
        'nrmse': nrmse(y_true, y_pred),
        'mae': mean_absolute_error(y_true, y_pred),
        'r2': r2_score(y_true, y_pred)
    }


def load_data():
    """Load payroll data."""
    df = pd.read_csv(CSV_PATH)
    df['month'] = pd.to_datetime(df['month'])
    df = df.sort_values(['employee_id', 'month']).reset_index(drop=True)
    return df


# =============================
# AR Models
# =============================
def ar_train_test(df, lags=2):
    """AR model with train/test split."""
    employees = df['employee_id'].unique()
    all_y_true, all_y_pred = [], []
    
    for emp in employees:
        sub = df[df['employee_id'] == emp].copy().sort_values('month')
        
        if len(sub) <= TEST_MONTHS + lags + 3:
            continue
        
        for lag in range(1, lags + 1):
            sub[f'y_lag{lag}'] = sub[TARGET_LABEL].shift(lag)
        sub = sub.dropna().copy()
        
        if len(sub) <= TEST_MONTHS + 3:
            continue
        
        sub['t'] = np.arange(len(sub))
        lag_cols = [f'y_lag{i}' for i in range(1, lags + 1)]
        X = sub[['t'] + lag_cols]
        y = sub[TARGET_LABEL].values
        
        test_size = TEST_MONTHS
        train_size = len(sub) - test_size
        
        X_train = X.iloc[:train_size]
        X_test = X.iloc[train_size:]
        y_train = y[:train_size]
        y_test = y[train_size:]
        
        scaler = StandardScaler()
        X_train_scaled = scaler.fit_transform(X_train)
        X_test_scaled = scaler.transform(X_test)
        
        model = LinearRegression()
        model.fit(X_train_scaled, y_train)
        
        y_pred = model.predict(X_test_scaled)
        all_y_true.extend(y_test)
        all_y_pred.extend(y_pred)
    
    return compute_metrics(all_y_true, all_y_pred)


def ar_recursive(df, lags=2):
    """AR model with recursive multi-step forecasting."""
    employees = df['employee_id'].unique()
    all_y_true, all_y_pred = [], []
    
    for emp in employees:
        sub = df[df['employee_id'] == emp].copy().sort_values('month')
        
        if len(sub) <= TEST_MONTHS + lags + 3:
            continue
        
        for lag in range(1, lags + 1):
            sub[f'y_lag{lag}'] = sub[TARGET_LABEL].shift(lag)
        sub = sub.dropna().copy()
        
        if len(sub) <= TEST_MONTHS + 3:
            continue
        
        sub['t'] = np.arange(len(sub))
        lag_cols = [f'y_lag{i}' for i in range(1, lags + 1)]
        X = sub[['t'] + lag_cols]
        y = sub[TARGET_LABEL].values
        
        test_size = TEST_MONTHS
        train_size = len(sub) - test_size
        
        X_train = X.iloc[:train_size]
        y_train = y[:train_size]
        y_test = y[train_size:]
        
        scaler = StandardScaler()
        X_train_scaled = scaler.fit_transform(X_train)
        
        model = LinearRegression()
        model.fit(X_train_scaled, y_train)
        
        # Recursive forecasting
        y_pred_recursive = []
        last_known_values = list(y[:train_size][-lags:])
        t_start = train_size
        
        for step in range(test_size):
            t_current = t_start + step
            features = [t_current] + last_known_values[::-1][:lags]
            X_pred = np.array(features).reshape(1, -1)
            X_pred_scaled = scaler.transform(X_pred)
            
            y_next = model.predict(X_pred_scaled)[0]
            y_pred_recursive.append(y_next)
            
            last_known_values.append(y_next)
            last_known_values = last_known_values[1:]
        
        all_y_true.extend(y_test)
        all_y_pred.extend(y_pred_recursive)
    
    return compute_metrics(all_y_true, all_y_pred)


def ar_time_series_cv(df, lags=2, n_splits=3):
    """AR model with time series CV."""
    employees = df['employee_id'].unique()
    fold_metrics = []
    
    for emp in employees:
        sub = df[df['employee_id'] == emp].copy().sort_values('month')
        
        if len(sub) <= 20:
            continue
        
        for lag in range(1, lags + 1):
            sub[f'y_lag{lag}'] = sub[TARGET_LABEL].shift(lag)
        sub = sub.dropna().copy()
        
        if len(sub) <= 15:
            continue
        
        sub['t'] = np.arange(len(sub))
        lag_cols = [f'y_lag{i}' for i in range(1, lags + 1)]
        X = sub[['t'] + lag_cols]
        y = sub[TARGET_LABEL].values
        
        tscv = TimeSeriesSplit(n_splits=n_splits, test_size=TEST_MONTHS)
        
        for train_idx, test_idx in tscv.split(X):
            X_train, X_test = X.iloc[train_idx], X.iloc[test_idx]
            y_train, y_test = y[train_idx], y[test_idx]
            
            scaler = StandardScaler()
            X_train_scaled = scaler.fit_transform(X_train)
            X_test_scaled = scaler.transform(X_test)
            
            model = LinearRegression()
            model.fit(X_train_scaled, y_train)
            y_pred = model.predict(X_test_scaled)
            
            fold_metrics.append(compute_metrics(y_test, y_pred))
    
    if not fold_metrics:
        return {k: np.nan for k in ['rmse', 'nrmse', 'mae', 'r2']}
    
    avg_metrics = {}
    for key in fold_metrics[0].keys():
        avg_metrics[key] = np.mean([m[key] for m in fold_metrics])
    
    return avg_metrics


# =============================
# Ridge/Lasso Models
# =============================
def ridge_lasso_train_test(df, model_type='ridge'):
    """Ridge/Lasso pooled panel model with employee and month fixed effects."""
    # Prepare data with month index
    df = df.copy()
    df['month'] = pd.to_datetime(df['month'])
    df = df.sort_values(['employee_id', 'month'])
    
    # Create month index (sequential time variable)
    df['month_idx'] = (df['month'] - df['month'].min()).dt.days // 30
    
    # Create calendar month for seasonality
    df['calendar_month'] = df['month'].dt.month
    
    # Split data by time
    unique_months = df['month_idx'].unique()
    unique_months = np.sort(unique_months)
    train_months = unique_months[:-TEST_MONTHS]
    test_months = unique_months[-TEST_MONTHS:]
    
    train_data = df[df['month_idx'].isin(train_months)].copy()
    test_data = df[df['month_idx'].isin(test_months)].copy()
    
    if len(train_data) < 50:
        return {'rmse': np.inf, 'nrmse': np.inf, 'mae': np.inf, 'r2': -np.inf}
    
    # Create employee dummies
    emp_dummies_train = pd.get_dummies(train_data['employee_id'], prefix='emp', drop_first=True)
    emp_dummies_test = pd.get_dummies(test_data['employee_id'], prefix='emp', drop_first=True)
    
    # Create month dummies for seasonality
    month_dummies_train = pd.get_dummies(train_data['calendar_month'], prefix='month', drop_first=True)
    month_dummies_test = pd.get_dummies(test_data['calendar_month'], prefix='month', drop_first=True)
    
    # Align columns
    for col in emp_dummies_train.columns:
        if col not in emp_dummies_test.columns:
            emp_dummies_test[col] = 0
    for col in emp_dummies_test.columns:
        if col not in emp_dummies_train.columns:
            emp_dummies_train[col] = 0
    emp_dummies_test = emp_dummies_test[emp_dummies_train.columns]
    
    for col in month_dummies_train.columns:
        if col not in month_dummies_test.columns:
            month_dummies_test[col] = 0
    for col in month_dummies_test.columns:
        if col not in month_dummies_train.columns:
            month_dummies_train[col] = 0
    month_dummies_test = month_dummies_test[month_dummies_train.columns]
    
    # Build feature matrices: time trend + employee dummies + month dummies
    X_train = pd.concat([
        train_data[['month_idx']].reset_index(drop=True),
        emp_dummies_train.reset_index(drop=True),
        month_dummies_train.reset_index(drop=True)
    ], axis=1)
    
    X_test = pd.concat([
        test_data[['month_idx']].reset_index(drop=True),
        emp_dummies_test.reset_index(drop=True),
        month_dummies_test.reset_index(drop=True)
    ], axis=1)
    
    y_train = train_data[TARGET_LABEL].values
    y_test = test_data[TARGET_LABEL].values
    
    # CV for alpha selection
    alphas = [0.1, 1.0, 10.0, 100.0]
    tscv = TimeSeriesSplit(n_splits=3)
    best_alpha = alphas[0]
    best_cv_score = np.inf
    
    for alpha in alphas:
        cv_scores = []
        for train_idx, val_idx in tscv.split(X_train):
            X_cv_train = X_train.iloc[train_idx]
            X_cv_val = X_train.iloc[val_idx]
            y_cv_train = y_train[train_idx]
            y_cv_val = y_train[val_idx]
            
            scaler = StandardScaler()
            X_cv_train_scaled = scaler.fit_transform(X_cv_train)
            X_cv_val_scaled = scaler.transform(X_cv_val)
            
            if model_type == 'ridge':
                model = Ridge(alpha=alpha)
            else:
                model = Lasso(alpha=alpha)
            
            model.fit(X_cv_train_scaled, y_cv_train)
            y_cv_pred = model.predict(X_cv_val_scaled)
            cv_scores.append(nrmse(y_cv_val, y_cv_pred))
        
        avg_cv_score = np.mean(cv_scores)
        if avg_cv_score < best_cv_score:
            best_cv_score = avg_cv_score
            best_alpha = alpha
    
    # Train final model
    scaler = StandardScaler()
    X_train_scaled = scaler.fit_transform(X_train)
    X_test_scaled = scaler.transform(X_test)
    
    if model_type == 'ridge':
        final_model = Ridge(alpha=best_alpha)
    else:
        final_model = Lasso(alpha=best_alpha)
    
    final_model.fit(X_train_scaled, y_train)
    y_pred = final_model.predict(X_test_scaled)
    
    return compute_metrics(y_test, y_pred)


def ridge_lasso_recursive(df, model_type='ridge'):
    """Ridge/Lasso pooled panel model with recursive month extrapolation."""
    # Prepare data with month index
    df = df.copy()
    df['month'] = pd.to_datetime(df['month'])
    df = df.sort_values(['employee_id', 'month'])
    
    # Create month index (sequential time variable)
    df['month_idx'] = (df['month'] - df['month'].min()).dt.days // 30
    
    # Create calendar month for seasonality
    df['calendar_month'] = df['month'].dt.month
    
    # Split data by time
    unique_months = df['month_idx'].unique()
    unique_months = np.sort(unique_months)
    train_months = unique_months[:-TEST_MONTHS]
    test_months = unique_months[-TEST_MONTHS:]
    
    train_data = df[df['month_idx'].isin(train_months)].copy()
    test_data = df[df['month_idx'].isin(test_months)].copy()
    
    if len(train_data) < 50:
        return {'rmse': np.inf, 'nrmse': np.inf, 'mae': np.inf, 'r2': -np.inf}
    
    # Create employee dummies
    emp_dummies_train = pd.get_dummies(train_data['employee_id'], prefix='emp', drop_first=True)
    
    # Create month dummies for seasonality
    month_dummies_train = pd.get_dummies(train_data['calendar_month'], prefix='month', drop_first=True)
    
    # Build feature matrices
    X_train = pd.concat([
        train_data[['month_idx']].reset_index(drop=True),
        emp_dummies_train.reset_index(drop=True),
        month_dummies_train.reset_index(drop=True)
    ], axis=1)
    
    y_train = train_data[TARGET_LABEL].values
    
    # CV for alpha selection
    alphas = [0.1, 1.0, 10.0, 100.0]
    tscv = TimeSeriesSplit(n_splits=3)
    best_alpha = alphas[0]
    best_cv_score = np.inf
    
    for alpha in alphas:
        cv_scores = []
        for train_idx, val_idx in tscv.split(X_train):
            X_cv_train = X_train.iloc[train_idx]
            X_cv_val = X_train.iloc[val_idx]
            y_cv_train = y_train[train_idx]
            y_cv_val = y_train[val_idx]
            
            scaler = StandardScaler()
            X_cv_train_scaled = scaler.fit_transform(X_cv_train)
            X_cv_val_scaled = scaler.transform(X_cv_val)
            
            if model_type == 'ridge':
                model = Ridge(alpha=alpha)
            else:
                model = Lasso(alpha=alpha)
            
            model.fit(X_cv_train_scaled, y_cv_train)
            y_cv_pred = model.predict(X_cv_val_scaled)
            cv_scores.append(nrmse(y_cv_val, y_cv_pred))
        
        avg_cv_score = np.mean(cv_scores)
        if avg_cv_score < best_cv_score:
            best_cv_score = avg_cv_score
            best_alpha = alpha
    
    # Train final model
    scaler = StandardScaler()
    X_train_scaled = scaler.fit_transform(X_train)
    
    if model_type == 'ridge':
        final_model = Ridge(alpha=best_alpha)
    else:
        final_model = Lasso(alpha=best_alpha)
    
    final_model.fit(X_train_scaled, y_train)
    
    # Recursive forecasting: predict future months for all employees
    all_y_pred = []
    all_y_true = []
    
    for future_month_idx in test_months:
        # Get all employees in this month
        month_test_data = test_data[test_data['month_idx'] == future_month_idx].copy()
        
        if len(month_test_data) == 0:
            continue
        
        # Create employee dummies for test
        emp_dummies_test = pd.get_dummies(month_test_data['employee_id'], prefix='emp', drop_first=True)
        
        # Align with training columns
        for col in emp_dummies_train.columns:
            if col not in emp_dummies_test.columns:
                emp_dummies_test[col] = 0
        emp_dummies_test = emp_dummies_test[emp_dummies_train.columns]
        
        # Create month dummies
        month_dummies_test = pd.get_dummies(month_test_data['calendar_month'], prefix='month', drop_first=True)
        for col in month_dummies_train.columns:
            if col not in month_dummies_test.columns:
                month_dummies_test[col] = 0
        month_dummies_test = month_dummies_test[month_dummies_train.columns]
        
        # Build feature matrix
        X_test = pd.concat([
            month_test_data[['month_idx']].reset_index(drop=True),
            emp_dummies_test.reset_index(drop=True),
            month_dummies_test.reset_index(drop=True)
        ], axis=1)
        
        X_test_scaled = scaler.transform(X_test)
        y_pred_month = final_model.predict(X_test_scaled)
        
        all_y_pred.extend(y_pred_month)
        all_y_true.extend(month_test_data[TARGET_LABEL].values)
    
    return compute_metrics(all_y_true, all_y_pred)


def ridge_lasso_time_series_cv(df, model_type='ridge', n_splits=3):
    """Ridge/Lasso with time series CV and nested CV for alpha selection."""
    employees = df['employee_id'].unique()
    fold_metrics = []
    
    alphas = [0.1, 1.0, 10.0, 100.0]
    
    for emp in employees:
        sub = df[df['employee_id'] == emp].copy().sort_values('month')
        
        if len(sub) <= 20:
            continue
        
        for lag in [1, 2, 3]:
            sub[f'y_lag{lag}'] = sub[TARGET_LABEL].shift(lag)
        sub = sub.dropna().copy()
        
        if len(sub) <= 15:
            continue
        
        sub['t'] = np.arange(len(sub))
        X = sub[['t', 'y_lag1', 'y_lag2', 'y_lag3']]
        y = sub[TARGET_LABEL].values
        
        outer_cv = TimeSeriesSplit(n_splits=n_splits, test_size=TEST_MONTHS)
        
        for train_idx, test_idx in outer_cv.split(X):
            X_train_full = X.iloc[train_idx]
            X_test = X.iloc[test_idx]
            y_train_full = y[train_idx]
            y_test = y[test_idx]
            
            # Nested CV for alpha selection
            inner_cv = TimeSeriesSplit(n_splits=2)
            best_alpha = alphas[0]
            best_cv_score = np.inf
            
            for alpha in alphas:
                cv_scores = []
                for inner_train_idx, inner_val_idx in inner_cv.split(X_train_full):
                    X_cv_train = X_train_full.iloc[inner_train_idx]
                    X_cv_val = X_train_full.iloc[inner_val_idx]
                    y_cv_train = y_train_full[inner_train_idx]
                    y_cv_val = y_train_full[inner_val_idx]
                    
                    if len(X_cv_val) == 0:
                        continue
                    
                    scaler = StandardScaler()
                    X_cv_train_scaled = scaler.fit_transform(X_cv_train)
                    X_cv_val_scaled = scaler.transform(X_cv_val)
                    
                    if model_type == 'ridge':
                        model = Ridge(alpha=alpha)
                    else:
                        model = Lasso(alpha=alpha)
                    
                    model.fit(X_cv_train_scaled, y_cv_train)
                    y_cv_pred = model.predict(X_cv_val_scaled)
                    cv_scores.append(nrmse(y_cv_val, y_cv_pred))
                
                if cv_scores:
                    avg_cv_score = np.mean(cv_scores)
                    if avg_cv_score < best_cv_score:
                        best_cv_score = avg_cv_score
                        best_alpha = alpha
            
            # Train with best alpha on full outer training set
            scaler = StandardScaler()
            X_train_scaled = scaler.fit_transform(X_train_full)
            X_test_scaled = scaler.transform(X_test)
            
            if model_type == 'ridge':
                final_model = Ridge(alpha=best_alpha)
            else:
                final_model = Lasso(alpha=best_alpha)
            
            final_model.fit(X_train_scaled, y_train_full)
            y_pred = final_model.predict(X_test_scaled)
            
            fold_metrics.append(compute_metrics(y_test, y_pred))
    
    if not fold_metrics:
        return {k: np.nan for k in ['rmse', 'nrmse', 'mae', 'r2']}
    
    avg_metrics = {}
    for key in fold_metrics[0].keys():
        avg_metrics[key] = np.mean([m[key] for m in fold_metrics])
    
    return avg_metrics


# =============================
# Pooled FE
# =============================
def pooled_fe_train_test(df):
    """Pooled FE with CV-based alpha selection on training set."""
    unique_months = np.sort(df['month'].unique())
    n_months = len(unique_months)
    
    test_start_idx = n_months - TEST_MONTHS
    
    df = df.sort_values(['employee_id', 'month']).copy()
    df['month_idx'] = df['month'].map({m: i for i, m in enumerate(unique_months)})
    
    train_df = df[df['month_idx'] < test_start_idx].copy()
    test_df = df[df['month_idx'] >= test_start_idx].copy()
    
    if len(train_df) == 0 or len(test_df) == 0:
        return {k: np.nan for k in ['rmse', 'nrmse', 'mae', 'r2']}
    
    def build_features(data):
        emp_dummies = pd.get_dummies(data['employee_id'], prefix='emp', drop_first=True)
        features = pd.concat([
            data[['month_idx']],
            emp_dummies
        ], axis=1)
        return features
    
    X_train = build_features(train_df)
    X_test = build_features(test_df)
    y_train = train_df[TARGET_LABEL].values
    y_test = test_df[TARGET_LABEL].values
    
    # Align columns
    all_cols = list(set(X_train.columns) | set(X_test.columns))
    for col in all_cols:
        if col not in X_train.columns:
            X_train[col] = 0
        if col not in X_test.columns:
            X_test[col] = 0
    
    X_train = X_train[all_cols]
    X_test = X_test[all_cols]
    
    # CV on training set for alpha selection
    alphas = [0, 0.1, 1.0, 10.0, 100.0]
    tscv = TimeSeriesSplit(n_splits=3)
    
    # Need to split train_df by time for CV
    train_months = sorted(train_df['month_idx'].unique())
    best_alpha = alphas[0]
    best_cv_score = np.inf
    
    for alpha in alphas:
        cv_scores = []
        for train_months_idx, val_months_idx in tscv.split(train_months):
            train_months_cv = [train_months[i] for i in train_months_idx]
            val_months_cv = [train_months[i] for i in val_months_idx]
            
            train_cv_df = train_df[train_df['month_idx'].isin(train_months_cv)]
            val_cv_df = train_df[train_df['month_idx'].isin(val_months_cv)]
            
            X_train_cv = build_features(train_cv_df)
            X_val_cv = build_features(val_cv_df)
            y_train_cv = train_cv_df[TARGET_LABEL].values
            y_val_cv = val_cv_df[TARGET_LABEL].values
            
            # Align columns
            cv_cols = list(set(X_train_cv.columns) | set(X_val_cv.columns))
            for col in cv_cols:
                if col not in X_train_cv.columns:
                    X_train_cv[col] = 0
                if col not in X_val_cv.columns:
                    X_val_cv[col] = 0
            
            X_train_cv = X_train_cv[cv_cols]
            X_val_cv = X_val_cv[cv_cols]
            
            if alpha == 0:
                model = LinearRegression()
            else:
                model = Ridge(alpha=alpha)
            
            model.fit(X_train_cv, y_train_cv)
            y_val_pred = model.predict(X_val_cv)
            cv_scores.append(nrmse(y_val_cv, y_val_pred))
        
        avg_cv_score = np.mean(cv_scores)
        if avg_cv_score < best_cv_score:
            best_cv_score = avg_cv_score
            best_alpha = alpha
    
    # Train with best alpha on full training set
    if best_alpha == 0:
        final_model = LinearRegression()
    else:
        final_model = Ridge(alpha=best_alpha)
    
    final_model.fit(X_train, y_train)
    y_pred = final_model.predict(X_test)
    
    return compute_metrics(y_test, y_pred)


def pooled_fe_recursive(df):
    """Pooled FE with recursive forecasting and CV-based alpha selection."""
    unique_months = np.sort(df['month'].unique())
    n_months = len(unique_months)
    
    test_start_idx = n_months - TEST_MONTHS
    
    df = df.sort_values(['employee_id', 'month']).copy()
    df['month_idx'] = df['month'].map({m: i for i, m in enumerate(unique_months)})
    
    train_df = df[df['month_idx'] < test_start_idx].copy()
    test_df = df[df['month_idx'] >= test_start_idx].copy()
    
    if len(train_df) == 0 or len(test_df) == 0:
        return {k: np.nan for k in ['rmse', 'nrmse', 'mae', 'r2']}
    
    def build_features(data):
        emp_dummies = pd.get_dummies(data['employee_id'], prefix='emp', drop_first=True)
        features = pd.concat([
            data[['month_idx']],
            emp_dummies
        ], axis=1)
        return features
    
    X_train = build_features(train_df)
    y_train = train_df[TARGET_LABEL].values
    
    # CV for alpha selection (same as train_test)
    alphas = [0, 0.1, 1.0, 10.0, 100.0]
    tscv = TimeSeriesSplit(n_splits=3)
    
    train_months = sorted(train_df['month_idx'].unique())
    best_alpha = alphas[0]
    best_cv_score = np.inf
    
    for alpha in alphas:
        cv_scores = []
        for train_months_idx, val_months_idx in tscv.split(train_months):
            train_months_cv = [train_months[i] for i in train_months_idx]
            val_months_cv = [train_months[i] for i in val_months_idx]
            
            train_cv_df = train_df[train_df['month_idx'].isin(train_months_cv)]
            val_cv_df = train_df[train_df['month_idx'].isin(val_months_cv)]
            
            X_train_cv = build_features(train_cv_df)
            X_val_cv = build_features(val_cv_df)
            y_train_cv = train_cv_df[TARGET_LABEL].values
            y_val_cv = val_cv_df[TARGET_LABEL].values
            
            cv_cols = list(set(X_train_cv.columns) | set(X_val_cv.columns))
            for col in cv_cols:
                if col not in X_train_cv.columns:
                    X_train_cv[col] = 0
                if col not in X_val_cv.columns:
                    X_val_cv[col] = 0
            
            X_train_cv = X_train_cv[cv_cols]
            X_val_cv = X_val_cv[cv_cols]
            
            if alpha == 0:
                model = LinearRegression()
            else:
                model = Ridge(alpha=alpha)
            
            model.fit(X_train_cv, y_train_cv)
            y_val_pred = model.predict(X_val_cv)
            cv_scores.append(nrmse(y_val_cv, y_val_pred))
        
        avg_cv_score = np.mean(cv_scores)
        if avg_cv_score < best_cv_score:
            best_cv_score = avg_cv_score
            best_alpha = alpha
    
    # Train with best alpha
    all_cols = X_train.columns
    
    if best_alpha == 0:
        final_model = LinearRegression()
    else:
        final_model = Ridge(alpha=best_alpha)
    
    final_model.fit(X_train, y_train)
    
    # Recursive forecasting (linear extrapolation)
    test_months_idx = sorted(test_df['month_idx'].unique())
    all_y_true, all_y_pred = [], []
    
    for month_idx in test_months_idx:
        month_test = test_df[test_df['month_idx'] == month_idx]
        X_month = build_features(month_test)
        
        for col in all_cols:
            if col not in X_month.columns:
                X_month[col] = 0
        X_month = X_month[all_cols]
        
        y_pred = final_model.predict(X_month)
        y_true = month_test[TARGET_LABEL].values
        
        all_y_true.extend(y_true)
        all_y_pred.extend(y_pred)
    
    return compute_metrics(all_y_true, all_y_pred)


# =============================
# Unpooled
# =============================
def unpooled_train_test(df):
    """Unpooled with CV-based alpha selection on training set."""
    employees = df['employee_id'].unique()
    all_y_true, all_y_pred = [], []
    
    alphas = [0, 0.1, 1.0, 10.0, 100.0]
    
    for emp in employees:
        sub = df[df['employee_id'] == emp].copy().sort_values('month')
        
        if len(sub) <= TEST_MONTHS + 3:
            continue
        
        sub['t'] = np.arange(len(sub))
        X = sub[['t']]
        y = sub[TARGET_LABEL].values
        
        test_size = TEST_MONTHS
        train_size = len(sub) - test_size
        
        X_train = X.iloc[:train_size]
        X_test = X.iloc[train_size:]
        y_train = y[:train_size]
        y_test = y[train_size:]
        
        # CV on training set for alpha selection
        tscv = TimeSeriesSplit(n_splits=3)
        best_alpha = alphas[0]
        best_cv_score = np.inf
        
        for alpha in alphas:
            cv_scores = []
            for train_idx, val_idx in tscv.split(X_train):
                X_cv_train = X_train.iloc[train_idx]
                X_cv_val = X_train.iloc[val_idx]
                y_cv_train = y_train[train_idx]
                y_cv_val = y_train[val_idx]
                
                if alpha == 0:
                    model = LinearRegression()
                else:
                    model = Ridge(alpha=alpha)
                
                model.fit(X_cv_train, y_cv_train)
                y_cv_pred = model.predict(X_cv_val)
                cv_scores.append(nrmse(y_cv_val, y_cv_pred))
            
            avg_cv_score = np.mean(cv_scores)
            if avg_cv_score < best_cv_score:
                best_cv_score = avg_cv_score
                best_alpha = alpha
        
        # Train with best alpha on full training set
        if best_alpha == 0:
            final_model = LinearRegression()
        else:
            final_model = Ridge(alpha=best_alpha)
        
        final_model.fit(X_train, y_train)
        y_pred = final_model.predict(X_test)
        
        all_y_true.extend(y_test)
        all_y_pred.extend(y_pred)
    
    return compute_metrics(all_y_true, all_y_pred)


def unpooled_recursive(df):
    """Unpooled with recursive forecasting and CV-based alpha selection."""
    employees = df['employee_id'].unique()
    all_y_true, all_y_pred = [], []
    
    alphas = [0, 0.1, 1.0, 10.0, 100.0]
    
    for emp in employees:
        sub = df[df['employee_id'] == emp].copy().sort_values('month')
        
        if len(sub) <= TEST_MONTHS + 3:
            continue
        
        sub['t'] = np.arange(len(sub))
        X = sub[['t']]
        y = sub[TARGET_LABEL].values
        
        test_size = TEST_MONTHS
        train_size = len(sub) - test_size
        
        X_train = X.iloc[:train_size]
        y_train = y[:train_size]
        y_test = y[train_size:]
        
        # CV on training set for alpha selection
        tscv = TimeSeriesSplit(n_splits=3)
        best_alpha = alphas[0]
        best_cv_score = np.inf
        
        for alpha in alphas:
            cv_scores = []
            for train_idx, val_idx in tscv.split(X_train):
                X_cv_train = X_train.iloc[train_idx]
                X_cv_val = X_train.iloc[val_idx]
                y_cv_train = y_train[train_idx]
                y_cv_val = y_train[val_idx]
                
                if alpha == 0:
                    model = LinearRegression()
                else:
                    model = Ridge(alpha=alpha)
                
                model.fit(X_cv_train, y_cv_train)
                y_cv_pred = model.predict(X_cv_val)
                cv_scores.append(nrmse(y_cv_val, y_cv_pred))
            
            avg_cv_score = np.mean(cv_scores)
            if avg_cv_score < best_cv_score:
                best_cv_score = avg_cv_score
                best_alpha = alpha
        
        # Train with best alpha
        if best_alpha == 0:
            final_model = LinearRegression()
        else:
            final_model = Ridge(alpha=best_alpha)
        
        final_model.fit(X_train, y_train)
        
        # Recursive forecasting (linear extrapolation)
        y_pred_recursive = []
        t_start = train_size
        
        for step in range(test_size):
            t_current = t_start + step
            X_pred = pd.DataFrame({'t': [t_current]})
            y_next = final_model.predict(X_pred)[0]
            y_pred_recursive.append(y_next)
        
        all_y_true.extend(y_test)
        all_y_pred.extend(y_pred_recursive)
    
    return compute_metrics(all_y_true, all_y_pred)


# =============================
# AR(2) + Regularization Models
# =============================
def ar2_ridge_train_test(df):
    """AR(2) model with Ridge regularization."""
    employees = df['employee_id'].unique()
    all_y_true, all_y_pred = [], []
    alphas = [0.1, 1.0, 10.0, 100.0]
    
    for emp in employees:
        sub = df[df['employee_id'] == emp].copy().sort_values('month')
        if len(sub) <= TEST_MONTHS + 5:
            continue
        
        # Create lags
        for lag in [1, 2]:
            sub[f'y_lag{lag}'] = sub[TARGET_LABEL].shift(lag)
        sub = sub.dropna().copy()
        if len(sub) <= TEST_MONTHS + 3:
            continue
        
        sub['t'] = np.arange(len(sub))
        X = sub[['t', 'y_lag1', 'y_lag2']]
        y = sub[TARGET_LABEL].values
        
        train_size = len(sub) - TEST_MONTHS
        X_train, X_test = X.iloc[:train_size], X.iloc[train_size:]
        y_train, y_test = y[:train_size], y[train_size:]
        
        # CV for alpha selection
        tscv = TimeSeriesSplit(n_splits=3)
        best_alpha, best_score = alphas[0], np.inf
        for alpha in alphas:
            cv_scores = []
            for train_idx, val_idx in tscv.split(X_train):
                scaler = StandardScaler()
                X_cv_tr = scaler.fit_transform(X_train.iloc[train_idx])
                X_cv_val = scaler.transform(X_train.iloc[val_idx])
                model = Ridge(alpha=alpha)
                model.fit(X_cv_tr, y_train[train_idx])
                cv_scores.append(nrmse(y_train[val_idx], model.predict(X_cv_val)))
            if np.mean(cv_scores) < best_score:
                best_score = np.mean(cv_scores)
                best_alpha = alpha
        
        # Train final model
        scaler = StandardScaler()
        X_train_sc = scaler.fit_transform(X_train)
        X_test_sc = scaler.transform(X_test)
        model = Ridge(alpha=best_alpha)
        model.fit(X_train_sc, y_train)
        all_y_pred.extend(model.predict(X_test_sc))
        all_y_true.extend(y_test)
    
    return compute_metrics(all_y_true, all_y_pred)


def ar2_ridge_recursive(df):
    """AR(2) with Ridge regularization and recursive forecasting."""
    employees = df['employee_id'].unique()
    all_y_true, all_y_pred = [], []
    alphas = [0.1, 1.0, 10.0, 100.0]
    
    for emp in employees:
        sub = df[df['employee_id'] == emp].copy().sort_values('month')
        if len(sub) <= TEST_MONTHS + 5:
            continue
        
        for lag in [1, 2]:
            sub[f'y_lag{lag}'] = sub[TARGET_LABEL].shift(lag)
        sub = sub.dropna().copy()
        if len(sub) <= TEST_MONTHS + 3:
            continue
        
        sub['t'] = np.arange(len(sub))
        X = sub[['t', 'y_lag1', 'y_lag2']]
        y = sub[TARGET_LABEL].values
        
        train_size = len(sub) - TEST_MONTHS
        X_train = X.iloc[:train_size]
        y_train = y[:train_size]
        y_test = y[train_size:]
        
        # CV for alpha
        tscv = TimeSeriesSplit(n_splits=3)
        best_alpha, best_score = alphas[0], np.inf
        for alpha in alphas:
            cv_scores = []
            for train_idx, val_idx in tscv.split(X_train):
                scaler = StandardScaler()
                X_cv_tr = scaler.fit_transform(X_train.iloc[train_idx])
                X_cv_val = scaler.transform(X_train.iloc[val_idx])
                model = Ridge(alpha=alpha)
                model.fit(X_cv_tr, y_train[train_idx])
                cv_scores.append(nrmse(y_train[val_idx], model.predict(X_cv_val)))
            if np.mean(cv_scores) < best_score:
                best_score = np.mean(cv_scores)
                best_alpha = alpha
        
        # Train and recursive forecast
        scaler = StandardScaler()
        X_train_sc = scaler.fit_transform(X_train)
        model = Ridge(alpha=best_alpha)
        model.fit(X_train_sc, y_train)
        
        last_vals = list(y[:train_size][-2:])
        y_pred_rec = []
        for step in range(TEST_MONTHS):
            features = [train_size + step] + last_vals[::-1]
            X_pred = scaler.transform(np.array(features).reshape(1, -1))
            y_next = model.predict(X_pred)[0]
            y_pred_rec.append(y_next)
            last_vals = [last_vals[1], y_next]
        
        all_y_pred.extend(y_pred_rec)
        all_y_true.extend(y_test)
    
    return compute_metrics(all_y_true, all_y_pred)


def ar2_lasso_train_test(df):
    """AR(2) model with Lasso regularization."""
    employees = df['employee_id'].unique()
    all_y_true, all_y_pred = [], []
    alphas = [0.1, 1.0, 10.0, 100.0]
    
    for emp in employees:
        sub = df[df['employee_id'] == emp].copy().sort_values('month')
        if len(sub) <= TEST_MONTHS + 5:
            continue
        
        for lag in [1, 2]:
            sub[f'y_lag{lag}'] = sub[TARGET_LABEL].shift(lag)
        sub = sub.dropna().copy()
        if len(sub) <= TEST_MONTHS + 3:
            continue
        
        sub['t'] = np.arange(len(sub))
        X = sub[['t', 'y_lag1', 'y_lag2']]
        y = sub[TARGET_LABEL].values
        
        train_size = len(sub) - TEST_MONTHS
        X_train, X_test = X.iloc[:train_size], X.iloc[train_size:]
        y_train, y_test = y[:train_size], y[train_size:]
        
        # CV for alpha
        tscv = TimeSeriesSplit(n_splits=3)
        best_alpha, best_score = alphas[0], np.inf
        for alpha in alphas:
            cv_scores = []
            for train_idx, val_idx in tscv.split(X_train):
                scaler = StandardScaler()
                X_cv_tr = scaler.fit_transform(X_train.iloc[train_idx])
                X_cv_val = scaler.transform(X_train.iloc[val_idx])
                model = Lasso(alpha=alpha, max_iter=10000)
                model.fit(X_cv_tr, y_train[train_idx])
                cv_scores.append(nrmse(y_train[val_idx], model.predict(X_cv_val)))
            if np.mean(cv_scores) < best_score:
                best_score = np.mean(cv_scores)
                best_alpha = alpha
        
        # Train final
        scaler = StandardScaler()
        X_train_sc = scaler.fit_transform(X_train)
        X_test_sc = scaler.transform(X_test)
        model = Lasso(alpha=best_alpha, max_iter=10000)
        model.fit(X_train_sc, y_train)
        all_y_pred.extend(model.predict(X_test_sc))
        all_y_true.extend(y_test)
    
    return compute_metrics(all_y_true, all_y_pred)


def ar2_lasso_recursive(df):
    """AR(2) with Lasso regularization and recursive forecasting."""
    employees = df['employee_id'].unique()
    all_y_true, all_y_pred = [], []
    alphas = [0.1, 1.0, 10.0, 100.0]
    
    for emp in employees:
        sub = df[df['employee_id'] == emp].copy().sort_values('month')
        if len(sub) <= TEST_MONTHS + 5:
            continue
        
        for lag in [1, 2]:
            sub[f'y_lag{lag}'] = sub[TARGET_LABEL].shift(lag)
        sub = sub.dropna().copy()
        if len(sub) <= TEST_MONTHS + 3:
            continue
        
        sub['t'] = np.arange(len(sub))
        X = sub[['t', 'y_lag1', 'y_lag2']]
        y = sub[TARGET_LABEL].values
        
        train_size = len(sub) - TEST_MONTHS
        X_train = X.iloc[:train_size]
        y_train = y[:train_size]
        y_test = y[train_size:]
        
        # CV for alpha
        tscv = TimeSeriesSplit(n_splits=3)
        best_alpha, best_score = alphas[0], np.inf
        for alpha in alphas:
            cv_scores = []
            for train_idx, val_idx in tscv.split(X_train):
                scaler = StandardScaler()
                X_cv_tr = scaler.fit_transform(X_train.iloc[train_idx])
                X_cv_val = scaler.transform(X_train.iloc[val_idx])
                model = Lasso(alpha=alpha, max_iter=10000)
                model.fit(X_cv_tr, y_train[train_idx])
                cv_scores.append(nrmse(y_train[val_idx], model.predict(X_cv_val)))
            if np.mean(cv_scores) < best_score:
                best_score = np.mean(cv_scores)
                best_alpha = alpha
        
        # Train and recursive forecast
        scaler = StandardScaler()
        X_train_sc = scaler.fit_transform(X_train)
        model = Lasso(alpha=best_alpha, max_iter=10000)
        model.fit(X_train_sc, y_train)
        
        last_vals = list(y[:train_size][-2:])
        y_pred_rec = []
        for step in range(TEST_MONTHS):
            features = [train_size + step] + last_vals[::-1]
            X_pred = scaler.transform(np.array(features).reshape(1, -1))
            y_next = model.predict(X_pred)[0]
            y_pred_rec.append(y_next)
            last_vals = [last_vals[1], y_next]
        
        all_y_pred.extend(y_pred_rec)
        all_y_true.extend(y_test)
    
    return compute_metrics(all_y_true, all_y_pred)


# =============================
# Main
# =============================
def main():
    """Run train/test only evaluation."""
    print("=" * 80)
    print("TRAIN/TEST ONLY EVALUATION")
    print("CV-Based Hyperparameter Selection on Training Set")
    print("=" * 80)
    
    df = load_data()
    print(f"\nLoaded {len(df)} records, {df['employee_id'].nunique()} employees\n")
    
    results = []
    
    # AR Models
    print("Evaluating AR(1)...")
    results.append({'Model': 'AR(1)', 'Method': 'Train/Test', **ar_train_test(df, lags=1)})
    results.append({'Model': 'AR(1)', 'Method': 'Recursive Multi-Step', **ar_recursive(df, lags=1)})
    results.append({'Model': 'AR(1)', 'Method': 'Time Series CV', **ar_time_series_cv(df, lags=1)})
    
    print("Evaluating AR(2)...")
    results.append({'Model': 'AR(2)', 'Method': 'Train/Test', **ar_train_test(df, lags=2)})
    results.append({'Model': 'AR(2)', 'Method': 'Recursive Multi-Step', **ar_recursive(df, lags=2)})
    results.append({'Model': 'AR(2)', 'Method': 'Time Series CV', **ar_time_series_cv(df, lags=2)})
    
    # AR(2) + Ridge
    print("Evaluating AR(2)+Ridge...")
    results.append({'Model': 'AR(2)+Ridge', 'Method': 'Train/Test', **ar2_ridge_train_test(df)})
    results.append({'Model': 'AR(2)+Ridge', 'Method': 'Recursive Multi-Step', **ar2_ridge_recursive(df)})
    
    # AR(2) + Lasso
    print("Evaluating AR(2)+Lasso...")
    results.append({'Model': 'AR(2)+Lasso', 'Method': 'Train/Test', **ar2_lasso_train_test(df)})
    results.append({'Model': 'AR(2)+Lasso', 'Method': 'Recursive Multi-Step', **ar2_lasso_recursive(df)})
    
    # Ridge (Panel)
    print("Evaluating Ridge (CV-based alpha)...")
    results.append({'Model': 'Ridge', 'Method': 'Train/Test', **ridge_lasso_train_test(df, 'ridge')})
    results.append({'Model': 'Ridge', 'Method': 'Recursive Multi-Step', **ridge_lasso_recursive(df, 'ridge')})
    results.append({'Model': 'Ridge', 'Method': 'Time Series CV', **ridge_lasso_time_series_cv(df, 'ridge')})
    
    # Lasso (Panel)
    print("Evaluating Lasso (CV-based alpha)...")
    results.append({'Model': 'Lasso', 'Method': 'Train/Test', **ridge_lasso_train_test(df, 'lasso')})
    results.append({'Model': 'Lasso', 'Method': 'Recursive Multi-Step', **ridge_lasso_recursive(df, 'lasso')})
    results.append({'Model': 'Lasso', 'Method': 'Time Series CV', **ridge_lasso_time_series_cv(df, 'lasso')})
    
    # Pooled FE
    print("Evaluating Pooled FE (CV-based alpha)...")
    results.append({'Model': 'Pooled FE', 'Method': 'Train/Test', **pooled_fe_train_test(df)})
    results.append({'Model': 'Pooled FE', 'Method': 'Recursive Multi-Step', **pooled_fe_recursive(df)})
    results.append({'Model': 'Pooled FE', 'Method': 'Time Series CV', 'rmse': np.nan, 'nrmse': np.nan, 'mae': np.nan, 'r2': np.nan})
    
    # Unpooled
    print("Evaluating Unpooled (CV-based alpha)...")
    results.append({'Model': 'Unpooled', 'Method': 'Train/Test', **unpooled_train_test(df)})
    results.append({'Model': 'Unpooled', 'Method': 'Recursive Multi-Step', **unpooled_recursive(df)})
    results.append({'Model': 'Unpooled', 'Method': 'Time Series CV', 'rmse': np.nan, 'nrmse': np.nan, 'mae': np.nan, 'r2': np.nan})
    
    # Save results
    results_df = pd.DataFrame(results)
    results_df.to_csv('outputs/train_test_only_evaluation.csv', index=False)
    print(f"\n✓ Results saved to outputs/train_test_only_evaluation.csv")
    
    # Print summary
    print("\n" + "=" * 80)
    print("RESULTS SUMMARY")
    print("=" * 80)
    for method in ['Train/Test', 'Recursive Multi-Step', 'Time Series CV']:
        print(f"\n{method}:")
        method_results = results_df[results_df['Method'] == method].copy()
        method_results = method_results.sort_values('nrmse')
        for _, row in method_results.iterrows():
            if not np.isnan(row['nrmse']):
                print(f"  {row['Model']:15s} - NRMSE: {row['nrmse']:.4f}, RMSE: {row['rmse']:.1f}, MAE: {row['mae']:.1f}")


if __name__ == '__main__':
    main()
