"""
HONEST Forecasting Model Evaluation
====================================
All models use validation-based hyperparameter selection for fair comparison.

Models:
- AR(1), AR(2): No hyperparameters (lag order is model definition)
- Ridge, Lasso: Validation-based alpha selection from [0.1, 1.0, 10.0, 100.0]
- Pooled FE: Validation-based alpha selection from [0, 0.1, 1.0, 10.0, 100.0]
- Unpooled: Validation-based alpha selection from [0, 0.1, 1.0, 10.0, 100.0]

Evaluation Methods:
1. Train/Val/Test Split
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

# Configuration
TARGET_LABEL = "salaire_brut"
# FIX_ME: Configuration - Update data path and test period if needed
# CSV_PATH: Location of the processed payroll data
# TEST_MONTHS: Number of months to reserve for testing evaluation
CSV_PATH = "outputs/payroll_long.csv"
TEST_MONTHS = 6
VAL_MONTHS = 3


# =============================
# Metrics
# =============================
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
# AR Models (No Hyperparameters)
# =============================
def ar_train_val_test(df, lags=2):
    """AR model with train/val/test split."""
    employees = df['employee_id'].unique()
    all_y_true, all_y_pred = [], []
    
    for emp in employees:
        sub = df[df['employee_id'] == emp].copy().sort_values('month')
        
        if len(sub) <= TEST_MONTHS + VAL_MONTHS + lags + 3:
            continue
        
        # Create lagged features
        for lag in range(1, lags + 1):
            sub[f'y_lag{lag}'] = sub[TARGET_LABEL].shift(lag)
        sub = sub.dropna().copy()
        
        if len(sub) <= TEST_MONTHS + VAL_MONTHS + 3:
            continue
        
        sub['t'] = np.arange(len(sub))
        lag_cols = [f'y_lag{i}' for i in range(1, lags + 1)]
        X = sub[['t'] + lag_cols]
        y = sub[TARGET_LABEL].values
        
        test_size = TEST_MONTHS
        val_size = VAL_MONTHS
        train_size = len(sub) - test_size - val_size
        
        X_train = X.iloc[:train_size]
        X_val = X.iloc[train_size:train_size+val_size]
        X_test = X.iloc[train_size+val_size:]
        y_train = y[:train_size]
        y_val = y[train_size:train_size+val_size]
        y_test = y[train_size+val_size:]
        
        scaler = StandardScaler()
        X_train_scaled = scaler.fit_transform(X_train)
        X_val_scaled = scaler.transform(X_val)
        X_test_scaled = scaler.transform(X_test)
        
        model = LinearRegression()
        X_train_val = np.vstack([X_train_scaled, X_val_scaled])
        y_train_val = np.concatenate([y_train, y_val])
        model.fit(X_train_val, y_train_val)
        
        y_pred = model.predict(X_test_scaled)
        all_y_true.extend(y_test)
        all_y_pred.extend(y_pred)
    
    return compute_metrics(all_y_true, all_y_pred)


def ar_recursive_multistep(df, lags=2):
    """AR model with RECURSIVE multi-step forecasting."""
    employees = df['employee_id'].unique()
    all_y_true, all_y_pred = [], []
    
    for emp in employees:
        sub = df[df['employee_id'] == emp].copy().sort_values('month')
        
        if len(sub) <= TEST_MONTHS + VAL_MONTHS + lags + 3:
            continue
        
        for lag in range(1, lags + 1):
            sub[f'y_lag{lag}'] = sub[TARGET_LABEL].shift(lag)
        sub = sub.dropna().copy()
        
        if len(sub) <= TEST_MONTHS + VAL_MONTHS + 3:
            continue
        
        sub['t'] = np.arange(len(sub))
        lag_cols = [f'y_lag{i}' for i in range(1, lags + 1)]
        X = sub[['t'] + lag_cols]
        y = sub[TARGET_LABEL].values
        
        test_size = TEST_MONTHS
        val_size = VAL_MONTHS
        train_size = len(sub) - test_size - val_size
        
        X_train = X.iloc[:train_size]
        X_val = X.iloc[train_size:train_size+val_size]
        y_train = y[:train_size]
        y_val = y[train_size:train_size+val_size]
        y_test = y[train_size+val_size:]
        
        scaler = StandardScaler()
        X_train_scaled = scaler.fit_transform(X_train)
        X_val_scaled = scaler.transform(X_val)
        
        model = LinearRegression()
        X_train_val = np.vstack([X_train_scaled, X_val_scaled])
        y_train_val = np.concatenate([y_train, y_val])
        model.fit(X_train_val, y_train_val)
        
        # RECURSIVE forecasting
        y_pred_recursive = []
        last_known_values = list(y[:train_size+val_size][-lags:])
        t_start = train_size + val_size
        
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
# Ridge/Lasso Models (Validation-Based Alpha Selection)
# =============================
def ridge_lasso_train_val_test(df, model_type='ridge'):
    """Ridge/Lasso with validation-based alpha selection."""
    employees = df['employee_id'].unique()
    all_y_true, all_y_pred = [], []
    
    alphas = [0.1, 1.0, 10.0, 100.0]
    
    for emp in employees:
        sub = df[df['employee_id'] == emp].copy().sort_values('month')
        
        if len(sub) <= TEST_MONTHS + VAL_MONTHS + 5:
            continue
        
        for lag in [1, 2, 3]:
            sub[f'y_lag{lag}'] = sub[TARGET_LABEL].shift(lag)
        sub = sub.dropna().copy()
        
        if len(sub) <= TEST_MONTHS + VAL_MONTHS + 3:
            continue
        
        sub['t'] = np.arange(len(sub))
        X = sub[['t', 'y_lag1', 'y_lag2', 'y_lag3']]
        y = sub[TARGET_LABEL].values
        
        test_size = TEST_MONTHS
        val_size = VAL_MONTHS
        train_size = len(sub) - test_size - val_size
        
        X_train = X.iloc[:train_size]
        X_val = X.iloc[train_size:train_size+val_size]
        X_test = X.iloc[train_size+val_size:]
        y_train = y[:train_size]
        y_val = y[train_size:train_size+val_size]
        y_test = y[train_size+val_size:]
        
        scaler = StandardScaler()
        X_train_scaled = scaler.fit_transform(X_train)
        X_val_scaled = scaler.transform(X_val)
        X_test_scaled = scaler.transform(X_test)
        
        # Validation-based alpha selection
        best_alpha = alphas[0]
        best_val_error = np.inf
        
        for alpha in alphas:
            if model_type == 'ridge':
                model = Ridge(alpha=alpha)
            else:
                model = Lasso(alpha=alpha)
            
            model.fit(X_train_scaled, y_train)
            y_val_pred = model.predict(X_val_scaled)
            val_error = nrmse(y_val, y_val_pred)
            
            if val_error < best_val_error:
                best_val_error = val_error
                best_alpha = alpha
        
        # Retrain with best alpha on train+val
        X_train_val = np.vstack([X_train_scaled, X_val_scaled])
        y_train_val = np.concatenate([y_train, y_val])
        
        if model_type == 'ridge':
            final_model = Ridge(alpha=best_alpha)
        else:
            final_model = Lasso(alpha=best_alpha)
        
        final_model.fit(X_train_val, y_train_val)
        y_pred = final_model.predict(X_test_scaled)
        
        all_y_true.extend(y_test)
        all_y_pred.extend(y_pred)
    
    return compute_metrics(all_y_true, all_y_pred)


def ridge_lasso_recursive_multistep(df, model_type='ridge'):
    """Ridge/Lasso with RECURSIVE forecasting and validation-based alpha selection."""
    employees = df['employee_id'].unique()
    all_y_true, all_y_pred = [], []
    
    alphas = [0.1, 1.0, 10.0, 100.0]
    
    for emp in employees:
        sub = df[df['employee_id'] == emp].copy().sort_values('month')
        
        if len(sub) <= TEST_MONTHS + VAL_MONTHS + 5:
            continue
        
        for lag in [1, 2, 3]:
            sub[f'y_lag{lag}'] = sub[TARGET_LABEL].shift(lag)
        sub = sub.dropna().copy()
        
        if len(sub) <= TEST_MONTHS + VAL_MONTHS + 3:
            continue
        
        sub['t'] = np.arange(len(sub))
        X = sub[['t', 'y_lag1', 'y_lag2', 'y_lag3']]
        y = sub[TARGET_LABEL].values
        
        test_size = TEST_MONTHS
        val_size = VAL_MONTHS
        train_size = len(sub) - test_size - val_size
        
        X_train = X.iloc[:train_size]
        X_val = X.iloc[train_size:train_size+val_size]
        y_train = y[:train_size]
        y_val = y[train_size:train_size+val_size]
        y_test = y[train_size+val_size:]
        
        scaler = StandardScaler()
        X_train_scaled = scaler.fit_transform(X_train)
        X_val_scaled = scaler.transform(X_val)
        
        # Validation-based alpha selection
        best_alpha = alphas[0]
        best_val_error = np.inf
        
        for alpha in alphas:
            if model_type == 'ridge':
                model = Ridge(alpha=alpha)
            else:
                model = Lasso(alpha=alpha)
            
            model.fit(X_train_scaled, y_train)
            y_val_pred = model.predict(X_val_scaled)
            val_error = nrmse(y_val, y_val_pred)
            
            if val_error < best_val_error:
                best_val_error = val_error
                best_alpha = alpha
        
        # Retrain with best alpha
        X_train_val = np.vstack([X_train_scaled, X_val_scaled])
        y_train_val = np.concatenate([y_train, y_val])
        
        if model_type == 'ridge':
            final_model = Ridge(alpha=best_alpha)
        else:
            final_model = Lasso(alpha=best_alpha)
        
        final_model.fit(X_train_val, y_train_val)
        
        # RECURSIVE forecasting
        y_pred_recursive = []
        last_known_values = list(y[:train_size+val_size][-3:])
        t_start = train_size + val_size
        
        for step in range(test_size):
            t_current = t_start + step
            features = [t_current] + last_known_values[::-1][:3]
            X_pred = np.array(features).reshape(1, -1)
            X_pred_scaled = scaler.transform(X_pred)
            
            y_next = final_model.predict(X_pred_scaled)[0]
            y_pred_recursive.append(y_next)
            
            last_known_values.append(y_next)
            last_known_values = last_known_values[1:]
        
        all_y_true.extend(y_test)
        all_y_pred.extend(y_pred_recursive)
    
    return compute_metrics(all_y_true, all_y_pred)


def ridge_lasso_time_series_cv(df, model_type='ridge', n_splits=3):
    """Ridge/Lasso with Time Series CV and validation-based alpha selection within each fold."""
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
        
        tscv = TimeSeriesSplit(n_splits=n_splits, test_size=TEST_MONTHS)
        
        for train_idx, test_idx in tscv.split(X):
            X_train_full, X_test = X.iloc[train_idx], X.iloc[test_idx]
            y_train_full, y_test = y[train_idx], y[test_idx]
            
            # Split train into train+val for alpha selection
            val_size = min(VAL_MONTHS, len(X_train_full) // 4)
            train_size = len(X_train_full) - val_size
            
            # Skip if not enough data for validation split
            if train_size < 3 or val_size < 1:
                continue
            
            X_train = X_train_full.iloc[:train_size]
            X_val = X_train_full.iloc[train_size:]
            y_train = y_train_full[:train_size]
            y_val = y_train_full[train_size:]
            
            if len(X_val) == 0:
                continue
            
            scaler = StandardScaler()
            X_train_scaled = scaler.fit_transform(X_train)
            X_val_scaled = scaler.transform(X_val)
            X_test_scaled = scaler.transform(X_test)
            
            # Select best alpha on validation set
            best_alpha = alphas[0]
            best_val_error = np.inf
            
            for alpha in alphas:
                if model_type == 'ridge':
                    model = Ridge(alpha=alpha)
                else:
                    model = Lasso(alpha=alpha)
                
                model.fit(X_train_scaled, y_train)
                y_val_pred = model.predict(X_val_scaled)
                val_error = nrmse(y_val, y_val_pred)
                
                if val_error < best_val_error:
                    best_val_error = val_error
                    best_alpha = alpha
            
            # Retrain with best alpha on full training set
            X_train_full_scaled = scaler.fit_transform(X_train_full)
            X_test_scaled = scaler.transform(X_test)
            
            if model_type == 'ridge':
                final_model = Ridge(alpha=best_alpha)
            else:
                final_model = Lasso(alpha=best_alpha)
            
            final_model.fit(X_train_full_scaled, y_train_full)
            y_pred = final_model.predict(X_test_scaled)
            
            fold_metrics.append(compute_metrics(y_test, y_pred))
    
    if not fold_metrics:
        return {k: np.nan for k in ['rmse', 'nrmse', 'mae', 'r2']}
    
    avg_metrics = {}
    for key in fold_metrics[0].keys():
        avg_metrics[key] = np.mean([m[key] for m in fold_metrics])
    
    return avg_metrics


# =============================
# Pooled FE (Validation-Based Alpha Selection)
# =============================
def pooled_fe_train_val_test(df):
    """Pooled FE with validation-based alpha selection."""
    unique_months = np.sort(df['month'].unique())
    n_months = len(unique_months)
    
    test_start_idx = n_months - TEST_MONTHS
    val_start_idx = test_start_idx - VAL_MONTHS
    
    df = df.sort_values(['employee_id', 'month']).copy()
    df['month_idx'] = df['month'].map({m: i for i, m in enumerate(unique_months)})
    
    train_df = df[df['month_idx'] < val_start_idx].copy()
    val_df = df[(df['month_idx'] >= val_start_idx) & (df['month_idx'] < test_start_idx)].copy()
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
    X_val = build_features(val_df)
    X_test = build_features(test_df)
    y_train = train_df[TARGET_LABEL].values
    y_val = val_df[TARGET_LABEL].values
    y_test = test_df[TARGET_LABEL].values
    
    # Align columns
    all_cols = list(set(X_train.columns) | set(X_val.columns) | set(X_test.columns))
    for col in all_cols:
        if col not in X_train.columns:
            X_train[col] = 0
        if col not in X_val.columns:
            X_val[col] = 0
        if col not in X_test.columns:
            X_test[col] = 0
    
    X_train = X_train[all_cols]
    X_val = X_val[all_cols]
    X_test = X_test[all_cols]
    
    # Validation-based alpha selection (including alpha=0 for OLS)
    alphas = [0, 0.1, 1.0, 10.0, 100.0]
    best_alpha = alphas[0]
    best_val_error = np.inf
    
    for alpha in alphas:
        if alpha == 0:
            model = LinearRegression()
        else:
            model = Ridge(alpha=alpha)
        
        model.fit(X_train, y_train)
        y_val_pred = model.predict(X_val)
        val_error = nrmse(y_val, y_val_pred)
        
        if val_error < best_val_error:
            best_val_error = val_error
            best_alpha = alpha
    
    # Retrain with best alpha on train+val
    X_train_val = pd.concat([X_train, X_val], ignore_index=True)
    y_train_val = np.concatenate([y_train, y_val])
    
    if best_alpha == 0:
        final_model = LinearRegression()
    else:
        final_model = Ridge(alpha=best_alpha)
    
    final_model.fit(X_train_val, y_train_val)
    y_pred = final_model.predict(X_test)
    
    return compute_metrics(y_test, y_pred)


def pooled_fe_recursive_multistep(df):
    """Pooled FE with RECURSIVE forecasting and validation-based alpha selection."""
    unique_months = np.sort(df['month'].unique())
    n_months = len(unique_months)
    
    test_start_idx = n_months - TEST_MONTHS
    val_start_idx = test_start_idx - VAL_MONTHS
    
    df = df.sort_values(['employee_id', 'month']).copy()
    df['month_idx'] = df['month'].map({m: i for i, m in enumerate(unique_months)})
    
    train_df = df[df['month_idx'] < val_start_idx].copy()
    val_df = df[(df['month_idx'] >= val_start_idx) & (df['month_idx'] < test_start_idx)].copy()
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
    X_val = build_features(val_df)
    y_train = train_df[TARGET_LABEL].values
    y_val = val_df[TARGET_LABEL].values
    
    # Align columns
    all_cols = list(set(X_train.columns) | set(X_val.columns))
    for col in all_cols:
        if col not in X_train.columns:
            X_train[col] = 0
        if col not in X_val.columns:
            X_val[col] = 0
    
    X_train = X_train[all_cols]
    X_val = X_val[all_cols]
    
    # Validation-based alpha selection
    alphas = [0, 0.1, 1.0, 10.0, 100.0]
    best_alpha = alphas[0]
    best_val_error = np.inf
    
    for alpha in alphas:
        if alpha == 0:
            model = LinearRegression()
        else:
            model = Ridge(alpha=alpha)
        
        model.fit(X_train, y_train)
        y_val_pred = model.predict(X_val)
        val_error = nrmse(y_val, y_val_pred)
        
        if val_error < best_val_error:
            best_val_error = val_error
            best_alpha = alpha
    
    # Retrain on train+val with best alpha
    X_train_val = pd.concat([X_train, X_val], ignore_index=True)
    y_train_val = np.concatenate([y_train, y_val])
    
    if best_alpha == 0:
        final_model = LinearRegression()
    else:
        final_model = Ridge(alpha=best_alpha)
    
    final_model.fit(X_train_val, y_train_val)
    
    # RECURSIVE forecasting (linear extrapolation)
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
# Unpooled Models (Validation-Based Alpha Selection)
# =============================
def unpooled_train_val_test(df):
    """Unpooled with validation-based alpha selection."""
    employees = df['employee_id'].unique()
    all_y_true, all_y_pred = [], []
    
    alphas = [0, 0.1, 1.0, 10.0, 100.0]
    
    for emp in employees:
        sub = df[df['employee_id'] == emp].copy().sort_values('month')
        
        if len(sub) <= TEST_MONTHS + VAL_MONTHS + 3:
            continue
        
        sub['t'] = np.arange(len(sub))
        X = sub[['t']]
        y = sub[TARGET_LABEL].values
        
        test_size = TEST_MONTHS
        val_size = VAL_MONTHS
        train_size = len(sub) - test_size - val_size
        
        X_train = X.iloc[:train_size]
        X_val = X.iloc[train_size:train_size+val_size]
        X_test = X.iloc[train_size+val_size:]
        y_train = y[:train_size]
        y_val = y[train_size:train_size+val_size]
        y_test = y[train_size+val_size:]
        
        # Validation-based alpha selection
        best_alpha = alphas[0]
        best_val_error = np.inf
        
        for alpha in alphas:
            if alpha == 0:
                model = LinearRegression()
            else:
                model = Ridge(alpha=alpha)
            
            model.fit(X_train, y_train)
            y_val_pred = model.predict(X_val)
            val_error = nrmse(y_val, y_val_pred)
            
            if val_error < best_val_error:
                best_val_error = val_error
                best_alpha = alpha
        
        # Retrain with best alpha on train+val
        X_train_val = pd.concat([X_train, X_val], ignore_index=True)
        y_train_val = np.concatenate([y_train, y_val])
        
        if best_alpha == 0:
            final_model = LinearRegression()
        else:
            final_model = Ridge(alpha=best_alpha)
        
        final_model.fit(X_train_val, y_train_val)
        y_pred = final_model.predict(X_test)
        
        all_y_true.extend(y_test)
        all_y_pred.extend(y_pred)
    
    return compute_metrics(all_y_true, all_y_pred)


def unpooled_recursive_multistep(df):
    """Unpooled with RECURSIVE forecasting and validation-based alpha selection."""
    employees = df['employee_id'].unique()
    all_y_true, all_y_pred = [], []
    
    alphas = [0, 0.1, 1.0, 10.0, 100.0]
    
    for emp in employees:
        sub = df[df['employee_id'] == emp].copy().sort_values('month')
        
        if len(sub) <= TEST_MONTHS + VAL_MONTHS + 3:
            continue
        
        sub['t'] = np.arange(len(sub))
        X = sub[['t']]
        y = sub[TARGET_LABEL].values
        
        test_size = TEST_MONTHS
        val_size = VAL_MONTHS
        train_size = len(sub) - test_size - val_size
        
        X_train = X.iloc[:train_size]
        X_val = X.iloc[train_size:train_size+val_size]
        y_train = y[:train_size]
        y_val = y[train_size:train_size+val_size]
        y_test = y[train_size+val_size:]
        
        # Validation-based alpha selection
        best_alpha = alphas[0]
        best_val_error = np.inf
        
        for alpha in alphas:
            if alpha == 0:
                model = LinearRegression()
            else:
                model = Ridge(alpha=alpha)
            
            model.fit(X_train, y_train)
            y_val_pred = model.predict(X_val)
            val_error = nrmse(y_val, y_val_pred)
            
            if val_error < best_val_error:
                best_val_error = val_error
                best_alpha = alpha
        
        # Retrain with best alpha
        X_train_val = pd.concat([X_train, X_val], ignore_index=True)
        y_train_val = np.concatenate([y_train, y_val])
        
        if best_alpha == 0:
            final_model = LinearRegression()
        else:
            final_model = Ridge(alpha=alpha)
        
        final_model.fit(X_train_val, y_train_val)
        
        # RECURSIVE forecasting (simple linear extrapolation)
        y_pred_recursive = []
        t_start = train_size + val_size
        
        for step in range(test_size):
            t_current = t_start + step
            X_pred = pd.DataFrame({'t': [t_current]})
            y_next = final_model.predict(X_pred)[0]
            y_pred_recursive.append(y_next)
        
        all_y_true.extend(y_test)
        all_y_pred.extend(y_pred_recursive)
    
    return compute_metrics(all_y_true, all_y_pred)


# =============================
# Main Evaluation
# =============================
def main():
    """Run honest evaluation with validation-based hyperparameter selection for all models."""
    print("=" * 80)
    print("HONEST FORECASTING EVALUATION")
    print("Validation-Based Hyperparameter Selection for ALL Models")
    print("=" * 80)
    
    df = load_data()
    print(f"\nLoaded {len(df)} records, {df['employee_id'].nunique()} employees\n")
    
    results = []
    
    # AR Models (no hyperparameters)
    print("Evaluating AR(1)...")
    results.append({
        'Model': 'AR(1)',
        'Method': 'Train/Val/Test',
        **ar_train_val_test(df, lags=1)
    })
    results.append({
        'Model': 'AR(1)',
        'Method': 'Recursive Multi-Step',
        **ar_recursive_multistep(df, lags=1)
    })
    results.append({
        'Model': 'AR(1)',
        'Method': 'Time Series CV',
        **ar_time_series_cv(df, lags=1)
    })
    
    print("Evaluating AR(2)...")
    results.append({
        'Model': 'AR(2)',
        'Method': 'Train/Val/Test',
        **ar_train_val_test(df, lags=2)
    })
    results.append({
        'Model': 'AR(2)',
        'Method': 'Recursive Multi-Step',
        **ar_recursive_multistep(df, lags=2)
    })
    results.append({
        'Model': 'AR(2)',
        'Method': 'Time Series CV',
        **ar_time_series_cv(df, lags=2)
    })
    
    # Ridge (validation-based alpha selection)
    print("Evaluating Ridge (validation-based alpha)...")
    results.append({
        'Model': 'Ridge',
        'Method': 'Train/Val/Test',
        **ridge_lasso_train_val_test(df, 'ridge')
    })
    results.append({
        'Model': 'Ridge',
        'Method': 'Recursive Multi-Step',
        **ridge_lasso_recursive_multistep(df, 'ridge')
    })
    results.append({
        'Model': 'Ridge',
        'Method': 'Time Series CV',
        **ridge_lasso_time_series_cv(df, 'ridge')
    })
    
    # Lasso (validation-based alpha selection)
    print("Evaluating Lasso (validation-based alpha)...")
    results.append({
        'Model': 'Lasso',
        'Method': 'Train/Val/Test',
        **ridge_lasso_train_val_test(df, 'lasso')
    })
    results.append({
        'Model': 'Lasso',
        'Method': 'Recursive Multi-Step',
        **ridge_lasso_recursive_multistep(df, 'lasso')
    })
    results.append({
        'Model': 'Lasso',
        'Method': 'Time Series CV',
        **ridge_lasso_time_series_cv(df, 'lasso')
    })
    
    # Pooled FE (validation-based alpha selection)
    print("Evaluating Pooled FE (validation-based alpha)...")
    results.append({
        'Model': 'Pooled FE',
        'Method': 'Train/Val/Test',
        **pooled_fe_train_val_test(df)
    })
    results.append({
        'Model': 'Pooled FE',
        'Method': 'Recursive Multi-Step',
        **pooled_fe_recursive_multistep(df)
    })
    results.append({
        'Model': 'Pooled FE',
        'Method': 'Time Series CV',
        'rmse': np.nan, 'nrmse': np.nan, 'mae': np.nan, 'r2': np.nan
    })
    
    # Unpooled (validation-based alpha selection)
    print("Evaluating Unpooled (validation-based alpha)...")
    results.append({
        'Model': 'Unpooled',
        'Method': 'Train/Val/Test',
        **unpooled_train_val_test(df)
    })
    results.append({
        'Model': 'Unpooled',
        'Method': 'Recursive Multi-Step',
        **unpooled_recursive_multistep(df)
    })
    results.append({
        'Model': 'Unpooled',
        'Method': 'Time Series CV',
        'rmse': np.nan, 'nrmse': np.nan, 'mae': np.nan, 'r2': np.nan
    })
    
    # Save results
    results_df = pd.DataFrame(results)
    results_df.to_csv('outputs/honest_forecasting_evaluation.csv', index=False)
    print(f"\n✓ Results saved to outputs/honest_forecasting_evaluation.csv")
    
    # Print summary
    print("\n" + "=" * 80)
    print("RESULTS SUMMARY")
    print("=" * 80)
    for method in ['Train/Val/Test', 'Recursive Multi-Step', 'Time Series CV']:
        print(f"\n{method}:")
        method_results = results_df[results_df['Method'] == method].copy()
        method_results = method_results.sort_values('nrmse')
        for _, row in method_results.iterrows():
            if not np.isnan(row['nrmse']):
                print(f"  {row['Model']:15s} - NRMSE: {row['nrmse']:.4f}, RMSE: {row['rmse']:.1f}, MAE: {row['mae']:.1f}")


if __name__ == '__main__':
    main()
