"""
Simplified Forecasting Model Comparison

Models:
- AR(1), AR(2), AR(3), AR(4)
- Ridge, Lasso
- Pooled Fixed Effects, Unpooled
- XGBoost

Evaluation Methods:
1. Train/Validation/Test Split
2. Time Series Cross-Validation
3. Nested Cross-Validation (with hyperparameter tuning)
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

try:
    from xgboost import XGBRegressor
    XGBOOST_AVAILABLE = True
except ImportError:
    XGBOOST_AVAILABLE = False
    print("⚠️  XGBoost not installed")


# =============================
# Configuration
# =============================
TARGET_LABEL = "salaire_brut"
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


# =============================
# Data Loading
# =============================
def load_data():
    """Load payroll data."""
    df = pd.read_csv(CSV_PATH)
    df['month'] = pd.to_datetime(df['month'])
    df = df.sort_values(['employee_id', 'month']).reset_index(drop=True)
    return df


# =============================
# AR Models
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
        lag_cols = [f'y_lag{lag}' for lag in range(1, lags + 1)]
        feature_cols = ['t'] + lag_cols
        X = sub[feature_cols].copy()
        y = sub[TARGET_LABEL].values
        
        # Split
        test_size = TEST_MONTHS
        val_size = VAL_MONTHS
        train_size = len(sub) - test_size - val_size
        
        X_train, X_val, X_test = X.iloc[:train_size], X.iloc[train_size:train_size+val_size], X.iloc[train_size+val_size:]
        y_train, y_val, y_test = y[:train_size], y[train_size:train_size+val_size], y[train_size+val_size:]
        
        model = LinearRegression()
        model.fit(X_train, y_train)
        y_pred = model.predict(X_test)
        
        all_y_true.extend(y_test)
        all_y_pred.extend(y_pred)
    
    return compute_metrics(all_y_true, all_y_pred)


def ar_time_series_cv(df, lags=2, n_splits=3):
    """AR model with time series CV."""
    employees = df['employee_id'].unique()
    fold_metrics = []
    
    for emp in employees:
        sub = df[df['employee_id'] == emp].copy().sort_values('month')
        
        if len(sub) <= 20:
            continue
        
        # Create lagged features
        for lag in range(1, lags + 1):
            sub[f'y_lag{lag}'] = sub[TARGET_LABEL].shift(lag)
        sub = sub.dropna().copy()
        
        if len(sub) <= 15:
            continue
        
        sub['t'] = np.arange(len(sub))
        lag_cols = [f'y_lag{lag}' for lag in range(1, lags + 1)]
        feature_cols = ['t'] + lag_cols
        X = sub[feature_cols].copy()
        y = sub[TARGET_LABEL].values
        
        tscv = TimeSeriesSplit(n_splits=n_splits, test_size=TEST_MONTHS)
        
        for train_idx, test_idx in tscv.split(X):
            X_train, X_test = X.iloc[train_idx], X.iloc[test_idx]
            y_train, y_test = y[train_idx], y[test_idx]
            
            model = LinearRegression()
            model.fit(X_train, y_train)
            y_pred = model.predict(X_test)
            
            fold_metrics.append(compute_metrics(y_test, y_pred))
    
    if not fold_metrics:
        return {k: np.nan for k in ['rmse', 'nrmse', 'mae', 'r2']}
    
    avg_metrics = {}
    for key in fold_metrics[0].keys():
        avg_metrics[key] = np.mean([m[key] for m in fold_metrics])
    
    return avg_metrics


def ar_nested_cv(df, lags=2, outer_splits=3, inner_splits=2):
    """AR model with nested CV (no hyperparameters, just for robust estimation)."""
    return ar_time_series_cv(df, lags, n_splits=outer_splits)


def ar_recursive_multistep(df, lags=2):
    """AR model with RECURSIVE multi-step forecasting (predictions feed back as inputs)."""
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
        lag_cols = [f'y_lag{lag}' for lag in range(1, lags + 1)]
        feature_cols = ['t'] + lag_cols
        X = sub[feature_cols].copy()
        y = sub[TARGET_LABEL].values
        
        # Split
        test_size = TEST_MONTHS
        val_size = VAL_MONTHS
        train_size = len(sub) - test_size - val_size
        
        X_train = X.iloc[:train_size]
        y_train = y[:train_size]
        y_test = y[train_size+val_size:]
        
        # Train model
        model = LinearRegression()
        model.fit(X_train, y_train)
        
        # RECURSIVE multi-step prediction
        y_pred_recursive = []
        last_known_values = list(y[:train_size+val_size][-lags:])  # Last 'lags' actual values before test
        t_start = train_size + val_size
        
        for step in range(test_size):
            t_current = t_start + step
            
            # Build feature vector with predictions fed back
            features = [t_current] + last_known_values[::-1][:lags]
            X_pred = np.array(features).reshape(1, -1)
            
            # Predict next value
            y_next = model.predict(X_pred)[0]
            y_pred_recursive.append(y_next)
            
            # Update history with PREDICTION (not actual)
            last_known_values.append(y_next)
            last_known_values = last_known_values[1:]  # Keep only last 'lags' values
        
        all_y_true.extend(y_test)
        all_y_pred.extend(y_pred_recursive)
    
    return compute_metrics(all_y_true, all_y_pred)


def ridge_lasso_recursive_multistep(df, model_type='ridge'):
    """Ridge/Lasso with RECURSIVE multi-step forecasting."""
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
        
        # Tune on validation set
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
        
        # Train with best alpha
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


# =============================
# Ridge/Lasso Models
# =============================
def ridge_lasso_train_val_test(df, model_type='ridge'):
    """Ridge/Lasso with train/val/test split."""
    employees = df['employee_id'].unique()
    all_y_true, all_y_pred = [], []
    
    # Hyperparameter tuning on validation set
    alphas = [0.1, 1.0, 10.0, 100.0]
    
    for emp in employees:
        sub = df[df['employee_id'] == emp].copy().sort_values('month')
        
        if len(sub) <= TEST_MONTHS + VAL_MONTHS + 5:
            continue
        
        # Create features
        for lag in [1, 2, 3]:
            sub[f'y_lag{lag}'] = sub[TARGET_LABEL].shift(lag)
        sub = sub.dropna().copy()
        
        if len(sub) <= TEST_MONTHS + VAL_MONTHS + 3:
            continue
        
        sub['t'] = np.arange(len(sub))
        X = sub[['t', 'y_lag1', 'y_lag2', 'y_lag3']]
        y = sub[TARGET_LABEL].values
        
        # Split
        test_size = TEST_MONTHS
        val_size = VAL_MONTHS
        train_size = len(sub) - test_size - val_size
        
        X_train = X.iloc[:train_size]
        X_val = X.iloc[train_size:train_size+val_size]
        X_test = X.iloc[train_size+val_size:]
        y_train = y[:train_size]
        y_val = y[train_size:train_size+val_size]
        y_test = y[train_size+val_size:]
        
        # Scale
        scaler = StandardScaler()
        X_train_scaled = scaler.fit_transform(X_train)
        X_val_scaled = scaler.transform(X_val)
        X_test_scaled = scaler.transform(X_test)
        
        # Tune on validation set
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
        
        # Train with best alpha on train+val, test on test set
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


def ridge_lasso_cv(df, model_type='ridge', n_splits=3):
    """Ridge/Lasso with time series CV."""
    employees = df['employee_id'].unique()
    fold_metrics = []
    
    alphas = [0.1, 1.0, 10.0, 100.0]
    
    for emp in employees:
        sub = df[df['employee_id'] == emp].copy().sort_values('month')
        
        if len(sub) <= 20:
            continue
        
        # Create features
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
            X_train, X_test = X.iloc[train_idx], X.iloc[test_idx]
            y_train, y_test = y[train_idx], y[test_idx]
            
            scaler = StandardScaler()
            X_train_scaled = scaler.fit_transform(X_train)
            X_test_scaled = scaler.transform(X_test)
            
            # Simple fixed alpha for CV
            if model_type == 'ridge':
                model = Ridge(alpha=1.0)
            else:
                model = Lasso(alpha=1.0)
            
            model.fit(X_train_scaled, y_train)
            y_pred = model.predict(X_test_scaled)
            
            fold_metrics.append(compute_metrics(y_test, y_pred))
    
    if not fold_metrics:
        return {k: np.nan for k in ['rmse', 'nrmse', 'mae', 'r2']}
    
    avg_metrics = {}
    for key in fold_metrics[0].keys():
        avg_metrics[key] = np.mean([m[key] for m in fold_metrics])
    
    return avg_metrics


def ridge_lasso_nested_cv(df, model_type='ridge', outer_splits=3, inner_splits=2):
    """Ridge/Lasso with nested CV for hyperparameter tuning."""
    employees = df['employee_id'].unique()
    outer_fold_metrics = []
    
    alphas = [0.1, 1.0, 10.0, 100.0]
    
    for emp in employees:
        sub = df[df['employee_id'] == emp].copy().sort_values('month')
        
        if len(sub) <= 20:
            continue
        
        # Create features
        for lag in [1, 2, 3]:
            sub[f'y_lag{lag}'] = sub[TARGET_LABEL].shift(lag)
        sub = sub.dropna().copy()
        
        if len(sub) <= 18:
            continue
        
        sub['t'] = np.arange(len(sub))
        X = sub[['t', 'y_lag1', 'y_lag2', 'y_lag3']]
        y = sub[TARGET_LABEL].values
        
        outer_cv = TimeSeriesSplit(n_splits=outer_splits, test_size=TEST_MONTHS)
        
        for outer_train_idx, outer_test_idx in outer_cv.split(X):
            X_outer_train = X.iloc[outer_train_idx]
            y_outer_train = y[outer_train_idx]
            X_outer_test = X.iloc[outer_test_idx]
            y_outer_test = y[outer_test_idx]
            
            # Inner CV for alpha selection
            best_alpha = alphas[0]
            best_val_error = np.inf
            
            inner_test_size = max(1, len(X_outer_train) // (inner_splits + 3))
            if len(X_outer_train) >= (inner_splits + 1) * inner_test_size + 2:
                inner_cv = TimeSeriesSplit(n_splits=inner_splits, test_size=inner_test_size)
                
                for alpha in alphas:
                    inner_errors = []
                    
                    for inner_train_idx, inner_val_idx in inner_cv.split(X_outer_train):
                        X_train = X_outer_train.iloc[inner_train_idx]
                        y_train = y_outer_train[inner_train_idx]
                        X_val = X_outer_train.iloc[inner_val_idx]
                        y_val = y_outer_train[inner_val_idx]
                        
                        scaler = StandardScaler()
                        X_train_scaled = scaler.fit_transform(X_train)
                        X_val_scaled = scaler.transform(X_val)
                        
                        if model_type == 'ridge':
                            model = Ridge(alpha=alpha)
                        else:
                            model = Lasso(alpha=alpha)
                        
                        model.fit(X_train_scaled, y_train)
                        y_pred = model.predict(X_val_scaled)
                        inner_errors.append(nrmse(y_val, y_pred))
                    
                    avg_error = np.mean(inner_errors)
                    if avg_error < best_val_error:
                        best_val_error = avg_error
                        best_alpha = alpha
            
            # Train on full outer training with best alpha
            scaler = StandardScaler()
            X_outer_train_scaled = scaler.fit_transform(X_outer_train)
            X_outer_test_scaled = scaler.transform(X_outer_test)
            
            if model_type == 'ridge':
                final_model = Ridge(alpha=best_alpha)
            else:
                final_model = Lasso(alpha=best_alpha)
            
            final_model.fit(X_outer_train_scaled, y_outer_train)
            y_pred = final_model.predict(X_outer_test_scaled)
            
            outer_fold_metrics.append(compute_metrics(y_outer_test, y_pred))
    
    if not outer_fold_metrics:
        return {k: np.nan for k in ['rmse', 'nrmse', 'mae', 'r2']}
    
    avg_metrics = {}
    for key in outer_fold_metrics[0].keys():
        avg_metrics[key] = np.mean([m[key] for m in outer_fold_metrics])
    
    return avg_metrics


# =============================
# Pooled/Unpooled Models
# =============================
def pooled_train_val_test(df):
    """Pooled fixed effects with train/val/test split."""
    # Get unique months
    unique_months = np.sort(df['month'].unique())
    
    # Split by time
    test_months = unique_months[-TEST_MONTHS:]
    val_months = unique_months[-(TEST_MONTHS+VAL_MONTHS):-TEST_MONTHS]
    train_months = unique_months[:-(TEST_MONTHS+VAL_MONTHS)]
    
    train_df = df[df['month'].isin(train_months)].copy()
    val_df = df[df['month'].isin(val_months)].copy()
    test_df = df[df['month'].isin(test_months)].copy()
    
    # Create features
    train_df['t'] = train_df['month'].map({m: i for i, m in enumerate(unique_months)})
    val_df['t'] = val_df['month'].map({m: i for i, m in enumerate(unique_months)})
    test_df['t'] = test_df['month'].map({m: i for i, m in enumerate(unique_months)})
    
    # Employee dummies
    emp_dummies_train = pd.get_dummies(train_df['employee_id'], prefix='emp', drop_first=True)
    emp_dummies_val = pd.get_dummies(val_df['employee_id'], prefix='emp', drop_first=True)
    emp_dummies_test = pd.get_dummies(test_df['employee_id'], prefix='emp', drop_first=True)
    
    # Align columns
    all_emp_cols = set(emp_dummies_train.columns) | set(emp_dummies_val.columns) | set(emp_dummies_test.columns)
    for col in all_emp_cols:
        if col not in emp_dummies_train.columns:
            emp_dummies_train[col] = 0
        if col not in emp_dummies_val.columns:
            emp_dummies_val[col] = 0
        if col not in emp_dummies_test.columns:
            emp_dummies_test[col] = 0
    
    X_train = pd.concat([train_df[['t']], emp_dummies_train], axis=1)
    X_val = pd.concat([val_df[['t']], emp_dummies_val], axis=1)
    X_test = pd.concat([test_df[['t']], emp_dummies_test], axis=1)
    
    y_train = train_df[TARGET_LABEL].values
    y_val = val_df[TARGET_LABEL].values
    y_test = test_df[TARGET_LABEL].values
    
    model = LinearRegression()
    model.fit(X_train, y_train)
    y_pred = model.predict(X_test)
    
    return compute_metrics(y_test, y_pred)


def unpooled_train_val_test(df):
    """Unpooled (per-employee) models with train/val/test split."""
    employees = df['employee_id'].unique()
    all_y_true, all_y_pred = [], []
    
    for emp in employees:
        sub = df[df['employee_id'] == emp].copy().sort_values('month')
        
        if len(sub) <= TEST_MONTHS + VAL_MONTHS + 3:
            continue
        
        sub['t'] = np.arange(len(sub))
        X = sub[['t']]
        y = sub[TARGET_LABEL].values
        
        # Split
        test_size = TEST_MONTHS
        val_size = VAL_MONTHS
        train_size = len(sub) - test_size - val_size
        
        X_train, X_val, X_test = X.iloc[:train_size], X.iloc[train_size:train_size+val_size], X.iloc[train_size+val_size:]
        y_train, y_val, y_test = y[:train_size], y[train_size:train_size+val_size], y[train_size+val_size:]
        
        model = LinearRegression()
        model.fit(X_train, y_train)
        y_pred = model.predict(X_test)
        
        all_y_true.extend(y_test)
        all_y_pred.extend(y_pred)
    
    return compute_metrics(all_y_true, all_y_pred)


# =============================
# XGBoost
# =============================
def create_ml_features(series, lags=[1, 2, 3]):
    """Create features for ML models."""
    data = []
    
    for i in range(max(lags), len(series)):
        features = {}
        
        for lag in lags:
            features[f'lag_{lag}'] = series.iloc[i - lag]
        
        features['rolling_mean_3'] = series.iloc[i-3:i].mean()
        features['rolling_std_3'] = series.iloc[i-3:i].std()
        features['time_index'] = i
        features['target'] = series.iloc[i]
        
        data.append(features)
    
    return pd.DataFrame(data)


def xgboost_train_val_test(df):
    """XGBoost with train/val/test split."""
    if not XGBOOST_AVAILABLE:
        return {k: np.nan for k in ['rmse', 'nrmse', 'mae', 'r2']}
    
    employees = df['employee_id'].unique()
    all_y_true, all_y_pred = [], []
    
    for emp in employees:
        sub = df[df['employee_id'] == emp].copy().sort_values('month')
        
        if len(sub) <= TEST_MONTHS + VAL_MONTHS + 5:
            continue
        
        ml_data = create_ml_features(sub[TARGET_LABEL].reset_index(drop=True))
        
        if len(ml_data) <= TEST_MONTHS + VAL_MONTHS + 3:
            continue
        
        # Split
        test_size = TEST_MONTHS
        val_size = VAL_MONTHS
        train_size = len(ml_data) - test_size - val_size
        
        train_data = ml_data.iloc[:train_size]
        val_data = ml_data.iloc[train_size:train_size+val_size]
        test_data = ml_data.iloc[train_size+val_size:]
        
        X_train = train_data.drop('target', axis=1)
        y_train = train_data['target']
        X_val = val_data.drop('target', axis=1)
        y_val = val_data['target']
        X_test = test_data.drop('target', axis=1)
        y_test = test_data['target']
        
        scaler = StandardScaler()
        X_train_scaled = scaler.fit_transform(X_train)
        X_val_scaled = scaler.transform(X_val)
        X_test_scaled = scaler.transform(X_test)
        
        model = XGBRegressor(n_estimators=100, learning_rate=0.1, max_depth=3, random_state=42)
        model.fit(X_train_scaled, y_train)
        y_pred = model.predict(X_test_scaled)
        
        all_y_true.extend(y_test)
        all_y_pred.extend(y_pred)
    
    return compute_metrics(all_y_true, all_y_pred)


def xgboost_recursive_multistep(df):
    """XGBoost with RECURSIVE multi-step forecasting."""
    if not XGBOOST_AVAILABLE:
        return {k: np.nan for k in ['rmse', 'nrmse', 'mae', 'r2']}
    
    employees = df['employee_id'].unique()
    all_y_true, all_y_pred = [], []
    
    for emp in employees:
        sub = df[df['employee_id'] == emp].copy().sort_values('month')
        
        if len(sub) <= TEST_MONTHS + VAL_MONTHS + 5:
            continue
        
        ml_data = create_ml_features(sub[TARGET_LABEL].reset_index(drop=True))
        
        if len(ml_data) <= TEST_MONTHS + VAL_MONTHS + 3:
            continue
        
        test_size = TEST_MONTHS
        val_size = VAL_MONTHS
        train_size = len(ml_data) - test_size - val_size
        
        train_data = ml_data.iloc[:train_size]
        val_data = ml_data.iloc[train_size:train_size+val_size]
        
        X_train = train_data.drop('target', axis=1)
        y_train = train_data['target']
        X_val = val_data.drop('target', axis=1)
        y_val = val_data['target']
        
        # Train model
        X_train_val = pd.concat([X_train, X_val])
        y_train_val = pd.concat([y_train, y_val])
        
        scaler = StandardScaler()
        X_train_val_scaled = scaler.fit_transform(X_train_val)
        
        model = XGBRegressor(n_estimators=100, learning_rate=0.1, max_depth=3, random_state=42)
        model.fit(X_train_val_scaled, y_train_val)
        
        # Get actual test values
        y_test = ml_data.iloc[train_size+val_size:]['target'].values
        
        # RECURSIVE forecasting
        y_pred_recursive = []
        history = list(sub[TARGET_LABEL].values[:train_size+val_size])
        
        for step in range(test_size):
            # Build features from history
            lag_1 = history[-1]
            lag_2 = history[-2]
            lag_3 = history[-3]
            rolling_mean = np.mean(history[-3:])
            rolling_std = np.std(history[-3:])
            time_index = len(history)
            
            features = pd.DataFrame([{
                'lag_1': lag_1,
                'lag_2': lag_2,
                'lag_3': lag_3,
                'rolling_mean_3': rolling_mean,
                'rolling_std_3': rolling_std,
                'time_index': time_index
            }])
            
            features_scaled = scaler.transform(features)
            y_next = model.predict(features_scaled)[0]
            y_pred_recursive.append(y_next)
            
            # Update history with PREDICTION
            history.append(y_next)
        
        all_y_true.extend(y_test)
        all_y_pred.extend(y_pred_recursive)
    
    return compute_metrics(all_y_true, all_y_pred)


def xgboost_cv(df, n_splits=3):
    """XGBoost with time series CV."""
    if not XGBOOST_AVAILABLE:
        return {k: np.nan for k in ['rmse', 'nrmse', 'mae', 'r2']}
    
    employees = df['employee_id'].unique()
    fold_metrics = []
    
    for emp in employees:
        sub = df[df['employee_id'] == emp].copy().sort_values('month')
        
        if len(sub) <= 20:
            continue
        
        ml_data = create_ml_features(sub[TARGET_LABEL].reset_index(drop=True))
        
        if len(ml_data) <= 15:
            continue
        
        X = ml_data.drop('target', axis=1)
        y = ml_data['target']
        
        tscv = TimeSeriesSplit(n_splits=n_splits, test_size=TEST_MONTHS)
        
        for train_idx, test_idx in tscv.split(X):
            X_train, X_test = X.iloc[train_idx], X.iloc[test_idx]
            y_train, y_test = y.iloc[train_idx], y.iloc[test_idx]
            
            scaler = StandardScaler()
            X_train_scaled = scaler.fit_transform(X_train)
            X_test_scaled = scaler.transform(X_test)
            
            model = XGBRegressor(n_estimators=100, learning_rate=0.1, max_depth=3, random_state=42)
            model.fit(X_train_scaled, y_train)
            y_pred = model.predict(X_test_scaled)
            
            fold_metrics.append(compute_metrics(y_test, y_pred))
    
    if not fold_metrics:
        return {k: np.nan for k in ['rmse', 'nrmse', 'mae', 'r2']}
    
    avg_metrics = {}
    for key in fold_metrics[0].keys():
        avg_metrics[key] = np.mean([m[key] for m in fold_metrics])
    
    return avg_metrics


def xgboost_nested_cv(df, outer_splits=3, inner_splits=2):
    """XGBoost with nested CV for hyperparameter tuning."""
    if not XGBOOST_AVAILABLE:
        return {k: np.nan for k in ['rmse', 'nrmse', 'mae', 'r2']}
    
    employees = df['employee_id'].unique()
    outer_fold_metrics = []
    
    param_grid = [
        {'n_estimators': 50, 'max_depth': 3, 'learning_rate': 0.1},
        {'n_estimators': 100, 'max_depth': 3, 'learning_rate': 0.1},
        {'n_estimators': 100, 'max_depth': 5, 'learning_rate': 0.05}
    ]
    
    for emp in employees:
        sub = df[df['employee_id'] == emp].copy().sort_values('month')
        
        if len(sub) <= 20:
            continue
        
        ml_data = create_ml_features(sub[TARGET_LABEL].reset_index(drop=True))
        
        if len(ml_data) <= 18:
            continue
        
        X = ml_data.drop('target', axis=1)
        y = ml_data['target']
        
        outer_cv = TimeSeriesSplit(n_splits=outer_splits, test_size=TEST_MONTHS)
        
        for outer_train_idx, outer_test_idx in outer_cv.split(X):
            X_outer_train = X.iloc[outer_train_idx]
            y_outer_train = y.iloc[outer_train_idx]
            X_outer_test = X.iloc[outer_test_idx]
            y_outer_test = y.iloc[outer_test_idx]
            
            # Inner CV for hyperparameter selection
            best_params = param_grid[0]
            best_val_error = np.inf
            
            inner_test_size = max(1, len(X_outer_train) // (inner_splits + 3))
            if len(X_outer_train) >= (inner_splits + 1) * inner_test_size + 2:
                inner_cv = TimeSeriesSplit(n_splits=inner_splits, test_size=inner_test_size)
                
                for params in param_grid:
                    inner_errors = []
                    
                    for inner_train_idx, inner_val_idx in inner_cv.split(X_outer_train):
                        X_train = X_outer_train.iloc[inner_train_idx]
                        y_train = y_outer_train.iloc[inner_train_idx]
                        X_val = X_outer_train.iloc[inner_val_idx]
                        y_val = y_outer_train.iloc[inner_val_idx]
                        
                        scaler = StandardScaler()
                        X_train_scaled = scaler.fit_transform(X_train)
                        X_val_scaled = scaler.transform(X_val)
                        
                        model = XGBRegressor(**params, random_state=42)
                        model.fit(X_train_scaled, y_train)
                        y_pred = model.predict(X_val_scaled)
                        inner_errors.append(nrmse(y_val, y_pred))
                    
                    avg_error = np.mean(inner_errors)
                    if avg_error < best_val_error:
                        best_val_error = avg_error
                        best_params = params
            
            # Train on full outer train with best params
            scaler = StandardScaler()
            X_outer_train_scaled = scaler.fit_transform(X_outer_train)
            X_outer_test_scaled = scaler.transform(X_outer_test)
            
            final_model = XGBRegressor(**best_params, random_state=42)
            final_model.fit(X_outer_train_scaled, y_outer_train)
            y_pred = final_model.predict(X_outer_test_scaled)
            
            outer_fold_metrics.append(compute_metrics(y_outer_test, y_pred))
    
    if not outer_fold_metrics:
        return {k: np.nan for k in ['rmse', 'nrmse', 'mae', 'r2']}
    
    avg_metrics = {}
    for key in outer_fold_metrics[0].keys():
        avg_metrics[key] = np.mean([m[key] for m in outer_fold_metrics])
    
    return avg_metrics


# =============================
# Main Comparison
# =============================
def main():
    print("=" * 80)
    print("FORECASTING MODEL COMPARISON")
    print("=" * 80)
    print("\nModels: AR(1-4), Ridge, Lasso, Pooled FE, Unpooled, XGBoost")
    print("Evaluation: Train/Val/Test, Time Series CV, Nested CV, Recursive Multi-Step")
    
    df = load_data()
    print(f"\n✓ Loaded {len(df)} records, {df['employee_id'].nunique()} employees")
    
    results = []
    
    # Method 1: Train/Val/Test Split (One-Step-Ahead)
    print("\n" + "=" * 80)
    print("METHOD 1: TRAIN/VALIDATION/TEST SPLIT (One-Step-Ahead)")
    print("=" * 80)
    
    models_to_test = [
        ('AR(1)', lambda: ar_train_val_test(df, lags=1)),
        ('AR(2)', lambda: ar_train_val_test(df, lags=2)),
        ('AR(3)', lambda: ar_train_val_test(df, lags=3)),
        ('AR(4)', lambda: ar_train_val_test(df, lags=4)),
        ('Ridge', lambda: ridge_lasso_train_val_test(df, 'ridge')),
        ('Lasso', lambda: ridge_lasso_train_val_test(df, 'lasso')),
        ('Pooled FE', lambda: pooled_train_val_test(df)),
        ('Unpooled', lambda: unpooled_train_val_test(df)),
    ]
    
    if XGBOOST_AVAILABLE:
        models_to_test.append(('XGBoost', lambda: xgboost_train_val_test(df)))
    
    for model_name, model_func in models_to_test:
        print(f"  {model_name}...", end=" ")
        metrics = model_func()
        results.append({
            'Model': model_name,
            'Method': 'Train/Val/Test',
            **metrics
        })
        print(f"NRMSE: {metrics['nrmse']:.4f}")
    
    # Method 1b: Recursive Multi-Step (Realistic Forecasting)
    print("\n" + "=" * 80)
    print("METHOD 1b: RECURSIVE MULTI-STEP FORECASTING (Predictions Feed Back)")
    print("=" * 80)
    
    recursive_models = [
        ('AR(1)', lambda: ar_recursive_multistep(df, lags=1)),
        ('AR(2)', lambda: ar_recursive_multistep(df, lags=2)),
        ('Ridge', lambda: ridge_lasso_recursive_multistep(df, 'ridge')),
        ('Lasso', lambda: ridge_lasso_recursive_multistep(df, 'lasso')),
    ]
    
    if XGBOOST_AVAILABLE:
        recursive_models.append(('XGBoost', lambda: xgboost_recursive_multistep(df)))
    
    for model_name, model_func in recursive_models:
        print(f"  {model_name}...", end=" ")
        metrics = model_func()
        results.append({
            'Model': model_name,
            'Method': 'Recursive Multi-Step',
            **metrics
        })
        print(f"NRMSE: {metrics['nrmse']:.4f}")
    
    # Method 2: Time Series CV
    print("\n" + "=" * 80)
    print("METHOD 2: TIME SERIES CROSS-VALIDATION")
    print("=" * 80)
    
    cv_models = [
        ('AR(1)', lambda: ar_time_series_cv(df, lags=1)),
        ('AR(2)', lambda: ar_time_series_cv(df, lags=2)),
        ('AR(3)', lambda: ar_time_series_cv(df, lags=3)),
        ('AR(4)', lambda: ar_time_series_cv(df, lags=4)),
        ('Ridge', lambda: ridge_lasso_cv(df, 'ridge')),
        ('Lasso', lambda: ridge_lasso_cv(df, 'lasso')),
    ]
    
    if XGBOOST_AVAILABLE:
        cv_models.append(('XGBoost', lambda: xgboost_cv(df)))
    
    for model_name, model_func in cv_models:
        print(f"  {model_name}...", end=" ")
        metrics = model_func()
        results.append({
            'Model': model_name,
            'Method': 'Time Series CV',
            **metrics
        })
        print(f"NRMSE: {metrics['nrmse']:.4f}")
    
    # Method 3: Nested CV
    print("\n" + "=" * 80)
    print("METHOD 3: NESTED CROSS-VALIDATION (Hyperparameter Tuning)")
    print("=" * 80)
    
    nested_models = [
        ('Ridge', lambda: ridge_lasso_nested_cv(df, 'ridge')),
        ('Lasso', lambda: ridge_lasso_nested_cv(df, 'lasso')),
    ]
    
    if XGBOOST_AVAILABLE:
        nested_models.append(('XGBoost', lambda: xgboost_nested_cv(df)))
    
    for model_name, model_func in nested_models:
        print(f"  {model_name}...", end=" ")
        metrics = model_func()
        results.append({
            'Model': model_name,
            'Method': 'Nested CV',
            **metrics
        })
        print(f"NRMSE: {metrics['nrmse']:.4f}")
    
    # Save results
    print("\n" + "=" * 80)
    print("RESULTS SUMMARY")
    print("=" * 80)
    
    results_df = pd.DataFrame(results)
    output_path = project_root / "outputs" / "forecasting_simplified_comparison.csv"
    results_df.to_csv(output_path, index=False)
    print(f"\n✓ Results saved to: {output_path}")
    
    # Display table
    print("\n" + "=" * 100)
    print(f"{'Model':<15} {'Method':<30} {'NRMSE':<10} {'RMSE':<10} {'MAE':<10} {'R²':<10}")
    print("=" * 100)
    
    for _, row in results_df.iterrows():
        print(f"{row['Model']:<15} {row['Method']:<30} "
              f"{row['nrmse']:<10.4f} {row['rmse']:<10.2f} "
              f"{row['mae']:<10.2f} {row['r2']:<10.4f}")
    
    # Best models
    print("\n" + "=" * 80)
    print("BEST MODELS BY EVALUATION METHOD")
    print("=" * 80)
    
    for method in results_df['Method'].unique():
        method_results = results_df[results_df['Method'] == method]
        best = method_results.loc[method_results['nrmse'].idxmin()]
        print(f"\n{method}:")
        print(f"  🏆 {best['Model']} - NRMSE: {best['nrmse']:.4f}")
    
    # Compare one-step vs recursive for AR models
    print("\n" + "=" * 80)
    print("ERROR ACCUMULATION ANALYSIS: One-Step vs Recursive")
    print("=" * 80)
    
    for model_name in ['AR(1)', 'AR(2)', 'Ridge', 'Lasso', 'XGBoost']:
        one_step = results_df[(results_df['Model'] == model_name) & 
                              (results_df['Method'] == 'Train/Val/Test')]
        recursive = results_df[(results_df['Model'] == model_name) & 
                               (results_df['Method'] == 'Recursive Multi-Step')]
        
        if len(one_step) > 0 and len(recursive) > 0:
            os_nrmse = one_step['nrmse'].values[0]
            rec_nrmse = recursive['nrmse'].values[0]
            increase = ((rec_nrmse - os_nrmse) / os_nrmse) * 100
            
            print(f"\n{model_name}:")
            print(f"  One-Step-Ahead: NRMSE = {os_nrmse:.4f}")
            print(f"  Recursive:      NRMSE = {rec_nrmse:.4f}")
            print(f"  ⚠️  Error Increase: +{increase:.1f}%")
    
    print("\n" + "=" * 80)
    print("ANALYSIS COMPLETE")
    print("=" * 80)


if __name__ == "__main__":
    main()
