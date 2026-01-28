"""
Comprehensive Forecasting Model Comparison

Compares multiple forecasting methods with three evaluation approaches:
1. One-step-ahead prediction (optimistic baseline)
2. Multi-step recursive prediction (realistic with error accumulation)
3. Cross-validation for ML models (train/val/test, CV, nested CV)

Models tested:
- AR(1), AR(2), AR(3), AR(4)
- Ridge Regression
- Lasso Regression
- XGBoost (new)
- Random Forest (new)
- SVR (new)
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
from sklearn.ensemble import RandomForestRegressor
from sklearn.svm import SVR
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import mean_squared_error, mean_absolute_error, r2_score
from sklearn.model_selection import TimeSeriesSplit

try:
    from xgboost import XGBRegressor
    XGBOOST_AVAILABLE = True
except ImportError:
    XGBOOST_AVAILABLE = False
    print("⚠️  XGBoost not installed. Install with: pip install xgboost")


# =============================
# Configuration
# =============================
TARGET_LABEL = "salaire_brut"
CSV_PATH = "outputs/payroll_long.csv"
TEST_MONTHS = 6
VAL_MONTHS = 3
LAGS = [1, 2, 3, 4]  # For feature engineering


# =============================
# Metrics
# =============================
def nrmse(y_true, y_pred):
    """Normalized RMSE for scale-independent comparison."""
    y_true = np.array(y_true)
    y_pred = np.array(y_pred)
    rmse = np.sqrt(mean_squared_error(y_true, y_pred))
    mean_true = np.mean(y_true)
    return rmse / mean_true if mean_true != 0 else np.inf


def mape(y_true, y_pred):
    """Mean Absolute Percentage Error."""
    y_true = np.array(y_true)
    y_pred = np.array(y_pred)
    mask = y_true != 0
    if mask.sum() == 0:
        return np.nan
    return np.mean(np.abs((y_true[mask] - y_pred[mask]) / y_true[mask])) * 100.0


def compute_metrics(y_true, y_pred):
    """Compute all metrics."""
    return {
        'rmse': np.sqrt(mean_squared_error(y_true, y_pred)),
        'nrmse': nrmse(y_true, y_pred),
        'mae': mean_absolute_error(y_true, y_pred),
        'mape': mape(y_true, y_pred),
        'r2': r2_score(y_true, y_pred)
    }


# =============================
# Data Loading
# =============================
def load_data():
    """Load and prepare payroll data."""
    df = pd.read_csv(CSV_PATH)
    df['month'] = pd.to_datetime(df['month'])
    df = df.sort_values(['employee_id', 'month']).reset_index(drop=True)
    return df


# =============================
# Feature Engineering for ML Models
# =============================
def create_ml_features(series, lags=[1, 2, 3]):
    """
    Create features from time series for ML models.
    
    Features:
    - Lagged values (t-1, t-2, t-3)
    - Rolling statistics
    - Time features
    """
    data = []
    
    for i in range(max(lags), len(series)):
        features = {}
        
        # Lagged features
        for lag in lags:
            features[f'lag_{lag}'] = series.iloc[i - lag]
        
        # Rolling statistics
        features['rolling_mean_3'] = series.iloc[i-3:i].mean()
        features['rolling_std_3'] = series.iloc[i-3:i].std()
        
        # Time index (trend)
        features['time_index'] = i
        
        # Target
        features['target'] = series.iloc[i]
        
        data.append(features)
    
    return pd.DataFrame(data)


# =============================
# AR Models (Statistical)
# =============================
def ar_one_step(df, lags=2):
    """AR model with one-step-ahead prediction."""
    employees = df['employee_id'].unique()
    all_y_true, all_y_pred = [], []
    
    for emp in employees:
        sub = df[df['employee_id'] == emp].copy().sort_values('month')
        
        if len(sub) <= TEST_MONTHS + lags + 3:
            continue
        
        # Create lagged features
        for lag in range(1, lags + 1):
            sub[f'y_lag{lag}'] = sub[TARGET_LABEL].shift(lag)
        sub = sub.dropna().copy()
        
        if len(sub) <= TEST_MONTHS + 3:
            continue
        
        sub['t'] = np.arange(len(sub))
        lag_cols = [f'y_lag{lag}' for lag in range(1, lags + 1)]
        X = sub[['t'] + lag_cols]
        y = sub[TARGET_LABEL].values
        
        split = len(sub) - TEST_MONTHS
        X_train, X_test = X.iloc[:split], X.iloc[split:]
        y_train, y_test = y[:split], y[split:]
        
        model = LinearRegression()
        model.fit(X_train, y_train)
        y_pred = model.predict(X_test)
        
        all_y_true.extend(y_test)
        all_y_pred.extend(y_pred)
    
    return compute_metrics(all_y_true, all_y_pred)


def ar_multi_step(df, lags=2, steps=6):
    """AR model with multi-step recursive prediction."""
    employees = df['employee_id'].unique()
    all_y_true, all_y_pred = [], []
    
    for emp in employees:
        sub = df[df['employee_id'] == emp].copy().sort_values('month')
        
        if len(sub) <= TEST_MONTHS + lags + 3:
            continue
        
        # Prepare training data
        for lag in range(1, lags + 1):
            sub[f'y_lag{lag}'] = sub[TARGET_LABEL].shift(lag)
        sub = sub.dropna().copy()
        
        if len(sub) <= TEST_MONTHS + 3:
            continue
        
        sub['t'] = np.arange(len(sub))
        lag_cols = [f'y_lag{lag}' for lag in range(1, lags + 1)]
        
        split = len(sub) - TEST_MONTHS
        X_train = sub.iloc[:split][['t'] + lag_cols]
        y_train = sub.iloc[:split][TARGET_LABEL].values
        y_test = sub.iloc[split:split+steps][TARGET_LABEL].values
        
        # Train model
        model = LinearRegression()
        model.fit(X_train, y_train)
        
        # Multi-step prediction
        predictions = []
        history = list(sub.iloc[split-lags:split][TARGET_LABEL].values)
        
        for step in range(min(steps, len(y_test))):
            # Create feature vector
            lag_features = [history[-(lag)] for lag in range(1, lags + 1)]
            t_value = split + step
            X_pred = np.array([[t_value] + lag_features])
            
            # Predict
            y_pred = model.predict(X_pred)[0]
            predictions.append(y_pred)
            history.append(y_pred)  # Use prediction for next step
        
        all_y_true.extend(y_test[:len(predictions)])
        all_y_pred.extend(predictions)
    
    return compute_metrics(all_y_true, all_y_pred)


# =============================
# Ridge/Lasso Models
# =============================
def ridge_lasso_one_step(df, alpha=1.0, model_type='ridge'):
    """Ridge or Lasso with one-step prediction."""
    employees = df['employee_id'].unique()
    all_y_true, all_y_pred = [], []
    
    for emp in employees:
        sub = df[df['employee_id'] == emp].copy().sort_values('month')
        
        if len(sub) <= TEST_MONTHS + 5:
            continue
        
        # Create features
        for lag in [1, 2, 3]:
            sub[f'y_lag{lag}'] = sub[TARGET_LABEL].shift(lag)
        sub = sub.dropna().copy()
        
        if len(sub) <= TEST_MONTHS + 3:
            continue
        
        sub['t'] = np.arange(len(sub))
        X = sub[['t', 'y_lag1', 'y_lag2', 'y_lag3']]
        y = sub[TARGET_LABEL].values
        
        split = len(sub) - TEST_MONTHS
        X_train, X_test = X.iloc[:split], X.iloc[split:]
        y_train, y_test = y[:split], y[split:]
        
        # Scale features
        scaler = StandardScaler()
        X_train_scaled = scaler.fit_transform(X_train)
        X_test_scaled = scaler.transform(X_test)
        
        # Train model
        if model_type == 'ridge':
            model = Ridge(alpha=alpha)
        else:
            model = Lasso(alpha=alpha)
        
        model.fit(X_train_scaled, y_train)
        y_pred = model.predict(X_test_scaled)
        
        all_y_true.extend(y_test)
        all_y_pred.extend(y_pred)
    
    return compute_metrics(all_y_true, all_y_pred)


# =============================
# ML Models (XGBoost, Random Forest, SVR)
# =============================
def ml_model_train_val_test(df, model_type='xgboost', lags=3):
    """
    ML model with train/val/test split.
    
    Split:
    - Train: All data except last (VAL_MONTHS + TEST_MONTHS)
    - Val: Next VAL_MONTHS
    - Test: Last TEST_MONTHS
    """
    employees = df['employee_id'].unique()
    all_y_true, all_y_pred = [], []
    
    for emp in employees:
        sub = df[df['employee_id'] == emp].copy().sort_values('month')
        
        if len(sub) <= TEST_MONTHS + VAL_MONTHS + lags + 3:
            continue
        
        # Create ML features
        ml_data = create_ml_features(sub[TARGET_LABEL].reset_index(drop=True), lags=[1, 2, 3])
        
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
        
        # Scale features
        scaler = StandardScaler()
        X_train_scaled = scaler.fit_transform(X_train)
        X_val_scaled = scaler.transform(X_val)
        X_test_scaled = scaler.transform(X_test)
        
        # Select and train model
        if model_type == 'xgboost' and XGBOOST_AVAILABLE:
            model = XGBRegressor(n_estimators=100, learning_rate=0.1, max_depth=3, random_state=42)
        elif model_type == 'random_forest':
            model = RandomForestRegressor(n_estimators=100, max_depth=5, random_state=42)
        elif model_type == 'svr':
            model = SVR(kernel='rbf', C=10, epsilon=0.1)
        else:
            continue
        
        model.fit(X_train_scaled, y_train)
        y_pred = model.predict(X_test_scaled)
        
        all_y_true.extend(y_test)
        all_y_pred.extend(y_pred)
    
    return compute_metrics(all_y_true, all_y_pred)


def ml_model_time_series_cv(df, model_type='xgboost', n_splits=3):
    """ML model with time series cross-validation."""
    employees = df['employee_id'].unique()
    fold_metrics = []
    
    for emp in employees:
        sub = df[df['employee_id'] == emp].copy().sort_values('month')
        
        if len(sub) <= 20:
            continue
        
        # Create ML features
        ml_data = create_ml_features(sub[TARGET_LABEL].reset_index(drop=True), lags=[1, 2, 3])
        
        if len(ml_data) <= 15:
            continue
        
        X = ml_data.drop('target', axis=1)
        y = ml_data['target']
        
        # Time series CV
        tscv = TimeSeriesSplit(n_splits=n_splits, test_size=TEST_MONTHS)
        
        for train_idx, test_idx in tscv.split(X):
            X_train, X_test = X.iloc[train_idx], X.iloc[test_idx]
            y_train, y_test = y.iloc[train_idx], y.iloc[test_idx]
            
            # Scale
            scaler = StandardScaler()
            X_train_scaled = scaler.fit_transform(X_train)
            X_test_scaled = scaler.transform(X_test)
            
            # Train
            if model_type == 'xgboost' and XGBOOST_AVAILABLE:
                model = XGBRegressor(n_estimators=100, learning_rate=0.1, max_depth=3, random_state=42)
            elif model_type == 'random_forest':
                model = RandomForestRegressor(n_estimators=100, max_depth=5, random_state=42)
            elif model_type == 'svr':
                model = SVR(kernel='rbf', C=10, epsilon=0.1)
            else:
                continue
            
            model.fit(X_train_scaled, y_train)
            y_pred = model.predict(X_test_scaled)
            
            fold_metrics.append(compute_metrics(y_test, y_pred))
    
    # Average across all folds
    if not fold_metrics:
        return {k: np.nan for k in ['rmse', 'nrmse', 'mae', 'mape', 'r2']}
    
    avg_metrics = {}
    for key in fold_metrics[0].keys():
        avg_metrics[key] = np.mean([m[key] for m in fold_metrics if not np.isnan(m[key])])
    
    return avg_metrics


def ml_model_nested_cv(df, model_type='xgboost', outer_splits=3, inner_splits=2):
    """
    ML model with nested cross-validation for hyperparameter tuning.
    
    Outer loop: Performance estimation
    Inner loop: Hyperparameter selection
    """
    employees = df['employee_id'].unique()
    outer_fold_metrics = []
    
    # Hyperparameter grid
    if model_type == 'xgboost':
        param_grid = [
            {'n_estimators': 50, 'max_depth': 3, 'learning_rate': 0.1},
            {'n_estimators': 100, 'max_depth': 3, 'learning_rate': 0.1},
            {'n_estimators': 100, 'max_depth': 5, 'learning_rate': 0.05}
        ]
    elif model_type == 'random_forest':
        param_grid = [
            {'n_estimators': 50, 'max_depth': 3},
            {'n_estimators': 100, 'max_depth': 5},
            {'n_estimators': 100, 'max_depth': 7}
        ]
    elif model_type == 'svr':
        param_grid = [
            {'C': 1, 'epsilon': 0.1},
            {'C': 10, 'epsilon': 0.1},
            {'C': 10, 'epsilon': 0.01}
        ]
    else:
        return {k: np.nan for k in ['rmse', 'nrmse', 'mae', 'mape', 'r2']}
    
    for emp in employees:
        sub = df[df['employee_id'] == emp].copy().sort_values('month')
        
        if len(sub) <= 20:
            continue
        
        ml_data = create_ml_features(sub[TARGET_LABEL].reset_index(drop=True), lags=[1, 2, 3])
        
        if len(ml_data) <= 18:  # Need at least 18 for outer CV with test_size=6
            continue
        
        X = ml_data.drop('target', axis=1)
        y = ml_data['target']
        
        # Outer CV
        outer_cv = TimeSeriesSplit(n_splits=outer_splits, test_size=TEST_MONTHS)
        
        for outer_train_idx, outer_test_idx in outer_cv.split(X):
            X_outer_train = X.iloc[outer_train_idx]
            y_outer_train = y.iloc[outer_train_idx]
            X_outer_test = X.iloc[outer_test_idx]
            y_outer_test = y.iloc[outer_test_idx]
            
            # Inner CV for hyperparameter selection
            best_params = None
            best_val_error = np.inf
            
            # Adjust test_size based on available data
            inner_test_size = max(1, len(X_outer_train) // (inner_splits + 3))
            if len(X_outer_train) < (inner_splits + 1) * inner_test_size + 2:
                # Not enough data for inner CV, use first param set
                best_params = param_grid[0]
            else:
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
                        
                        # Train with current params
                        if model_type == 'xgboost' and XGBOOST_AVAILABLE:
                            model = XGBRegressor(**params, random_state=42)
                        elif model_type == 'random_forest':
                            model = RandomForestRegressor(**params, random_state=42)
                        elif model_type == 'svr':
                            model = SVR(kernel='rbf', **params)
                        
                        model.fit(X_train_scaled, y_train)
                        y_pred = model.predict(X_val_scaled)
                        inner_errors.append(nrmse(y_val, y_pred))
                    
                    avg_error = np.mean(inner_errors)
                    if avg_error < best_val_error:
                        best_val_error = avg_error
                        best_params = params
            
            # Train on full outer training set with best params
            scaler = StandardScaler()
            X_outer_train_scaled = scaler.fit_transform(X_outer_train)
            X_outer_test_scaled = scaler.transform(X_outer_test)
            
            if model_type == 'xgboost' and XGBOOST_AVAILABLE:
                final_model = XGBRegressor(**best_params, random_state=42)
            elif model_type == 'random_forest':
                final_model = RandomForestRegressor(**best_params, random_state=42)
            elif model_type == 'svr':
                final_model = SVR(kernel='rbf', **best_params)
            
            final_model.fit(X_outer_train_scaled, y_outer_train)
            y_pred = final_model.predict(X_outer_test_scaled)
            
            outer_fold_metrics.append(compute_metrics(y_outer_test, y_pred))
    
    # Average across all outer folds
    if not outer_fold_metrics:
        return {k: np.nan for k in ['rmse', 'nrmse', 'mae', 'mape', 'r2']}
    
    avg_metrics = {}
    for key in outer_fold_metrics[0].keys():
        avg_metrics[key] = np.mean([m[key] for m in outer_fold_metrics if not np.isnan(m[key])])
    
    return avg_metrics


# =============================
# Main Comparison
# =============================
def main():
    print("=" * 80)
    print("COMPREHENSIVE FORECASTING MODEL COMPARISON")
    print("=" * 80)
    
    # Load data
    print("\nLoading data...")
    df = load_data()
    print(f"✓ Loaded {len(df)} records, {df['employee_id'].nunique()} employees")
    
    results = []
    
    # =============================
    # PART 1: One-Step-Ahead (Optimistic)
    # =============================
    print("\n" + "=" * 80)
    print("PART 1: ONE-STEP-AHEAD PREDICTION (Optimistic Baseline)")
    print("=" * 80)
    
    print("\n1. AR Models...")
    for lags in [1, 2, 3, 4]:
        print(f"   - AR({lags})...", end=" ")
        metrics = ar_one_step(df, lags=lags)
        results.append({
            'Model': f'AR({lags})',
            'Method': 'One-Step-Ahead',
            **metrics
        })
        print(f"NRMSE: {metrics['nrmse']:.4f}")
    
    print("\n2. Ridge Regression...")
    metrics = ridge_lasso_one_step(df, alpha=1.0, model_type='ridge')
    results.append({
        'Model': 'Ridge',
        'Method': 'One-Step-Ahead',
        **metrics
    })
    print(f"   NRMSE: {metrics['nrmse']:.4f}")
    
    print("\n3. Lasso Regression...")
    metrics = ridge_lasso_one_step(df, alpha=1.0, model_type='lasso')
    results.append({
        'Model': 'Lasso',
        'Method': 'One-Step-Ahead',
        **metrics
    })
    print(f"   NRMSE: {metrics['nrmse']:.4f}")
    
    print("\n4. ML Models (Train/Val/Test Split)...")
    
    if XGBOOST_AVAILABLE:
        print("   - XGBoost...", end=" ")
        metrics = ml_model_train_val_test(df, model_type='xgboost')
        results.append({
            'Model': 'XGBoost',
            'Method': 'One-Step (Train/Val/Test)',
            **metrics
        })
        print(f"NRMSE: {metrics['nrmse']:.4f}")
    
    print("   - Random Forest...", end=" ")
    metrics = ml_model_train_val_test(df, model_type='random_forest')
    results.append({
        'Model': 'Random Forest',
        'Method': 'One-Step (Train/Val/Test)',
        **metrics
    })
    print(f"NRMSE: {metrics['nrmse']:.4f}")
    
    print("   - SVR...", end=" ")
    metrics = ml_model_train_val_test(df, model_type='svr')
    results.append({
        'Model': 'SVR',
        'Method': 'One-Step (Train/Val/Test)',
        **metrics
    })
    print(f"NRMSE: {metrics['nrmse']:.4f}")
    
    # =============================
    # PART 2: Multi-Step Recursive (Realistic)
    # =============================
    print("\n" + "=" * 80)
    print("PART 2: MULTI-STEP RECURSIVE PREDICTION (Realistic with Error Accumulation)")
    print("=" * 80)
    
    print("\n1. AR Models...")
    for lags in [1, 2, 3, 4]:
        print(f"   - AR({lags})...", end=" ")
        metrics = ar_multi_step(df, lags=lags, steps=6)
        results.append({
            'Model': f'AR({lags})',
            'Method': 'Multi-Step (6 months)',
            **metrics
        })
        print(f"NRMSE: {metrics['nrmse']:.4f}")
    
    # =============================
    # PART 3: Cross-Validation (ML Models Only)
    # =============================
    print("\n" + "=" * 80)
    print("PART 3: CROSS-VALIDATION (ML Models Only)")
    print("=" * 80)
    
    print("\n1. Time Series Cross-Validation...")
    
    if XGBOOST_AVAILABLE:
        print("   - XGBoost...", end=" ")
        metrics = ml_model_time_series_cv(df, model_type='xgboost', n_splits=3)
        results.append({
            'Model': 'XGBoost',
            'Method': 'Time Series CV',
            **metrics
        })
        print(f"NRMSE: {metrics['nrmse']:.4f}")
    
    print("   - Random Forest...", end=" ")
    metrics = ml_model_time_series_cv(df, model_type='random_forest', n_splits=3)
    results.append({
        'Model': 'Random Forest',
        'Method': 'Time Series CV',
        **metrics
    })
    print(f"NRMSE: {metrics['nrmse']:.4f}")
    
    print("   - SVR...", end=" ")
    metrics = ml_model_time_series_cv(df, model_type='svr', n_splits=3)
    results.append({
        'Model': 'SVR',
        'Method': 'Time Series CV',
        **metrics
    })
    print(f"NRMSE: {metrics['nrmse']:.4f}")
    
    print("\n2. Nested Cross-Validation (with Hyperparameter Tuning)...")
    
    if XGBOOST_AVAILABLE:
        print("   - XGBoost...", end=" ")
        metrics = ml_model_nested_cv(df, model_type='xgboost', outer_splits=3, inner_splits=2)
        results.append({
            'Model': 'XGBoost',
            'Method': 'Nested CV',
            **metrics
        })
        print(f"NRMSE: {metrics['nrmse']:.4f}")
    
    print("   - Random Forest...", end=" ")
    metrics = ml_model_nested_cv(df, model_type='random_forest', outer_splits=3, inner_splits=2)
    results.append({
        'Model': 'Random Forest',
        'Method': 'Nested CV',
        **metrics
    })
    print(f"NRMSE: {metrics['nrmse']:.4f}")
    
    print("   - SVR...", end=" ")
    metrics = ml_model_nested_cv(df, model_type='svr', outer_splits=3, inner_splits=2)
    results.append({
        'Model': 'SVR',
        'Method': 'Nested CV',
        **metrics
    })
    print(f"NRMSE: {metrics['nrmse']:.4f}")
    
    # =============================
    # Results Summary
    # =============================
    print("\n" + "=" * 80)
    print("RESULTS SUMMARY")
    print("=" * 80)
    
    results_df = pd.DataFrame(results)
    
    # Save to CSV
    output_path = project_root / "outputs" / "forecasting_comparison_comprehensive.csv"
    results_df.to_csv(output_path, index=False)
    print(f"\n✓ Results saved to: {output_path}")
    
    # Display formatted table
    print("\n" + "=" * 100)
    print(f"{'Model':<20} {'Method':<30} {'NRMSE':<10} {'RMSE':<10} {'MAE':<10} {'R²':<10}")
    print("=" * 100)
    
    for _, row in results_df.iterrows():
        print(f"{row['Model']:<20} {row['Method']:<30} "
              f"{row['nrmse']:<10.4f} {row['rmse']:<10.2f} "
              f"{row['mae']:<10.2f} {row['r2']:<10.4f}")
    
    # Best models
    print("\n" + "=" * 80)
    print("BEST MODELS BY METHOD")
    print("=" * 80)
    
    for method in results_df['Method'].unique():
        method_results = results_df[results_df['Method'] == method]
        best = method_results.loc[method_results['nrmse'].idxmin()]
        print(f"\n{method}:")
        print(f"  🏆 {best['Model']} - NRMSE: {best['nrmse']:.4f}")
    
    print("\n" + "=" * 80)
    print("ANALYSIS COMPLETE")
    print("=" * 80)
    print("\nKey Insights:")
    print("1. One-step-ahead gives OPTIMISTIC estimates (uses true recent values)")
    print("2. Multi-step shows REALISTIC performance (error accumulation)")
    print("3. Cross-validation provides ROBUST estimates across time periods")
    print("4. Nested CV gives UNBIASED estimates when tuning hyperparameters")


if __name__ == "__main__":
    main()
