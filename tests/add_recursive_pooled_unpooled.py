"""
Add Recursive Multi-Step Evaluation for Pooled FE and Unpooled Models
This extends the simplified_forecasting_comparison.py to include missing evaluations
"""

import sys
from pathlib import Path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pandas as pd
import numpy as np
import warnings
warnings.filterwarnings('ignore')

from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_squared_error, mean_absolute_error, r2_score

# Configuration
TARGET_LABEL = "salaire_brut"
CSV_PATH = "outputs/payroll_long.csv"
TEST_MONTHS = 6
VAL_MONTHS = 3


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


def pooled_fe_recursive_multistep(df):
    """Pooled FE with RECURSIVE multi-step forecasting.
    
    Traditional Pooled FE without lagged features - relies only on time trend and employee fixed effects.
    """
    # Get unique months
    unique_months = np.sort(df['month'].unique())
    n_months = len(unique_months)
    
    # Split by time (no lags needed)
    test_start_idx = n_months - TEST_MONTHS
    val_start_idx = test_start_idx - VAL_MONTHS
    
    df = df.sort_values(['employee_id', 'month']).copy()
    df['month_idx'] = df['month'].map({m: i for i, m in enumerate(unique_months)})
    
    train_df = df[df['month_idx'] < val_start_idx].copy()
    val_df = df[(df['month_idx'] >= val_start_idx) & (df['month_idx'] < test_start_idx)].copy()
    test_df = df[df['month_idx'] >= test_start_idx].copy()
    
    if len(train_df) == 0 or len(test_df) == 0:
        return {k: np.nan for k in ['rmse', 'nrmse', 'mae', 'r2']}
    
    # Features: time index + employee dummies only (no lags)
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
    
    # Train model on train+val with Ridge regularization (alpha=0.1)
    X_train_val = pd.concat([X_train, X_val], ignore_index=True)
    y_train_val = np.concatenate([y_train, y_val])
    
    # Align columns
    all_cols = X_train_val.columns
    
    from sklearn.linear_model import Ridge as RidgeReg
    model = RidgeReg(alpha=0.1)
    model.fit(X_train_val, y_train_val)
    
    # RECURSIVE forecasting (no lags, just linear extrapolation)
    test_months_idx = sorted(test_df['month_idx'].unique())
    
    all_y_true = []
    all_y_pred = []
    
    # Forecast month by month
    for month_idx in test_months_idx:
        month_data = test_df[test_df['month_idx'] == month_idx].copy()
        
        # Build features
        features_list = []
        y_true_list = []
        
        for _, row in month_data.iterrows():
            emp_id = row['employee_id']
            
            # Build feature row
            feature_row = {'month_idx': month_idx}
            
            # Add employee dummies
            for col in all_cols:
                if col.startswith('emp_'):
                    emp_num = int(col.split('_')[1])
                    feature_row[col] = 1 if emp_id == emp_num else 0
            
            features_list.append(feature_row)
            y_true_list.append(row[TARGET_LABEL])
        
        if len(features_list) == 0:
            continue
        
        # Create DataFrame with all required columns
        X_month = pd.DataFrame(features_list)
        
        # Ensure all columns from training are present
        for col in all_cols:
            if col not in X_month.columns:
                X_month[col] = 0
        
        X_month = X_month[all_cols]  # Reorder columns
        
        # Predict
        y_pred_month = model.predict(X_month)
        
        all_y_true.extend(y_true_list)
        all_y_pred.extend(y_pred_month)
    
    if len(all_y_true) == 0:
        return {k: np.nan for k in ['rmse', 'nrmse', 'mae', 'r2']}
    
    return compute_metrics(all_y_true, all_y_pred)


def unpooled_recursive_multistep(df):
    """Unpooled (per-employee) with RECURSIVE multi-step forecasting."""
    employees = df['employee_id'].unique()
    all_y_true, all_y_pred = [], []
    
    for emp in employees:
        sub = df[df['employee_id'] == emp].copy().sort_values('month')
        
        if len(sub) <= TEST_MONTHS + VAL_MONTHS + 5:
            continue
        
        # Create lagged features
        for lag in [1, 2]:
            sub[f'y_lag{lag}'] = sub[TARGET_LABEL].shift(lag)
        sub = sub.dropna().copy()
        
        if len(sub) <= TEST_MONTHS + VAL_MONTHS + 3:
            continue
        
        sub['t'] = np.arange(len(sub))
        y = sub[TARGET_LABEL].values
        
        # Split
        test_size = TEST_MONTHS
        val_size = VAL_MONTHS
        train_size = len(sub) - test_size - val_size
        
        if train_size < 3:
            continue
        
        # Train on train+val with lags
        X_train_val = sub.iloc[:train_size+val_size][['t', 'y_lag1', 'y_lag2']].values
        y_train_val = y[:train_size+val_size]
        y_test = y[train_size+val_size:]
        
        # Fit model
        model = LinearRegression()
        model.fit(X_train_val, y_train_val)
        
        # RECURSIVE forecasting
        y_pred_recursive = []
        last_known_values = list(y[:train_size+val_size][-2:])  # Last 2 values
        t_start = train_size + val_size
        
        for step in range(test_size):
            t_current = t_start + step
            # Features: [time, lag1, lag2]
            features = [t_current, last_known_values[-1], last_known_values[-2]]
            X_pred = np.array(features).reshape(1, -1)
            
            y_next = model.predict(X_pred)[0]
            y_pred_recursive.append(y_next)
            
            # Update history with PREDICTION
            last_known_values.append(y_next)
            last_known_values = last_known_values[-2:]  # Keep last 2
        
        all_y_true.extend(y_test)
        all_y_pred.extend(y_pred_recursive)
    
    if len(all_y_true) == 0:
        return {k: np.nan for k in ['rmse', 'nrmse', 'mae', 'r2']}
    
    return compute_metrics(all_y_true, all_y_pred)


def main():
    print("\n" + "="*80)
    print("ADDING RECURSIVE MULTI-STEP FOR POOLED FE AND UNPOOLED")
    print("="*80 + "\n")
    
    # Load data
    df = load_data()
    print(f"✓ Loaded {len(df)} records, {df['employee_id'].nunique()} employees\n")
    
    # Evaluate Pooled FE Recursive
    print("Evaluating Pooled FE (Recursive Multi-Step)...")
    pooled_metrics = pooled_fe_recursive_multistep(df)
    print(f"  NRMSE: {pooled_metrics['nrmse']:.4f}")
    print(f"  RMSE: {pooled_metrics['rmse']:.1f}")
    print(f"  MAE: {pooled_metrics['mae']:.1f}")
    
    # Evaluate Unpooled Recursive
    print("\nEvaluating Unpooled (Recursive Multi-Step)...")
    unpooled_metrics = unpooled_recursive_multistep(df)
    print(f"  NRMSE: {unpooled_metrics['nrmse']:.4f}")
    print(f"  RMSE: {unpooled_metrics['rmse']:.1f}")
    print(f"  MAE: {unpooled_metrics['mae']:.1f}")
    
    # Load existing results
    existing_results = pd.read_csv('outputs/forecasting_simplified_comparison.csv')
    
    # Add new results
    new_rows = [
        {
            'Model': 'Pooled FE',
            'Method': 'Recursive Multi-Step',
            'rmse': pooled_metrics['rmse'],
            'nrmse': pooled_metrics['nrmse'],
            'mae': pooled_metrics['mae'],
            'r2': pooled_metrics['r2']
        },
        {
            'Model': 'Unpooled',
            'Method': 'Recursive Multi-Step',
            'rmse': unpooled_metrics['rmse'],
            'nrmse': unpooled_metrics['nrmse'],
            'mae': unpooled_metrics['mae'],
            'r2': unpooled_metrics['r2']
        }
    ]
    
    # Remove old Pooled FE and Unpooled Recursive entries if they exist
    existing_results = existing_results[
        ~((existing_results['Model'].isin(['Pooled FE', 'Unpooled'])) & 
          (existing_results['Method'] == 'Recursive Multi-Step'))
    ]
    
    # Append new rows
    updated_results = pd.concat([existing_results, pd.DataFrame(new_rows)], ignore_index=True)
    
    # Save
    updated_results.to_csv('outputs/forecasting_simplified_comparison.csv', index=False)
    print(f"\n✓ Updated results saved to: outputs/forecasting_simplified_comparison.csv")
    
    print("\n" + "="*80)
    print("COMPLETE")
    print("="*80)


if __name__ == '__main__':
    main()
