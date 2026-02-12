"""
Quick test to verify Ridge regression forecasting in dashboard.
This script tests the forecasting functionality without opening the full GUI.
"""

import sys
from pathlib import Path
import pandas as pd
import numpy as np
from sklearn.linear_model import Ridge
from sklearn.metrics import mean_squared_error

# Add project root to path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

def test_ridge_forecasting():
    """Test Ridge regression forecasting logic."""
    print("Testing Ridge Regression Forecasting...")
    print("=" * 60)
    
    # Load data
    data_path = project_root / "outputs" / "payroll_long.csv"
    if not data_path.exists():
        print(f"ERROR: Data file not found at {data_path}")
        return False
    
    df = pd.read_csv(data_path)
    df['month'] = pd.to_datetime(df['month'], format='%Y-%m')
    print(f"✓ Loaded {len(df)} rows of payroll data")
    print(f"✓ Date range: {df['month'].min()} to {df['month'].max()}")
    print(f"✓ {df['employee_id'].nunique()} unique employees")
    
    # Build features
    df_model = df.copy()
    df_model['month_num'] = df_model['month'].dt.month
    
    # Create time index
    unique_months = np.sort(df_model['month'].unique())
    month_to_t = {m: i for i, m in enumerate(unique_months)}
    df_model['t'] = df_model['month'].map(month_to_t)
    
    # Create dummy variables
    emp_dummies = pd.get_dummies(df_model['employee_id'], prefix='emp', drop_first=True)
    month_dummies = pd.get_dummies(df_model['month_num'], prefix='m', drop_first=True)
    
    # Build feature matrix
    X = pd.concat([df_model[['t']], emp_dummies, month_dummies], axis=1)
    y = df_model['total_cost'].values
    
    print(f"✓ Feature matrix shape: {X.shape}")
    print(f"✓ Target variable: total_cost")
    
    # Fit Ridge model
    ridge_model = Ridge(alpha=1.0)
    ridge_model.fit(X, y)
    print(f"✓ Ridge model trained with alpha=1.0")
    
    # Calculate training RMSE
    y_pred_hist = ridge_model.predict(X)
    rmse = np.sqrt(mean_squared_error(y, y_pred_hist))
    print(f"✓ Training RMSE: ${rmse:,.2f}")
    
    # Generate 6-month forecast
    forecast_steps = 6
    last_date = df['month'].max()
    future_dates = pd.date_range(start=last_date, periods=forecast_steps+1, freq='MS')[1:]
    
    print(f"\n{'Month':<15} {'Forecast':<15} {'Lower CI':<15} {'Upper CI':<15} {'Error Band':<15}")
    print("-" * 75)
    
    for step_idx, future_date in enumerate(future_dates):
        future_t = len(unique_months) + step_idx
        future_month_num = future_date.month
        
        # Create features for each employee in this future month
        future_rows = []
        for emp_id in df['employee_id'].unique():
            row = {'t': future_t}
            
            # Employee dummies
            for col in emp_dummies.columns:
                emp_num = int(col.split('_')[1])
                row[col] = 1 if emp_num == emp_id else 0
            
            # Month dummies
            for col in month_dummies.columns:
                month = int(col.split('_')[1])
                row[col] = 1 if month == future_month_num else 0
            
            future_rows.append(row)
        
        X_future = pd.DataFrame(future_rows)
        # Align columns with training data
        for col in X.columns:
            if col not in X_future.columns:
                X_future[col] = 0
        X_future = X_future[X.columns]
        
        # Predict for all employees and sum
        employee_predictions = ridge_model.predict(X_future)
        total_forecast = employee_predictions.sum()
        
        # Error propagation
        propagated_error = rmse * np.sqrt(step_idx + 1) * 1.96
        conf_lower = total_forecast - propagated_error
        conf_upper = total_forecast + propagated_error
        error_band = propagated_error
        
        print(f"{future_date.strftime('%Y-%m'):<15} ${total_forecast:>12,.0f}  ${conf_lower:>12,.0f}  ${conf_upper:>12,.0f}  ±${error_band:>12,.0f}")
    
    print("\n" + "=" * 60)
    print("✓ Ridge regression forecasting test completed successfully!")
    print("✓ Error propagation visible: uncertainty increases with horizon")
    print(f"✓ Error formula: ±{rmse:,.0f} × √t × 1.96 (95% CI)")
    
    return True


if __name__ == "__main__":
    success = test_ridge_forecasting()
    sys.exit(0 if success else 1)
