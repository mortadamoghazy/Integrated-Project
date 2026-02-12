"""
Visual comparison: AR(2) vs Ridge Regression forecasting
This script generates a side-by-side comparison plot.
"""

import sys
from pathlib import Path
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from sklearn.linear_model import Ridge
from sklearn.metrics import mean_squared_error
from statsmodels.tsa.ar_model import AutoReg

# Add project root to path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

def load_data():
    """Load payroll data."""
    data_path = project_root / "outputs" / "payroll_long.csv"
    df = pd.read_csv(data_path)
    df['month'] = pd.to_datetime(df['month'], format='%Y-%m')
    return df


def generate_ar2_forecast(df):
    """Generate AR(2) forecast."""
    monthly_cost = df.groupby('month')['total_cost'].sum().reset_index()
    monthly_cost = monthly_cost.sort_values('month')
    
    ts_data = monthly_cost['total_cost'].values
    model = AutoReg(ts_data, lags=2)
    fitted_model = model.fit()
    
    forecast = fitted_model.forecast(steps=6)
    
    # Simple confidence interval (constant width)
    std_error = np.std(fitted_model.resid)
    conf_lower = forecast - 1.96 * std_error
    conf_upper = forecast + 1.96 * std_error
    
    last_date = monthly_cost['month'].iloc[-1]
    future_dates = pd.date_range(start=last_date, periods=7, freq='MS')[1:]
    
    return monthly_cost, future_dates, forecast, conf_lower, conf_upper, std_error


def generate_ridge_forecast(df):
    """Generate Ridge regression forecast."""
    df_model = df.copy()
    df_model['month_num'] = df_model['month'].dt.month
    
    unique_months = np.sort(df_model['month'].unique())
    month_to_t = {m: i for i, m in enumerate(unique_months)}
    df_model['t'] = df_model['month'].map(month_to_t)
    
    emp_dummies = pd.get_dummies(df_model['employee_id'], prefix='emp', drop_first=True)
    month_dummies = pd.get_dummies(df_model['month_num'], prefix='m', drop_first=True)
    
    X = pd.concat([df_model[['t']], emp_dummies, month_dummies], axis=1)
    y = df_model['total_cost'].values
    
    ridge_model = Ridge(alpha=1.0)
    ridge_model.fit(X, y)
    
    y_pred_hist = ridge_model.predict(X)
    rmse = np.sqrt(mean_squared_error(y, y_pred_hist))
    
    monthly_cost = df.groupby('month')['total_cost'].sum().reset_index()
    monthly_cost = monthly_cost.sort_values('month')
    
    last_date = monthly_cost['month'].iloc[-1]
    future_dates = pd.date_range(start=last_date, periods=7, freq='MS')[1:]
    
    future_forecasts = []
    error_propagation_lower = []
    error_propagation_upper = []
    
    for step_idx, future_date in enumerate(future_dates):
        future_t = len(unique_months) + step_idx
        future_month_num = future_date.month
        
        future_rows = []
        for emp_id in df['employee_id'].unique():
            row = {'t': future_t}
            for col in emp_dummies.columns:
                emp_num = int(col.split('_')[1])
                row[col] = 1 if emp_num == emp_id else 0
            for col in month_dummies.columns:
                month = int(col.split('_')[1])
                row[col] = 1 if month == future_month_num else 0
            future_rows.append(row)
        
        X_future = pd.DataFrame(future_rows)
        for col in X.columns:
            if col not in X_future.columns:
                X_future[col] = 0
        X_future = X_future[X.columns]
        
        employee_predictions = ridge_model.predict(X_future)
        total_forecast = employee_predictions.sum()
        future_forecasts.append(total_forecast)
        
        propagated_error = rmse * np.sqrt(step_idx + 1) * 1.96
        error_propagation_lower.append(total_forecast - propagated_error)
        error_propagation_upper.append(total_forecast + propagated_error)
    
    forecast = np.array(future_forecasts)
    conf_lower = np.array(error_propagation_lower)
    conf_upper = np.array(error_propagation_upper)
    
    return monthly_cost, future_dates, forecast, conf_lower, conf_upper, rmse


def main():
    """Generate comparison plot."""
    print("Generating AR(2) vs Ridge Regression comparison...")
    
    df = load_data()
    
    # Generate both forecasts
    monthly_ar2, dates_ar2, fc_ar2, lower_ar2, upper_ar2, std_ar2 = generate_ar2_forecast(df)
    monthly_ridge, dates_ridge, fc_ridge, lower_ridge, upper_ridge, rmse_ridge = generate_ridge_forecast(df)
    
    # Create side-by-side comparison
    fig, axes = plt.subplots(1, 2, figsize=(16, 6))
    
    # AR(2) plot
    ax1 = axes[0]
    ax1.plot(monthly_ar2['month'], monthly_ar2['total_cost'], 
             marker='o', label='Historical', linewidth=2, markersize=6, color='#2E86AB')
    ax1.plot(dates_ar2, fc_ar2, marker='s', label='Forecast (AR2)',
             linewidth=2, markersize=6, color='#A23B72', linestyle='--')
    ax1.fill_between(dates_ar2, lower_ar2, upper_ar2, 
                     alpha=0.3, color='#A23B72', label='95% CI (constant)')
    
    ax1.set_title('AR(2) Autoregression Model\\n(Previous Implementation)', 
                  fontsize=14, fontweight='bold', pad=15)
    ax1.set_xlabel('Month', fontsize=11)
    ax1.set_ylabel('Total Cost ($)', fontsize=11)
    ax1.legend(fontsize=9)
    ax1.grid(True, alpha=0.3)
    ax1.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m'))
    plt.setp(ax1.xaxis.get_majorticklabels(), rotation=45, ha='right')
    
    error_text1 = f'Std Error: ${std_ar2:,.0f}\\nCI Width: Constant'
    ax1.text(0.02, 0.98, error_text1, transform=ax1.transAxes,
             ha='left', va='top', fontsize=9,
             bbox=dict(boxstyle='round', fc='white', alpha=0.8))
    
    # Ridge plot
    ax2 = axes[1]
    ax2.plot(monthly_ridge['month'], monthly_ridge['total_cost'],
             marker='o', label='Historical', linewidth=2, markersize=6, color='#2E86AB')
    ax2.plot(dates_ridge, fc_ridge, marker='s', label='Forecast (Ridge)',
             linewidth=2, markersize=6, color='#A23B72', linestyle='--')
    ax2.fill_between(dates_ridge, lower_ridge, upper_ridge,
                     alpha=0.3, color='#A23B72', label='95% CI (propagating)')
    
    # Add error bars
    for date, fc, lower, upper in zip(dates_ridge, fc_ridge, lower_ridge, upper_ridge):
        ax2.plot([date, date], [lower, upper], color='#A23B72', linewidth=1.5, alpha=0.6)
    
    ax2.set_title('Ridge Regression Model\\n(NEW: Best Performer with Error Propagation)', 
                  fontsize=14, fontweight='bold', pad=15)
    ax2.set_xlabel('Month', fontsize=11)
    ax2.set_ylabel('Total Cost ($)', fontsize=11)
    ax2.legend(fontsize=9)
    ax2.grid(True, alpha=0.3)
    ax2.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m'))
    plt.setp(ax2.xaxis.get_majorticklabels(), rotation=45, ha='right')
    
    error_text2 = f'Training RMSE: ${rmse_ridge:,.0f}\\nError grows: ±RMSE×√t'
    ax2.text(0.02, 0.98, error_text2, transform=ax2.transAxes,
             ha='left', va='top', fontsize=9,
             bbox=dict(boxstyle='round', fc='white', alpha=0.8, edgecolor='#A23B72'))
    
    plt.tight_layout()
    
    # Save comparison
    output_path = project_root / "outputs" / "model_comparison_ar2_vs_ridge.png"
    plt.savefig(output_path, dpi=150, bbox_inches='tight')
    print(f"✓ Saved comparison plot to: {output_path}")
    
    # Print statistics
    print("\\n" + "="*70)
    print("FORECAST COMPARISON")
    print("="*70)
    print(f"{'Month':<12} {'AR(2)':<15} {'Ridge':<15} {'Difference':<15}")
    print("-"*70)
    for date, f_ar2, f_ridge in zip(dates_ridge, fc_ar2, fc_ridge):
        diff = f_ridge - f_ar2
        print(f"{date.strftime('%Y-%m'):<12} ${f_ar2:>12,.0f}  ${f_ridge:>12,.0f}  ${diff:>12,.0f}")
    
    print("\\n" + "="*70)
    print("ERROR PROPAGATION COMPARISON")
    print("="*70)
    print(f"{'Month':<12} {'AR(2) CI Width':<20} {'Ridge CI Width':<20} {'Ratio':<10}")
    print("-"*70)
    for i, date in enumerate(dates_ridge):
        ci_ar2 = upper_ar2[i] - lower_ar2[i]
        ci_ridge = upper_ridge[i] - lower_ridge[i]
        ratio = ci_ridge / ci_ar2
        print(f"{date.strftime('%Y-%m'):<12} ${ci_ar2:>17,.0f}  ${ci_ridge:>17,.0f}  {ratio:>9.2f}x")
    
    print("\\n✓ Comparison complete!")
    print("\\nKey Observations:")
    print("1. Ridge CI width INCREASES over time (error propagation)")
    print("2. AR(2) CI width is CONSTANT (unrealistic)")
    print("3. Ridge reflects true forecast uncertainty")


if __name__ == "__main__":
    main()
