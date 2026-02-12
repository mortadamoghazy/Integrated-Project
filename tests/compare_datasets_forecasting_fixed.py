"""
Comprehensive Forecasting Comparison: Startup (30) vs SME (250)
Compare AR(1), AR(2), Ridge, Lasso, Pooled FE, Unpooled models
"""

import pandas as pd
import numpy as np
from sklearn.linear_model import Ridge, Lasso, LinearRegression
from statsmodels.tsa.ar_model import AutoReg
from sklearn.metrics import mean_squared_error, mean_absolute_error
import warnings
warnings.filterwarnings('ignore')


def load_startup_data():
    """Load startup (30 employees) payroll data"""
    df = pd.read_csv('outputs/payroll_long.csv')
    df['month'] = pd.to_datetime(df['month'])
    return df


def load_sme_data():
    """Load SME (250 employees) payroll data from Excel"""
    excel_path = 'synthetic_payroll_sme.xlsx'
    
    # Read all sheets (organized by month)
    xl_file = pd.ExcelFile(excel_path)
    dfs = []
    
    for sheet in xl_file.sheet_names:
        # Sheet names are months like '2023-01', '2023-02', etc.
        df_sheet = pd.read_excel(excel_path, sheet_name=sheet)
        df_sheet['month'] = pd.to_datetime(sheet + '-01')  # Convert sheet name to date
        df_sheet = df_sheet[['employee_id', 'month', 'total_cost']]
        dfs.append(df_sheet)
    
    df = pd.concat(dfs, ignore_index=True)
    df = df.sort_values(['employee_id', 'month']).reset_index(drop=True)
    return df


def build_panel_features(df):
    """Build panel regression features: time trend + employee dummies + seasonality"""
    df = df.copy()
    
    # Time features
    df = df.sort_values(['employee_id', 'month'])
    df['time_idx'] = df.groupby('employee_id').cumcount()
    
    # Month dummies for seasonality
    df['month_num'] = df['month'].dt.month
    month_dummies = pd.get_dummies(df['month_num'], prefix='month', drop_first=True)
    
    # Employee dummies
    emp_dummies = pd.get_dummies(df['employee_id'], prefix='emp', drop_first=True)
    
    # Combine features
    X = pd.concat([
        df[['time_idx']],
        emp_dummies,
        month_dummies
    ], axis=1)
    
    y = df['total_cost'].values
    
    return X, y, df


def evaluate_ar_models(df, dataset_name):
    """Evaluate AR(1) and AR(2) models"""
    results = []
    
    # Group by employee
    grouped = df.groupby('employee_id')
    
    for p in [1, 2]:  # AR(1) and AR(2) only
        all_rmse = []
        all_mae = []
        all_nrmse = []
        
        for emp_id, emp_df in grouped:
            emp_df = emp_df.sort_values('month').reset_index(drop=True)
            y = emp_df['total_cost'].values
            
            # Train/Val/Test split
            n = len(y)
            test_size = 6
            train_end = n - test_size
            
            if train_end <= p + 5:  # Need enough data
                continue
            
            y_train = y[:train_end]
            y_test = y[train_end:]
            
            try:
                # Fit AR model
                model = AutoReg(y_train, lags=p, trend='c')
                model_fit = model.fit()
                
                # Forecast
                y_pred = model_fit.predict(start=train_end, end=n-1)
                
                # Metrics
                rmse = np.sqrt(mean_squared_error(y_test, y_pred))
                mae = mean_absolute_error(y_test, y_pred)
                nrmse = rmse / (y_test.max() - y_test.min()) if y_test.max() != y_test.min() else 0
                
                all_rmse.append(rmse)
                all_mae.append(mae)
                all_nrmse.append(nrmse)
                
            except:
                continue
        
        # Average metrics
        if len(all_rmse) > 0:
            results.append({
                'Dataset': dataset_name,
                'Model': f'AR({p})',
                'Method': 'Train/Val/Test',
                'RMSE': np.mean(all_rmse),
                'MAE': np.mean(all_mae),
                'NRMSE': np.mean(all_nrmse)
            })
    
    return results


def evaluate_panel_models(df, dataset_name):
    """Evaluate Ridge, Lasso, Pooled FE, and Unpooled models"""
    results = []
    
    # Build panel features
    X, y, df_features = build_panel_features(df)
    
    # Train/test split (last 6 months as test)
    n_months = df['month'].nunique()
    test_months = 6
    train_months = n_months - test_months
    
    # Get train/test indices
    months_sorted = sorted(df['month'].unique())
    train_months_set = set(months_sorted[:train_months])
    
    train_idx = df['month'].isin(train_months_set)
    test_idx = ~train_idx
    
    X_train, X_test = X[train_idx], X[test_idx]
    y_train, y_test = y[train_idx], y[test_idx]
    
    # Model configurations
    models_config = [
        ('Ridge', Ridge(alpha=1.0)),
        ('Lasso', Lasso(alpha=1.0, max_iter=5000)),
        ('Pooled FE', Ridge(alpha=0.1))
    ]
    
    # Evaluate each model
    for name, model in models_config:
        model.fit(X_train, y_train)
        y_pred = model.predict(X_test)
        
        rmse = np.sqrt(mean_squared_error(y_test, y_pred))
        mae = mean_absolute_error(y_test, y_pred)
        
        # NRMSE calculation
        y_range = y_test.max() - y_test.min()
        nrmse = rmse / y_range if y_range > 0 else 0
        
        results.append({
            'Dataset': dataset_name,
            'Model': name,
            'Method': 'Train/Val/Test',
            'RMSE': rmse,
            'MAE': mae,
            'NRMSE': nrmse
        })
    
    # Unpooled model (per-employee linear regression)
    grouped = df.groupby('employee_id')
    all_rmse = []
    all_mae = []
    all_nrmse = []
    
    for emp_id, emp_df in grouped:
        emp_df = emp_df.sort_values('month').reset_index(drop=True)
        
        # Time trend for this employee
        emp_df['time_idx'] = range(len(emp_df))
        X_emp = emp_df[['time_idx']].values
        y_emp = emp_df['total_cost'].values
        
        # Train/test split
        n = len(y_emp)
        train_end = n - test_months
        
        if train_end <= 1:
            continue
        
        X_train_emp = X_emp[:train_end]
        X_test_emp = X_emp[train_end:]
        y_train_emp = y_emp[:train_end]
        y_test_emp = y_emp[train_end:]
        
        # Fit linear regression
        model = LinearRegression()
        model.fit(X_train_emp, y_train_emp)
        y_pred_emp = model.predict(X_test_emp)
        
        # Metrics
        rmse = np.sqrt(mean_squared_error(y_test_emp, y_pred_emp))
        mae = mean_absolute_error(y_test_emp, y_pred_emp)
        y_range = y_test_emp.max() - y_test_emp.min()
        nrmse = rmse / y_range if y_range > 0 else 0
        
        all_rmse.append(rmse)
        all_mae.append(mae)
        all_nrmse.append(nrmse)
    
    # Average unpooled metrics
    if len(all_rmse) > 0:
        results.append({
            'Dataset': dataset_name,
            'Model': 'Unpooled',
            'Method': 'Train/Val/Test',
            'RMSE': np.mean(all_rmse),
            'MAE': np.mean(all_mae),
            'NRMSE': np.mean(all_nrmse)
        })
    
    return results


def generate_latex_table(df_results):
    """Generate LaTeX table showing side-by-side comparison"""
    # Pivot to get side-by-side comparison
    startup = df_results[df_results['Dataset'] == 'Startup (30)'].copy()
    sme = df_results[df_results['Dataset'] == 'SME (250)'].copy()
    
    # Sort by model order
    model_order = ['AR(1)', 'AR(2)', 'Ridge', 'Lasso', 'Pooled FE', 'Unpooled']
    startup['Model'] = pd.Categorical(startup['Model'], categories=model_order, ordered=True)
    sme['Model'] = pd.Categorical(sme['Model'], categories=model_order, ordered=True)
    startup = startup.sort_values('Model')
    sme = sme.sort_values('Model')
    
    # Find best models by NRMSE for each dataset
    startup_best = startup.loc[startup['NRMSE'].idxmin(), 'Model']
    sme_best = sme.loc[sme['NRMSE'].idxmin(), 'Model']
    
    # Start LaTeX table
    latex = r"""
\begin{table}[htbp]
\centering
\caption{Forecasting Performance Comparison: Startup (30 Employees) vs SME (250 Employees)}
\label{tab:forecasting_comparison}
\resizebox{\textwidth}{!}{%
\begin{tabular}{@{} l l ccc ccc @{}}
\toprule
\multirow{2}{*}{\textbf{Model}} & \multirow{2}{*}{\textbf{Method}} & 
\multicolumn{3}{c}{\textbf{Startup (30 Employees)}} & 
\multicolumn{3}{c}{\textbf{SME (250 Employees)}} \\
\cmidrule(lr){3-5} \cmidrule(lr){6-8}
& & \textbf{RMSE} & \textbf{MAE} & \textbf{NRMSE} & 
\textbf{RMSE} & \textbf{MAE} & \textbf{NRMSE} \\
\midrule
"""
    
    # Add rows for each model
    for model in model_order:
        s_row = startup[startup['Model'] == model]
        sme_row = sme[sme['Model'] == model]
        
        if len(s_row) == 0 or len(sme_row) == 0:
            continue
            
        method = s_row['Method'].values[0]
        
        # Highlight best models (Pooled FE for both)
        if model == startup_best or model == sme_best:
            latex += r"\textbf{" + model + r"} & " + method + r" & "
            latex += r"\textbf{" + f"{s_row['RMSE'].values[0]:.1f}" + r"} & "
            latex += r"\textbf{" + f"{s_row['MAE'].values[0]:.1f}" + r"} & "
            latex += r"\textbf{" + f"{s_row['NRMSE'].values[0]:.4f}" + r"} & "
            latex += r"\textbf{" + f"{sme_row['RMSE'].values[0]:.1f}" + r"} & "
            latex += r"\textbf{" + f"{sme_row['MAE'].values[0]:.1f}" + r"} & "
            latex += r"\textbf{" + f"{sme_row['NRMSE'].values[0]:.4f}" + r"} \\"
        else:
            latex += model + r" & " + method + r" & "
            latex += f"{s_row['RMSE'].values[0]:.1f}" + r" & "
            latex += f"{s_row['MAE'].values[0]:.1f}" + r" & "
            latex += f"{s_row['NRMSE'].values[0]:.4f}" + r" & "
            latex += f"{sme_row['RMSE'].values[0]:.1f}" + r" & "
            latex += f"{sme_row['MAE'].values[0]:.1f}" + r" & "
            latex += f"{sme_row['NRMSE'].values[0]:.4f}" + r" \\"
        
        latex += "\n"
    
    # Close table
    latex += r"""\bottomrule
\end{tabular}%
}
\end{table}
"""
    
    return latex


def main():
    """Main execution"""
    print("\n" + "="*60)
    print("COMPREHENSIVE FORECASTING COMPARISON")
    print("Startup (30 Employees) vs SME (250 Employees)")
    print("="*60 + "\n")
    
    # Load both datasets
    print("Loading datasets...")
    df_startup = load_startup_data()
    df_sme = load_sme_data()
    print(f"✓ Startup: {len(df_startup)} rows, {df_startup['employee_id'].nunique()} employees")
    print(f"✓ SME: {len(df_sme)} rows, {df_sme['employee_id'].nunique()} employees")
    
    # Results storage
    results = []
    
    # Evaluate AR models on both datasets
    print("\nEvaluating AR models...")
    results.extend(evaluate_ar_models(df_startup, 'Startup (30)'))
    results.extend(evaluate_ar_models(df_sme, 'SME (250)'))
    
    # Evaluate panel models on both datasets
    print("Evaluating panel models...")
    results.extend(evaluate_panel_models(df_startup, 'Startup (30)'))
    results.extend(evaluate_panel_models(df_sme, 'SME (250)'))
    
    # Convert to DataFrame
    df_results = pd.DataFrame(results)
    
    # Save to CSV
    output_path = 'outputs/dataset_comparison_results.csv'
    df_results.to_csv(output_path, index=False)
    print(f"\n✓ Results saved to: {output_path}\n")
    
    # Generate LaTeX table
    print("\nLATEX TABLE")
    print("="*60)
    latex = generate_latex_table(df_results)
    print(latex)
    
    # Save LaTeX to file
    latex_path = 'outputs/forecasting_comparison_table.tex'
    with open(latex_path, 'w') as f:
        f.write(latex)
    print(f"\n✓ LaTeX table saved to: {latex_path}")
    
    # Print summary
    print("\n" + "="*60)
    print("SUMMARY")
    print("="*60)
    print("\nStartup (30 Employees):")
    startup_results = df_results[df_results['Dataset'] == 'Startup (30)']
    print(startup_results[['Model', 'RMSE', 'MAE', 'NRMSE']].to_string(index=False))
    
    print("\nSME (250 Employees):")
    sme_results = df_results[df_results['Dataset'] == 'SME (250)']
    print(sme_results[['Model', 'RMSE', 'MAE', 'NRMSE']].to_string(index=False))
    
    # Find best models
    startup_best = startup_results.loc[startup_results['NRMSE'].idxmin()]
    sme_best = sme_results.loc[sme_results['NRMSE'].idxmin()]
    
    print("\n" + "="*60)
    print("BEST MODELS (by NRMSE)")
    print("="*60)
    print(f"Startup: {startup_best['Model']} (NRMSE={startup_best['NRMSE']:.4f})")
    print(f"SME:     {sme_best['Model']} (NRMSE={sme_best['NRMSE']:.4f})")


if __name__ == '__main__':
    main()
