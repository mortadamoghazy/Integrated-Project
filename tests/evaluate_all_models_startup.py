"""
Complete Forecasting Evaluation: All Models on Startup Dataset
Includes AR(1-4), Ridge, Lasso, Pooled FE, Unpooled, XGBoost
"""

import pandas as pd
import numpy as np
from sklearn.linear_model import Ridge, Lasso, LinearRegression
from sklearn.ensemble import GradientBoostingRegressor
from statsmodels.tsa.ar_model import AutoReg
from sklearn.metrics import mean_squared_error, mean_absolute_error
import warnings
warnings.filterwarnings('ignore')


def load_startup_data():
    """Load startup (30 employees) payroll data"""
    df = pd.read_csv('outputs/payroll_long.csv')
    df['month'] = pd.to_datetime(df['month'])
    return df


def build_panel_features(df):
    """Build panel regression features"""
    df = df.copy()
    df = df.sort_values(['employee_id', 'month'])
    df['time_idx'] = df.groupby('employee_id').cumcount()
    df['month_num'] = df['month'].dt.month
    month_dummies = pd.get_dummies(df['month_num'], prefix='month', drop_first=True)
    emp_dummies = pd.get_dummies(df['employee_id'], prefix='emp', drop_first=True)
    X = pd.concat([df[['time_idx']], emp_dummies, month_dummies], axis=1)
    y = df['total_cost'].values
    return X, y, df


def evaluate_ar_models(df):
    """Evaluate AR(1-4) models"""
    results = []
    grouped = df.groupby('employee_id')
    
    for p in [1, 2, 3, 4]:
        all_rmse = []
        all_mae = []
        all_nrmse = []
        
        for emp_id, emp_df in grouped:
            emp_df = emp_df.sort_values('month').reset_index(drop=True)
            y = emp_df['total_cost'].values
            n = len(y)
            test_size = 6
            train_end = n - test_size
            
            if train_end <= p + 5:
                continue
            
            y_train = y[:train_end]
            y_test = y[train_end:]
            
            try:
                model = AutoReg(y_train, lags=p, trend='c')
                model_fit = model.fit()
                y_pred = model_fit.predict(start=train_end, end=n-1)
                
                rmse = np.sqrt(mean_squared_error(y_test, y_pred))
                mae = mean_absolute_error(y_test, y_pred)
                nrmse = rmse / (y_test.max() - y_test.min()) if y_test.max() != y_test.min() else 0
                
                all_rmse.append(rmse)
                all_mae.append(mae)
                all_nrmse.append(nrmse)
            except:
                continue
        
        if len(all_rmse) > 0:
            results.append({
                'Model': f'AR({p})',
                'Method': 'Train/Val/Test',
                'RMSE': np.mean(all_rmse),
                'MAE': np.mean(all_mae),
                'NRMSE': np.mean(all_nrmse)
            })
    
    return results


def evaluate_panel_models(df):
    """Evaluate Ridge, Lasso, Pooled FE, Unpooled, XGBoost"""
    results = []
    X, y, df_features = build_panel_features(df)
    
    # Train/test split
    n_months = df['month'].nunique()
    test_months = 6
    train_months = n_months - test_months
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
        ('Pooled FE', Ridge(alpha=0.1)),
        ('XGBoost', GradientBoostingRegressor(n_estimators=100, max_depth=3, random_state=42))
    ]
    
    for name, model in models_config:
        model.fit(X_train, y_train)
        y_pred = model.predict(X_test)
        
        rmse = np.sqrt(mean_squared_error(y_test, y_pred))
        mae = mean_absolute_error(y_test, y_pred)
        y_range = y_test.max() - y_test.min()
        nrmse = rmse / y_range if y_range > 0 else 0
        
        results.append({
            'Model': name,
            'Method': 'Train/Val/Test',
            'RMSE': rmse,
            'MAE': mae,
            'NRMSE': nrmse
        })
    
    # Unpooled model
    grouped = df.groupby('employee_id')
    all_rmse = []
    all_mae = []
    all_nrmse = []
    test_months_val = 6
    
    for emp_id, emp_df in grouped:
        emp_df = emp_df.sort_values('month').reset_index(drop=True)
        emp_df['time_idx'] = range(len(emp_df))
        X_emp = emp_df[['time_idx']].values
        y_emp = emp_df['total_cost'].values
        n = len(y_emp)
        train_end = n - test_months_val
        
        if train_end <= 1:
            continue
        
        X_train_emp = X_emp[:train_end]
        X_test_emp = X_emp[train_end:]
        y_train_emp = y_emp[:train_end]
        y_test_emp = y_emp[train_end:]
        
        model = LinearRegression()
        model.fit(X_train_emp, y_train_emp)
        y_pred_emp = model.predict(X_test_emp)
        
        rmse = np.sqrt(mean_squared_error(y_test_emp, y_pred_emp))
        mae = mean_absolute_error(y_test_emp, y_pred_emp)
        y_range = y_test_emp.max() - y_test_emp.min()
        nrmse = rmse / y_range if y_range > 0 else 0
        
        all_rmse.append(rmse)
        all_mae.append(mae)
        all_nrmse.append(nrmse)
    
    if len(all_rmse) > 0:
        results.append({
            'Model': 'Unpooled',
            'Method': 'Train/Val/Test',
            'RMSE': np.mean(all_rmse),
            'MAE': np.mean(all_mae),
            'NRMSE': np.mean(all_nrmse)
        })
    
    return results


def generate_latex_table(df_results):
    """Generate LaTeX table for all models on Startup dataset only"""
    model_order = ['AR(1)', 'AR(2)', 'AR(3)', 'AR(4)', 'Ridge', 'Lasso', 'Pooled FE', 'Unpooled', 'XGBoost']
    df_results['Model'] = pd.Categorical(df_results['Model'], categories=model_order, ordered=True)
    df_results = df_results.sort_values('Model')
    
    # Find best model
    best_model = df_results.loc[df_results['NRMSE'].idxmin(), 'Model']
    
    latex = r"""
\begin{table}[htbp]
\centering
\caption{Forecasting Model Performance Comparison: Startup (30 Employees)}
\label{tab:forecasting_comparison}
\begin{tabular}{@{} l l ccc @{}}
\toprule
\textbf{Model} & \textbf{Method} & \textbf{RMSE (€)} & \textbf{MAE (€)} & \textbf{NRMSE} \\
\midrule
"""
    
    for _, row in df_results.iterrows():
        model = row['Model']
        method = row['Method']
        rmse = row['RMSE']
        mae = row['MAE']
        nrmse = row['NRMSE']
        
        if model == best_model:
            latex += f"\\textbf{{{model}}} & {method} & \\textbf{{{rmse:.1f}}} & \\textbf{{{mae:.1f}}} & \\textbf{{{nrmse:.4f}}} \\\\\n"
        else:
            latex += f"{model} & {method} & {rmse:.1f} & {mae:.1f} & {nrmse:.4f} \\\\\n"
    
    latex += r"""\bottomrule
\end{tabular}
\end{table}
"""
    
    return latex


def main():
    print("\n" + "="*60)
    print("COMPLETE MODEL EVALUATION: STARTUP DATASET")
    print("="*60 + "\n")
    
    # Load data
    print("Loading dataset...")
    df = load_startup_data()
    print(f"✓ Startup: {len(df)} rows, {df['employee_id'].nunique()} employees\n")
    
    # Evaluate all models
    results = []
    
    print("Evaluating AR models (1-4)...")
    results.extend(evaluate_ar_models(df))
    
    print("Evaluating panel models + XGBoost...")
    results.extend(evaluate_panel_models(df))
    
    # Convert to DataFrame
    df_results = pd.DataFrame(results)
    
    # Save CSV
    output_path = 'outputs/all_models_startup_results.csv'
    df_results.to_csv(output_path, index=False)
    print(f"\n✓ Results saved to: {output_path}\n")
    
    # Generate LaTeX table
    print("\nLATEX TABLE")
    print("="*60)
    latex = generate_latex_table(df_results)
    print(latex)
    
    # Save LaTeX
    latex_path = 'outputs/forecasting_comparison_table.tex'
    with open(latex_path, 'w') as f:
        f.write(latex)
    print(f"\n✓ LaTeX table saved to: {latex_path}")
    
    # Print summary
    print("\n" + "="*60)
    print("RESULTS SUMMARY")
    print("="*60)
    print(df_results.to_string(index=False))
    
    best_idx = df_results['NRMSE'].idxmin()
    best = df_results.loc[best_idx]
    print("\n" + "="*60)
    print("BEST MODEL (by NRMSE)")
    print("="*60)
    print(f"{best['Model']}: NRMSE={best['NRMSE']:.4f}, RMSE=€{best['RMSE']:.1f}, MAE=€{best['MAE']:.1f}")


if __name__ == '__main__':
    main()
