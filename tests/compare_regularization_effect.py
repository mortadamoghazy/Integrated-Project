"""
Compare Regularization Effect on Pooled and Unpooled Models
Evaluate with and without Ridge regularization (different alpha values)
"""

import pandas as pd
import numpy as np
import warnings
warnings.filterwarnings('ignore')

from sklearn.linear_model import LinearRegression, Ridge
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


def pooled_train_val_test(df, alpha=0):
    """Pooled model with train/val/test split.
    
    Args:
        alpha: Ridge regularization parameter. 0 = OLS (unregularized)
    """
    unique_months = np.sort(df['month'].unique())
    
    test_months = unique_months[-TEST_MONTHS:]
    val_months = unique_months[-(TEST_MONTHS+VAL_MONTHS):-TEST_MONTHS]
    train_months = unique_months[:-(TEST_MONTHS+VAL_MONTHS)]
    
    train_df = df[df['month'].isin(train_months)].copy()
    val_df = df[df['month'].isin(val_months)].copy()
    test_df = df[df['month'].isin(test_months)].copy()
    
    train_df['t'] = train_df['month'].map({m: i for i, m in enumerate(unique_months)})
    val_df['t'] = val_df['month'].map({m: i for i, m in enumerate(unique_months)})
    test_df['t'] = test_df['month'].map({m: i for i, m in enumerate(unique_months)})
    
    emp_dummies_train = pd.get_dummies(train_df['employee_id'], prefix='emp', drop_first=True)
    emp_dummies_val = pd.get_dummies(val_df['employee_id'], prefix='emp', drop_first=True)
    emp_dummies_test = pd.get_dummies(test_df['employee_id'], prefix='emp', drop_first=True)
    
    all_emp_cols = set(emp_dummies_train.columns) | set(emp_dummies_val.columns) | set(emp_dummies_test.columns)
    for col in all_emp_cols:
        if col not in emp_dummies_train.columns:
            emp_dummies_train[col] = 0
        if col not in emp_dummies_val.columns:
            emp_dummies_val[col] = 0
        if col not in emp_dummies_test.columns:
            emp_dummies_test[col] = 0
    
    X_train = pd.concat([train_df[['t']], emp_dummies_train], axis=1)
    X_test = pd.concat([test_df[['t']], emp_dummies_test], axis=1)
    
    y_train = train_df[TARGET_LABEL].values
    y_test = test_df[TARGET_LABEL].values
    
    if alpha == 0:
        model = LinearRegression()
    else:
        model = Ridge(alpha=alpha)
    
    model.fit(X_train, y_train)
    y_pred = model.predict(X_test)
    
    return compute_metrics(y_test, y_pred)


def unpooled_train_val_test(df, alpha=0):
    """Unpooled (per-employee) with train/val/test split.
    
    Args:
        alpha: Ridge regularization parameter. 0 = OLS (unregularized)
    """
    employees = df['employee_id'].unique()
    all_y_true, all_y_pred = [], []
    
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
        
        X_train, X_test = X.iloc[:train_size], X.iloc[train_size+val_size:]
        y_train, y_test = y[:train_size], y[train_size+val_size:]
        
        if alpha == 0:
            model = LinearRegression()
        else:
            model = Ridge(alpha=alpha)
        
        model.fit(X_train, y_train)
        y_pred = model.predict(X_test)
        
        all_y_true.extend(y_test)
        all_y_pred.extend(y_pred)
    
    return compute_metrics(all_y_true, all_y_pred)


def pooled_recursive(df, alpha=0):
    """Pooled with recursive multi-step forecasting."""
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
        features = pd.concat([data[['month_idx']], emp_dummies], axis=1)
        return features
    
    X_train = build_features(train_df)
    X_val = build_features(val_df)
    y_train = train_df[TARGET_LABEL].values
    y_val = val_df[TARGET_LABEL].values
    
    X_train_val = pd.concat([X_train, X_val], ignore_index=True)
    y_train_val = np.concatenate([y_train, y_val])
    all_cols = X_train_val.columns
    
    if alpha == 0:
        model = LinearRegression()
    else:
        model = Ridge(alpha=alpha)
    
    model.fit(X_train_val, y_train_val)
    
    test_months_idx = sorted(test_df['month_idx'].unique())
    all_y_true = []
    all_y_pred = []
    
    for month_idx in test_months_idx:
        month_data = test_df[test_df['month_idx'] == month_idx].copy()
        features_list = []
        y_true_list = []
        
        for _, row in month_data.iterrows():
            emp_id = row['employee_id']
            feature_row = {'month_idx': month_idx}
            
            for col in all_cols:
                if col.startswith('emp_'):
                    emp_num = int(col.split('_')[1])
                    feature_row[col] = 1 if emp_id == emp_num else 0
            
            features_list.append(feature_row)
            y_true_list.append(row[TARGET_LABEL])
        
        if len(features_list) == 0:
            continue
        
        X_month = pd.DataFrame(features_list)
        for col in all_cols:
            if col not in X_month.columns:
                X_month[col] = 0
        X_month = X_month[all_cols]
        
        y_pred_month = model.predict(X_month)
        all_y_true.extend(y_true_list)
        all_y_pred.extend(y_pred_month)
    
    if len(all_y_true) == 0:
        return {k: np.nan for k in ['rmse', 'nrmse', 'mae', 'r2']}
    
    return compute_metrics(all_y_true, all_y_pred)


def unpooled_recursive(df, alpha=0):
    """Unpooled with recursive multi-step forecasting."""
    employees = df['employee_id'].unique()
    all_y_true, all_y_pred = [], []
    
    for emp in employees:
        sub = df[df['employee_id'] == emp].copy().sort_values('month')
        
        if len(sub) <= TEST_MONTHS + VAL_MONTHS + 5:
            continue
        
        for lag in [1, 2]:
            sub[f'y_lag{lag}'] = sub[TARGET_LABEL].shift(lag)
        sub = sub.dropna().copy()
        
        if len(sub) <= TEST_MONTHS + VAL_MONTHS + 3:
            continue
        
        sub['t'] = np.arange(len(sub))
        y = sub[TARGET_LABEL].values
        
        test_size = TEST_MONTHS
        val_size = VAL_MONTHS
        train_size = len(sub) - test_size - val_size
        
        if train_size < 3:
            continue
        
        X_train_val = sub.iloc[:train_size+val_size][['t', 'y_lag1', 'y_lag2']].values
        y_train_val = y[:train_size+val_size]
        y_test = y[train_size+val_size:]
        
        if alpha == 0:
            model = LinearRegression()
        else:
            model = Ridge(alpha=alpha)
        
        model.fit(X_train_val, y_train_val)
        
        y_pred_recursive = []
        last_known_values = list(y[:train_size+val_size][-2:])
        t_start = train_size + val_size
        
        for step in range(test_size):
            t_current = t_start + step
            features = [t_current, last_known_values[-1], last_known_values[-2]]
            X_pred = np.array(features).reshape(1, -1)
            
            y_next = model.predict(X_pred)[0]
            y_pred_recursive.append(y_next)
            last_known_values.append(y_next)
            last_known_values = last_known_values[-2:]
        
        all_y_true.extend(y_test)
        all_y_pred.extend(y_pred_recursive)
    
    if len(all_y_true) == 0:
        return {k: np.nan for k in ['rmse', 'nrmse', 'mae', 'r2']}
    
    return compute_metrics(all_y_true, all_y_pred)


def main():
    print("\n" + "="*80)
    print("REGULARIZATION EFFECT: POOLED vs UNPOOLED MODELS")
    print("="*80 + "\n")
    
    df = load_data()
    print(f"✓ Loaded {len(df)} records, {df['employee_id'].nunique()} employees\n")
    
    results = []
    alphas = [0, 0.1, 1.0, 10.0, 100.0]
    
    # Pooled models
    print("Evaluating POOLED models...")
    for alpha in alphas:
        reg_type = "Unregularized (OLS)" if alpha == 0 else f"Ridge (α={alpha})"
        print(f"  {reg_type}...")
        
        # Train/Val/Test
        metrics_tvt = pooled_train_val_test(df, alpha)
        results.append({
            'Model': 'Pooled FE',
            'Regularization': reg_type,
            'Alpha': alpha,
            'Method': 'Train/Val/Test',
            'RMSE': metrics_tvt['rmse'],
            'MAE': metrics_tvt['mae'],
            'NRMSE': metrics_tvt['nrmse'],
            'R²': metrics_tvt['r2']
        })
        
        # Recursive Multi-Step
        metrics_rec = pooled_recursive(df, alpha)
        results.append({
            'Model': 'Pooled FE',
            'Regularization': reg_type,
            'Alpha': alpha,
            'Method': 'Recursive Multi-Step',
            'RMSE': metrics_rec['rmse'],
            'MAE': metrics_rec['mae'],
            'NRMSE': metrics_rec['nrmse'],
            'R²': metrics_rec['r2']
        })
    
    # Unpooled models
    print("\nEvaluating UNPOOLED models...")
    for alpha in alphas:
        reg_type = "Unregularized (OLS)" if alpha == 0 else f"Ridge (α={alpha})"
        print(f"  {reg_type}...")
        
        # Train/Val/Test
        metrics_tvt = unpooled_train_val_test(df, alpha)
        results.append({
            'Model': 'Unpooled',
            'Regularization': reg_type,
            'Alpha': alpha,
            'Method': 'Train/Val/Test',
            'RMSE': metrics_tvt['rmse'],
            'MAE': metrics_tvt['mae'],
            'NRMSE': metrics_tvt['nrmse'],
            'R²': metrics_tvt['r2']
        })
        
        # Recursive Multi-Step
        metrics_rec = unpooled_recursive(df, alpha)
        results.append({
            'Model': 'Unpooled',
            'Regularization': reg_type,
            'Alpha': alpha,
            'Method': 'Recursive Multi-Step',
            'RMSE': metrics_rec['rmse'],
            'MAE': metrics_rec['mae'],
            'NRMSE': metrics_rec['nrmse'],
            'R²': metrics_rec['r2']
        })
    
    # Create DataFrame
    df_results = pd.DataFrame(results)
    
    # Save results
    output_path = 'outputs/regularization_comparison.csv'
    df_results.to_csv(output_path, index=False)
    print(f"\n✓ Results saved to: {output_path}")
    
    # Generate LaTeX table
    print("\n" + "="*80)
    print("LATEX TABLE: REGULARIZATION EFFECT")
    print("="*80 + "\n")
    
    latex = generate_latex_table(df_results)
    print(latex)
    
    # Save LaTeX
    latex_path = 'outputs/regularization_comparison_table.tex'
    with open(latex_path, 'w') as f:
        f.write(latex)
    print(f"\n✓ LaTeX table saved to: {latex_path}")
    
    # Print summary
    print("\n" + "="*80)
    print("SUMMARY: BEST CONFIGURATIONS")
    print("="*80)
    
    for model in ['Pooled FE', 'Unpooled']:
        for method in ['Train/Val/Test', 'Recursive Multi-Step']:
            subset = df_results[(df_results['Model'] == model) & (df_results['Method'] == method)]
            best_idx = subset['NRMSE'].idxmin()
            best = subset.loc[best_idx]
            print(f"\n{model} - {method}:")
            print(f"  Best: {best['Regularization']}")
            print(f"  NRMSE: {best['NRMSE']:.4f}, RMSE: {best['RMSE']:.1f}, MAE: {best['MAE']:.1f}")


def generate_latex_table(df):
    """Generate comprehensive LaTeX table."""
    
    latex = r"""
\begin{table}[htbp]
\centering
\caption{Regularization Effect on Pooled and Unpooled Models}
\label{tab:regularization_effect}
\resizebox{\textwidth}{!}{%
\begin{tabular}{@{} l l cccc cccc @{}}
\toprule
& & \multicolumn{4}{c}{\textbf{Train/Val/Test}} & \multicolumn{4}{c}{\textbf{Recursive Multi-Step}} \\
\cmidrule(lr){3-6} \cmidrule(lr){7-10}
\textbf{Model} & \textbf{Regularization} & \textbf{RMSE} & \textbf{MAE} & \textbf{NRMSE} & \textbf{R²} & \textbf{RMSE} & \textbf{MAE} & \textbf{NRMSE} & \textbf{R²} \\
\midrule
"""
    
    # Pooled models
    pooled_df = df[df['Model'] == 'Pooled FE']
    alphas = sorted(pooled_df['Alpha'].unique())
    
    for i, alpha in enumerate(alphas):
        reg_type = "OLS" if alpha == 0 else f"Ridge ($\\alpha={alpha}$)"
        
        tvt = pooled_df[(pooled_df['Alpha'] == alpha) & (pooled_df['Method'] == 'Train/Val/Test')].iloc[0]
        rec = pooled_df[(pooled_df['Alpha'] == alpha) & (pooled_df['Method'] == 'Recursive Multi-Step')].iloc[0]
        
        if i == 0:
            model_name = "\\textbf{Pooled FE}"
        else:
            model_name = ""
        
        latex += f"{model_name} & {reg_type} & "
        latex += f"{tvt['RMSE']:.1f} & {tvt['MAE']:.1f} & {tvt['NRMSE']:.4f} & {tvt['R²']:.4f} & "
        latex += f"{rec['RMSE']:.1f} & {rec['MAE']:.1f} & {rec['NRMSE']:.4f} & {rec['R²']:.4f} \\\\\n"
    
    latex += "\\midrule\n"
    
    # Unpooled models
    unpooled_df = df[df['Model'] == 'Unpooled']
    
    for i, alpha in enumerate(alphas):
        reg_type = "OLS" if alpha == 0 else f"Ridge ($\\alpha={alpha}$)"
        
        tvt = unpooled_df[(unpooled_df['Alpha'] == alpha) & (unpooled_df['Method'] == 'Train/Val/Test')].iloc[0]
        rec = unpooled_df[(unpooled_df['Alpha'] == alpha) & (unpooled_df['Method'] == 'Recursive Multi-Step')].iloc[0]
        
        if i == 0:
            model_name = "\\textbf{Unpooled}"
        else:
            model_name = ""
        
        latex += f"{model_name} & {reg_type} & "
        latex += f"{tvt['RMSE']:.1f} & {tvt['MAE']:.1f} & {tvt['NRMSE']:.4f} & {tvt['R²']:.4f} & "
        latex += f"{rec['RMSE']:.1f} & {rec['MAE']:.1f} & {rec['NRMSE']:.4f} & {rec['R²']:.4f} \\\\\n"
    
    latex += r"""\bottomrule
\end{tabular}%
}
\end{table}

\vspace{0.3cm}
\noindent \textbf{Note:} OLS = Ordinary Least Squares (unregularized). Ridge regularization with different $\alpha$ values shows the trade-off between bias and variance.
"""
    
    return latex


if __name__ == '__main__':
    main()
