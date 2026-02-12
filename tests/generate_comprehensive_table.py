"""
Generate comprehensive LaTeX table with all models and evaluation methods
"""

import pandas as pd
import numpy as np

# Load results
df = pd.read_csv('outputs/forecasting_simplified_comparison.csv')

# Define model order
model_order = ['AR(1)', 'AR(2)', 'AR(3)', 'AR(4)', 'Ridge', 'Lasso', 'Pooled FE', 'Unpooled', 'XGBoost']
method_order = ['Train/Val/Test', 'Recursive Multi-Step', 'Time Series CV', 'Nested CV']

# Create pivot table structure
results_dict = {}
for model in model_order:
    results_dict[model] = {}
    for method in method_order:
        row = df[(df['Model'] == model) & (df['Method'] == method)]
        if len(row) > 0:
            results_dict[model][method] = {
                'rmse': row['rmse'].values[0],
                'mae': row['mae'].values[0],
                'nrmse': row['nrmse'].values[0]
            }
        else:
            results_dict[model][method] = None

# Find best NRMSE for each method
best_by_method = {}
for method in method_order:
    method_data = df[df['Method'] == method]
    if len(method_data) > 0:
        best_idx = method_data['nrmse'].idxmin()
        best_by_method[method] = method_data.loc[best_idx, 'Model']

print("Best models by method:")
for method, model in best_by_method.items():
    nrmse = df[(df['Model'] == model) & (df['Method'] == method)]['nrmse'].values[0]
    print(f"  {method}: {model} (NRMSE={nrmse:.4f})")

# Generate LaTeX table
latex = r"""
\begin{table}[htbp]
\centering
\caption{Comprehensive Forecasting Model Evaluation: All Models and Methods}
\label{tab:forecasting_comprehensive}
\resizebox{\textwidth}{!}{%
\begin{tabular}{@{} l cccc cccc cccc @{}}
\toprule
& \multicolumn{4}{c}{\textbf{Train/Val/Test}} & \multicolumn{4}{c}{\textbf{Recursive Multi-Step}} & \multicolumn{4}{c}{\textbf{Time Series CV}} \\
\cmidrule(lr){2-5} \cmidrule(lr){6-9} \cmidrule(lr){10-13}
\textbf{Model} & \textbf{RMSE} & \textbf{MAE} & \textbf{NRMSE} & \textbf{Rank} & \textbf{RMSE} & \textbf{MAE} & \textbf{NRMSE} & \textbf{Rank} & \textbf{RMSE} & \textbf{MAE} & \textbf{NRMSE} & \textbf{Rank} \\
\midrule
"""

# Add rows for each model
for model in model_order:
    # Skip Nested CV as it only has 3 models
    row_data = []
    
    for method in ['Train/Val/Test', 'Recursive Multi-Step', 'Time Series CV']:
        if results_dict[model][method] is not None:
            data = results_dict[model][method]
            rmse = data['rmse']
            mae = data['mae']
            nrmse = data['nrmse']
            
            # Calculate rank for this method
            method_df = df[df['Method'] == method].sort_values('nrmse')
            rank = method_df[method_df['Model'] == model].index[0] - method_df.index[0] + 1
            
            # Check if best
            is_best = (model == best_by_method[method])
            
            if is_best:
                row_data.append(f"\\textbf{{{rmse:.1f}}} & \\textbf{{{mae:.1f}}} & \\textbf{{{nrmse:.4f}}} & \\textbf{{{rank}}}")
            else:
                row_data.append(f"{rmse:.1f} & {mae:.1f} & {nrmse:.4f} & {rank}")
        else:
            row_data.append("-- & -- & -- & --")
    
    # Check if this model is best in any method
    is_best_somewhere = model in best_by_method.values()
    model_name = f"\\textbf{{{model}}}" if is_best_somewhere else model
    
    latex += f"{model_name} & " + " & ".join(row_data) + " \\\\\n"

latex += r"""\bottomrule
\end{tabular}%
}
\end{table}

\vspace{0.5cm}
\noindent \textbf{Note:} Rank indicates relative performance within each evaluation method (1=best). Bold values indicate the best model for each method. Nested CV results available only for Ridge, Lasso, and XGBoost (not shown).
"""

print("\n" + "="*80)
print("LATEX TABLE")
print("="*80)
print(latex)

# Save to file
with open('outputs/forecasting_comparison_table.tex', 'w') as f:
    f.write(latex)

print("\n✓ Table saved to: outputs/forecasting_comparison_table.tex")

# Generate simpler version (just NRMSE)
latex_simple = r"""
\begin{table}[htbp]
\centering
\caption{Forecasting Model Performance: NRMSE Comparison Across Evaluation Methods}
\label{tab:forecasting_nrmse}
\begin{tabular}{@{} l ccc @{}}
\toprule
\textbf{Model} & \textbf{Train/Val/Test} & \textbf{Recursive Multi-Step} & \textbf{Time Series CV} \\
\midrule
"""

for model in model_order:
    row_data = []
    for method in ['Train/Val/Test', 'Recursive Multi-Step', 'Time Series CV']:
        if results_dict[model][method] is not None:
            nrmse = results_dict[model][method]['nrmse']
            is_best = (model == best_by_method[method])
            if is_best:
                row_data.append(f"\\textbf{{{nrmse:.4f}}}")
            else:
                row_data.append(f"{nrmse:.4f}")
        else:
            row_data.append("--")
    
    is_best_somewhere = model in best_by_method.values()
    model_name = f"\\textbf{{{model}}}" if is_best_somewhere else model
    latex_simple += f"{model_name} & " + " & ".join(row_data) + " \\\\\n"

latex_simple += r"""\bottomrule
\end{tabular}
\end{table}

\vspace{0.3cm}
\noindent \textbf{Evaluation Methods:}
\begin{itemize}
    \item \textbf{Train/Val/Test}: Traditional split with 6-month test period (one-step-ahead forecasts)
    \item \textbf{Recursive Multi-Step}: Realistic forecasting where predictions feed back as inputs (error accumulation)
    \item \textbf{Time Series CV}: Cross-validation with expanding window (3 splits, 6-month test size)
\end{itemize}
"""

print("\n" + "="*80)
print("SIMPLIFIED TABLE (NRMSE ONLY)")
print("="*80)
print(latex_simple)

# Save simplified version
with open('outputs/forecasting_comparison_table_simple.tex', 'w') as f:
    f.write(latex_simple)

print("\n✓ Simplified table saved to: outputs/forecasting_comparison_table_simple.tex")
