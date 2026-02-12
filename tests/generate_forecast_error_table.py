"""
Generate LaTeX Table for Recursive Forecast Error Propagation
==============================================================
Shows how forecast uncertainty grows at each step using Ridge regression.
"""

import pandas as pd
import numpy as np
from pathlib import Path

# Project root
project_root = Path(__file__).parent.parent
output_dir = project_root / "outputs"

print("=" * 80)
print("RECURSIVE FORECAST ERROR PROPAGATION TABLE")
print("=" * 80)

# Load train/test evaluation results to get Ridge RMSE
results_file = output_dir / "train_test_only_evaluation.csv"

if not results_file.exists():
    print(f"\n❌ Error: {results_file} not found")
    print("   Please run train_test_only_evaluation.py first")
    exit(1)

results_df = pd.read_csv(results_file)
print(f"\n✓ Loaded forecasting evaluation results")

# Get Ridge Train/Test RMSE
ridge_row = results_df[(results_df['Model'] == 'Ridge') & (results_df['Method'] == 'Train/Test')]
if ridge_row.empty:
    print("\n❌ Error: Ridge Train/Test results not found")
    exit(1)

rmse = ridge_row['rmse'].values[0]  # lowercase column name
print(f"✓ Ridge Train/Test RMSE: €{rmse:.2f}")

# Calculate recursive error propagation for 6 steps
forecast_steps = 6
print(f"\n1. Calculating error propagation for {forecast_steps} forecast steps...")

error_data = []
for t in range(1, forecast_steps + 1):
    error_margin = rmse * np.sqrt(t) * 1.96
    ci_width = 2 * error_margin
    percentage_increase = ((error_margin / (rmse * 1.96)) - 1) * 100
    
    error_data.append({
        'Step': t,
        'Time_Horizon': f'{t} month(s)',
        'Error_Margin': error_margin,
        'CI_Width': ci_width,
        'Percentage_Increase': percentage_increase
    })
    
    print(f"   Step {t}: ±€{error_margin:,.2f} (width: €{ci_width:,.2f}, +{percentage_increase:.1f}%)")

error_df = pd.DataFrame(error_data)

# Generate LaTeX table
print("\n2. Generating LaTeX table...")

latex_table = r"""\begin{table}[htbp]
\centering
\caption{Recursive Forecast Error Propagation (Ridge Regression)}
\label{tab:forecast_error_propagation}
\begin{tabular}{ccccc}
\toprule
\textbf{Forecast} & \textbf{Time} & \textbf{Error Margin} & \textbf{CI Width} & \textbf{Growth} \\
\textbf{Step} & \textbf{Horizon} & \textbf{(€)} & \textbf{(€)} & \textbf{(\%)} \\
\midrule
"""

for idx, row in error_df.iterrows():
    step = int(row['Step'])
    time_horizon = row['Time_Horizon']
    error_margin = row['Error_Margin']
    ci_width = row['CI_Width']
    percentage = row['Percentage_Increase']
    
    # Highlight first and last steps
    if step == 1:
        latex_table += f"    {step} & {time_horizon} & ±{error_margin:,.2f} & {ci_width:,.2f} & — \\\\\n"
    else:
        latex_table += f"    {step} & {time_horizon} & ±{error_margin:,.2f} & {ci_width:,.2f} & +{percentage:.1f}\\% \\\\\n"

latex_table += r"""\bottomrule
\end{tabular}
\begin{tablenotes}
    \small
    \item \textbf{Error Margin}: Confidence interval margin calculated as $\pm\text{RMSE} \times \sqrt{t} \times 1.96$ where $t$ is the forecast step.
    \item \textbf{CI Width}: Total width of 95\% confidence interval (2 × Error Margin).
    \item \textbf{Growth}: Percentage increase in error margin compared to step 1 (baseline).
    \item Train/Test RMSE: €""" + f"{rmse:.2f}" + r""" (Ridge Regression with $\alpha=10.0$).
    \item Formula assumes independent forecast errors; actual recursive error may differ.
\end{tablenotes}
\end{table}
"""

# Save LaTeX table
table_file = output_dir / "forecast_error_propagation_table.tex"
with open(table_file, 'w') as f:
    f.write(latex_table)
print(f"   ✓ Saved: {table_file}")

# Save CSV for reference
csv_file = output_dir / "forecast_error_propagation.csv"
error_df.to_csv(csv_file, index=False)
print(f"   ✓ Saved: {csv_file}")

# Summary
print("\n" + "=" * 80)
print("SUMMARY")
print("=" * 80)
print(f"\nRidge Regression Recursive Forecast Error Propagation:")
print(f"   • Base RMSE (Train/Test): €{rmse:.2f}")
print(f"   • Forecast horizon: {forecast_steps} months")
print(f"   • Error at step 1: ±€{error_data[0]['Error_Margin']:,.2f}")
print(f"   • Error at step {forecast_steps}: ±€{error_data[-1]['Error_Margin']:,.2f}")
print(f"   • Total growth: +{error_data[-1]['Percentage_Increase']:.1f}%")
print(f"   • Formula: CI_t = ±RMSE × √t × 1.96")

print("\n" + "=" * 80)
print("TABLE GENERATION COMPLETE")
print("=" * 80)
print(f"\nOutput files:")
print(f"   1. {table_file.name} - LaTeX table for thesis")
print(f"   2. {csv_file.name} - Raw data (CSV)")
