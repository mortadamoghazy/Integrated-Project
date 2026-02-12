"""
Generate LaTeX Tables for Clustering Evaluation
================================================
Creates publication-ready LaTeX tables from clustering evaluation results.
Similar structure to forecasting LaTeX table generation.
"""

import pandas as pd
from pathlib import Path

# Project root
project_root = Path(__file__).parent.parent
output_dir = project_root / "outputs"

print("=" * 80)
print("CLUSTERING EVALUATION LATEX TABLE GENERATION")
print("=" * 80)

# Load results
results_file = output_dir / "clustering_evaluation.csv"
profiles_file = output_dir / "clustering_profiles_optimal.csv"

if not results_file.exists():
    print(f"\n❌ Error: {results_file} not found")
    print("   Please run evaluate_clustering.py first")
    exit(1)

results_df = pd.read_csv(results_file)
print(f"\n✓ Loaded clustering evaluation results (K = {results_df['K'].iloc[0]:.0f})")

if profiles_file.exists():
    profiles_df = pd.read_csv(profiles_file)
    optimal_k = len(profiles_df)
    print(f"✓ Loaded cluster profiles for optimal K={optimal_k}")
else:
    print(f"\n⚠ Warning: {profiles_file} not found")
    profiles_df = None
    optimal_k = None

# ============================================================================
# Table 1: Comprehensive Metrics for K=3
# ============================================================================

print("\n1. Generating clustering metrics table for K=3...")

latex_table1 = r"""\begin{table}[htbp]
\centering
\caption{K-Means Clustering Validation Metrics (K=3)}
\label{tab:clustering_metrics}
\begin{tabular}{cccccc}
\toprule
\textbf{K} & \textbf{Silhouette} & \textbf{Davies-} & \textbf{Calinski-} & \textbf{Inertia} & \textbf{Avg Cluster} \\
           & \textbf{Score}      & \textbf{Bouldin} & \textbf{Harabasz}  &                  & \textbf{Size}        \\
\midrule
"""

for idx, row in results_df.iterrows():
    k = int(row['K'])
    silhouette = row['Silhouette_Score']
    db = row['Davies_Bouldin_Index']
    ch = row['Calinski_Harabasz_Index']
    inertia = row['Inertia']
    avg_size = row['Avg_Cluster_Size']
    
    latex_table1 += f"    {k} & {silhouette:.4f} & {db:.4f} & {ch:.2f} & {inertia:.2f} & {avg_size:.1f} \\\\\n"

latex_table1 += r"""\bottomrule
\end{tabular}
\begin{tablenotes}
    \small
    \item \textbf{Silhouette Score}: Higher is better (range: -1 to 1). Measures cluster cohesion and separation.
    \item \textbf{Davies-Bouldin Index}: Lower is better. Measures cluster compactness and separation.
    \item \textbf{Calinski-Harabasz Index}: Higher is better. Ratio of between-cluster to within-cluster variance.
\end{tablenotes}
\end{table}
"""

# Save table 1
table1_file = output_dir / "clustering_metrics_table.tex"
with open(table1_file, 'w') as f:
    f.write(latex_table1)
print(f"   ✓ Saved: {table1_file}")

# ============================================================================
# Table 2: Detailed Cluster Profiles for Optimal K
# ============================================================================

if profiles_df is not None:
    print("\n2. Generating cluster profiles table...")
    
    latex_table2 = r"""\begin{table}[htbp]
\centering
\caption{Detailed Cluster Profiles (K=3)}
\label{tab:cluster_profiles}
\resizebox{\textwidth}{!}{%
\begin{tabular}{lcccccc}
\toprule
\textbf{Cluster Tier} & \textbf{Size} & \textbf{Avg Cost} & \textbf{Cost Range} & \textbf{Avg Salary} & \textbf{Avg Contrib.} & \textbf{Total/Month} \\
                      & (employees)   & (€/month)         & (€)                 & (€)                 & (€)                   & (€)                  \\
\midrule
"""
    
    for idx, row in profiles_df.iterrows():
        tier_name = row['Tier_Name']
        size = int(row['Size'])
        avg_cost = row['Avg_Cost']
        std_cost = row['Std_Cost']
        min_cost = row['Min_Cost']
        max_cost = row['Max_Cost']
        avg_salary = row['Avg_Salary']
        avg_contrib = row['Avg_Contributions']
        total_monthly = avg_cost * size
        
        cost_range = f"{min_cost:,.0f}--{max_cost:,.0f}"
        
        latex_table2 += f"    {tier_name} & {size} & {avg_cost:,.2f} $\\pm$ {std_cost:,.2f} & {cost_range} & {avg_salary:,.2f} & {avg_contrib:,.2f} & {total_monthly:,.2f} \\\\\n"
    
    # Add total row
    total_employees = profiles_df['Size'].sum()
    grand_total = (profiles_df['Avg_Cost'] * profiles_df['Size']).sum()
    
    latex_table2 += r"""\midrule
    \textbf{Total} & """ + f"\\textbf{{{int(total_employees)}}} & — & — & — & — & \\textbf{{{grand_total:,.2f}}} \\\\\n"
    
    latex_table2 += r"""\bottomrule
\end{tabular}%
}
\begin{tablenotes}
    \small
    \item Clusters sorted by average cost (ascending).
    \item \textbf{Avg Cost}: Mean ± standard deviation per employee per month.
    \item \textbf{Cost Range}: Minimum and maximum employee costs within cluster.
    \item \textbf{Total/Month}: Total monthly payroll cost for all employees in cluster.
\end{tablenotes}
\end{table}
"""
    
    # Save table 2
    table2_file = output_dir / "clustering_profiles_table.tex"
    with open(table2_file, 'w') as f:
        f.write(latex_table2)
    print(f"   ✓ Saved: {table2_file}")
else:
    table2_file = None

# ============================================================================
# Table 3: Simplified Summary Table
# ============================================================================

print("\n3. Generating simplified summary table...")

latex_table3 = r"""\begin{table}[htbp]
\centering
\caption{K-Means Clustering Summary (K=3)}
\label{tab:clustering_summary}
\begin{tabular}{cccc}
\toprule
\textbf{K} & \textbf{Silhouette} & \textbf{Davies-Bouldin} & \textbf{Calinski-Harabasz} \\
\midrule
"""

for idx, row in results_df.iterrows():
    k = int(row['K'])
    silhouette = row['Silhouette_Score']
    db = row['Davies_Bouldin_Index']
    ch = row['Calinski_Harabasz_Index']
    
    latex_table3 += f"    {k} & {silhouette:.4f} & {db:.4f} & {ch:.2f} \\\\\n"

latex_table3 += r"""\bottomrule
\end{tabular}
\end{table}
"""

# Save table 3
table3_file = output_dir / "clustering_comparison_table.tex"
with open(table3_file, 'w') as f:
    f.write(latex_table3)
print(f"   ✓ Saved: {table3_file}")

# ============================================================================
# Summary
# ============================================================================

print("\n" + "=" * 80)
print("LATEX TABLE GENERATION COMPLETE")
print("=" * 80)
print(f"\nGenerated tables:")
print(f"   1. {table1_file.name} - Clustering metrics for K=3")
if table2_file:
    print(f"   2. {table2_file.name} - Detailed cluster profiles for K=3")
print(f"   3. {table3_file.name} - Simplified clustering summary")

print(f"\nK=3 Validation Metrics:")
optimal_row = results_df.iloc[0]
print(f"   • Silhouette Score: {optimal_row['Silhouette_Score']:.4f}")
print(f"   • Davies-Bouldin Index: {optimal_row['Davies_Bouldin_Index']:.4f}")
print(f"   • Calinski-Harabasz Index: {optimal_row['Calinski_Harabasz_Index']:.2f}")
