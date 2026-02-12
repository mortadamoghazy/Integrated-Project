"""
Plot Forecasting Model Evaluation Results
==========================================
Generate publication-quality plots from train/test evaluation results.
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path

# Set style
plt.style.use('seaborn-v0_8-darkgrid')
sns.set_palette("husl")

# Load results
results_df = pd.read_csv('outputs/train_test_only_evaluation.csv')

# Create output directory for plots
output_dir = Path('outputs/plots')
output_dir.mkdir(exist_ok=True)

print("=" * 80)
print("GENERATING FORECASTING RESULT PLOTS")
print("=" * 80)

# Define model order and colors
model_order = ['Ridge', 'Lasso', 'AR(2)', 'AR(1)', 'Unpooled', 'Pooled FE']
colors = {
    'Ridge': '#2E86AB',
    'Lasso': '#A23B72', 
    'AR(2)': '#F18F01',
    'AR(1)': '#C73E1D',
    'Unpooled': '#6A994E',
    'Pooled FE': '#BC4749'
}

# ===========================
# Plot 1: NRMSE Comparison (Main Result)
# ===========================
print("\n1. Generating NRMSE comparison plot...")

fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))

# Train/Test
train_test = results_df[results_df['Method'] == 'Train/Test'].copy()
train_test = train_test.set_index('Model').reindex(model_order)

bars1 = ax1.barh(range(len(train_test)), train_test['nrmse'], 
                  color=[colors[m] for m in train_test.index], alpha=0.8)
ax1.set_yticks(range(len(train_test)))
ax1.set_yticklabels(train_test.index)
ax1.set_xlabel('NRMSE', fontsize=12, fontweight='bold')
ax1.set_title('Train/Test Split', fontsize=14, fontweight='bold')
ax1.grid(True, alpha=0.3, axis='x')
ax1.invert_yaxis()

# Add value labels
for i, (idx, row) in enumerate(train_test.iterrows()):
    ax1.text(row['nrmse'] + 0.002, i, f"{row['nrmse']:.4f}", 
             va='center', fontsize=10, fontweight='bold')

# Add rank badges
for i, (idx, row) in enumerate(train_test.iterrows()):
    if i == 0:  # Best model
        ax1.text(0.001, i, '★', va='center', fontsize=16, color='gold')

# Recursive Multi-Step
recursive = results_df[results_df['Method'] == 'Recursive Multi-Step'].copy()
recursive = recursive.set_index('Model').reindex(model_order)

bars2 = ax2.barh(range(len(recursive)), recursive['nrmse'],
                  color=[colors[m] for m in recursive.index], alpha=0.8)
ax2.set_yticks(range(len(recursive)))
ax2.set_yticklabels(recursive.index)
ax2.set_xlabel('NRMSE', fontsize=12, fontweight='bold')
ax2.set_title('Recursive Multi-Step', fontsize=14, fontweight='bold')
ax2.grid(True, alpha=0.3, axis='x')
ax2.invert_yaxis()

# Add value labels
for i, (idx, row) in enumerate(recursive.iterrows()):
    ax2.text(row['nrmse'] + 0.002, i, f"{row['nrmse']:.4f}",
             va='center', fontsize=10, fontweight='bold')

# Add rank badges
for i, (idx, row) in enumerate(recursive.iterrows()):
    if i == 0:  # Best model
        ax2.text(0.001, i, '★', va='center', fontsize=16, color='gold')

plt.suptitle('Forecasting Model Performance Comparison (NRMSE)', 
             fontsize=16, fontweight='bold', y=1.02)
plt.tight_layout()
plt.savefig(output_dir / 'nrmse_comparison.png', dpi=300, bbox_inches='tight')
print(f"   ✓ Saved: {output_dir / 'nrmse_comparison.png'}")
plt.close()


# ===========================
# Plot 2: All Metrics Comparison
# ===========================
print("\n2. Generating all metrics comparison plot...")

fig, axes = plt.subplots(2, 3, figsize=(18, 10))
metrics = ['rmse', 'mae', 'nrmse']
methods = ['Train/Test', 'Recursive Multi-Step']

for method_idx, method in enumerate(methods):
    method_data = results_df[results_df['Method'] == method].copy()
    method_data = method_data.set_index('Model').reindex(model_order)
    
    for metric_idx, metric in enumerate(metrics):
        ax = axes[method_idx, metric_idx]
        
        bars = ax.barh(range(len(method_data)), method_data[metric],
                       color=[colors[m] for m in method_data.index], alpha=0.8)
        ax.set_yticks(range(len(method_data)))
        ax.set_yticklabels(method_data.index if metric_idx == 0 else [])
        ax.set_xlabel(metric.upper(), fontsize=11, fontweight='bold')
        ax.set_title(f'{method} - {metric.upper()}', fontsize=12, fontweight='bold')
        ax.grid(True, alpha=0.3, axis='x')
        ax.invert_yaxis()
        
        # Add value labels
        for i, (idx, row) in enumerate(method_data.iterrows()):
            if metric == 'nrmse':
                label = f"{row[metric]:.4f}"
            else:
                label = f"{row[metric]:.1f}"
            ax.text(row[metric] * 1.02, i, label, va='center', fontsize=9)

plt.suptitle('Comprehensive Forecasting Model Evaluation', 
             fontsize=16, fontweight='bold', y=0.995)
plt.tight_layout()
plt.savefig(output_dir / 'all_metrics_comparison.png', dpi=300, bbox_inches='tight')
print(f"   ✓ Saved: {output_dir / 'all_metrics_comparison.png'}")
plt.close()


# ===========================
# Plot 3: Error Degradation (Train/Test → Recursive)
# ===========================
print("\n3. Generating error degradation analysis...")

train_test = results_df[results_df['Method'] == 'Train/Test'].set_index('Model')
recursive = results_df[results_df['Method'] == 'Recursive Multi-Step'].set_index('Model')

degradation = pd.DataFrame({
    'Model': model_order,
    'Train/Test': [train_test.loc[m, 'nrmse'] for m in model_order],
    'Recursive': [recursive.loc[m, 'nrmse'] for m in model_order],
})

degradation['Degradation'] = ((degradation['Recursive'] - degradation['Train/Test']) / 
                               degradation['Train/Test'] * 100)

fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))

# Connected line plot showing degradation
for i, model in enumerate(model_order):
    row = degradation[degradation['Model'] == model].iloc[0]
    ax1.plot([0, 1], [row['Train/Test'], row['Recursive']], 
             marker='o', linewidth=2, markersize=10, 
             color=colors[model], label=model, alpha=0.8)

ax1.set_xticks([0, 1])
ax1.set_xticklabels(['Train/Test', 'Recursive Multi-Step'], fontsize=11)
ax1.set_ylabel('NRMSE', fontsize=12, fontweight='bold')
ax1.set_title('Error Propagation Across Methods', fontsize=14, fontweight='bold')
ax1.legend(loc='upper left', fontsize=10)
ax1.grid(True, alpha=0.3)

# Degradation percentage
bars = ax2.barh(range(len(degradation)), degradation['Degradation'],
                color=[colors[m] for m in degradation['Model']], alpha=0.8)
ax2.set_yticks(range(len(degradation)))
ax2.set_yticklabels(degradation['Model'])
ax2.set_xlabel('Error Increase (%)', fontsize=12, fontweight='bold')
ax2.set_title('Train/Test → Recursive Degradation', fontsize=14, fontweight='bold')
ax2.grid(True, alpha=0.3, axis='x')
ax2.invert_yaxis()
ax2.axvline(0, color='black', linewidth=0.8)

# Add value labels
for i, (idx, row) in enumerate(degradation.iterrows()):
    label = f"{row['Degradation']:.1f}%"
    x_pos = row['Degradation'] + (2 if row['Degradation'] > 0 else -2)
    ha = 'left' if row['Degradation'] > 0 else 'right'
    ax2.text(x_pos, i, label, va='center', ha=ha, fontsize=10, fontweight='bold')

plt.tight_layout()
plt.savefig(output_dir / 'error_degradation.png', dpi=300, bbox_inches='tight')
print(f"   ✓ Saved: {output_dir / 'error_degradation.png'}")
plt.close()


# ===========================
# Plot 4: Model Rankings
# ===========================
print("\n4. Generating model rankings plot...")

# Create ranking matrix
ranking_data = []
for method in ['Train/Test', 'Recursive Multi-Step']:
    method_results = results_df[results_df['Method'] == method].copy()
    method_results = method_results.sort_values('nrmse')
    for rank, (_, row) in enumerate(method_results.iterrows(), 1):
        ranking_data.append({
            'Method': method,
            'Model': row['Model'],
            'Rank': rank
        })

ranking_df = pd.DataFrame(ranking_data)
ranking_pivot = ranking_df.pivot(index='Model', columns='Method', values='Rank')
ranking_pivot = ranking_pivot.reindex(model_order)

fig, ax = plt.subplots(figsize=(10, 6))

x = np.arange(len(model_order))
width = 0.35

bars1 = ax.bar(x - width/2, ranking_pivot['Train/Test'], width, 
               label='Train/Test', color='#2E86AB', alpha=0.8)
bars2 = ax.bar(x + width/2, ranking_pivot['Recursive Multi-Step'], width,
               label='Recursive Multi-Step', color='#A23B72', alpha=0.8)

ax.set_xlabel('Model', fontsize=12, fontweight='bold')
ax.set_ylabel('Rank (1 = Best)', fontsize=12, fontweight='bold')
ax.set_title('Model Rankings Across Evaluation Methods', fontsize=14, fontweight='bold')
ax.set_xticks(x)
ax.set_xticklabels(model_order, rotation=45, ha='right')
ax.legend(fontsize=10)
ax.set_ylim(0, 7)
ax.invert_yaxis()
ax.grid(True, alpha=0.3, axis='y')

# Add value labels
for bars in [bars1, bars2]:
    for bar in bars:
        height = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2., height,
                f'{int(height)}', ha='center', va='bottom', fontsize=9, fontweight='bold')

plt.tight_layout()
plt.savefig(output_dir / 'model_rankings.png', dpi=300, bbox_inches='tight')
print(f"   ✓ Saved: {output_dir / 'model_rankings.png'}")
plt.close()


# ===========================
# Plot 5: Ridge Dominance (Simplified)
# ===========================
print("\n5. Generating Ridge dominance plot...")

fig, ax = plt.subplots(figsize=(10, 6))

methods = ['Train/Test', 'Recursive Multi-Step']
ridge_data = results_df[results_df['Model'] == 'Ridge'].set_index('Method')

# Get all models' NRMSE for comparison
all_nrmse = []
for method in methods:
    method_data = results_df[results_df['Method'] == method]
    all_nrmse.append(method_data['nrmse'].values)

# Box plot for context
positions = [0, 1]
bp = ax.boxplot(all_nrmse, positions=positions, widths=0.3,
                patch_artist=True, showfliers=False,
                boxprops=dict(facecolor='lightgray', alpha=0.5),
                medianprops=dict(color='red', linewidth=2))

# Ridge performance line
ridge_values = [ridge_data.loc[m, 'nrmse'] for m in methods]
ax.plot(positions, ridge_values, marker='o', linewidth=3, markersize=12,
        color='#2E86AB', label='Ridge', zorder=10)

# Add value labels
for i, (pos, val) in enumerate(zip(positions, ridge_values)):
    ax.text(pos, val + 0.005, f'{val:.4f}', ha='center', va='bottom',
            fontsize=11, fontweight='bold', color='#2E86AB')

ax.set_xticks(positions)
ax.set_xticklabels(methods, fontsize=12)
ax.set_ylabel('NRMSE', fontsize=12, fontweight='bold')
ax.set_title('Ridge Regression: Consistent Best Performance', fontsize=14, fontweight='bold')
ax.legend(fontsize=11, loc='upper left')
ax.grid(True, alpha=0.3, axis='y')

# Add annotation
ax.text(0.5, 0.95, 'Ridge outperforms all competitors\nacross both evaluation methods',
        transform=ax.transAxes, ha='center', va='top',
        fontsize=11, style='italic',
        bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.3))

plt.tight_layout()
plt.savefig(output_dir / 'ridge_dominance.png', dpi=300, bbox_inches='tight')
print(f"   ✓ Saved: {output_dir / 'ridge_dominance.png'}")
plt.close()


# ===========================
# Summary Statistics
# ===========================
print("\n" + "=" * 80)
print("SUMMARY STATISTICS")
print("=" * 80)

for method in ['Train/Test', 'Recursive Multi-Step']:
    print(f"\n{method}:")
    method_data = results_df[results_df['Method'] == method].sort_values('nrmse')
    
    best = method_data.iloc[0]
    worst = method_data.iloc[-1]
    
    print(f"  Best:  {best['Model']:12s} - NRMSE = {best['nrmse']:.4f}")
    print(f"  Worst: {worst['Model']:12s} - NRMSE = {worst['nrmse']:.4f}")
    print(f"  Range: {worst['nrmse'] - best['nrmse']:.4f}")
    print(f"  Ridge advantage over 2nd: {(method_data.iloc[1]['nrmse'] - best['nrmse']) / best['nrmse'] * 100:.1f}%")

print("\n" + "=" * 80)
print("PLOT GENERATION COMPLETE")
print("=" * 80)
print(f"\nAll plots saved to: {output_dir.absolute()}")
print("\nGenerated plots:")
print("  1. nrmse_comparison.png - Main NRMSE comparison")
print("  2. all_metrics_comparison.png - All metrics (RMSE, MAE, NRMSE)")
print("  3. error_degradation.png - Error propagation analysis")
print("  4. model_rankings.png - Ranking comparison")
print("  5. ridge_dominance.png - Ridge performance highlight")
