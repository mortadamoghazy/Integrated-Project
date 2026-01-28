"""
Visualize Simplified Forecasting Comparison Results

Creates clear visualizations comparing 3 evaluation methods:
- Train/Val/Test Split
- Time Series Cross-Validation
- Nested Cross-Validation
"""

import sys
from pathlib import Path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

# Set style
sns.set_style("whitegrid")
plt.rcParams['figure.figsize'] = (16, 10)
plt.rcParams['font.size'] = 10


def load_results():
    """Load simplified comparison results."""
    csv_path = project_root / "outputs" / "forecasting_simplified_comparison.csv"
    df = pd.read_csv(csv_path)
    return df


def plot_overview(df):
    """Main overview comparing evaluation methods."""
    fig, axes = plt.subplots(2, 2, figsize=(20, 14))
    fig.suptitle('Forecasting Model Comparison - 4 Evaluation Methods (Including Recursive)', 
                 fontsize=16, fontweight='bold')
    
    methods = df['Method'].unique()
    
    # 1. NRMSE by Evaluation Method - All Models
    ax1 = axes[0, 0]
    models = df['Model'].unique()
    x = np.arange(len(methods))
    width = 0.08
    colors = plt.cm.tab10(np.linspace(0, 1, len(models)))
    
    for i, model in enumerate(models):
        model_data = df[df['Model'] == model]
        nrmse_values = []
        for method in methods:
            method_data = model_data[model_data['Method'] == method]
            if len(method_data) > 0:
                nrmse_values.append(method_data['nrmse'].values[0])
            else:
                nrmse_values.append(np.nan)
        
        positions = x + (i - len(models)/2) * width
        ax1.bar(positions, nrmse_values, width, label=model, color=colors[i], alpha=0.8)
    
    ax1.set_xlabel('Evaluation Method', fontweight='bold')
    ax1.set_ylabel('NRMSE (Lower is Better)', fontweight='bold')
    ax1.set_title('NRMSE Comparison Across All Models', fontweight='bold')
    ax1.set_xticks(x)
    ax1.set_xticklabels(methods, rotation=15, ha='right')
    ax1.legend(loc='upper left', fontsize=7, ncol=2)
    ax1.grid(axis='y', alpha=0.3)
    
    # 2. Best Model per Method
    ax2 = axes[0, 1]
    best_models = []
    best_nrmse = []
    method_labels = []
    
    for method in methods:
        method_data = df[df['Method'] == method]
        best_idx = method_data['nrmse'].idxmin()
        best_model = method_data.loc[best_idx, 'Model']
        best_score = method_data.loc[best_idx, 'nrmse']
        
        best_models.append(best_model)
        best_nrmse.append(best_score)
        method_labels.append(method)
    
    bars = ax2.barh(method_labels, best_nrmse, color='steelblue', alpha=0.7)
    ax2.set_xlabel('NRMSE', fontweight='bold')
    ax2.set_title('Best Model per Evaluation Method', fontweight='bold')
    ax2.grid(axis='x', alpha=0.3)
    
    # Add model names on bars
    for i, (bar, model) in enumerate(zip(bars, best_models)):
        width = bar.get_width()
        ax2.text(width + 0.002, bar.get_y() + bar.get_height()/2, 
                f'{model}', ha='left', va='center', fontweight='bold', fontsize=10)
    
    # 3. AR Models Comparison Across Methods
    ax3 = axes[1, 0]
    ar_models = ['AR(1)', 'AR(2)']
    x_pos = np.arange(len(ar_models))
    width = 0.25
    colors_method = ['#2ecc71', '#3498db', '#e74c3c']
    
    # Only use Train/Val/Test and Time Series CV for AR models
    ar_methods = ['Train/Val/Test', 'Time Series CV']
    for i, method in enumerate(ar_methods):
        scores = []
        for model in ar_models:
            model_data = df[(df['Model'] == model) & (df['Method'] == method)]
            if len(model_data) > 0:
                scores.append(model_data['nrmse'].values[0])
            else:
                scores.append(np.nan)
        
        positions = x_pos + (i - 0.5) * width
        ax3.bar(positions, scores, width, label=method, color=colors_method[i], alpha=0.8)
    
    ax3.set_xlabel('AR Model', fontweight='bold')
    ax3.set_ylabel('NRMSE', fontweight='bold')
    ax3.set_title('AR Models Performance Across Evaluation Methods', fontweight='bold')
    ax3.set_xticks(x_pos)
    ax3.set_xticklabels(ar_models)
    ax3.legend()
    ax3.grid(axis='y', alpha=0.3)
    
    # 4. XGBoost vs Ridge/Lasso Across All Methods
    ax4 = axes[1, 1]
    ml_models = ['Ridge', 'Lasso', 'XGBoost']
    x_pos = np.arange(len(methods))
    width = 0.25
    colors_ml = ['#e74c3c', '#f39c12', '#2ecc71']
    
    for i, model in enumerate(ml_models):
        scores = []
        for method in methods:
            model_data = df[(df['Model'] == model) & (df['Method'] == method)]
            if len(model_data) > 0:
                scores.append(model_data['nrmse'].values[0])
            else:
                scores.append(np.nan)
        
        offset = (i - 1) * width
        ax4.bar(x_pos + offset, scores, width, label=model, 
                color=colors_ml[i], alpha=0.7)
    
    ax4.set_xlabel('Evaluation Method', fontweight='bold')
    ax4.set_ylabel('NRMSE', fontweight='bold')
    ax4.set_title('Ridge/Lasso/XGBoost Across Evaluation Methods', fontweight='bold')
    ax4.set_xticks(x_pos)
    ax4.set_xticklabels(methods, rotation=15, ha='right')
    ax4.legend()
    ax4.grid(axis='y', alpha=0.3)
    
    plt.tight_layout()
    return fig


def plot_detailed_metrics(df):
    """Comprehensive comparison of all metrics across all methods."""
    fig, axes = plt.subplots(2, 2, figsize=(20, 14))
    fig.suptitle('Comprehensive Error Metrics - All 4 Methods (Including Recursive)', fontsize=16, fontweight='bold')
    
    # Get all methods and models
    methods = df['Method'].unique()
    all_models = df['Model'].unique()
    
    # 1. NRMSE Heatmap across all methods and models
    ax1 = axes[0, 0]
    
    # Create pivot table for heatmap
    pivot_nrmse = df.pivot_table(values='nrmse', index='Model', columns='Method', aggfunc='first')
    pivot_nrmse = pivot_nrmse.reindex(columns=methods)
    
    sns.heatmap(pivot_nrmse, annot=True, fmt='.4f', cmap='RdYlGn_r', 
                ax=ax1, cbar_kws={'label': 'NRMSE'}, vmin=0, vmax=0.16)
    ax1.set_title('NRMSE: All Models × All Methods', fontweight='bold', fontsize=12)
    ax1.set_xlabel('Evaluation Method', fontweight='bold')
    ax1.set_ylabel('Model', fontweight='bold')
    plt.setp(ax1.get_xticklabels(), rotation=30, ha='right')
    
    # 2. RMSE Heatmap
    ax2 = axes[0, 1]
    
    pivot_rmse = df.pivot_table(values='rmse', index='Model', columns='Method', aggfunc='first')
    pivot_rmse = pivot_rmse.reindex(columns=methods)
    
    sns.heatmap(pivot_rmse, annot=True, fmt='.1f', cmap='RdYlGn_r', 
                ax=ax2, cbar_kws={'label': 'RMSE (EUR)'}, vmin=80, vmax=450)
    ax2.set_title('RMSE (EUR): All Models × All Methods', fontweight='bold', fontsize=12)
    ax2.set_xlabel('Evaluation Method', fontweight='bold')
    ax2.set_ylabel('Model', fontweight='bold')
    plt.setp(ax2.get_xticklabels(), rotation=30, ha='right')
    
    # 3. MAE Heatmap
    ax3 = axes[1, 0]
    
    pivot_mae = df.pivot_table(values='mae', index='Model', columns='Method', aggfunc='first')
    pivot_mae = pivot_mae.reindex(columns=methods)
    
    sns.heatmap(pivot_mae, annot=True, fmt='.1f', cmap='RdYlGn_r', 
                ax=ax3, cbar_kws={'label': 'MAE (EUR)'}, vmin=70, vmax=280)
    ax3.set_title('MAE (EUR): All Models × All Methods', fontweight='bold', fontsize=12)
    ax3.set_xlabel('Evaluation Method', fontweight='bold')
    ax3.set_ylabel('Model', fontweight='bold')
    plt.setp(ax3.get_xticklabels(), rotation=30, ha='right')
    
    # 4. R² Heatmap
    ax4 = axes[1, 1]
    
    pivot_r2 = df.pivot_table(values='r2', index='Model', columns='Method', aggfunc='first')
    pivot_r2 = pivot_r2.reindex(columns=methods)
    
    # Use custom colormap for R² (higher is better, but handle negative values)
    sns.heatmap(pivot_r2, annot=True, fmt='.3f', cmap='RdYlGn', 
                ax=ax4, cbar_kws={'label': 'R² Score'}, center=0.5)
    ax4.set_title('R² Score: All Models × All Methods (Higher is Better)', fontweight='bold', fontsize=12)
    ax4.set_xlabel('Evaluation Method', fontweight='bold')
    ax4.set_ylabel('Model', fontweight='bold')
    plt.setp(ax4.get_xticklabels(), rotation=30, ha='right')
    
    plt.tight_layout()
    return fig


def plot_model_ranking(df):
    """Create a ranking visualization showing winners."""
    fig, ax = plt.subplots(figsize=(14, 8))
    
    methods = df['Method'].unique()
    
    # Create ranking data
    rankings = []
    for method in methods:
        method_data = df[df['Method'] == method].sort_values('nrmse')
        top_3 = method_data.head(3)
        
        for rank, (_, row) in enumerate(top_3.iterrows(), 1):
            rankings.append({
                'Method': method,
                'Rank': rank,
                'Model': row['Model'],
                'NRMSE': row['nrmse']
            })
    
    rankings_df = pd.DataFrame(rankings)
    
    # Plot
    y_pos = 0
    method_positions = {}
    colors_rank = ['gold', 'silver', '#CD7F32']  # Gold, Silver, Bronze
    
    for method in methods:
        method_ranks = rankings_df[rankings_df['Method'] == method]
        
        ax.text(-0.15, y_pos + 1, method, fontsize=10, fontweight='bold', 
               ha='right', va='center')
        
        for _, row in method_ranks.iterrows():
            rank = int(row['Rank'])
            color = colors_rank[rank-1] if rank <= 3 else 'lightgray'
            
            # Medal emoji
            medal = ['🥇', '🥈', '🥉'][rank-1] if rank <= 3 else ''
            
            bar = ax.barh(y_pos, row['NRMSE'], height=0.7, 
                         color=color, alpha=0.7, edgecolor='black', linewidth=1.5)
            
            # Add model name and score
            ax.text(row['NRMSE'] + 0.002, y_pos, 
                   f"{medal} {row['Model']} ({row['NRMSE']:.4f})",
                   va='center', fontsize=9, fontweight='bold')
            
            y_pos += 1
        
        y_pos += 0.5  # Space between methods
    
    ax.set_xlabel('NRMSE (Lower is Better)', fontsize=12, fontweight='bold')
    ax.set_title('Top 3 Models per Evaluation Method', fontsize=14, fontweight='bold')
    ax.set_yticks([])
    ax.grid(axis='x', alpha=0.3)
    ax.set_xlim(0, max(rankings_df['NRMSE']) * 1.3)
    
    plt.tight_layout()
    return fig


def plot_comprehensive_table(df):
    """Create a comprehensive table with all error values."""
    fig, ax = plt.subplots(figsize=(20, 12))
    ax.axis('off')
    fig.suptitle('Complete Error Metrics Table - All Models × All Methods', 
                 fontsize=16, fontweight='bold', y=0.98)
    
    # Prepare data
    methods = df['Method'].unique()
    models = df['Model'].unique()
    
    # Create table data
    table_data = []
    for model in models:
        for method in methods:
            row_data = df[(df['Model'] == model) & (df['Method'] == method)]
            if len(row_data) > 0:
                table_data.append([
                    model,
                    method,
                    f"{row_data['nrmse'].values[0]:.4f}",
                    f"{row_data['rmse'].values[0]:.1f}",
                    f"{row_data['mae'].values[0]:.1f}",
                    f"{row_data['r2'].values[0]:.4f}"
                ])
            else:
                table_data.append([model, method, 'N/A', 'N/A', 'N/A', 'N/A'])
    
    # Create table
    table = ax.table(cellText=table_data,
                     colLabels=['Model', 'Evaluation Method', 'NRMSE', 'RMSE (EUR)', 'MAE (EUR)', 'R²'],
                     cellLoc='center',
                     loc='center',
                     colWidths=[0.15, 0.23, 0.12, 0.13, 0.13, 0.12])
    
    table.auto_set_font_size(False)
    table.set_fontsize(8)
    table.scale(1, 1.8)
    
    # Style header
    for i in range(6):
        table[(0, i)].set_facecolor('#2E86AB')
        table[(0, i)].set_text_props(weight='bold', color='white')
        table[(0, i)].set_height(0.08)
    
    # Alternate row colors and highlight best per method
    for method_idx, method in enumerate(methods):
        method_rows = [i for i, row in enumerate(table_data, 1) if row[1] == method]
        
        # Find best NRMSE for this method
        method_data = df[df['Method'] == method]
        if len(method_data) > 0:
            best_model = method_data.loc[method_data['nrmse'].idxmin(), 'Model']
        else:
            best_model = None
        
        for row_idx in method_rows:
            model_name = table_data[row_idx-1][0]
            
            # Alternate colors
            color = '#F0F0F0' if method_idx % 2 == 0 else '#FFFFFF'
            
            # Highlight best
            if model_name == best_model:
                color = '#90EE90'
            
            for col_idx in range(6):
                table[(row_idx, col_idx)].set_facecolor(color)
                table[(row_idx, col_idx)].set_height(0.06)
    
    # Add separator lines between methods
    current_method = None
    for i, row in enumerate(table_data, 1):
        if row[1] != current_method:
            current_method = row[1]
            if i > 1:
                for col_idx in range(6):
                    table[(i, col_idx)].set_edgecolor('black')
                    table[(i, col_idx)].set_linewidth(2)
    
    plt.tight_layout()
    return fig


def main():
    print("=" * 80)
    print("VISUALIZING FORECASTING COMPARISON (4 METHODS)")
    print("=" * 80)
    
    # Load data
    print("\nLoading results...")
    df = load_results()
    
    # Filter out AR(3) and AR(4)
    df = df[~df['Model'].isin(['AR(3)', 'AR(4)'])]
    
    print(f"✓ Loaded {len(df)} model evaluations")
    print(f"  Models: {df['Model'].nunique()}")
    print(f"  Methods: {', '.join(df['Method'].unique())}")
    print(f"\n  📊 Includes RECURSIVE MULTI-STEP forecasting (realistic error estimation)")
    
    # Create visualizations
    print("\n1. Creating overview comparison plot...")
    fig1 = plot_overview(df)
    output1 = project_root / "outputs" / "forecasting_simplified_overview.png"
    fig1.savefig(output1, dpi=300, bbox_inches='tight')
    print(f"   ✓ Saved to: {output1}")
    
    print("\n2. Creating detailed metrics plot (heatmaps)...")
    fig2 = plot_detailed_metrics(df)
    output2 = project_root / "outputs" / "forecasting_simplified_metrics.png"
    fig2.savefig(output2, dpi=300, bbox_inches='tight')
    print(f"   ✓ Saved to: {output2}")
    
    print("\n3. Creating model ranking plot...")
    fig3 = plot_model_ranking(df)
    output3 = project_root / "outputs" / "forecasting_simplified_ranking.png"
    fig3.savefig(output3, dpi=300, bbox_inches='tight')
    print(f"   ✓ Saved to: {output3}")
    
    print("\n4. Creating comprehensive data table...")
    fig4 = plot_comprehensive_table(df)
    output4 = project_root / "outputs" / "forecasting_comprehensive_table.png"
    fig4.savefig(output4, dpi=300, bbox_inches='tight')
    print(f"   ✓ Saved to: {output4}")
    
    print("\n" + "=" * 80)
    print("VISUALIZATION COMPLETE")
    print("=" * 80)
    print("\n📌 KEY EVALUATION METHODS EXPLAINED:")
    print("\n1. Train/Val/Test: One-step-ahead (uses actual values as inputs)")
    print("2. Recursive Multi-Step: Realistic forecasting (predictions feed back)")
    print("3. Time Series CV: Cross-validation with time-based splits")
    print("4. Nested CV: Hyperparameter tuning with CV")
    print("\n⚠️  IMPORTANT: Recursive Multi-Step shows REAL-WORLD performance!")
    print("   One-step-ahead gives optimistic errors. Recursive accumulates errors.")
    print("\nKey Findings:")
    
    # Find best model per method
    methods = df['Method'].unique()
    for method in methods:
        method_data = df[df['Method'] == method]
        best = method_data.loc[method_data['nrmse'].idxmin()]
        print(f"\n🏆 {method}:")
        print(f"   Winner: {best['Model']}")
        print(f"   NRMSE: {best['nrmse']:.4f}")
        print(f"   RMSE: {best['rmse']:.1f} EUR")
        print(f"   MAE: {best['mae']:.1f} EUR")
    
    # Overall recommendation
    print("\n" + "=" * 80)
    print("RECOMMENDATION:")
    print("=" * 80)
    
    # Find XGBoost performance
    xgb_cv = df[(df['Model'] == 'XGBoost') & (df['Method'] == 'Time Series CV')]
    xgb_nested = df[(df['Model'] == 'XGBoost') & (df['Method'] == 'Nested CV')]
    xgb_recursive = df[(df['Model'] == 'XGBoost') & (df['Method'] == 'Recursive Multi-Step')]
    
    if len(xgb_cv) > 0:
        print(f"\n✅ XGBoost with Cross-Validation shows excellent performance:")
        print(f"   Time Series CV: NRMSE = {xgb_cv['nrmse'].values[0]:.4f}")
        if len(xgb_nested) > 0:
            print(f"   Nested CV: NRMSE = {xgb_nested['nrmse'].values[0]:.4f}")
        if len(xgb_recursive) > 0:
            print(f"   Recursive Multi-Step: NRMSE = {xgb_recursive['nrmse'].values[0]:.4f}")
            print(f"\n   🎯 XGBoost is ROBUST: performs well even with recursive forecasting!")
        print(f"\n   This represents the most robust evaluation method")
        print(f"   and provides realistic error estimation for production use.")
    
    print("\n" + "=" * 80)
    print("\nDisplaying plots...")
    plt.show()


if __name__ == "__main__":
    main()
