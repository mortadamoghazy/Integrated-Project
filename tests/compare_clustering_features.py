"""
Compare clustering with 1 feature (mean only) vs 4 features.
Analyze feature importance and contribution to clustering.
"""

import sys
from pathlib import Path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pandas as pd
import numpy as np
from sklearn.cluster import KMeans
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import silhouette_score, davies_bouldin_score, calinski_harabasz_score
import matplotlib.pyplot as plt


def cluster_with_features(df, feature_cols):
    """Run clustering with specified features."""
    # Aggregate features per employee
    agg_dict = {'total_cost': ['mean', 'std', 'min', 'max']}
    features = df.groupby('employee_id').agg(agg_dict)
    features.columns = ['_'.join(col).strip() for col in features.columns.values]
    features = features.replace([np.inf, -np.inf], 0).fillna(0)
    
    # Select only specified features
    X = features[feature_cols]
    
    # Standardize
    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(X)
    
    # Cluster
    kmeans = KMeans(n_clusters=3, random_state=42, n_init=10)
    labels = kmeans.fit_predict(X_scaled)
    
    # Metrics
    silhouette = silhouette_score(X_scaled, labels)
    davies_bouldin = davies_bouldin_score(X_scaled, labels)
    calinski = calinski_harabasz_score(X_scaled, labels)
    
    # Variance explained
    total_var = np.var(X_scaled, axis=0).sum()
    within_var = sum([np.var(X_scaled[labels == i], axis=0).sum() * (labels == i).sum() 
                      for i in range(3)]) / len(X_scaled)
    variance_explained = (total_var - within_var) / total_var * 100
    
    return {
        'labels': labels,
        'silhouette': silhouette,
        'davies_bouldin': davies_bouldin,
        'calinski': calinski,
        'variance_explained': variance_explained,
        'scaler': scaler,
        'kmeans': kmeans,
        'X_scaled': X_scaled,
        'features': X
    }


def analyze_feature_importance(df):
    """Analyze which features contribute most to clustering."""
    print("=" * 80)
    print("FEATURE IMPORTANCE ANALYSIS")
    print("=" * 80)
    
    # Get all 4 features
    agg_dict = {'total_cost': ['mean', 'std', 'min', 'max']}
    features = df.groupby('employee_id').agg(agg_dict)
    features.columns = ['_'.join(col).strip() for col in features.columns.values]
    features = features.replace([np.inf, -np.inf], 0).fillna(0)
    
    # Standardize
    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(features)
    
    # Cluster
    kmeans = KMeans(n_clusters=3, random_state=42, n_init=10)
    labels = kmeans.fit_predict(X_scaled)
    
    print("\n1. VARIANCE BY FEATURE")
    print("-" * 80)
    print("How much each feature varies across all employees:")
    for i, col in enumerate(features.columns):
        variance = np.var(features[col])
        std_variance = np.var(X_scaled[:, i])
        print(f"  {col:20s}: Raw Var = {variance:10,.2f}, Scaled Var = {std_variance:.4f}")
    
    print("\n2. BETWEEN-CLUSTER VARIANCE BY FEATURE")
    print("-" * 80)
    print("How much each feature differs BETWEEN clusters (higher = more discriminative):")
    
    importance_scores = {}
    for i, col in enumerate(features.columns):
        # Calculate variance between cluster centers for this feature
        cluster_means = [X_scaled[labels == c, i].mean() for c in range(3)]
        between_var = np.var(cluster_means)
        
        # Calculate average within-cluster variance for this feature
        within_var = np.mean([np.var(X_scaled[labels == c, i]) for c in range(3)])
        
        # F-ratio: between variance / within variance (higher = more important)
        f_ratio = between_var / within_var if within_var > 0 else 0
        
        importance_scores[col] = f_ratio
        
        print(f"  {col:20s}: Between = {between_var:.4f}, Within = {within_var:.4f}, F-ratio = {f_ratio:.4f}")
    
    print("\n3. CENTROID DISTANCES BY FEATURE")
    print("-" * 80)
    print("Separation between cluster centers for each feature:")
    
    centroids = kmeans.cluster_centers_
    for i, col in enumerate(features.columns):
        # Calculate pairwise distances between centroids for this feature
        dist_01 = abs(centroids[0, i] - centroids[1, i])
        dist_02 = abs(centroids[0, i] - centroids[2, i])
        dist_12 = abs(centroids[1, i] - centroids[2, i])
        avg_dist = (dist_01 + dist_02 + dist_12) / 3
        
        print(f"  {col:20s}: Avg distance = {avg_dist:.4f}")
        print(f"    C0↔C1: {dist_01:.4f}, C0↔C2: {dist_02:.4f}, C1↔C2: {dist_12:.4f}")
    
    print("\n4. FEATURE IMPORTANCE RANKING")
    print("-" * 80)
    sorted_features = sorted(importance_scores.items(), key=lambda x: x[1], reverse=True)
    
    print("Based on F-ratio (between-cluster variance / within-cluster variance):\n")
    for rank, (feature, score) in enumerate(sorted_features, 1):
        percentage = (score / sum(importance_scores.values())) * 100
        print(f"  {rank}. {feature:20s}: {score:.4f} ({percentage:.1f}% of total)")
    
    return sorted_features


def compare_scenarios(df):
    """Compare different feature combinations."""
    print("\n" + "=" * 80)
    print("CLUSTERING COMPARISON: 1 FEATURE vs 4 FEATURES")
    print("=" * 80)
    
    scenarios = [
        ("Mean Only", ['total_cost_mean']),
        ("Mean + Std", ['total_cost_mean', 'total_cost_std']),
        ("Mean + Min/Max", ['total_cost_mean', 'total_cost_min', 'total_cost_max']),
        ("All 4 Features", ['total_cost_mean', 'total_cost_std', 'total_cost_min', 'total_cost_max'])
    ]
    
    results = []
    
    for name, features in scenarios:
        result = cluster_with_features(df, features)
        results.append({
            'name': name,
            'n_features': len(features),
            'features': features,
            **result
        })
    
    # Print comparison table
    print("\nQUALITY METRICS COMPARISON:")
    print("-" * 80)
    print(f"{'Scenario':<20s} {'Features':>10s} {'Silhouette':>12s} {'Davies-B':>10s} {'Calinski':>10s} {'Var Exp':>10s}")
    print("-" * 80)
    
    for r in results:
        print(f"{r['name']:<20s} {r['n_features']:>10d} {r['silhouette']:>12.4f} "
              f"{r['davies_bouldin']:>10.4f} {r['calinski']:>10.2f} {r['variance_explained']:>9.1f}%")
    
    print("\n" + "=" * 80)
    print("INTERPRETATION:")
    print("=" * 80)
    
    # Find best scenario
    best_silhouette = max(results, key=lambda x: x['silhouette'])
    best_davies = min(results, key=lambda x: x['davies_bouldin'])
    best_variance = max(results, key=lambda x: x['variance_explained'])
    
    print(f"\n✓ Best Silhouette Score (separation): {best_silhouette['name']}")
    print(f"  → Score: {best_silhouette['silhouette']:.4f}")
    
    print(f"\n✓ Best Davies-Bouldin (low overlap): {best_davies['name']}")
    print(f"  → Score: {best_davies['davies_bouldin']:.4f}")
    
    print(f"\n✓ Best Variance Explained: {best_variance['name']}")
    print(f"  → Percentage: {best_variance['variance_explained']:.1f}%")
    
    # Overall recommendation
    print("\n" + "=" * 80)
    print("RECOMMENDATION:")
    print("=" * 80)
    
    mean_only = results[0]
    all_four = results[3]
    
    improvement_silhouette = (all_four['silhouette'] - mean_only['silhouette']) / mean_only['silhouette'] * 100
    improvement_variance = all_four['variance_explained'] - mean_only['variance_explained']
    
    print(f"\nUsing 4 features vs Mean only:")
    print(f"  • Silhouette improvement: {improvement_silhouette:+.1f}%")
    print(f"  • Additional variance explained: {improvement_variance:+.1f}%")
    print(f"  • Davies-Bouldin: {all_four['davies_bouldin']:.4f} vs {mean_only['davies_bouldin']:.4f}")
    
    if improvement_silhouette > 5:
        print(f"\n✓ VERDICT: 4 features is SIGNIFICANTLY BETTER")
        print(f"  The additional features (std, min, max) capture important cost patterns")
        print(f"  that the mean alone misses (volatility and range).")
    elif improvement_silhouette > 0:
        print(f"\n✓ VERDICT: 4 features is SLIGHTLY BETTER")
        print(f"  Minor improvement, but std/min/max add useful information.")
    else:
        print(f"\n⚠ VERDICT: Mean alone is SUFFICIENT")
        print(f"  Additional features don't significantly improve clustering.")
    
    return results


def visualize_comparison(results):
    """Create visualization comparing different feature sets."""
    fig, axes = plt.subplots(2, 2, figsize=(14, 10))
    
    names = [r['name'] for r in results]
    colors = ['#E74C3C', '#F39C12', '#3498DB', '#2ECC71']
    
    # 1. Silhouette scores
    ax1 = axes[0, 0]
    bars = ax1.bar(names, [r['silhouette'] for r in results], color=colors, alpha=0.8, edgecolor='black')
    ax1.axhline(y=0.5, color='green', linestyle='--', linewidth=2, alpha=0.5, label='Good threshold')
    ax1.set_ylabel('Silhouette Score', fontweight='bold')
    ax1.set_title('Cluster Separation Quality', fontweight='bold', fontsize=12)
    ax1.set_ylim(0, 1)
    ax1.legend()
    ax1.grid(True, alpha=0.3, axis='y')
    for bar in bars:
        height = bar.get_height()
        ax1.text(bar.get_x() + bar.get_width()/2., height,
                f'{height:.3f}', ha='center', va='bottom', fontweight='bold')
    plt.setp(ax1.xaxis.get_majorticklabels(), rotation=15, ha='right')
    
    # 2. Variance explained
    ax2 = axes[0, 1]
    bars = ax2.bar(names, [r['variance_explained'] for r in results], color=colors, alpha=0.8, edgecolor='black')
    ax2.set_ylabel('Variance Explained (%)', fontweight='bold')
    ax2.set_title('Information Captured by Clustering', fontweight='bold', fontsize=12)
    ax2.set_ylim(0, 100)
    ax2.grid(True, alpha=0.3, axis='y')
    for bar in bars:
        height = bar.get_height()
        ax2.text(bar.get_x() + bar.get_width()/2., height,
                f'{height:.1f}%', ha='center', va='bottom', fontweight='bold')
    plt.setp(ax2.xaxis.get_majorticklabels(), rotation=15, ha='right')
    
    # 3. Davies-Bouldin (lower is better)
    ax3 = axes[1, 0]
    bars = ax3.bar(names, [r['davies_bouldin'] for r in results], color=colors, alpha=0.8, edgecolor='black')
    ax3.axhline(y=1.0, color='green', linestyle='--', linewidth=2, alpha=0.5, label='Good threshold')
    ax3.set_ylabel('Davies-Bouldin Index', fontweight='bold')
    ax3.set_title('Cluster Overlap (Lower = Better)', fontweight='bold', fontsize=12)
    ax3.legend()
    ax3.grid(True, alpha=0.3, axis='y')
    for bar in bars:
        height = bar.get_height()
        ax3.text(bar.get_x() + bar.get_width()/2., height,
                f'{height:.3f}', ha='center', va='bottom', fontweight='bold')
    plt.setp(ax3.xaxis.get_majorticklabels(), rotation=15, ha='right')
    
    # 4. Summary comparison
    ax4 = axes[1, 1]
    ax4.axis('off')
    
    # Create comparison table
    comparison_text = "SUMMARY\n" + "="*40 + "\n\n"
    comparison_text += f"Best Overall: {results[-1]['name']}\n\n"
    comparison_text += "Metrics (4 features vs Mean only):\n"
    comparison_text += f"  Silhouette: {results[-1]['silhouette']:.3f} vs {results[0]['silhouette']:.3f}\n"
    comparison_text += f"  Davies-B:   {results[-1]['davies_bouldin']:.3f} vs {results[0]['davies_bouldin']:.3f}\n"
    comparison_text += f"  Var Expl:   {results[-1]['variance_explained']:.1f}% vs {results[0]['variance_explained']:.1f}%\n\n"
    
    improvement = ((results[-1]['silhouette'] - results[0]['silhouette']) / results[0]['silhouette']) * 100
    comparison_text += f"Improvement: {improvement:+.1f}%\n\n"
    
    if improvement > 5:
        comparison_text += "✓ 4 features SIGNIFICANTLY better\n"
        comparison_text += "  Use all features for clustering"
    elif improvement > 0:
        comparison_text += "✓ 4 features slightly better\n"
        comparison_text += "  Use all features for best results"
    else:
        comparison_text += "⚠ Mean alone sufficient\n"
        comparison_text += "  Additional features don't help"
    
    ax4.text(0.1, 0.9, comparison_text, transform=ax4.transAxes,
            fontsize=11, verticalalignment='top', family='monospace',
            bbox=dict(boxstyle='round', facecolor='lightblue', alpha=0.3))
    
    plt.suptitle('Feature Comparison for Employee Clustering', fontsize=14, fontweight='bold')
    plt.tight_layout()
    
    return fig


def main():
    # Load data
    data_path = project_root / "outputs" / "payroll_long.csv"
    df = pd.read_csv(data_path)
    df['month'] = pd.to_datetime(df['month'])
    
    print(f"Loaded {len(df)} records, {df['employee_id'].nunique()} employees\n")
    
    # Analyze feature importance
    feature_ranking = analyze_feature_importance(df)
    
    # Compare different scenarios
    results = compare_scenarios(df)
    
    # Create visualization
    print("\nGenerating comparison plots...")
    fig = visualize_comparison(results)
    
    output_path = project_root / "outputs" / "feature_comparison.png"
    fig.savefig(output_path, dpi=150, bbox_inches='tight')
    print(f"✓ Plot saved to: {output_path}")
    
    plt.show()
    
    print("\n✓ Analysis complete!")


if __name__ == "__main__":
    main()
