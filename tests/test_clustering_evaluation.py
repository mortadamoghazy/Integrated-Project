"""
test_clustering_evaluation.py

Comprehensive evaluation script for K-Means clustering with detailed metrics.
Tests clustering quality using multiple statistical measures.
"""

import sys
from pathlib import Path

# Ensure project root is on path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from sklearn.metrics import (
    silhouette_score, 
    silhouette_samples,
    davies_bouldin_score,
    calinski_harabasz_score
)
from scipy.spatial.distance import cdist
from src.features.analytics.employee_clustering import EmployeeClustering


def evaluate_clustering_comprehensive(df: pd.DataFrame, n_clusters: int = 3):
    """
    Comprehensive evaluation of clustering with multiple metrics.
    
    Args:
        df: Payroll DataFrame
        n_clusters: Number of clusters (default 3)
        
    Returns:
        Dictionary with all evaluation metrics
    """
    print("=" * 80)
    print(f"COMPREHENSIVE K-MEANS CLUSTERING EVALUATION (K={n_clusters})")
    print("=" * 80)
    
    # Initialize and fit clustering
    clustering = EmployeeClustering(df, n_clusters=n_clusters)
    labels = clustering.fit()
    
    # Get features and scaled features
    features = clustering.features_df
    X_scaled = clustering.scaler.transform(features)
    
    print(f"\n{'Dataset Information':-^80}")
    print(f"Total Employees: {len(features)}")
    print(f"Features Used: {list(features.columns)}")
    print(f"Number of Clusters: {n_clusters}")
    
    # ==================== CLUSTER QUALITY METRICS ====================
    print(f"\n{'CLUSTER QUALITY METRICS':-^80}")
    
    # 1. Silhouette Score (Higher is better: -1 to 1)
    silhouette_avg = silhouette_score(X_scaled, labels)
    print(f"\n1. Silhouette Score: {silhouette_avg:.4f}")
    print(f"   Interpretation:")
    if silhouette_avg > 0.7:
        print(f"   ✓ EXCELLENT - Strong, well-separated clusters")
    elif silhouette_avg > 0.5:
        print(f"   ✓ GOOD - Reasonable cluster structure")
    elif silhouette_avg > 0.25:
        print(f"   ⚠ FAIR - Weak structure, some overlap")
    else:
        print(f"   ✗ POOR - No clear clustering")
    
    # 2. Davies-Bouldin Index (Lower is better: 0 to ∞)
    davies_bouldin = davies_bouldin_score(X_scaled, labels)
    print(f"\n2. Davies-Bouldin Index: {davies_bouldin:.4f}")
    print(f"   Interpretation:")
    if davies_bouldin < 1.0:
        print(f"   ✓ EXCELLENT - Well-separated, compact clusters")
    elif davies_bouldin < 2.0:
        print(f"   ✓ GOOD - Acceptable separation")
    else:
        print(f"   ⚠ FAIR - Clusters may be overlapping")
    
    # 3. Calinski-Harabasz Score (Higher is better: 0 to ∞)
    calinski = calinski_harabasz_score(X_scaled, labels)
    print(f"\n3. Calinski-Harabasz Score: {calinski:.2f}")
    print(f"   Interpretation: Higher = denser, better separated clusters")
    
    # 4. Inertia (Within-cluster sum of squares)
    inertia = clustering.kmeans.inertia_
    print(f"\n4. Inertia (WCSS): {inertia:.2f}")
    print(f"   Lower = tighter clusters")
    
    # ==================== PER-CLUSTER ANALYSIS ====================
    print(f"\n{'PER-CLUSTER ANALYSIS':-^80}")
    
    # Silhouette scores per cluster
    silhouette_vals = silhouette_samples(X_scaled, labels)
    
    cluster_profiles = clustering.get_cluster_profiles()
    
    for _, row in cluster_profiles.iterrows():
        cluster_id = int(row['cluster_id'])
        cluster_mask = labels == cluster_id
        cluster_silhouette = silhouette_vals[cluster_mask].mean()
        
        print(f"\n--- Cluster {cluster_id} ---")
        print(f"  Size: {int(row['size'])} employees ({int(row['size'])/len(features)*100:.1f}%)")
        print(f"  Avg Total Cost: ${row['avg_total_cost']:,.2f}/month")
        print(f"  Avg Gross Salary: ${row['avg_gross_salary']:,.2f}")
        print(f"  Avg Contributions: ${row['avg_contributions']:,.2f}")
        print(f"  Avg Benefits: ${row['avg_benefits']:,.2f}")
        print(f"  Cost Volatility: {row['avg_volatility']:.4f}")
        print(f"  Cost Growth Rate: {row['avg_growth_rate']*100:+.2f}%")
        print(f"  Cluster Silhouette: {cluster_silhouette:.4f}")
        
        if cluster_silhouette < silhouette_avg - 0.1:
            print(f"  ⚠ Warning: This cluster has below-average cohesion")
    
    # ==================== DISTANCE ANALYSIS ====================
    print(f"\n{'DISTANCE ANALYSIS':-^80}")
    
    # Distance from each point to cluster centers
    centroids = clustering.kmeans.cluster_centers_
    distances = cdist(X_scaled, centroids, 'euclidean')
    
    # Average distance to own centroid per cluster
    print(f"\nAverage distance to cluster centroid:")
    for cluster_id in range(n_clusters):
        cluster_mask = labels == cluster_id
        avg_dist = distances[cluster_mask, cluster_id].mean()
        print(f"  Cluster {cluster_id}: {avg_dist:.4f}")
    
    # Inter-centroid distances
    inter_distances = cdist(centroids, centroids, 'euclidean')
    print(f"\nInter-centroid distances (separation):")
    for i in range(n_clusters):
        for j in range(i+1, n_clusters):
            print(f"  Cluster {i} ↔ Cluster {j}: {inter_distances[i, j]:.4f}")
    
    # ==================== VARIANCE EXPLAINED ====================
    print(f"\n{'VARIANCE ANALYSIS':-^80}")
    
    # Total variance
    total_variance = np.var(X_scaled, axis=0).sum()
    
    # Within-cluster variance
    within_variance = 0
    for cluster_id in range(n_clusters):
        cluster_mask = labels == cluster_id
        cluster_data = X_scaled[cluster_mask]
        within_variance += np.var(cluster_data, axis=0).sum() * cluster_data.shape[0]
    within_variance /= len(X_scaled)
    
    # Between-cluster variance
    between_variance = total_variance - within_variance
    variance_ratio = between_variance / total_variance * 100
    
    print(f"Total Variance: {total_variance:.4f}")
    print(f"Within-Cluster Variance: {within_variance:.4f}")
    print(f"Between-Cluster Variance: {between_variance:.4f}")
    print(f"Variance Explained by Clustering: {variance_ratio:.2f}%")
    
    # ==================== CLUSTER STABILITY ====================
    print(f"\n{'CLUSTER STABILITY ANALYSIS':-^80}")
    
    # Check for outliers (points far from their centroid)
    print(f"\nPotential outliers (>2 std from centroid):")
    for cluster_id in range(n_clusters):
        cluster_mask = labels == cluster_id
        cluster_distances = distances[cluster_mask, cluster_id]
        threshold = cluster_distances.mean() + 2 * cluster_distances.std()
        outliers = features.index[cluster_mask][cluster_distances > threshold]
        
        if len(outliers) > 0:
            print(f"  Cluster {cluster_id}: {list(outliers)}")
        else:
            print(f"  Cluster {cluster_id}: None")
    
    # ==================== BUSINESS METRICS ====================
    print(f"\n{'BUSINESS METRICS':-^80}")
    
    # Total payroll per cluster
    print(f"\nMonthly payroll by cluster:")
    for _, row in cluster_profiles.iterrows():
        cluster_id = int(row['cluster_id'])
        total_payroll = row['avg_total_cost'] * row['size']
        print(f"  Cluster {cluster_id}: ${total_payroll:,.2f} ({int(row['size'])} employees)")
    
    overall_total = cluster_profiles['avg_total_cost'] * cluster_profiles['size']
    print(f"  TOTAL: ${overall_total.sum():,.2f}")
    
    # Cost distribution
    print(f"\nCost tier interpretation:")
    sorted_profiles = cluster_profiles.sort_values('avg_total_cost')
    tiers = ['Low-Cost (Entry)', 'Mid-Cost (Experienced)', 'High-Cost (Senior)']
    for idx, (_, row) in enumerate(sorted_profiles.iterrows()):
        tier_name = tiers[idx] if idx < len(tiers) else f"Tier {idx+1}"
        print(f"  Cluster {int(row['cluster_id'])}: {tier_name}")
        print(f"    → ${row['avg_total_cost']:,.0f}/month average")
    
    # ==================== SUMMARY ====================
    print(f"\n{'OVERALL ASSESSMENT':-^80}")
    
    quality_score = 0
    assessments = []
    
    # Scoring based on metrics
    if silhouette_avg > 0.5:
        quality_score += 3
        assessments.append("✓ Good cluster separation (Silhouette)")
    elif silhouette_avg > 0.25:
        quality_score += 2
        assessments.append("⚠ Moderate separation (Silhouette)")
    else:
        quality_score += 1
        assessments.append("✗ Weak separation (Silhouette)")
    
    if davies_bouldin < 1.5:
        quality_score += 3
        assessments.append("✓ Low cluster overlap (Davies-Bouldin)")
    elif davies_bouldin < 2.5:
        quality_score += 2
        assessments.append("⚠ Some overlap (Davies-Bouldin)")
    else:
        quality_score += 1
        assessments.append("✗ High overlap (Davies-Bouldin)")
    
    if variance_ratio > 60:
        quality_score += 3
        assessments.append("✓ High variance explained")
    elif variance_ratio > 40:
        quality_score += 2
        assessments.append("⚠ Moderate variance explained")
    else:
        quality_score += 1
        assessments.append("✗ Low variance explained")
    
    print(f"\nQuality Score: {quality_score}/9")
    for assessment in assessments:
        print(f"  {assessment}")
    
    if quality_score >= 8:
        print(f"\n🌟 EXCELLENT: Clustering is highly effective")
    elif quality_score >= 6:
        print(f"\n✓ GOOD: Clustering provides useful segmentation")
    elif quality_score >= 4:
        print(f"\n⚠ FAIR: Clustering has some value but may need adjustment")
    else:
        print(f"\n✗ POOR: Consider different number of clusters or features")
    
    print("\n" + "=" * 80)
    
    # Return all metrics
    return {
        'silhouette_score': silhouette_avg,
        'davies_bouldin': davies_bouldin,
        'calinski_harabasz': calinski,
        'inertia': inertia,
        'variance_explained': variance_ratio,
        'cluster_profiles': cluster_profiles,
        'quality_score': quality_score,
        'clustering_object': clustering
    }


def plot_detailed_evaluation(metrics: dict):
    """Create detailed evaluation plots."""
    clustering = metrics['clustering_object']
    
    fig = plt.figure(figsize=(16, 12))
    gs = fig.add_gridspec(3, 3, hspace=0.3, wspace=0.3)
    
    # 1. Cluster scatter (1D with jitter)
    ax1 = fig.add_subplot(gs[0, :2])
    colors = ['#2E86AB', '#A23B72', '#F18F01']
    
    # Get feature values
    feature_values = clustering.features_df['total_cost_mean'].values
    np.random.seed(42)
    y_jitter = np.random.normal(0, 0.05, len(feature_values))
    
    for cluster_id in range(clustering.n_clusters):
        mask = clustering.cluster_labels == cluster_id
        cluster_values = feature_values[mask]
        cluster_jitter = y_jitter[mask]
        ax1.scatter(cluster_values, cluster_jitter,
                   c=colors[cluster_id], label=f'Cluster {cluster_id}',
                   s=100, alpha=0.7, edgecolors='black')
    
    # Plot centroids
    centroids_scaled = clustering.kmeans.cluster_centers_
    centroids_original = clustering.scaler.inverse_transform(centroids_scaled)
    for i, centroid in enumerate(centroids_original):
        ax1.axvline(x=centroid[0], color=colors[i], linestyle='--', linewidth=2, alpha=0.5)
        ax1.scatter(centroid[0], 0, marker='X', s=300, c='red', 
                   edgecolors='black', linewidth=2, zorder=10)
    
    ax1.set_xlabel('Average Total Cost ($)', fontweight='bold')
    ax1.set_ylabel('Jitter')
    ax1.set_yticks([])
    ax1.set_title('Employee Clustering (K=3)', fontweight='bold', fontsize=12)
    ax1.legend()
    ax1.grid(True, alpha=0.3, axis='x')
    ax1.legend()
    ax1.grid(True, alpha=0.3)
    
    # 2. Silhouette plot
    ax2 = fig.add_subplot(gs[0, 2])
    from sklearn.metrics import silhouette_samples
    X_scaled = clustering.scaler.transform(clustering.features_df)
    silhouette_vals = silhouette_samples(X_scaled, clustering.cluster_labels)
    
    y_lower = 10
    for i in range(clustering.n_clusters):
        cluster_silhouette_vals = silhouette_vals[clustering.cluster_labels == i]
        cluster_silhouette_vals.sort()
        
        size_cluster_i = cluster_silhouette_vals.shape[0]
        y_upper = y_lower + size_cluster_i
        
        ax2.fill_betweenx(np.arange(y_lower, y_upper), 0, cluster_silhouette_vals,
                         facecolor=colors[i], edgecolor=colors[i], alpha=0.7)
        ax2.text(-0.05, y_lower + 0.5 * size_cluster_i, str(i))
        y_lower = y_upper + 10
    
    ax2.axvline(x=metrics['silhouette_score'], color="red", linestyle="--",
               label=f'Average: {metrics["silhouette_score"]:.3f}')
    ax2.set_xlabel("Silhouette Coefficient")
    ax2.set_ylabel("Cluster")
    ax2.set_title("Silhouette Plot", fontweight='bold')
    ax2.legend()
    
    # 3. Cluster sizes
    ax3 = fig.add_subplot(gs[1, 0])
    profiles = metrics['cluster_profiles']
    ax3.bar(profiles['cluster_id'], profiles['size'],
           color=colors, alpha=0.8, edgecolor='black')
    ax3.set_xlabel('Cluster ID')
    ax3.set_ylabel('Number of Employees')
    ax3.set_title('Cluster Sizes', fontweight='bold')
    ax3.grid(True, alpha=0.3, axis='y')
    for i, (_, row) in enumerate(profiles.iterrows()):
        ax3.text(i, row['size'], str(int(row['size'])),
                ha='center', va='bottom', fontweight='bold')
    
    # 4. Average costs
    ax4 = fig.add_subplot(gs[1, 1])
    ax4.barh(profiles['cluster_id'], profiles['avg_total_cost'],
            color=colors, alpha=0.8, edgecolor='black')
    ax4.set_ylabel('Cluster ID')
    ax4.set_xlabel('Average Total Cost ($)')
    ax4.set_title('Average Cost per Cluster', fontweight='bold')
    ax4.grid(True, alpha=0.3, axis='x')
    
    # 5. Metrics summary
    ax5 = fig.add_subplot(gs[1, 2])
    ax5.axis('off')
    metrics_text = f"""QUALITY METRICS
    
Silhouette: {metrics['silhouette_score']:.3f}
Davies-Bouldin: {metrics['davies_bouldin']:.3f}
Calinski-Harabasz: {metrics['calinski_harabasz']:.1f}

Variance Explained: {metrics['variance_explained']:.1f}%

Quality Score: {metrics['quality_score']}/9
"""
    ax5.text(0.1, 0.9, metrics_text, transform=ax5.transAxes,
            fontsize=10, verticalalignment='top', family='monospace',
            bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.5))
    
    # 6. Feature distributions by cluster (single feature)
    features = clustering.features_df
    feature_names = features.columns
    
    # For single feature, create a single centered plot
    ax = fig.add_subplot(gs[2, 0:2])
    for cluster_id in range(clustering.n_clusters):
        mask = clustering.cluster_labels == cluster_id
        cluster_data = features[mask][feature_names[0]]
        ax.hist(cluster_data, bins=15, alpha=0.6, label=f'Cluster {cluster_id}',
               color=colors[cluster_id], edgecolor='black')
    ax.set_xlabel('Average Total Cost ($)', fontsize=11, fontweight='bold')
    ax.set_ylabel('Number of Employees', fontsize=11, fontweight='bold')
    ax.set_title('Total Cost Distribution by Cluster', fontsize=12, fontweight='bold')
    ax.legend(fontsize=10)
    ax.grid(True, alpha=0.3)
    
    plt.suptitle('Comprehensive Clustering Evaluation (1-Feature)', fontsize=14, fontweight='bold', y=0.995)
    
    return fig


def main():
    """Run comprehensive clustering evaluation."""
    # Load data
    data_path = project_root / "outputs" / "payroll_long.csv"
    print(f"Loading data from: {data_path}")
    
    df = pd.read_csv(data_path)
    df['month'] = pd.to_datetime(df['month'])
    
    print(f"Data loaded: {len(df)} records, {df['employee_id'].nunique()} employees\n")
    
    # Run comprehensive evaluation with K=3
    metrics = evaluate_clustering_comprehensive(df, n_clusters=3)
    
    # Generate detailed plots
    print("\nGenerating detailed evaluation plots...")
    fig = plot_detailed_evaluation(metrics)
    
    # Save plot
    output_path = project_root / "outputs" / "clustering_evaluation_detailed.png"
    fig.savefig(output_path, dpi=150, bbox_inches='tight')
    print(f"✓ Detailed plot saved to: {output_path}")
    
    plt.show()
    
    print("\n✓ Evaluation complete!")


if __name__ == "__main__":
    main()
