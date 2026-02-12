"""
Clustering Validation and Visualization
========================================
Generate comprehensive validation plots for K-Means employee clustering.
"""

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')  # Non-interactive backend
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path
from sklearn.cluster import KMeans
from sklearn.metrics import silhouette_score, silhouette_samples, davies_bouldin_score, calinski_harabasz_score
from sklearn.preprocessing import StandardScaler
import warnings
warnings.filterwarnings('ignore')

# Set style
plt.style.use('seaborn-v0_8-darkgrid')
sns.set_palette("husl")

# Load data
print("=" * 80)
print("CLUSTERING VALIDATION AND VISUALIZATION")
print("=" * 80)

# Load raw data
df = pd.read_csv('outputs/payroll_long.csv')
df['month'] = pd.to_datetime(df['month'])

# Calculate employee-level features
print("\nCalculating employee-level features...")
emp_features = df.groupby('employee_id').agg({
    'total_cost': ['mean', 'std', 'min', 'max'],
    'salaire_brut': 'mean',
    'cot_patronale': 'mean',
    'avantages': 'mean'
}).reset_index()

emp_features.columns = ['employee_id', 'total_cost_mean', 'total_cost_std', 
                        'total_cost_min', 'total_cost_max', 
                        'salaire_brut_mean', 'cot_patronale_mean', 'avantages_mean']

# Handle missing values
emp_features = emp_features.fillna(0)

# Prepare features for clustering
feature_cols = ['total_cost_mean', 'total_cost_std', 'salaire_brut_mean', 
                'cot_patronale_mean', 'avantages_mean']
X = emp_features[feature_cols].values

# Standardize features
scaler = StandardScaler()
X_scaled = scaler.fit_transform(X)

# Create output directory
output_dir = Path('outputs/plots')
output_dir.mkdir(exist_ok=True)

print(f"Loaded {len(emp_features)} employees with {len(feature_cols)} features")

# ===========================
# Plot 1: Elbow Method
# ===========================
print("\n1. Generating Elbow Method plot...")

inertias = []
silhouette_scores = []
k_range = range(2, 11)

for k in k_range:
    kmeans = KMeans(n_clusters=k, random_state=42, n_init=10)
    kmeans.fit(X_scaled)
    inertias.append(kmeans.inertia_)
    silhouette_scores.append(silhouette_score(X_scaled, kmeans.labels_))

fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 5))

# Inertia plot
ax1.plot(k_range, inertias, marker='o', linewidth=2, markersize=10, color='#2E86AB')
ax1.set_xlabel('Number of Clusters (k)', fontsize=12, fontweight='bold')
ax1.set_ylabel('Inertia (Within-Cluster Sum of Squares)', fontsize=12, fontweight='bold')
ax1.set_title('Elbow Method for Optimal k', fontsize=14, fontweight='bold')
ax1.grid(True, alpha=0.3)
ax1.axvline(x=3, color='red', linestyle='--', linewidth=2, alpha=0.7, label='Selected k=3')
ax1.legend(fontsize=10)

# Add value labels
for x, y in zip(k_range, inertias):
    ax1.annotate(f'{y:.0f}', (x, y), textcoords="offset points", 
                xytext=(0, 10), ha='center', fontsize=9)

# Silhouette plot
ax2.plot(k_range, silhouette_scores, marker='s', linewidth=2, markersize=10, color='#A23B72')
ax2.set_xlabel('Number of Clusters (k)', fontsize=12, fontweight='bold')
ax2.set_ylabel('Silhouette Score', fontsize=12, fontweight='bold')
ax2.set_title('Silhouette Score vs Number of Clusters', fontsize=14, fontweight='bold')
ax2.grid(True, alpha=0.3)
ax2.axvline(x=3, color='red', linestyle='--', linewidth=2, alpha=0.7, label='Selected k=3')
ax2.legend(fontsize=10)

# Add value labels
for x, y in zip(k_range, silhouette_scores):
    ax2.annotate(f'{y:.3f}', (x, y), textcoords="offset points",
                xytext=(0, 10), ha='center', fontsize=9)

plt.tight_layout()
plt.savefig(output_dir / 'clustering_elbow_method.png', dpi=300, bbox_inches='tight')
print(f"   ✓ Saved: {output_dir / 'clustering_elbow_method.png'}")
plt.close()


# ===========================
# Plot 2: Silhouette Analysis (k=3)
# ===========================
print("\n2. Generating Silhouette Analysis plot...")

kmeans = KMeans(n_clusters=3, random_state=42, n_init=10)
cluster_labels = kmeans.fit_predict(X_scaled)

silhouette_avg = silhouette_score(X_scaled, cluster_labels)
sample_silhouette_values = silhouette_samples(X_scaled, cluster_labels)

fig, ax = plt.subplots(figsize=(10, 7))

y_lower = 10
colors = ['#2E86AB', '#A23B72', '#F18F01']

for i in range(3):
    # Aggregate silhouette scores for samples in cluster i
    ith_cluster_silhouette_values = sample_silhouette_values[cluster_labels == i]
    ith_cluster_silhouette_values.sort()
    
    size_cluster_i = ith_cluster_silhouette_values.shape[0]
    y_upper = y_lower + size_cluster_i
    
    ax.fill_betweenx(np.arange(y_lower, y_upper),
                     0, ith_cluster_silhouette_values,
                     facecolor=colors[i], edgecolor=colors[i], alpha=0.7)
    
    # Label the silhouette plots with their cluster numbers at the middle
    ax.text(-0.05, y_lower + 0.5 * size_cluster_i, f'Cluster {i}',
            fontsize=12, fontweight='bold')
    
    y_lower = y_upper + 10

ax.set_xlabel('Silhouette Coefficient', fontsize=12, fontweight='bold')
ax.set_ylabel('Cluster', fontsize=12, fontweight='bold')
ax.set_title(f'Silhouette Analysis (k=3, Average Score={silhouette_avg:.3f})', 
             fontsize=14, fontweight='bold')

# Average silhouette score line
ax.axvline(x=silhouette_avg, color='red', linestyle='--', linewidth=2,
          label=f'Average Silhouette Score: {silhouette_avg:.3f}')

ax.set_yticks([])
ax.set_xlim([-0.1, 1])
ax.legend(fontsize=10, loc='upper right')
ax.grid(True, alpha=0.3, axis='x')

plt.tight_layout()
plt.savefig(output_dir / 'clustering_silhouette_analysis.png', dpi=300, bbox_inches='tight')
print(f"   ✓ Saved: {output_dir / 'clustering_silhouette_analysis.png'}")
plt.close()


# ===========================
# Plot 3: Cluster Profiles
# ===========================
print("\n3. Generating Cluster Profiles plot...")

# Add cluster labels to features
emp_features['cluster'] = cluster_labels

# Calculate cluster statistics
cluster_stats = emp_features.groupby('cluster').agg({
    'total_cost_mean': ['mean', 'std', 'count'],
    'salaire_brut_mean': 'mean',
    'cot_patronale_mean': 'mean'
}).reset_index()

cluster_stats.columns = ['cluster', 'cost_mean', 'cost_std', 'count',
                         'salaire_mean', 'contrib_mean']

# Sort by cost
cluster_stats = cluster_stats.sort_values('cost_mean')
cluster_stats['tier'] = ['Low-Cost\n(Entry Level)', 'Mid-Cost\n(Experienced)', 'High-Cost\n(Senior)']

fig, axes = plt.subplots(2, 2, figsize=(14, 10))

# Cluster sizes
ax = axes[0, 0]
bars = ax.bar(cluster_stats['tier'], cluster_stats['count'], 
              color=colors, alpha=0.8, edgecolor='black')
ax.set_ylabel('Number of Employees', fontsize=11, fontweight='bold')
ax.set_title('Cluster Sizes', fontsize=13, fontweight='bold')
ax.grid(True, alpha=0.3, axis='y')
for bar, count in zip(bars, cluster_stats['count']):
    height = bar.get_height()
    ax.text(bar.get_x() + bar.get_width()/2., height,
            f'{int(count)}', ha='center', va='bottom', fontsize=11, fontweight='bold')

# Average cost per cluster
ax = axes[0, 1]
bars = ax.bar(cluster_stats['tier'], cluster_stats['cost_mean'],
              yerr=cluster_stats['cost_std'], color=colors, alpha=0.8,
              edgecolor='black', capsize=5)
ax.set_ylabel('Average Total Cost (€)', fontsize=11, fontweight='bold')
ax.set_title('Average Cost per Cluster', fontsize=13, fontweight='bold')
ax.grid(True, alpha=0.3, axis='y')
for bar, cost in zip(bars, cluster_stats['cost_mean']):
    height = bar.get_height()
    ax.text(bar.get_x() + bar.get_width()/2., height + 200,
            f'€{cost:.0f}', ha='center', va='bottom', fontsize=10, fontweight='bold')

# Cost distribution per cluster
ax = axes[1, 0]
for i, (cluster, tier) in enumerate(zip(cluster_stats['cluster'], cluster_stats['tier'])):
    cluster_costs = emp_features[emp_features['cluster'] == cluster]['total_cost_mean']
    ax.hist(cluster_costs, bins=10, alpha=0.6, color=colors[i], 
            label=tier, edgecolor='black')
ax.set_xlabel('Total Cost (€)', fontsize=11, fontweight='bold')
ax.set_ylabel('Number of Employees', fontsize=11, fontweight='bold')
ax.set_title('Cost Distribution by Cluster', fontsize=13, fontweight='bold')
ax.legend(fontsize=9)
ax.grid(True, alpha=0.3, axis='y')

# Cluster composition (salary vs contributions)
ax = axes[1, 1]
x_pos = np.arange(len(cluster_stats))
width = 0.35
bars1 = ax.bar(x_pos - width/2, cluster_stats['salaire_mean'], width,
               label='Gross Salary', color='#2E86AB', alpha=0.8, edgecolor='black')
bars2 = ax.bar(x_pos + width/2, cluster_stats['contrib_mean'], width,
               label='Employer Contributions', color='#A23B72', alpha=0.8, edgecolor='black')
ax.set_ylabel('Amount (€)', fontsize=11, fontweight='bold')
ax.set_title('Salary vs Contributions by Cluster', fontsize=13, fontweight='bold')
ax.set_xticks(x_pos)
ax.set_xticklabels(cluster_stats['tier'])
ax.legend(fontsize=10)
ax.grid(True, alpha=0.3, axis='y')

plt.suptitle('Employee Clustering Profiles (k=3)', fontsize=16, fontweight='bold', y=0.995)
plt.tight_layout()
plt.savefig(output_dir / 'clustering_profiles.png', dpi=300, bbox_inches='tight')
print(f"   ✓ Saved: {output_dir / 'clustering_profiles.png'}")
plt.close()


# ===========================
# Plot 4: Validation Metrics Comparison
# ===========================
print("\n4. Generating Validation Metrics plot...")

# Calculate multiple validation metrics for k=2 to 10
validation_metrics = {
    'k': [],
    'Silhouette Score': [],
    'Davies-Bouldin Index': [],
    'Calinski-Harabasz Index': []
}

for k in range(2, 11):
    kmeans = KMeans(n_clusters=k, random_state=42, n_init=10)
    labels = kmeans.fit_predict(X_scaled)
    
    validation_metrics['k'].append(k)
    validation_metrics['Silhouette Score'].append(silhouette_score(X_scaled, labels))
    validation_metrics['Davies-Bouldin Index'].append(davies_bouldin_score(X_scaled, labels))
    validation_metrics['Calinski-Harabasz Index'].append(calinski_harabasz_score(X_scaled, labels))

metrics_df = pd.DataFrame(validation_metrics)

fig, axes = plt.subplots(1, 3, figsize=(16, 5))

# Silhouette Score (higher is better)
ax = axes[0]
ax.plot(metrics_df['k'], metrics_df['Silhouette Score'], 
        marker='o', linewidth=2, markersize=10, color='#2E86AB')
ax.axvline(x=3, color='red', linestyle='--', linewidth=2, alpha=0.7)
ax.set_xlabel('Number of Clusters (k)', fontsize=11, fontweight='bold')
ax.set_ylabel('Silhouette Score', fontsize=11, fontweight='bold')
ax.set_title('Silhouette Score\n(Higher is Better)', fontsize=12, fontweight='bold')
ax.grid(True, alpha=0.3)

# Davies-Bouldin Index (lower is better)
ax = axes[1]
ax.plot(metrics_df['k'], metrics_df['Davies-Bouldin Index'],
        marker='s', linewidth=2, markersize=10, color='#A23B72')
ax.axvline(x=3, color='red', linestyle='--', linewidth=2, alpha=0.7)
ax.set_xlabel('Number of Clusters (k)', fontsize=11, fontweight='bold')
ax.set_ylabel('Davies-Bouldin Index', fontsize=11, fontweight='bold')
ax.set_title('Davies-Bouldin Index\n(Lower is Better)', fontsize=12, fontweight='bold')
ax.grid(True, alpha=0.3)

# Calinski-Harabasz Index (higher is better)
ax = axes[2]
ax.plot(metrics_df['k'], metrics_df['Calinski-Harabasz Index'],
        marker='^', linewidth=2, markersize=10, color='#F18F01')
ax.axvline(x=3, color='red', linestyle='--', linewidth=2, alpha=0.7)
ax.set_xlabel('Number of Clusters (k)', fontsize=11, fontweight='bold')
ax.set_ylabel('Calinski-Harabasz Index', fontsize=11, fontweight='bold')
ax.set_title('Calinski-Harabasz Index\n(Higher is Better)', fontsize=12, fontweight='bold')
ax.grid(True, alpha=0.3)

plt.suptitle('Clustering Validation Metrics', fontsize=16, fontweight='bold', y=1.02)
plt.tight_layout()
plt.savefig(output_dir / 'clustering_validation_metrics.png', dpi=300, bbox_inches='tight')
print(f"   ✓ Saved: {output_dir / 'clustering_validation_metrics.png'}")
plt.close()


# ===========================
# Plot 5: 2D Visualization (PCA)
# ===========================
print("\n5. Generating 2D PCA visualization...")

from sklearn.decomposition import PCA

# Apply PCA for 2D visualization
pca = PCA(n_components=2)
X_pca = pca.fit_transform(X_scaled)

fig, ax = plt.subplots(figsize=(10, 8))

# Plot each cluster
for i in range(3):
    cluster_mask = cluster_labels == i
    ax.scatter(X_pca[cluster_mask, 0], X_pca[cluster_mask, 1],
              c=colors[i], label=cluster_stats.iloc[i]['tier'],
              s=150, alpha=0.7, edgecolors='black', linewidth=1.5)

# Plot cluster centers
centers_pca = pca.transform(kmeans.cluster_centers_)
ax.scatter(centers_pca[:, 0], centers_pca[:, 1],
          c='red', marker='X', s=400, edgecolors='black',
          linewidth=2, label='Centroids', zorder=10)

ax.set_xlabel(f'First Principal Component ({pca.explained_variance_ratio_[0]:.1%} variance)',
             fontsize=12, fontweight='bold')
ax.set_ylabel(f'Second Principal Component ({pca.explained_variance_ratio_[1]:.1%} variance)',
             fontsize=12, fontweight='bold')
ax.set_title('Employee Clustering: 2D PCA Projection', fontsize=14, fontweight='bold')
ax.legend(fontsize=11, loc='best')
ax.grid(True, alpha=0.3)

# Add annotation
total_variance = pca.explained_variance_ratio_[0] + pca.explained_variance_ratio_[1]
ax.text(0.02, 0.98, f'Total variance explained: {total_variance:.1%}',
        transform=ax.transAxes, va='top', fontsize=10,
        bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.5))

plt.tight_layout()
plt.savefig(output_dir / 'clustering_pca_visualization.png', dpi=300, bbox_inches='tight')
print(f"   ✓ Saved: {output_dir / 'clustering_pca_visualization.png'}")
plt.close()


# ===========================
# Summary Statistics
# ===========================
print("\n" + "=" * 80)
print("CLUSTERING VALIDATION SUMMARY")
print("=" * 80)

print(f"\nOptimal number of clusters: k = 3")
print(f"\nValidation Metrics (k=3):")
print(f"  Silhouette Score:        {silhouette_avg:.4f} (range: -1 to 1, higher better)")
print(f"  Davies-Bouldin Index:    {davies_bouldin_score(X_scaled, cluster_labels):.4f} (lower better)")
print(f"  Calinski-Harabasz Index: {calinski_harabasz_score(X_scaled, cluster_labels):.2f} (higher better)")

print("\nCluster Profiles:")
for idx, row in cluster_stats.iterrows():
    print(f"\n  {row['tier'].replace(chr(10), ' ')}:")
    print(f"    Size: {int(row['count'])} employees ({row['count']/30*100:.1f}%)")
    print(f"    Avg Cost: €{row['cost_mean']:.2f} ± €{row['cost_std']:.2f}")
    print(f"    Total Monthly: €{row['cost_mean'] * row['count']:.2f}")

total_monthly = (cluster_stats['cost_mean'] * cluster_stats['count']).sum()
print(f"\n  Total Monthly Payroll: €{total_monthly:.2f}")

print("\n" + "=" * 80)
print("CLUSTERING VISUALIZATION COMPLETE")
print("=" * 80)
print(f"\nAll plots saved to: {output_dir.absolute()}")
print("\nGenerated plots:")
print("  1. clustering_elbow_method.png - Elbow method and silhouette scores")
print("  2. clustering_silhouette_analysis.png - Detailed silhouette analysis")
print("  3. clustering_profiles.png - Cluster profiles and distributions")
print("  4. clustering_validation_metrics.png - Multiple validation metrics")
print("  5. clustering_pca_visualization.png - 2D PCA visualization")
