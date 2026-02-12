"""
K-Means Clustering Evaluation
===============================
Comprehensive evaluation of K-Means clustering for employee cost segmentation.
Similar structure to forecasting evaluation with multiple metrics and K values.
"""

import pandas as pd
import numpy as np
from pathlib import Path
from sklearn.cluster import KMeans
from sklearn.metrics import silhouette_score, davies_bouldin_score, calinski_harabasz_score
from sklearn.preprocessing import StandardScaler
import warnings
warnings.filterwarnings('ignore')

# Project root
project_root = Path(__file__).parent.parent
output_dir = project_root / "outputs"

print("=" * 80)
print("K-MEANS CLUSTERING EVALUATION")
print("=" * 80)

# Load data
print("\n1. Loading payroll data...")
df = pd.read_csv(output_dir / 'payroll_long.csv')
df['month'] = pd.to_datetime(df['month'])

# Calculate employee-level features
print("2. Calculating employee-level features...")
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

print(f"   • {len(emp_features)} employees")
print(f"   • 7 features per employee")

# Prepare features for clustering
feature_cols = ['total_cost_mean', 'total_cost_std', 'salaire_brut_mean', 
                'cot_patronale_mean', 'avantages_mean']
X = emp_features[feature_cols].values

# Standardize features
scaler = StandardScaler()
X_scaled = scaler.fit_transform(X)

print(f"   • Features standardized (mean=0, std=1)")

# Evaluate only K = 3
print("\n3. Evaluating clustering for K = 3...")
k_range = [3]
results = []

for k in k_range:
    print(f"\n   Evaluating K = {k}:")
    
    # Fit K-Means
    kmeans = KMeans(n_clusters=k, random_state=42, n_init=10, max_iter=300)
    labels = kmeans.fit_predict(X_scaled)
    
    # Calculate validation metrics
    silhouette = silhouette_score(X_scaled, labels)
    davies_bouldin = davies_bouldin_score(X_scaled, labels)
    calinski_harabasz = calinski_harabasz_score(X_scaled, labels)
    inertia = kmeans.inertia_
    
    # Calculate cluster sizes
    unique, counts = np.unique(labels, return_counts=True)
    cluster_sizes = dict(zip(unique, counts))
    min_size = min(counts)
    max_size = max(counts)
    avg_size = np.mean(counts)
    
    # Calculate cluster cost statistics (in original scale)
    emp_features['cluster'] = labels
    cluster_stats = emp_features.groupby('cluster')['total_cost_mean'].agg(['mean', 'std', 'min', 'max'])
    
    avg_cluster_cost_mean = cluster_stats['mean'].mean()
    avg_cluster_cost_std = cluster_stats['std'].mean()
    
    # Store results
    results.append({
        'K': k,
        'Silhouette_Score': silhouette,
        'Davies_Bouldin_Index': davies_bouldin,
        'Calinski_Harabasz_Index': calinski_harabasz,
        'Inertia': inertia,
        'Min_Cluster_Size': min_size,
        'Max_Cluster_Size': max_size,
        'Avg_Cluster_Size': avg_size,
        'Avg_Cluster_Cost_Mean': avg_cluster_cost_mean,
        'Avg_Cluster_Cost_Std': avg_cluster_cost_std
    })
    
    print(f"      Silhouette Score:        {silhouette:.4f}")
    print(f"      Davies-Bouldin Index:    {davies_bouldin:.4f}")
    print(f"      Calinski-Harabasz Index: {calinski_harabasz:.2f}")
    print(f"      Inertia:                 {inertia:.2f}")
    print(f"      Cluster sizes: min={min_size}, max={max_size}, avg={avg_size:.1f}")

# Convert to DataFrame
results_df = pd.DataFrame(results)

# Set optimal K to 3 (only one value evaluated)
print("\n" + "=" * 80)
print("4. CLUSTERING ANALYSIS FOR K = 3")
print("=" * 80)

optimal_k = 3

# Add ranking columns (always rank 1 since only one K value)
results_df['Silhouette_Rank'] = 1
results_df['Davies_Bouldin_Rank'] = 1
results_df['Calinski_Harabasz_Rank'] = 1
results_df['Composite_Rank'] = 1.0
results_df['Overall_Rank'] = 1

# Detailed analysis for K=3
print("\n" + "=" * 80)
print(f"5. DETAILED ANALYSIS FOR K = 3")
print("=" * 80)

# Refit with optimal K
kmeans_optimal = KMeans(n_clusters=optimal_k, random_state=42, n_init=10, max_iter=300)
labels_optimal = kmeans_optimal.fit_predict(X_scaled)

emp_features['cluster'] = labels_optimal

# Cluster profiles
print(f"\nCluster Profiles:")
cluster_profiles = emp_features.groupby('cluster').agg({
    'employee_id': 'count',
    'total_cost_mean': ['mean', 'std', 'min', 'max'],
    'salaire_brut_mean': 'mean',
    'cot_patronale_mean': 'mean',
    'avantages_mean': 'mean'
}).round(2)

cluster_profiles.columns = ['Size', 'Avg_Cost', 'Std_Cost', 'Min_Cost', 'Max_Cost',
                            'Avg_Salary', 'Avg_Contributions', 'Avg_Benefits']

# Sort by average cost
cluster_profiles = cluster_profiles.sort_values('Avg_Cost')
cluster_profiles.reset_index(inplace=True)

# Assign tier names
tier_names = ['Low-Cost (Entry Level)', 'Mid-Cost (Experienced)', 'High-Cost (Senior)']
if optimal_k > 3:
    tier_names.extend([f'Tier {i+1}' for i in range(3, optimal_k)])
elif optimal_k < 3:
    tier_names = tier_names[:optimal_k]

cluster_profiles['Tier_Name'] = tier_names[:optimal_k]

print("\n" + "-" * 80)
for idx, row in cluster_profiles.iterrows():
    print(f"\n{row['Tier_Name']} (Cluster {row['cluster']}):")
    print(f"   • Size: {int(row['Size'])} employees ({row['Size']/len(emp_features)*100:.1f}%)")
    print(f"   • Avg Cost: €{row['Avg_Cost']:,.2f} ± €{row['Std_Cost']:,.2f}")
    print(f"   • Cost Range: €{row['Min_Cost']:,.2f} - €{row['Max_Cost']:,.2f}")
    print(f"   • Total Monthly: €{row['Avg_Cost'] * row['Size']:,.2f}")
    print(f"   • Avg Salary: €{row['Avg_Salary']:,.2f}")
    print(f"   • Avg Contributions: €{row['Avg_Contributions']:,.2f}")
    print(f"   • Avg Benefits: €{row['Avg_Benefits']:,.2f}")

total_monthly = (cluster_profiles['Avg_Cost'] * cluster_profiles['Size']).sum()
print(f"\nTotal Monthly Payroll: €{total_monthly:,.2f}")
print("-" * 80)

# Save results
print("\n" + "=" * 80)
print("6. SAVING RESULTS")
print("=" * 80)

# Save comprehensive results
results_file = output_dir / "clustering_evaluation.csv"
results_df.to_csv(results_file, index=False)
print(f"\n   ✓ Saved: {results_file}")

# Save cluster profiles for optimal K
profiles_file = output_dir / "clustering_profiles_optimal.csv"
cluster_profiles.to_csv(profiles_file, index=False)
print(f"   ✓ Saved: {profiles_file}")

# Save employee assignments for optimal K
emp_features_sorted = emp_features.sort_values(['cluster', 'total_cost_mean'])
assignments_file = output_dir / "clustering_employee_assignments.csv"
emp_features_sorted[['employee_id', 'total_cost_mean', 'cluster']].to_csv(assignments_file, index=False)
print(f"   ✓ Saved: {assignments_file}")

# Summary statistics
print("\n" + "=" * 80)
print("7. SUMMARY STATISTICS")
print("=" * 80)

print(f"\nK-Means Clustering Evaluation Summary:")
print(f"   • K evaluated: 3")
print(f"   • Number of employees: {len(emp_features)}")
print(f"   • Features used: {len(feature_cols)}")

print(f"\nK=3 Validation Metrics:")
optimal_row = results_df[results_df['K'] == optimal_k].iloc[0]
print(f"   • Silhouette Score:        {optimal_row['Silhouette_Score']:.4f}")
print(f"   • Davies-Bouldin Index:    {optimal_row['Davies_Bouldin_Index']:.4f}")
print(f"   • Calinski-Harabasz Index: {optimal_row['Calinski_Harabasz_Index']:.2f}")
print(f"   • Inertia:                 {optimal_row['Inertia']:.2f}")

print(f"\nCluster Quality Interpretation:")
silhouette_val = optimal_row['Silhouette_Score']
if silhouette_val > 0.7:
    quality = "Excellent"
elif silhouette_val > 0.5:
    quality = "Good"
elif silhouette_val > 0.3:
    quality = "Moderate"
else:
    quality = "Poor"
print(f"   • Silhouette Score {silhouette_val:.4f}: {quality} cluster separation")

db_val = optimal_row['Davies_Bouldin_Index']
if db_val < 0.5:
    db_quality = "Excellent"
elif db_val < 1.0:
    db_quality = "Good"
elif db_val < 1.5:
    db_quality = "Moderate"
else:
    db_quality = "Poor"
print(f"   • Davies-Bouldin Index {db_val:.4f}: {db_quality} cluster compactness")

print("\n" + "=" * 80)
print("CLUSTERING EVALUATION COMPLETE")
print("=" * 80)
print(f"\nOutput files:")
print(f"   1. {results_file.name} - K=3 clustering metrics")
print(f"   2. {profiles_file.name} - Detailed profiles for K=3")
print(f"   3. {assignments_file.name} - Employee-cluster assignments")
