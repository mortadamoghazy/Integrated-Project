"""
run_clustering.py

Run K-Means clustering analysis on payroll data.
"""

import sys
from pathlib import Path

# Ensure project root is on path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pandas as pd
from src.features.analytics.employee_clustering import EmployeeClustering


def main():
    """Run clustering analysis."""
    print("=" * 60)
    print("K-MEANS CLUSTERING ANALYSIS - EMPLOYEE SEGMENTATION")
    print("=" * 60)
    
    # Load data
    data_path = project_root / "outputs" / "payroll_long.csv"
    print(f"\nLoading data from: {data_path}")
    
    df = pd.read_csv(data_path)
    df['month'] = pd.to_datetime(df['month'])
    
    print(f"✓ Data loaded: {len(df)} records, {df['employee_id'].nunique()} employees")
    
    # Initialize clustering
    n_clusters = 3
    print(f"\nInitializing K-Means with K={n_clusters} clusters...")
    clustering = EmployeeClustering(df, n_clusters=n_clusters)
    
    # Engineer features
    print("Engineering features...")
    features = clustering.engineer_features()
    print(f"✓ Created {len(features.columns)} features:")
    for col in features.columns:
        print(f"  - {col}")
    
    # Fit model
    print(f"\nFitting K-Means model...")
    labels = clustering.fit()
    print(f"✓ Clustering complete!")
    print(f"  Silhouette Score: {clustering.silhouette:.3f}")
    print(f"  (Score > 0.5 = Good, > 0.7 = Excellent)")
    
    # Get profiles
    print("\n" + "=" * 60)
    print("CLUSTER PROFILES")
    print("=" * 60)
    
    profiles = clustering.get_cluster_profiles()
    for _, row in profiles.iterrows():
        print(f"\n{'─' * 60}")
        print(f"CLUSTER {int(row['cluster_id'])}")
        print(f"{'─' * 60}")
        print(f"  Size:                {int(row['size'])} employees")
        print(f"  Avg Total Cost:      ${row['avg_total_cost']:,.2f}/month")
        print(f"  Avg Gross Salary:    ${row['avg_gross_salary']:,.2f}/month")
        print(f"  Avg Contributions:   ${row['avg_contributions']:,.2f}/month")
        print(f"  Avg Benefits:        ${row['avg_benefits']:,.2f}/month")
        print(f"  Cost Volatility:     {row['avg_volatility']*100:.1f}%")
        print(f"  Cost Growth Rate:    {row['avg_growth_rate']*100:.1f}%")
        print(f"  Employee IDs:        {', '.join(map(str, row['employee_ids']))}")
    
    # Business interpretation
    print("\n" + "=" * 60)
    print("BUSINESS INSIGHTS")
    print("=" * 60)
    
    # Sort by cost
    sorted_profiles = profiles.sort_values('avg_total_cost')
    
    low_cost = sorted_profiles.iloc[0]
    high_cost = sorted_profiles.iloc[-1]
    
    print(f"\n💡 Low-Cost Tier (Cluster {int(low_cost['cluster_id'])}):")
    print(f"   - {int(low_cost['size'])} employees @ ${low_cost['avg_total_cost']:,.0f}/month avg")
    print(f"   - Suitable for: Entry-level, standard positions")
    print(f"   - Budget planning: ${low_cost['avg_total_cost']*int(low_cost['size']):,.0f}/month total")
    
    print(f"\n💰 High-Cost Tier (Cluster {int(high_cost['cluster_id'])}):")
    print(f"   - {int(high_cost['size'])} employees @ ${high_cost['avg_total_cost']:,.0f}/month avg")
    print(f"   - Likely: Senior staff, specialists, or variable comp")
    print(f"   - Retention focus: {int(high_cost['size'])} key employees")
    
    if len(profiles) > 2:
        mid_cost = sorted_profiles.iloc[1]
        print(f"\n📊 Mid-Tier (Cluster {int(mid_cost['cluster_id'])}):")
        print(f"   - {int(mid_cost['size'])} employees @ ${mid_cost['avg_total_cost']:,.0f}/month avg")
        print(f"   - Growth potential: Promotion candidates")
    
    # Find most volatile cluster
    most_volatile = profiles.loc[profiles['avg_volatility'].idxmax()]
    print(f"\n⚠️  Most Volatile Cluster: {int(most_volatile['cluster_id'])}")
    print(f"   - Volatility: {most_volatile['avg_volatility']*100:.1f}%")
    print(f"   - Action: Investigate compensation structure variability")
    
    print("\n" + "=" * 60)
    print("VISUALIZATIONS")
    print("=" * 60)
    print("\nGenerating plots...")
    
    import matplotlib.pyplot as plt
    
    # Generate all visualizations
    fig1 = clustering.plot_elbow_silhouette()
    fig2 = clustering.plot_clusters_2d()
    fig3 = clustering.plot_cluster_profiles()
    
    print("✓ 3 figures created")
    print("\nShowing plots... (close to continue)")
    plt.show()
    
    # Save results
    output_path = project_root / "outputs" / "clustering_results.csv"
    
    # Add cluster labels to original data
    emp_clusters = clustering.features_df.copy()
    emp_clusters['cluster'] = clustering.cluster_labels
    emp_clusters.to_csv(output_path)
    
    print(f"\n✓ Results saved to: {output_path}")
    print("\n" + "=" * 60)
    print("ANALYSIS COMPLETE")
    print("=" * 60)


if __name__ == "__main__":
    main()
