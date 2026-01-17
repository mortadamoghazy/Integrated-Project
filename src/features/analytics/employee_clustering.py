"""
employee_clustering.py

K-Means clustering for employee cost segmentation.
Discovers natural groupings in payroll data using unsupervised machine learning.
"""

import pandas as pd
import numpy as np
from sklearn.cluster import KMeans
from sklearn.preprocessing import StandardScaler
from sklearn.decomposition import PCA
from sklearn.metrics import silhouette_score
import matplotlib.pyplot as plt
from pathlib import Path
import sys

# Ensure project root is on path
project_root = Path(__file__).parent.parent.parent.parent
sys.path.insert(0, str(project_root))


class EmployeeClustering:
    """K-Means clustering for employee segmentation."""
    
    def __init__(self, df: pd.DataFrame, n_clusters: int = 3):
        """
        Initialize clustering model.
        
        Args:
            df: DataFrame with employee payroll data (long format)
            n_clusters: Number of clusters (default 3)
        """
        self.df = df.copy()
        self.n_clusters = n_clusters
        self.scaler = StandardScaler()
        self.kmeans = None
        self.features_df = None
        self.cluster_labels = None
        self.pca = None
        self.silhouette = None
        
    def engineer_features(self):
        """
        Create features for clustering from raw payroll data.
        
        Features:
        - Average total cost per employee
        - Cost volatility (std/mean)
        - Average gross salary
        - Average employer contributions
        - Average benefits
        - Contribution ratio (contributions/salary)
        - Cost growth rate
        """
        # Aggregate per employee
        agg_dict = {
            'total_cost': ['mean', 'std', 'min', 'max'],
            'salaire_brut': ['mean', 'std'],
            'cot_patronale': ['mean'],
            'avantages': ['mean'],
            'net_paye': ['mean']
        }
        
        features = self.df.groupby('employee_id').agg(agg_dict)
        features.columns = ['_'.join(col).strip() for col in features.columns.values]
        
        # Cost volatility (coefficient of variation)
        features['cost_volatility'] = features['total_cost_std'] / features['total_cost_mean']
        features['cost_volatility'] = features['cost_volatility'].fillna(0)
        
        # Contribution ratio
        features['contribution_ratio'] = features['cot_patronale_mean'] / features['salaire_brut_mean']
        features['contribution_ratio'] = features['contribution_ratio'].fillna(0)
        
        # Cost growth rate
        first_costs = self.df.groupby('employee_id').first()['total_cost']
        last_costs = self.df.groupby('employee_id').last()['total_cost']
        features['cost_growth_rate'] = ((last_costs - first_costs) / first_costs).fillna(0)
        
        # Number of months worked
        features['months_worked'] = self.df.groupby('employee_id')['month'].nunique()
        
        # Replace inf values
        features = features.replace([np.inf, -np.inf], 0)
        
        self.features_df = features
        return features
    
    def find_optimal_clusters(self, max_k: int = 8):
        """
        Find optimal number of clusters using Elbow method and Silhouette score.
        
        Args:
            max_k: Maximum number of clusters to test
            
        Returns:
            Dictionary with inertias and silhouette scores
        """
        if self.features_df is None:
            self.engineer_features()
        
        # Normalize features
        X_scaled = self.scaler.fit_transform(self.features_df)
        
        inertias = []
        silhouettes = []
        K_range = range(2, min(max_k + 1, len(self.features_df)))
        
        for k in K_range:
            kmeans = KMeans(n_clusters=k, random_state=42, n_init=10)
            labels = kmeans.fit_predict(X_scaled)
            inertias.append(kmeans.inertia_)
            silhouettes.append(silhouette_score(X_scaled, labels))
        
        return {
            'K_values': list(K_range),
            'inertias': inertias,
            'silhouettes': silhouettes
        }
    
    def fit(self):
        """
        Fit K-Means clustering model.
        
        Returns:
            Cluster labels for each employee
        """
        if self.features_df is None:
            self.engineer_features()
        
        # Normalize features
        X_scaled = self.scaler.fit_transform(self.features_df)
        
        # Fit K-Means
        self.kmeans = KMeans(n_clusters=self.n_clusters, random_state=42, n_init=10)
        self.cluster_labels = self.kmeans.fit_predict(X_scaled)
        
        # Calculate silhouette score
        self.silhouette = silhouette_score(X_scaled, self.cluster_labels)
        
        # PCA for visualization
        self.pca = PCA(n_components=2)
        self.pca_coords = self.pca.fit_transform(X_scaled)
        
        return self.cluster_labels
    
    def get_cluster_profiles(self):
        """
        Get detailed profiles for each cluster.
        
        Returns:
            DataFrame with cluster statistics
        """
        if self.cluster_labels is None:
            self.fit()
        
        # Add cluster labels to features
        features_with_clusters = self.features_df.copy()
        features_with_clusters['cluster'] = self.cluster_labels
        
        # Calculate statistics per cluster
        profiles = []
        for cluster_id in range(self.n_clusters):
            cluster_data = features_with_clusters[features_with_clusters['cluster'] == cluster_id]
            
            profile = {
                'cluster_id': cluster_id,
                'size': len(cluster_data),
                'avg_total_cost': cluster_data['total_cost_mean'].mean(),
                'avg_gross_salary': cluster_data['salaire_brut_mean'].mean(),
                'avg_contributions': cluster_data['cot_patronale_mean'].mean(),
                'avg_benefits': cluster_data['avantages_mean'].mean(),
                'avg_volatility': cluster_data['cost_volatility'].mean(),
                'avg_growth_rate': cluster_data['cost_growth_rate'].mean(),
                'employee_ids': cluster_data.index.tolist()
            }
            profiles.append(profile)
        
        # Sort by average cost
        profiles_df = pd.DataFrame(profiles).sort_values('avg_total_cost')
        
        return profiles_df
    
    def predict(self, employee_features):
        """
        Predict cluster for new employee data.
        
        Args:
            employee_features: Feature vector or DataFrame
            
        Returns:
            Cluster label
        """
        if self.kmeans is None:
            raise ValueError("Model not fitted. Call fit() first.")
        
        # Ensure features match training
        X_scaled = self.scaler.transform(employee_features)
        return self.kmeans.predict(X_scaled)
    
    def plot_elbow_silhouette(self, max_k: int = 8):
        """
        Plot Elbow curve and Silhouette scores to find optimal K.
        
        Args:
            max_k: Maximum number of clusters to test
            
        Returns:
            matplotlib figure
        """
        results = self.find_optimal_clusters(max_k)
        
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 5))
        
        # Elbow curve
        ax1.plot(results['K_values'], results['inertias'], 
                marker='o', linewidth=2, markersize=8, color='#2E86AB')
        ax1.set_xlabel('Number of Clusters (K)', fontsize=12)
        ax1.set_ylabel('Inertia (Within-Cluster Sum of Squares)', fontsize=12)
        ax1.set_title('Elbow Method', fontsize=14, fontweight='bold')
        ax1.grid(True, alpha=0.3)
        
        # Silhouette scores
        ax2.plot(results['K_values'], results['silhouettes'], 
                marker='s', linewidth=2, markersize=8, color='#A23B72')
        ax2.axhline(y=0.5, color='red', linestyle='--', label='Good threshold (0.5)')
        ax2.set_xlabel('Number of Clusters (K)', fontsize=12)
        ax2.set_ylabel('Silhouette Score', fontsize=12)
        ax2.set_title('Silhouette Analysis', fontsize=14, fontweight='bold')
        ax2.legend()
        ax2.grid(True, alpha=0.3)
        
        plt.tight_layout()
        return fig
    
    def plot_clusters_2d(self):
        """
        Plot clusters in 2D using PCA.
        
        Returns:
            matplotlib figure
        """
        if self.cluster_labels is None:
            self.fit()
        
        fig, ax = plt.subplots(figsize=(12, 8))
        
        # Color map
        colors = ['#2E86AB', '#A23B72', '#F18F01', '#C73E1D', '#6A994E']
        
        # Plot each cluster
        for cluster_id in range(self.n_clusters):
            mask = self.cluster_labels == cluster_id
            cluster_points = self.pca_coords[mask]
            color = colors[cluster_id % len(colors)]
            
            ax.scatter(cluster_points[:, 0], cluster_points[:, 1], 
                      c=color, label=f'Cluster {cluster_id}',
                      s=150, alpha=0.7, edgecolors='black', linewidth=1.5)
            
            # Annotate employee IDs
            for i, (x, y) in enumerate(cluster_points):
                emp_id = self.features_df.index[mask].tolist()[i]
                ax.annotate(str(emp_id), (x, y), fontsize=9, 
                           ha='center', va='center', fontweight='bold')
        
        # Plot centroids
        centroids_scaled = self.kmeans.cluster_centers_
        centroids_pca = self.pca.transform(centroids_scaled)
        ax.scatter(centroids_pca[:, 0], centroids_pca[:, 1], 
                  marker='X', s=400, c='red', edgecolors='black', 
                  linewidth=2, label='Centroids', zorder=10)
        
        ax.set_xlabel(f'PC1 ({self.pca.explained_variance_ratio_[0]:.1%} variance)', 
                     fontsize=12)
        ax.set_ylabel(f'PC2 ({self.pca.explained_variance_ratio_[1]:.1%} variance)', 
                     fontsize=12)
        ax.set_title(f'Employee Clusters (K-Means, K={self.n_clusters})\n'
                    f'Silhouette Score: {self.silhouette:.3f}', 
                    fontsize=14, fontweight='bold', pad=20)
        ax.legend(loc='best', fontsize=10)
        ax.grid(True, alpha=0.3)
        
        plt.tight_layout()
        return fig
    
    def plot_cluster_profiles(self):
        """
        Plot cluster profiles showing characteristics.
        
        Returns:
            matplotlib figure
        """
        profiles = self.get_cluster_profiles()
        
        fig, axes = plt.subplots(2, 2, figsize=(14, 10))
        colors = ['#2E86AB', '#A23B72', '#F18F01', '#C73E1D', '#6A994E']
        
        # 1. Cluster sizes
        ax1 = axes[0, 0]
        bars1 = ax1.bar(profiles['cluster_id'], profiles['size'], 
                       color=[colors[i % len(colors)] for i in range(len(profiles))],
                       alpha=0.8, edgecolor='black')
        ax1.set_xlabel('Cluster ID', fontsize=11)
        ax1.set_ylabel('Number of Employees', fontsize=11)
        ax1.set_title('Cluster Sizes', fontsize=12, fontweight='bold')
        ax1.grid(True, alpha=0.3, axis='y')
        
        # Add value labels
        for bar in bars1:
            height = bar.get_height()
            ax1.text(bar.get_x() + bar.get_width()/2., height,
                    f'{int(height)}', ha='center', va='bottom', fontweight='bold')
        
        # 2. Average total cost
        ax2 = axes[0, 1]
        bars2 = ax2.barh(profiles['cluster_id'], profiles['avg_total_cost'],
                        color=[colors[i % len(colors)] for i in range(len(profiles))],
                        alpha=0.8, edgecolor='black')
        ax2.set_ylabel('Cluster ID', fontsize=11)
        ax2.set_xlabel('Average Total Cost ($)', fontsize=11)
        ax2.set_title('Average Cost per Cluster', fontsize=12, fontweight='bold')
        ax2.grid(True, alpha=0.3, axis='x')
        
        # Add value labels
        for bar in bars2:
            width = bar.get_width()
            ax2.text(width, bar.get_y() + bar.get_height()/2.,
                    f'${width:,.0f}', ha='left', va='center', fontweight='bold')
        
        # 3. Cost components comparison
        ax3 = axes[1, 0]
        x = np.arange(len(profiles))
        width = 0.25
        
        ax3.bar(x - width, profiles['avg_gross_salary'], width, 
               label='Gross Salary', color='#2E86AB', alpha=0.8)
        ax3.bar(x, profiles['avg_contributions'], width,
               label='Contributions', color='#A23B72', alpha=0.8)
        ax3.bar(x + width, profiles['avg_benefits'], width,
               label='Benefits', color='#F18F01', alpha=0.8)
        
        ax3.set_xlabel('Cluster ID', fontsize=11)
        ax3.set_ylabel('Amount ($)', fontsize=11)
        ax3.set_title('Cost Components by Cluster', fontsize=12, fontweight='bold')
        ax3.set_xticks(x)
        ax3.set_xticklabels(profiles['cluster_id'])
        ax3.legend()
        ax3.grid(True, alpha=0.3, axis='y')
        
        # 4. Volatility and growth
        ax4 = axes[1, 1]
        ax4_twin = ax4.twinx()
        
        line1 = ax4.plot(profiles['cluster_id'], profiles['avg_volatility'] * 100, 
                        marker='o', linewidth=2, markersize=8, color='#C73E1D',
                        label='Volatility (%)')
        line2 = ax4_twin.plot(profiles['cluster_id'], profiles['avg_growth_rate'] * 100,
                             marker='s', linewidth=2, markersize=8, color='#6A994E',
                             label='Growth Rate (%)')
        
        ax4.set_xlabel('Cluster ID', fontsize=11)
        ax4.set_ylabel('Cost Volatility (%)', fontsize=11, color='#C73E1D')
        ax4_twin.set_ylabel('Growth Rate (%)', fontsize=11, color='#6A994E')
        ax4.set_title('Volatility & Growth by Cluster', fontsize=12, fontweight='bold')
        ax4.tick_params(axis='y', labelcolor='#C73E1D')
        ax4_twin.tick_params(axis='y', labelcolor='#6A994E')
        ax4.grid(True, alpha=0.3)
        
        # Combined legend
        lines = line1 + line2
        labels = [l.get_label() for l in lines]
        ax4.legend(lines, labels, loc='upper left')
        
        plt.tight_layout()
        return fig


def main():
    """Example usage."""
    # Load data
    data_path = project_root / "outputs" / "payroll_long.csv"
    df = pd.read_csv(data_path)
    df['month'] = pd.to_datetime(df['month'])
    
    # Initialize clustering
    clustering = EmployeeClustering(df, n_clusters=3)
    
    # Fit model
    print("Fitting K-Means clustering...")
    labels = clustering.fit()
    print(f"✓ Clustering complete! Silhouette score: {clustering.silhouette:.3f}")
    
    # Get profiles
    print("\n=== Cluster Profiles ===")
    profiles = clustering.get_cluster_profiles()
    for _, row in profiles.iterrows():
        print(f"\nCluster {int(row['cluster_id'])}:")
        print(f"  Size: {int(row['size'])} employees")
        print(f"  Avg Total Cost: ${row['avg_total_cost']:,.2f}")
        print(f"  Avg Gross Salary: ${row['avg_gross_salary']:,.2f}")
        print(f"  Volatility: {row['avg_volatility']*100:.1f}%")
        print(f"  Growth Rate: {row['avg_growth_rate']*100:.1f}%")
        print(f"  Employees: {row['employee_ids']}")
    
    # Plot
    print("\nGenerating visualizations...")
    fig1 = clustering.plot_elbow_silhouette()
    fig2 = clustering.plot_clusters_2d()
    fig3 = clustering.plot_cluster_profiles()
    
    plt.show()


if __name__ == "__main__":
    main()
