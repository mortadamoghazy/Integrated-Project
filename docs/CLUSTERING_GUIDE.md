# Employee Clustering Analysis Guide

## Overview

The **K-Means Clustering** module provides unsupervised machine learning for automatic employee segmentation based on compensation patterns. This helps identify natural groupings in your workforce without manual categorization.

## 🤖 What is K-Means Clustering?

K-Means is a **machine learning algorithm** that:
- Discovers patterns in data without labeled examples (unsupervised learning)
- Optimizes cluster assignments by minimizing within-cluster variance
- Learns optimal cluster centers from the data itself
- Can predict which cluster new employees would belong to

### Mathematical Foundation

K-Means solves an optimization problem:

```
minimize: J = Σ Σ ||x - μᵢ||²
          i=1..k  x∈Cᵢ
```

Where:
- `k` = number of clusters
- `Cᵢ` = cluster i  
- `x` = data point (employee)
- `μᵢ` = centroid (center) of cluster i

This is similar to how linear regression minimizes squared errors!

## 📊 Features Used for Clustering

The algorithm creates these features from your payroll data:

### Cost Metrics:
- **Average Total Cost**: Mean monthly cost per employee
- **Cost Volatility**: Standard deviation / mean (coefficient of variation)
- **Cost Range**: Min and max cost across months
- **Cost Growth Rate**: (Latest cost - First cost) / First cost

### Compensation Components:
- **Average Gross Salary**: Mean `salaire_brut`
- **Average Employer Contributions**: Mean `cot_patronale`  
- **Average Benefits**: Mean `avantages`
- **Average Net Pay**: Mean `net_paye`

### Derived Features:
- **Contribution Ratio**: Contributions / Gross Salary
- **Months Worked**: Number of months in dataset

Total: **13 features** per employee

## 🎯 Business Use Cases

### 1. **Workforce Segmentation**
Automatically group employees into cost tiers:
- **Low-tier**: Entry-level, standard positions
- **Mid-tier**: Experienced workers
- **High-tier**: Senior staff, specialists

### 2. **Budget Planning**
- Know average cost per tier
- Estimate costs for new hires based on cluster
- Allocate budgets by employee segment

### 3. **Anomaly Detection**
- Identify employees with unusual compensation patterns
- Detect volatility in specific groups
- Ensure compensation fairness

### 4. **Retention Strategy**
- Focus retention efforts on high-cost clusters
- Understand promotion paths between clusters
- Identify growth potential candidates

### 5. **Hiring Decisions**
- Predict which cluster a new hire will belong to
- Estimate realistic compensation packages
- Maintain balanced workforce composition

## 🚀 How to Use

### Method 1: Dashboard (GUI)

1. Launch the dashboard:
   ```bash
   python scripts/run_dashboard.py
   ```

2. Click **"10. Employee Clustering (K-Means ML)"**

3. Configure settings:
   - **Number of Clusters (K)**: Choose 2-8 (default: 3)
   - **Visualization Type**:
     - `2D Cluster Map (PCA)`: See employees plotted with cluster assignments
     - `Cluster Profiles`: Bar charts showing cluster characteristics
     - `Elbow & Silhouette`: Optimization curves to find best K

4. Click **"Run Clustering"**

5. Review results:
   - Cluster summary popup shows statistics
   - Visualization displays in dashboard
   - Silhouette score indicates quality (>0.5 = good, >0.7 = excellent)

### Method 2: Command Line Script

Run standalone analysis:
```bash
python scripts/run_clustering.py
```

This generates:
- Detailed cluster profiles in terminal
- 3 visualization figures (Elbow/Silhouette, 2D map, profiles)
- CSV output: `outputs/clustering_results.csv`

### Method 3: Python API

Use in your own code:

```python
from src.features.analytics.employee_clustering import EmployeeClustering
import pandas as pd

# Load data
df = pd.read_csv("outputs/payroll_long.csv")
df['month'] = pd.to_datetime(df['month'])

# Initialize with K=3 clusters
clustering = EmployeeClustering(df, n_clusters=3)

# Fit model
labels = clustering.fit()
print(f"Silhouette Score: {clustering.silhouette:.3f}")

# Get cluster profiles
profiles = clustering.get_cluster_profiles()
print(profiles)

# Predict cluster for new employee
new_emp_features = [[3500, 200, 3400, 3600, ...]]  # 13 features
cluster = clustering.predict(new_emp_features)
print(f"New employee belongs to Cluster {cluster}")

# Generate visualizations
fig1 = clustering.plot_clusters_2d()
fig2 = clustering.plot_cluster_profiles()
plt.show()
```

## 📈 Understanding Results

### Silhouette Score

Measures clustering quality (range: -1 to 1):
- **> 0.7**: Excellent - strong, distinct clusters
- **0.5 - 0.7**: Good - reasonable structure
- **0.25 - 0.5**: Fair - weak structure, consider different K
- **< 0.25**: Poor - no clear clustering

**Your data achieved: 0.381** (Fair)
- Reasonable but overlapping clusters
- Consider 2 or 4 clusters as alternatives

### Elbow Method

Plot shows "inertia" (total within-cluster variance) vs. K:
- Look for "elbow" where curve bends
- Diminishing returns after elbow point
- Elbow at K=3 or K=4 suggests optimal value

### Cluster Profiles

Example interpretation:

```
Cluster 0: 17 employees @ $2,717/month avg
  → Low-cost tier (entry-level)
  → Budget: $46K/month total
  
Cluster 2: 10 employees @ $4,564/month avg  
  → Mid-tier (experienced)
  → Promotion candidates
  
Cluster 1: 3 employees @ $4,701/month avg
  → High-cost (senior/specialists)
  → 13% volatility - investigate why
  → 47.6% growth rate - recent promotions?
```

## 🔬 Technical Details

### Algorithm Steps

1. **Feature Engineering**: Extract 13 features per employee
2. **Normalization**: StandardScaler (zero mean, unit variance)
3. **Initialization**: K-Means++ (smart centroid placement)
4. **Iteration**:
   - Assign each employee to nearest centroid
   - Recalculate centroids as cluster means
   - Repeat until convergence
5. **Evaluation**: Calculate silhouette score
6. **Dimensionality Reduction**: PCA for 2D visualization

### Parameters

- **`n_clusters`** (K): Number of groups (2-8 recommended)
- **`random_state=42`**: Reproducibility
- **`n_init=10`**: Run algorithm 10 times, keep best result

### Why It's Machine Learning

✅ **Learns from data**: Discovers patterns without hand-coded rules  
✅ **Optimizes objective**: Minimizes within-cluster variance  
✅ **Generalizes**: Can predict cluster for new employees  
✅ **Has parameters**: Centroids learned via iterative optimization  
✅ **Needs tuning**: Hyperparameter K selection  

It's **unsupervised ML** - learns structure without labels!

## 📊 Output Files

### `clustering_results.csv`
Contains per-employee data with cluster assignments:

```csv
employee_id,total_cost_mean,total_cost_std,...,cluster
00010,5247.83,0.0,...,2
00020,5197.35,328.18,...,2
00030,3584.11,141.76,...,0
...
```

## 🎨 Visualizations

### 1. 2D Cluster Map (PCA)
- Employees plotted in 2D space
- Colors indicate clusters
- Red X marks show centroids
- Axes show % variance explained

### 2. Cluster Profiles
- Cluster sizes (bar chart)
- Average costs (horizontal bars)
- Cost component breakdown (grouped bars)
- Volatility & growth (dual-axis line chart)

### 3. Elbow & Silhouette
- Elbow curve (inertia vs K)
- Silhouette scores (quality vs K)
- Helps choose optimal K

## ⚠️ Limitations

### Data Requirements
- **Minimum**: 2 employees (K < N)
- **Recommended**: 20+ employees for meaningful clusters
- **Ideal**: 100+ employees for statistical reliability

### Assumptions
- Clusters are spherical (circular in feature space)
- All features equally important (use feature scaling)
- Number of clusters known/guessed in advance

### Sensitivity
- Results depend on K choice
- Sensitive to outliers (high-cost employees)
- May need to run multiple times with different K

## 🔄 Comparison with Other Methods

| Method | Type | Pros | Cons |
|--------|------|------|------|
| **K-Means** | Unsupervised ML | Fast, scalable, interpretable | Needs K, assumes spherical |
| **DBSCAN** | Unsupervised ML | Finds outliers, no K needed | Sensitive to density |
| **Hierarchical** | Unsupervised ML | Shows dendrogram | Slow on large data |
| **Manual Rules** | Rule-based | Simple, explicit | Inflexible, no learning |

## 🎓 Further Learning

### Key Concepts
- **Unsupervised Learning**: Learning without labels
- **Clustering**: Grouping similar data points
- **Centroid**: Center point of a cluster
- **Inertia**: Sum of squared distances to centroids
- **Silhouette Score**: Clustering quality metric
- **PCA**: Principal Component Analysis (dimensionality reduction)

### Next Steps
1. Try different K values (2, 4, 5)
2. Experiment with feature selection
3. Compare with hierarchical clustering
4. Use clusters for predictive modeling

## 📞 Support

For questions or issues with clustering:
1. Check silhouette score (< 0.25 = poor clustering)
2. Try different K values
3. Inspect cluster profiles for business logic
4. Review feature distributions

---

**Remember**: Clustering is **exploratory** - use business judgment to interpret and validate results!
