# Payroll Analytics Dashboard Guide

## Overview

The **Payroll Analytics Dashboard** provides 10 comprehensive visualizations for strategic decision-making in small-medium enterprises (SMEs). All forecasting uses **AR(2) autoregression** model for accurate predictions.

## Accessing the Dashboard

### Method 1: From Excel Control Panel (Recommended)
1. Open `Control_Panel_New.xlsm`
2. Click **Button 4: Visualize Data**
3. Dashboard opens automatically

### Method 2: From Command Line
```bash
python scripts/run_dashboard.py
```

### Method 3: Direct Script
```bash
python src/features/analytics/dashboard_gui.py
```

## Data Source

- **Default:** Uses synthetic payroll data from `outputs/payroll_long.csv`
- **Purpose:** Synthetic data provides statistically valid insights without privacy concerns
- **Real data:** Used only for methodology validation, not for actual analysis

## 10 Visualization Options

### 1. **Total Payroll Cost Trend** 📈
**Purpose:** Track overall payroll spending over time

**Use Cases:**
- Budget planning and monitoring
- Identify cost spikes or anomalies
- Present to board/investors

**Metrics Shown:**
- Monthly total payroll cost
- Trend line with data points
- Dollar amounts labeled

**Decision Support:**
- "Are we growing sustainably?"
- "When did costs increase significantly?"
- "What's our average monthly burn?"

---

### 2. **Cost Forecasting (AR2)** 🔮
**Purpose:** Predict future payroll costs 6 months ahead

**Use Cases:**
- Cash flow planning
- Budget allocation
- Runway estimation

**Metrics Shown:**
- Historical data (solid line)
- 6-month forecast (dashed line)
- 95% confidence interval (shaded area)

**Model:** Autoregression order 2 (AR2)
- Uses past 2 months to predict next month
- Most accurate model from your testing
- Shows prediction uncertainty

**Decision Support:**
- "How much cash do we need for Q2?"
- "Can we afford 3 new hires?"
- "What's our payroll in 6 months?"

---

### 3. **Cost per Employee Trend** 👤💰
**Purpose:** Monitor efficiency and per-capita costs

**Use Cases:**
- Efficiency tracking
- Compare against industry benchmarks
- Identify cost inflation

**Metrics Shown:**
- Average cost per employee (top chart)
- Employee headcount (bottom chart)

**Decision Support:**
- "Is cost per employee rising?"
- "Are we hiring efficiently?"
- "Economies of scale working?"

---

### 4. **Payroll Cost Breakdown** 📊
**Purpose:** Understand where money actually goes

**Use Cases:**
- Tax/contribution optimization
- Negotiate benefit packages
- Understand true employment cost

**Metrics Shown:**
- Gross Salary (bottom layer)
- Employer Contributions (middle layer)
- Benefits (top layer)
- Stacked area chart showing proportions

**Decision Support:**
- "What's our tax burden?"
- "How much are benefits costing?"
- "Can we optimize contribution structure?"

---

### 5. **Employee Cost Distribution** 📦
**Purpose:** Identify salary equity and outliers

**Use Cases:**
- Compensation fairness analysis
- Identify compression issues
- Detect cost outliers

**Metrics Shown:**
- Box plot showing quartiles
- Histogram showing distribution
- Mean and median lines

**Decision Support:**
- "Is pay equitable?"
- "Do we have outliers?"
- "What's typical employee cost?"

---

### 6. **Top 10 Most Expensive Employees** 🏆
**Purpose:** Identify cost concentration and retention risks

**Use Cases:**
- Retention planning for key employees
- Budget impact of losing top talent
- Cost concentration analysis

**Metrics Shown:**
- Horizontal bar chart
- Employee IDs and total costs
- Latest month data

**Decision Support:**
- "Who are our most expensive employees?"
- "What if we lose our top 3?"
- "Is cost too concentrated?"

---

### 7. **Gross vs Net Salary** 💵➡️💰
**Purpose:** Compare gross salary to take-home pay

**Use Cases:**
- Understand employee perspective
- Tax burden analysis
- Compensation negotiations

**Metrics Shown:**
- Gross salary (solid line)
- Net pay (dashed line)
- Multiple employees comparison
- Interactive employee selection

**Flexibility:**
- Select specific employees to compare
- Multi-employee overlay
- Time series for each

**Decision Support:**
- "What's the take-home ratio?"
- "How much do taxes impact net pay?"
- "Compare compensation across employees"

---

### 8. **Headcount Trend** 👥
**Purpose:** Track workforce growth/reduction

**Use Cases:**
- Growth rate monitoring
- Capacity planning
- Turnover analysis

**Metrics Shown:**
- Employee count per month
- Trend line
- Labeled data points

**Decision Support:**
- "How fast are we growing?"
- "Did we have turnover spikes?"
- "Headcount vs revenue trends?"

---

### 9. **Monthly Burn Rate Dashboard** 🔥
**Purpose:** Comprehensive monthly cash analysis

**Use Cases:**
- Cash management
- Runway calculations
- Financial reporting

**Metrics Shown:**
- Monthly total burn (bar chart)
- Cost components breakdown (pie chart)
- Average per employee trend (line chart)
- Cumulative burn (area chart)

**Decision Support:**
- "What's our monthly burn?"
- "How long is our runway?"
- "Where can we cut costs?"

---

### 10. **Year-over-Year Comparison** 📅
**Purpose:** Compare performance across years

**Use Cases:**
- Identify seasonal patterns
- Measure year-over-year growth
- Budget accuracy assessment

**Metrics Shown:**
- Multiple year overlays
- Monthly comparison
- Growth trends

**Requirements:**
- Needs at least 2 years of data
- Automatically groups by year

**Decision Support:**
- "Are we growing year-over-year?"
- "Any seasonal patterns?"
- "How accurate was last year's budget?"

---

## Features

### Interactive Selection
- **Plot 7 (Gross vs Net):** Select specific employees from dropdown
- Multiple selection supported
- Default: First 5 employees

### Data Refresh
- Click "🔄 Refresh Data" to reload latest CSV
- Automatic data validation
- Shows data summary (records, employees, date range)

### AR(2) Forecasting
- **All forecasts use AR(2) model**
- Chosen based on your regression testing (most accurate)
- Provides 6-month ahead predictions
- Includes confidence intervals

## Data Requirements

The dashboard expects these columns in CSV:
- `employee_id`: Unique employee identifier
- `month`: Date in YYYY-MM format
- `salaire_brut`: Gross salary
- `cot_salariale`: Employee contributions
- `cot_patronale`: Employer contributions
- `net_imposable`: Taxable net
- `pas`: Tax withholding
- `net_paye`: Net pay
- `avantages`: Benefits
- `total_cost`: Total employment cost

## Troubleshooting

### "Data file not found"
**Solution:** Generate synthetic data first
```bash
python src/data generation/generate_synthetic_payroll.py
```

### "Need at least 3 months for AR(2)"
**Solution:** Generate more months of data (minimum 3 required for autoregression)

### "Need at least 2 years for YoY"
**Solution:** Plot #10 requires 2+ years. Generate longer time series or skip this plot.

### Plot doesn't show
**Solution:** 
1. Check if data loaded (info panel on left)
2. Click "Refresh Data"
3. Verify CSV exists in outputs/

## Technical Details

### Dependencies
- `pandas`: Data manipulation
- `numpy`: Numerical operations
- `matplotlib`: Plotting
- `statsmodels`: AR(2) forecasting model
- `tkinter`: GUI framework (built into Python)

### AR(2) Model Explained
```python
from statsmodels.tsa.ar_model import AutoReg

# Fit AR(2): uses past 2 months to predict next
model = AutoReg(time_series_data, lags=2)
fitted_model = model.fit()

# Forecast 6 months ahead
forecast = fitted_model.forecast(steps=6)
```

**Why AR(2)?**
- Your testing showed AR(2) most accurate (vs AR1, AR3, AR4, pooled FE)
- Captures short-term dependencies
- Statistically consistent
- Low NRMSE (Normalized Root Mean Squared Error)

## Use Cases by Role

### CEO/CFO
- **Primary plots:** 1, 2, 9
- Focus: Overall cost, forecasting, burn rate
- Questions: "Can we afford growth?", "What's our runway?"

### HR Manager
- **Primary plots:** 3, 5, 6, 7
- Focus: Per-employee costs, equity, compensation
- Questions: "Is pay fair?", "Retention risks?"

### Operations Manager
- **Primary plots:** 4, 8
- Focus: Cost breakdown, headcount
- Questions: "Where's the money going?", "Hiring pace?"

### Board Presentation
- **Primary plots:** 1, 2, 10
- Focus: Trends, forecasts, YoY growth
- Questions: "Are we growing?", "On track?"

## Best Practices

1. **Regular Review:** Check dashboard weekly/monthly
2. **Forecasting:** Use Plot 2 for quarterly planning
3. **Equity:** Use Plot 5 monthly to ensure fair compensation
4. **Retention:** Use Plot 6 to identify key employees
5. **Comparison:** Use Plot 7 to understand employee perspective

## Future Enhancements

Potential additions:
- Export plots to PDF/PowerPoint
- Custom date range selection
- Role-based filtering (by department)
- Scenario analysis ("What if we hire 5 people?")
- Automated email reports

## Support

For issues or questions:
1. Check data exists: `outputs/payroll_long.csv`
2. Verify Python packages installed: `pip list | grep -E "pandas|matplotlib|statsmodels"`
3. Review error messages in terminal
4. Check data format matches expected columns

---

**Remember:** This dashboard uses **synthetic data** for analysis, ensuring privacy while providing statistically valid business insights. The AR(2) forecasting model was validated as the most accurate through your extensive testing.
