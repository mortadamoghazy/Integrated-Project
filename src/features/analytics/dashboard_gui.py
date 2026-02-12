"""
dashboard_gui.py

Interactive dashboard for payroll data visualization.
Provides 10 key plots for CEO/Manager/HR decision making.
All forecasting uses AR(2) with Ridge Regularization and recursive error amplification.
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.dates as mdates
from datetime import datetime
from typing import Optional, List
import warnings
warnings.filterwarnings('ignore')

# Ensure project root is on path
project_root = Path(__file__).parent.parent.parent.parent
sys.path.insert(0, str(project_root))

from sklearn.linear_model import Ridge
from src.features.analytics.employee_clustering import EmployeeClustering


class PayrollDashboard:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Payroll Data Visualization Dashboard")
        self.root.geometry("1400x800")
        
        # FIX_ME: Data file path - Update if your payroll_long.csv is located elsewhere
        # This CSV should contain processed payroll data with columns:
        # employee_id, month, salaire_brut, total_cost, etc.
        # Default location: outputs/payroll_long.csv
        self.df: Optional[pd.DataFrame] = None
        self.data_path = project_root / "outputs" / "payroll_long.csv"

        # Selection controls (populated after load)
        self.emp_var = tk.StringVar(value='All')
        self.month_var = tk.StringVar(value='All')
        
        # Style
        style = ttk.Style()
        style.theme_use('clam')
        
        self.setup_ui()
        self.load_data()
        
    def setup_ui(self):
        """Setup the main UI layout."""
        # Main container
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Left panel - plot selection
        left_panel = ttk.Frame(main_frame, padding=5)
        left_panel.pack(side=tk.LEFT, fill=tk.Y)
        
        # Title
        title = ttk.Label(left_panel, text="📊 Payroll Analytics", 
                         font=("Arial", 16, "bold"))
        title.pack(pady=10)
        
        # Data info
        self.info_label = ttk.Label(left_panel, text="Loading data...", 
                                    font=("Arial", 9))
        self.info_label.pack(pady=5)
        
        # Separator
        ttk.Separator(left_panel, orient='horizontal').pack(fill='x', pady=10)
        
        # Plot selection
        plots_label = ttk.Label(left_panel, text="Select Visualization:", 
                               font=("Arial", 11, "bold"))
        plots_label.pack(pady=5)
        
        # Plot buttons
        self.plot_buttons = [
            ("1. Total Payroll Cost Trend", self.plot_total_cost_trend),
            ("2. Cost Forecasting (Ar(2) Ridge)", self.plot_cost_forecast),
            ("3. Cost per Employee Trend", self.plot_cost_per_employee),
            ("4. Payroll Cost Breakdown", self.plot_cost_breakdown),
            ("5. Employee Cost Distribution", self.plot_cost_distribution),
            ("6. Top 10 Most Expensive Employees", self.plot_top_employees),
            ("7. Gross vs Net Salary", self.plot_gross_vs_net),
            ("8. Headcount Trend", self.plot_headcount_trend),
            ("9. Year-over-Year Comparison", self.plot_yoy_comparison),
            ("10. Employee Clustering (K-Means ML)", self.plot_clustering),
        ]
        
        for text, command in self.plot_buttons:
            btn = ttk.Button(left_panel, text=text, command=command, width=35)
            btn.pack(pady=3, padx=5)

        # Scope selectors: Employee and Month
        ttk.Separator(left_panel, orient='horizontal').pack(fill='x', pady=10)
        scope_label = ttk.Label(left_panel, text="Plot Scope:", font=("Arial", 11, "bold"))
        scope_label.pack(pady=(5, 2))

        ttk.Label(left_panel, text="Employee:", font=("Arial", 9)).pack(anchor='w', padx=5)
        self.emp_cb = ttk.Combobox(left_panel, textvariable=self.emp_var, state='readonly')
        self.emp_cb['values'] = ['All']
        self.emp_cb.current(0)
        self.emp_cb.pack(fill='x', padx=5, pady=2)

        ttk.Label(left_panel, text="Month:", font=("Arial", 9)).pack(anchor='w', padx=5)
        self.month_cb = ttk.Combobox(left_panel, textvariable=self.month_var, state='readonly')
        self.month_cb['values'] = ['All']
        self.month_cb.current(0)
        self.month_cb.pack(fill='x', padx=5, pady=(2, 6))
        
        # Refresh data button
        ttk.Separator(left_panel, orient='horizontal').pack(fill='x', pady=10)
        refresh_btn = ttk.Button(left_panel, text="🔄 Refresh Data", 
                                command=self.load_data)
        refresh_btn.pack(pady=5)
        
        # Right panel - plot display
        right_panel = ttk.Frame(main_frame)
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(10, 0))
        
        # Canvas for matplotlib
        self.plot_frame = ttk.Frame(right_panel)
        self.plot_frame.pack(fill=tk.BOTH, expand=True)
        
        # Initial message
        welcome = ttk.Label(self.plot_frame, 
                           text="Welcome to Payroll Analytics Dashboard\n\n"
                                "Select a visualization from the left panel",
                           font=("Arial", 14))
        welcome.place(relx=0.5, rely=0.5, anchor='center')
        
    def load_data(self):
        """Load payroll data from CSV."""
        try:
            if not self.data_path.exists():
                self.info_label.config(text="❌ Data file not found")
                messagebox.showerror("Data Error", 
                                   f"Payroll data not found at:\n{self.data_path}\n\n"
                                   "Please run data generation first.")
                return
            
            self.df = pd.read_csv(self.data_path)
            self.df['month'] = pd.to_datetime(self.df['month'])
            self.df = self.df.sort_values('month')
            
            # Calculate derived metrics
            if 'total_cost' not in self.df.columns:
                # Total cost = gross salary + employer contributions + benefits
                employer_contrib_cols = [c for c in self.df.columns if 'cot_patronale' in c.lower()]
                if employer_contrib_cols:
                    self.df['total_cost'] = (self.df['salaire_brut'] + 
                                            self.df[employer_contrib_cols[0]] + 
                                            self.df.get('avantages', 0))
                else:
                    self.df['total_cost'] = self.df['salaire_brut'] * 1.3  # Approximate
            
            employees = self.df['employee_id'].nunique()
            months = self.df['month'].nunique()
            date_range = f"{self.df['month'].min().strftime('%Y-%m')} to {self.df['month'].max().strftime('%Y-%m')}"
            
            self.info_label.config(
                text=f"✅ Loaded: {len(self.df)} records\n"
                     f"👥 {employees} employees\n"
                     f"📅 {months} months\n"
                     f"📆 {date_range}"
            )
            
            messagebox.showinfo("Data Loaded", 
                              f"Successfully loaded payroll data\n\n"
                              f"Records: {len(self.df)}\n"
                              f"Employees: {employees}\n"
                              f"Period: {date_range}")

            # Populate employee and month selectors
            emp_list = ['All'] + sorted(self.df['employee_id'].unique())
            months = ['All'] + sorted(self.df['month'].dt.strftime('%Y-%m').unique())
            try:
                self.emp_cb['values'] = emp_list
                self.month_cb['values'] = months
            except Exception:
                pass
            
        except Exception as e:
            self.info_label.config(text="❌ Error loading data")
            messagebox.showerror("Load Error", f"Failed to load data:\n{e}")
    
    def clear_plot(self):
        """Clear the current plot."""
        for widget in self.plot_frame.winfo_children():
            widget.destroy()

    def _apply_scope_filters(self, df: pd.DataFrame) -> pd.DataFrame:
        """Apply selected employee/month scope filters to a dataframe."""
        df2 = df.copy()
        emp_sel = self.emp_var.get() if hasattr(self, 'emp_var') else 'All'
        month_sel = self.month_var.get() if hasattr(self, 'month_var') else 'All'

        if emp_sel and emp_sel != 'All':
            # Employee IDs may be numeric in the dataframe but the combobox
            # stores strings; compare using string form for robustness.
            df2 = df2[df2['employee_id'].astype(str) == str(emp_sel)]

        if month_sel and month_sel != 'All':
            try:
                # month_sel is in 'YYYY-MM' format; match using string comparison
                df2 = df2[df2['month'].dt.strftime('%Y-%m') == month_sel]
            except Exception:
                pass

        return df2

    def _set_scope_controls(self, emp_enabled: bool = True, month_enabled: bool = True):
        """Enable or disable employee/month scope controls depending on plot context.

        When disabling a control, reset its selection to 'All' to avoid filtering by an
        invalid value.
        """
        try:
            if emp_enabled:
                self.emp_cb.config(state='readonly')
            else:
                self.emp_cb.config(state='disabled')
                self.emp_var.set('All')

            if month_enabled:
                self.month_cb.config(state='readonly')
            else:
                self.month_cb.config(state='disabled')
                self.month_var.set('All')
        except Exception:
            # If controls not yet created or any error, silently continue
            pass
    
    def plot_total_cost_trend(self):
        """Plot 1: Total payroll cost trend over time."""
        if self.df is None:
            messagebox.showwarning("No Data", "Please load data first")
            return
        
        # enable both selectors for this plot
        self._set_scope_controls(emp_enabled=True, month_enabled=True)

        self.clear_plot()
        
        # Apply scope filters (employee/month)
        df_used = self._apply_scope_filters(self.df)

        if df_used.empty:
            messagebox.showinfo("No Data", "No data available for the selected employee/month scope.")
            return

        # Calculate monthly totals
        monthly_cost = df_used.groupby('month')['total_cost'].sum().reset_index()
        
        # Create plot
        fig, ax = plt.subplots(figsize=(12, 6))
        ax.plot(monthly_cost['month'], monthly_cost['total_cost'], 
               marker='o', linewidth=2, markersize=8, color='#2E86AB')
        ax.fill_between(monthly_cost['month'], monthly_cost['total_cost'], 
                        alpha=0.3, color='#2E86AB')
        
        ax.set_title('Total Payroll Cost Trend', fontsize=16, fontweight='bold', pad=20)
        ax.set_xlabel('Month', fontsize=12)
        ax.set_ylabel('Total Cost (€)', fontsize=12)
        ax.grid(True, alpha=0.3)
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m'))
        plt.xticks(rotation=45, ha='right')
        
        # Add value labels (reduce density to avoid overlap)
        n_points = len(monthly_cost)
        max_labels = 12
        step = max(1, int(np.ceil(n_points / max_labels)))
        for i, (x, y) in enumerate(zip(monthly_cost['month'], monthly_cost['total_cost'])):
            # only label a subset of points (roughly `max_labels`), always include last
            if (i % step == 0) or (i == n_points - 1):
                # alternate upward/downward offsets to reduce collisions
                offset = 10 if ((i // step) % 2 == 0) else -12
                va = 'bottom' if offset > 0 else 'top'
                ax.annotate(f'€{y:,.0f}', (x, y), textcoords="offset points",
                           xytext=(0, offset), ha='center', va=va, fontsize=8,
                           bbox=dict(boxstyle='round,pad=0.2', fc='white', alpha=0.6, linewidth=0))

        # Add a little vertical margin so top labels don't get clipped
        ymin, ymax = ax.get_ylim()
        yrange = ymax - ymin
        ax.set_ylim(ymin - 0.02 * yrange, ymax + 0.06 * yrange)
        
        plt.tight_layout()
        
        # Embed in tkinter
        canvas = FigureCanvasTkAgg(fig, self.plot_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
    
    def plot_cost_forecast(self):
        """Plot 2: Cost forecasting using AR(2) with Ridge Regularization and recursive error amplification."""
        if self.df is None:
            messagebox.showwarning("No Data", "Please load data first")
            return
        
        # allow both emp/month scope for forecasting
        self._set_scope_controls(emp_enabled=True, month_enabled=True)

        self.clear_plot()

        # Apply scope filters (employee/month)
        df_used = self._apply_scope_filters(self.df)

        if df_used.empty:
            messagebox.showinfo("No Data", "No data available for the selected employee/month scope.")
            return

        # Calculate monthly totals
        monthly_cost = df_used.groupby('month')['total_cost'].sum().reset_index()
        monthly_cost = monthly_cost.sort_values('month')

        # Fit Ridge Regression with panel features
        try:
            # Prepare features: time index, employee fixed effects, lagged values
            monthly_cost['time_idx'] = range(len(monthly_cost))
            monthly_cost['month_num'] = monthly_cost['month'].dt.month
            
            # Create lagged features
            for lag in [1, 2, 3]:
                monthly_cost[f'lag_{lag}'] = monthly_cost['total_cost'].shift(lag)
            
            # Drop NaN rows from lagging
            train_data = monthly_cost.dropna()
            
            if len(train_data) < 5:
                messagebox.showerror("Insufficient Data", "Need at least 5 months of data for AR(2)+Ridge forecasting")
                return
            
            # Prepare training features
            feature_cols = ['time_idx', 'month_num', 'lag_1', 'lag_2', 'lag_3']
            X_train = train_data[feature_cols].values
            y_train = train_data['total_cost'].values
            
            # Fit Ridge model
            model = Ridge(alpha=10.0)  # Using optimal alpha from AR(2)+Ridge evaluation
            model.fit(X_train, y_train)
            
            # Calculate RMSE for error propagation
            y_pred_train = model.predict(X_train)
            rmse = np.sqrt(np.mean((y_train - y_pred_train) ** 2))
            
            # Recursive forecasting 6 months ahead
            forecast_steps = 6
            forecast = []
            last_values = list(monthly_cost['total_cost'].values[-3:])  # Last 3 values for lags
            last_time = monthly_cost['time_idx'].iloc[-1]
            last_month = monthly_cost['month'].iloc[-1]
            
            for step in range(1, forecast_steps + 1):
                next_time = last_time + step
                next_month = (last_month + pd.DateOffset(months=step)).month
                
                # Create features for next step
                X_next = np.array([[
                    next_time,
                    next_month,
                    last_values[-1],  # lag_1
                    last_values[-2],  # lag_2
                    last_values[-3]   # lag_3
                ]])
                
                # Predict
                pred = model.predict(X_next)[0]
                forecast.append(pred)
                
                # Update last_values for next iteration (recursive)
                last_values.append(pred)
                last_values.pop(0)
            
            forecast = np.array(forecast)
            
            # Create future dates
            last_date = monthly_cost['month'].iloc[-1]
            future_dates = pd.date_range(start=last_date, periods=forecast_steps+1, freq='MS')[1:]
            
            # Recursive error amplification: CI_t = ±RMSE × √t × 1.96
            conf_intervals = []
            for t in range(1, forecast_steps + 1):
                error_margin = rmse * np.sqrt(t) * 1.96
                conf_intervals.append(error_margin)
            
            conf_intervals = np.array(conf_intervals)
            conf_upper = forecast + conf_intervals
            conf_lower = forecast - conf_intervals
            
            # Create plot
            fig, ax = plt.subplots(figsize=(12, 6))
            
            # Historical data
            ax.plot(monthly_cost['month'], monthly_cost['total_cost'], marker='o', label='Historical',
                    linewidth=2, markersize=8, color='#2E86AB')
            
            # Forecast
            ax.plot(future_dates, forecast, marker='s', label='Forecast (AR(2)+Ridge)',
                    linewidth=2, markersize=8, color='#A23B72', linestyle='--')
            
            # Confidence intervals with recursive error amplification
            ax.fill_between(future_dates, conf_lower, conf_upper, 
                           alpha=0.2, color='#A23B72')
            
            ax.set_title('Payroll Cost Forecasting', 
                        fontsize=16, fontweight='bold', pad=20)
            ax.set_xlabel('Month', fontsize=12)
            ax.set_ylabel('Total Cost (€)', fontsize=12)
            ax.legend(fontsize=10)
            ax.grid(True, alpha=0.3)
            ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m'))
            plt.xticks(rotation=45, ha='right')
            
            plt.tight_layout()
            
            # Embed in tkinter
            canvas = FigureCanvasTkAgg(fig, self.plot_frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
            
        except Exception as e:
            messagebox.showerror("Forecast Error", 
                               f"Failed to generate forecast:\n{e}\n\n"
                               "Need at least 5 months of data for AR(2)+Ridge forecasting")
    
    def plot_cost_per_employee(self):
        """Plot 3: Average cost per employee over time."""
        if self.df is None:
            messagebox.showwarning("No Data", "Please load data first")
            return
        
        # this plot can be scoped by employee/month
        self._set_scope_controls(emp_enabled=True, month_enabled=True)

        self.clear_plot()
        
        # Apply scope filters (employee/month)
        df_used = self._apply_scope_filters(self.df)

        if df_used.empty:
            messagebox.showinfo("No Data", "No data available for the selected employee/month scope.")
            return

        # Calculate monthly averages
        monthly_avg = df_used.groupby('month').agg({
            'total_cost': 'mean',
            'employee_id': 'count'
        }).reset_index()
        monthly_avg.columns = ['month', 'avg_cost', 'employee_count']
        
        # Create plot
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 8))
        
        # Average cost per employee
        ax1.plot(monthly_avg['month'], monthly_avg['avg_cost'], 
                marker='o', linewidth=2, markersize=8, color='#F18F01')
        ax1.set_title('Average Cost per Employee', fontsize=14, fontweight='bold')
        ax1.set_ylabel('Average Cost (€)', fontsize=11)
        ax1.grid(True, alpha=0.3)
        ax1.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m'))
        
        # Employee count
        ax2.bar(monthly_avg['month'], monthly_avg['employee_count'], 
               color='#C73E1D', alpha=0.7)
        ax2.set_title('Number of Employees', fontsize=14, fontweight='bold')
        ax2.set_xlabel('Month', fontsize=11)
        ax2.set_ylabel('Employee Count', fontsize=11)
        ax2.grid(True, alpha=0.3, axis='y')
        ax2.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m'))
        
        plt.setp(ax1.xaxis.get_majorticklabels(), rotation=45, ha='right')
        plt.setp(ax2.xaxis.get_majorticklabels(), rotation=45, ha='right')
        
        plt.tight_layout()
        
        canvas = FigureCanvasTkAgg(fig, self.plot_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
    
    def plot_cost_breakdown(self):
        """Plot 4: Payroll cost breakdown - stacked area chart."""
        if self.df is None:
            messagebox.showwarning("No Data", "Please load data first")
            return
        
        # allow scoping for breakdown
        self._set_scope_controls(emp_enabled=True, month_enabled=True)

        self.clear_plot()

        # Apply scope filters (employee/month)
        df_used = self._apply_scope_filters(self.df)

        if df_used.empty:
            messagebox.showinfo("No Data", "No data available for the selected employee/month scope.")
            return

        # Calculate monthly component totals
        monthly_breakdown = df_used.groupby('month').agg({
            'salaire_brut': 'sum',
            'cot_patronale': 'sum',
            'avantages': 'sum'
        }).reset_index()
        
        # Create plot
        fig, ax = plt.subplots(figsize=(12, 6))
        
        ax.fill_between(monthly_breakdown['month'], 0, 
                       monthly_breakdown['salaire_brut'],
                       label='Gross Salary', alpha=0.8, color='#2E86AB')
        ax.fill_between(monthly_breakdown['month'], 
                       monthly_breakdown['salaire_brut'],
                       monthly_breakdown['salaire_brut'] + monthly_breakdown['cot_patronale'],
                       label='Employer Contributions', alpha=0.8, color='#A23B72')
        ax.fill_between(monthly_breakdown['month'],
                       monthly_breakdown['salaire_brut'] + monthly_breakdown['cot_patronale'],
                       monthly_breakdown['salaire_brut'] + monthly_breakdown['cot_patronale'] + monthly_breakdown['avantages'],
                       label='Benefits', alpha=0.8, color='#F18F01')
        
        ax.set_title('Payroll Cost Breakdown', fontsize=16, fontweight='bold', pad=20)
        ax.set_xlabel('Month', fontsize=12)
        ax.set_ylabel('Total Cost (€)', fontsize=12)
        ax.legend(loc='upper left', fontsize=10)
        ax.grid(True, alpha=0.3, axis='y')
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m'))
        plt.xticks(rotation=45, ha='right')
        
        plt.tight_layout()
        
        canvas = FigureCanvasTkAgg(fig, self.plot_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
    
    def plot_cost_distribution(self):
        """Plot 5: Employee cost distribution - box plot and histogram."""
        if self.df is None:
            messagebox.showwarning("No Data", "Please load data first")
            return
        
        # This plot supports scoping by employee and month
        self._set_scope_controls(emp_enabled=True, month_enabled=True)

        self.clear_plot()

        # Apply scope filters (employee/month)
        df_used = self._apply_scope_filters(self.df)

        if df_used.empty:
            messagebox.showinfo("No Data", "No data available for the selected employee/month scope.")
            return

        # Determine month to use: if month selected, df_used already filtered to that month
        if self.month_var.get() and self.month_var.get() != 'All':
            latest_data = df_used.copy()
            latest_month = latest_data['month'].max() if not latest_data.empty else None
        else:
            latest_month = df_used['month'].max()
            latest_data = df_used[df_used['month'] == latest_month]
        
        # Create single histogram plot (removed boxplot)
        fig, ax = plt.subplots(figsize=(12, 6))

        ax.hist(latest_data['total_cost'], bins=20, color='#A23B72', 
            alpha=0.7, edgecolor='black')
        mean_val = latest_data['total_cost'].mean()
        med_val = latest_data['total_cost'].median()
        ax.axvline(mean_val, color='red', linestyle='--', linewidth=2, label=f"Mean: €{mean_val:,.0f}")
        ax.axvline(med_val, color='green', linestyle='--', linewidth=2, label=f"Median: €{med_val:,.0f}")
        ax.set_title('Employee Cost Distribution (Latest Month)', fontsize=14, fontweight='bold')
        ax.set_xlabel('Total Cost (€)', fontsize=11)
        ax.set_ylabel('Number of Employees', fontsize=11)
        ax.legend(fontsize=9)
        ax.grid(True, alpha=0.3, axis='y')

        plt.tight_layout()

        canvas = FigureCanvasTkAgg(fig, self.plot_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
    
    def plot_top_employees(self):
        """Plot 6: Top 10 most expensive employees."""
        if self.df is None:
            messagebox.showwarning("No Data", "Please load data first")
            return
        
        # For Top 10 employees we do NOT allow selecting specific employees
        # (that would be contradictory). Allow selecting month only.
        self._set_scope_controls(emp_enabled=False, month_enabled=True)

        self.clear_plot()

        # Apply scope filters (employee/month)
        df_used = self._apply_scope_filters(self.df)

        if df_used.empty:
            messagebox.showinfo("No Data", "No data available for the selected employee/month scope.")
            return

        # If the user selected 'All' months, compute per-employee average across
        # all months and pick the top 10 by that average. Otherwise pick top 10
        # within the selected month.
        if self.month_var.get() and self.month_var.get() != 'All':
            try:
                sel_month = pd.to_datetime(self.month_var.get() + '-01')
            except Exception:
                sel_month = df_used['month'].max()
            month_data = df_used[df_used['month'] == sel_month].copy()
            if month_data.empty:
                messagebox.showinfo("No Data", "No records for the selected month.")
                return
            emp_tot = month_data.groupby('employee_id')['total_cost'].sum().reset_index()
            top10 = emp_tot.nlargest(10, 'total_cost').sort_values('total_cost')
            title_period = sel_month.strftime('%Y-%m')
        else:
            # average across all months for each employee
            emp_avg = df_used.groupby('employee_id')['total_cost'].mean().reset_index()
            top10 = emp_avg.nlargest(10, 'total_cost').sort_values('total_cost')
            title_period = 'All months (avg)'
        
        # Create plot
        fig, ax = plt.subplots(figsize=(12, 8))
        
        bars = ax.barh(range(len(top10)), top10['total_cost'], color='#C73E1D', alpha=0.7)
        ax.set_yticks(range(len(top10)))
        ax.set_yticklabels(top10['employee_id'])
        ax.set_xlabel('Total Cost (€)', fontsize=12)
        ax.set_ylabel('Employee ID', fontsize=12)
        ax.set_title(f'Top 10 Most Expensive Employees ({title_period})', 
                fontsize=16, fontweight='bold', pad=20)
        ax.grid(True, alpha=0.3, axis='x')
        
        # Add value labels
        for i, (idx, row) in enumerate(top10.iterrows()):
            ax.text(row['total_cost'], i, f"  €{row['total_cost']:,.0f}", 
                   va='center', fontsize=9)
        
        plt.tight_layout()
        
        canvas = FigureCanvasTkAgg(fig, self.plot_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
    
    def plot_gross_vs_net(self):
        """Plot 7: Gross vs Net salary with employee selection."""
        if self.df is None:
            messagebox.showwarning("No Data", "Please load data first")
            return
        
        # For Gross vs Net we use an internal employee selector dialog, disable
        # the global employee combobox to avoid confusion.
        self._set_scope_controls(emp_enabled=False, month_enabled=True)

        # Employee selection dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Select Employees")
        dialog.geometry("400x500")
        
        ttk.Label(dialog, text="Select employees to compare:", 
                 font=("Arial", 11, "bold")).pack(pady=10)
        
        # Listbox with scrollbar
        frame = ttk.Frame(dialog)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        scrollbar = ttk.Scrollbar(frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        listbox = tk.Listbox(frame, selectmode=tk.MULTIPLE, 
                            yscrollcommand=scrollbar.set, height=15)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)
        
        # Populate with employees
        employees = sorted(self.df['employee_id'].unique())
        for emp in employees:
            listbox.insert(tk.END, emp)
        
        # Select first 5 by default
        for i in range(min(5, len(employees))):
            listbox.selection_set(i)
        
        def plot_selected():
            selected_indices = listbox.curselection()
            if not selected_indices:
                messagebox.showwarning("No Selection", "Please select at least one employee")
                return
            
            selected_employees = [employees[i] for i in selected_indices]
            dialog.destroy()
            self._plot_gross_vs_net_chart(selected_employees)
        
        ttk.Button(dialog, text="Plot", command=plot_selected).pack(pady=10)
        
    def _plot_gross_vs_net_chart(self, employees: List[str]):
        """Create the actual gross vs net plot."""
        self.clear_plot()
        
        # Filter data
        emp_data = self.df[self.df['employee_id'].isin(employees)].copy()
        
        # Create plot
        fig, ax = plt.subplots(figsize=(12, 6))
        
        for emp in employees:
            emp_subset = emp_data[emp_data['employee_id'] == emp].sort_values('month')
            ax.plot(emp_subset['month'], emp_subset['salaire_brut'], 
                   marker='o', label=f'{emp} (Gross)', linewidth=2)
            ax.plot(emp_subset['month'], emp_subset['net_paye'], 
                   marker='s', label=f'{emp} (Net)', linewidth=2, linestyle='--')
        
        ax.set_title('Gross vs Net Salary Comparison', 
                    fontsize=16, fontweight='bold', pad=20)
        ax.set_xlabel('Month', fontsize=12)
        ax.set_ylabel('Salary (€)', fontsize=12)
        ax.legend(bbox_to_anchor=(1.05, 1), loc='upper left', fontsize=9)
        ax.grid(True, alpha=0.3)
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m'))
        plt.xticks(rotation=45, ha='right')
        
        plt.tight_layout()
        
        canvas = FigureCanvasTkAgg(fig, self.plot_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
    
    def plot_headcount_trend(self):
        """Plot 8: Employee headcount trend."""
        if self.df is None:
            messagebox.showwarning("No Data", "Please load data first")
            return
        
        self.clear_plot()
        
        # Calculate monthly headcount
        headcount = self.df.groupby('month')['employee_id'].nunique().reset_index()
        headcount.columns = ['month', 'headcount']
        
        # Create plot
        fig, ax = plt.subplots(figsize=(12, 6))
        
        ax.plot(headcount['month'], headcount['headcount'], 
               marker='o', linewidth=3, markersize=10, color='#2E86AB')
        ax.fill_between(headcount['month'], headcount['headcount'], 
                       alpha=0.3, color='#2E86AB')
        
        # Add value labels (reduce density to avoid overlap)
        n_points = len(headcount)
        max_labels = 12
        step = max(1, int(np.ceil(n_points / max_labels)))
        for i, (x, y) in enumerate(zip(headcount['month'], headcount['headcount'])):
            if (i % step == 0) or (i == n_points - 1):
                offset = 8 if ((i // step) % 2 == 0) else -10
                va = 'bottom' if offset > 0 else 'top'
                ax.annotate(f'{int(y)}', (x, y), textcoords="offset points",
                           xytext=(0, offset), ha='center', va=va, fontsize=10, fontweight='bold',
                           bbox=dict(boxstyle='round,pad=0.2', fc='white', alpha=0.6, linewidth=0))

        # Small y-margin so labels don't overlap axes
        ymin, ymax = ax.get_ylim()
        yrange = ymax - ymin
        ax.set_ylim(max(0, ymin - 0.02 * yrange), ymax + 0.05 * yrange)
        
        ax.set_title('Employee Headcount Trend', fontsize=16, fontweight='bold', pad=20)
        ax.set_xlabel('Month', fontsize=12)
        ax.set_ylabel('Number of Employees', fontsize=12)
        ax.grid(True, alpha=0.3)
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m'))
        plt.xticks(rotation=45, ha='right')
        
        # Set y-axis to integer
        ax.yaxis.set_major_locator(plt.MaxNLocator(integer=True))
        
        plt.tight_layout()
        
        canvas = FigureCanvasTkAgg(fig, self.plot_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
    
    # Plot 9 (Monthly Burn Rate Dashboard) removed per user request.
    
    def plot_yoy_comparison(self):
        """Plot 10: Year-over-Year comparison."""
        if self.df is None:
            messagebox.showwarning("No Data", "Please load data first")
            return
        
        self.clear_plot()
        
        # Extract year and month
        df_yoy = self.df.copy()
        df_yoy['year'] = df_yoy['month'].dt.year
        df_yoy['month_num'] = df_yoy['month'].dt.month
        
        # Get unique years
        years = sorted(df_yoy['year'].unique())
        
        if len(years) < 2:
            messagebox.showinfo("Insufficient Data", 
                              "Need at least 2 years of data for YoY comparison")
            return
        
        # Create plot
        fig, ax = plt.subplots(figsize=(12, 6))
        
        for year in years:
            year_data = df_yoy[df_yoy['year'] == year].groupby('month_num')['total_cost'].sum()
            ax.plot(year_data.index, year_data.values, 
                   marker='o', linewidth=2, markersize=8, label=str(year))
        
        ax.set_title('Year-over-Year Comparison', fontsize=16, fontweight='bold', pad=20)
        ax.set_xlabel('Month', fontsize=12)
        ax.set_ylabel('Total Cost (€)', fontsize=12)
        ax.set_xticks(range(1, 13))
        ax.set_xticklabels(['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                           'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'])
        ax.legend(title='Year', fontsize=10)
        ax.grid(True, alpha=0.3)
        
        plt.tight_layout()
        
        canvas = FigureCanvasTkAgg(fig, self.plot_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
    
    def plot_clustering(self):
        """Plot 10: Employee clustering using K-Means ML."""
        if self.df is None:
            messagebox.showwarning("No Data", "Please load data first")
            return
        
        # Clustering works on all employees, disable individual employee selector
        self._set_scope_controls(emp_enabled=False, month_enabled=False)
        
        self.clear_plot()
        
        # Run clustering immediately with K=3 and show simple HR view
        try:
            clustering = EmployeeClustering(self.df, n_clusters=3)
            clustering.fit()
            
            # Always use simplified HR view in GUI
            fig = clustering.plot_hr_simple()
            
            # Show cluster summary with employee assignments
            profiles = clustering.get_cluster_profiles()
            sorted_profiles = profiles.sort_values('avg_total_cost')
            tier_names = ['Low-Cost (Entry Level)', 'Mid-Cost (Experienced)', 'High-Cost (Senior)']
            
            summary = "Employee Cost Clustering Results\n"
            summary += "=" * 50 + "\n\n"
            
            for idx, (_, row) in enumerate(sorted_profiles.iterrows()):
                tier_name = tier_names[idx] if idx < len(tier_names) else f"Tier {idx+1}"
                summary += f"{tier_name}\n"
                summary += f"  • {int(row['size'])} employees\n"
                summary += f"  • €{row['avg_total_cost']:,.0f} per employee/month\n"
                summary += f"  • €{row['avg_total_cost'] * row['size']:,.0f} total/month\n"
                
                # Add employee list
                emp_list = ', '.join(str(emp) for emp in row['employee_ids'])
                summary += f"  • Employees: {emp_list}\n\n"
            
            total_monthly = (profiles['avg_total_cost'] * profiles['size']).sum()
            summary += f"Total Monthly Payroll: €{total_monthly:,.0f}\n"
            summary += f"Clustering Quality: {clustering.silhouette:.2f}/1.00"
            
            messagebox.showinfo("Clustering Complete", summary)
            
            # Display figure
            canvas = FigureCanvasTkAgg(fig, self.plot_frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
            
        except Exception as e:
            messagebox.showerror("Clustering Error", 
                               f"Failed to perform clustering:\n{str(e)}")
    
    def run(self):
        """Run the dashboard."""
        self.root.mainloop()


def main():
    """Main entry point."""
    dashboard = PayrollDashboard()
    dashboard.run()


if __name__ == "__main__":
    main()
