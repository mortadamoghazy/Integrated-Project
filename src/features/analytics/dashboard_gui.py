"""
dashboard_gui.py

Interactive dashboard for payroll data visualization.
Provides 10 key plots for CEO/Manager/HR decision making.
All forecasting uses AR(2) model (autoregression order 2).
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

from statsmodels.tsa.ar_model import AutoReg


class PayrollDashboard:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Payroll Data Visualization Dashboard")
        self.root.geometry("1400x800")
        
        # Data
        self.df: Optional[pd.DataFrame] = None
        self.data_path = project_root / "outputs" / "payroll_long.csv"
        
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
            ("2. Cost Forecasting (AR2)", self.plot_cost_forecast),
            ("3. Cost per Employee Trend", self.plot_cost_per_employee),
            ("4. Payroll Cost Breakdown", self.plot_cost_breakdown),
            ("5. Employee Cost Distribution", self.plot_cost_distribution),
            ("6. Top 10 Most Expensive Employees", self.plot_top_employees),
            ("7. Gross vs Net Salary", self.plot_gross_vs_net),
            ("8. Headcount Trend", self.plot_headcount_trend),
            ("9. Monthly Burn Rate Dashboard", self.plot_burn_rate),
            ("10. Year-over-Year Comparison", self.plot_yoy_comparison),
        ]
        
        for text, command in self.plot_buttons:
            btn = ttk.Button(left_panel, text=text, command=command, width=35)
            btn.pack(pady=3, padx=5)
        
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
            
        except Exception as e:
            self.info_label.config(text="❌ Error loading data")
            messagebox.showerror("Load Error", f"Failed to load data:\n{e}")
    
    def clear_plot(self):
        """Clear the current plot."""
        for widget in self.plot_frame.winfo_children():
            widget.destroy()
    
    def plot_total_cost_trend(self):
        """Plot 1: Total payroll cost trend over time."""
        if self.df is None:
            messagebox.showwarning("No Data", "Please load data first")
            return
        
        self.clear_plot()
        
        # Calculate monthly totals
        monthly_cost = self.df.groupby('month')['total_cost'].sum().reset_index()
        
        # Create plot
        fig, ax = plt.subplots(figsize=(12, 6))
        ax.plot(monthly_cost['month'], monthly_cost['total_cost'], 
               marker='o', linewidth=2, markersize=8, color='#2E86AB')
        ax.fill_between(monthly_cost['month'], monthly_cost['total_cost'], 
                        alpha=0.3, color='#2E86AB')
        
        ax.set_title('Total Payroll Cost Trend', fontsize=16, fontweight='bold', pad=20)
        ax.set_xlabel('Month', fontsize=12)
        ax.set_ylabel('Total Cost ($)', fontsize=12)
        ax.grid(True, alpha=0.3)
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m'))
        plt.xticks(rotation=45, ha='right')
        
        # Add value labels
        for x, y in zip(monthly_cost['month'], monthly_cost['total_cost']):
            ax.annotate(f'${y:,.0f}', (x, y), textcoords="offset points", 
                       xytext=(0, 10), ha='center', fontsize=8)
        
        plt.tight_layout()
        
        # Embed in tkinter
        canvas = FigureCanvasTkAgg(fig, self.plot_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
    
    def plot_cost_forecast(self):
        """Plot 2: Cost forecasting using AR(2) model."""
        if self.df is None:
            messagebox.showwarning("No Data", "Please load data first")
            return
        
        self.clear_plot()
        
        # Calculate monthly totals
        monthly_cost = self.df.groupby('month')['total_cost'].sum().reset_index()
        monthly_cost = monthly_cost.sort_values('month')
        
        # Fit AR(2) model
        try:
            ts_data = monthly_cost['total_cost'].values
            model = AutoReg(ts_data, lags=2)
            fitted_model = model.fit()
            
            # Forecast 6 months ahead
            forecast_steps = 6
            forecast = fitted_model.forecast(steps=forecast_steps)
            
            # Create future dates
            last_date = monthly_cost['month'].iloc[-1]
            future_dates = pd.date_range(start=last_date, periods=forecast_steps+1, 
                                        freq='MS')[1:]
            
            # Create plot
            fig, ax = plt.subplots(figsize=(12, 6))
            
            # Historical data
            ax.plot(monthly_cost['month'], monthly_cost['total_cost'], 
                   marker='o', label='Historical', linewidth=2, markersize=8, 
                   color='#2E86AB')
            
            # Forecast
            ax.plot(future_dates, forecast, marker='s', label='Forecast (AR2)', 
                   linewidth=2, markersize=8, color='#A23B72', linestyle='--')
            
            # Confidence interval (simple estimation)
            std_error = np.std(fitted_model.resid)
            conf_upper = forecast + 1.96 * std_error
            conf_lower = forecast - 1.96 * std_error
            ax.fill_between(future_dates, conf_lower, conf_upper, 
                           alpha=0.2, color='#A23B72', label='95% CI')
            
            ax.set_title('Payroll Cost Forecasting (AR2 Model)', 
                        fontsize=16, fontweight='bold', pad=20)
            ax.set_xlabel('Month', fontsize=12)
            ax.set_ylabel('Total Cost ($)', fontsize=12)
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
                               "Need at least 3 months of data for AR(2)")
    
    def plot_cost_per_employee(self):
        """Plot 3: Average cost per employee over time."""
        if self.df is None:
            messagebox.showwarning("No Data", "Please load data first")
            return
        
        self.clear_plot()
        
        # Calculate monthly averages
        monthly_avg = self.df.groupby('month').agg({
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
        ax1.set_ylabel('Average Cost ($)', fontsize=11)
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
        
        self.clear_plot()
        
        # Calculate monthly component totals
        monthly_breakdown = self.df.groupby('month').agg({
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
        ax.set_ylabel('Total Cost ($)', fontsize=12)
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
        
        self.clear_plot()
        
        # Get latest month data
        latest_month = self.df['month'].max()
        latest_data = self.df[self.df['month'] == latest_month]
        
        # Create plot
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))
        
        # Box plot
        bp = ax1.boxplot([latest_data['total_cost']], vert=True, patch_artist=True,
                         labels=[latest_month.strftime('%Y-%m')])
        bp['boxes'][0].set_facecolor('#2E86AB')
        bp['boxes'][0].set_alpha(0.7)
        ax1.set_title('Employee Cost Distribution (Latest Month)', 
                     fontsize=14, fontweight='bold')
        ax1.set_ylabel('Total Cost ($)', fontsize=11)
        ax1.grid(True, alpha=0.3, axis='y')
        
        # Histogram
        ax2.hist(latest_data['total_cost'], bins=20, color='#A23B72', 
                alpha=0.7, edgecolor='black')
        ax2.axvline(latest_data['total_cost'].mean(), color='red', 
                   linestyle='--', linewidth=2, label=f"Mean: ${latest_data['total_cost'].mean():,.0f}")
        ax2.axvline(latest_data['total_cost'].median(), color='green', 
                   linestyle='--', linewidth=2, label=f"Median: ${latest_data['total_cost'].median():,.0f}")
        ax2.set_title('Cost Distribution Histogram', fontsize=14, fontweight='bold')
        ax2.set_xlabel('Total Cost ($)', fontsize=11)
        ax2.set_ylabel('Number of Employees', fontsize=11)
        ax2.legend(fontsize=9)
        ax2.grid(True, alpha=0.3, axis='y')
        
        plt.tight_layout()
        
        canvas = FigureCanvasTkAgg(fig, self.plot_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
    
    def plot_top_employees(self):
        """Plot 6: Top 10 most expensive employees."""
        if self.df is None:
            messagebox.showwarning("No Data", "Please load data first")
            return
        
        self.clear_plot()
        
        # Get latest month data
        latest_month = self.df['month'].max()
        latest_data = self.df[self.df['month'] == latest_month].copy()
        
        # Top 10
        top10 = latest_data.nlargest(10, 'total_cost').sort_values('total_cost')
        
        # Create plot
        fig, ax = plt.subplots(figsize=(12, 8))
        
        bars = ax.barh(range(len(top10)), top10['total_cost'], color='#C73E1D', alpha=0.7)
        ax.set_yticks(range(len(top10)))
        ax.set_yticklabels(top10['employee_id'])
        ax.set_xlabel('Total Cost ($)', fontsize=12)
        ax.set_ylabel('Employee ID', fontsize=12)
        ax.set_title(f'Top 10 Most Expensive Employees ({latest_month.strftime("%Y-%m")})', 
                    fontsize=16, fontweight='bold', pad=20)
        ax.grid(True, alpha=0.3, axis='x')
        
        # Add value labels
        for i, (idx, row) in enumerate(top10.iterrows()):
            ax.text(row['total_cost'], i, f"  ${row['total_cost']:,.0f}", 
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
        ax.set_ylabel('Salary ($)', fontsize=12)
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
        
        # Add value labels
        for x, y in zip(headcount['month'], headcount['headcount']):
            ax.annotate(f'{int(y)}', (x, y), textcoords="offset points", 
                       xytext=(0, 10), ha='center', fontsize=10, fontweight='bold')
        
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
    
    def plot_burn_rate(self):
        """Plot 9: Monthly burn rate dashboard."""
        if self.df is None:
            messagebox.showwarning("No Data", "Please load data first")
            return
        
        self.clear_plot()
        
        # Calculate components
        monthly_data = self.df.groupby('month').agg({
            'total_cost': 'sum',
            'salaire_brut': 'sum',
            'cot_patronale': 'sum',
            'avantages': 'sum',
            'employee_id': 'nunique'
        }).reset_index()
        
        # Create dashboard
        fig = plt.figure(figsize=(14, 10))
        gs = fig.add_gridspec(3, 2, hspace=0.3, wspace=0.3)
        
        # Total burn rate
        ax1 = fig.add_subplot(gs[0, :])
        ax1.bar(monthly_data['month'], monthly_data['total_cost'], 
               color='#C73E1D', alpha=0.7, edgecolor='black')
        ax1.set_title('Monthly Burn Rate', fontsize=14, fontweight='bold')
        ax1.set_ylabel('Total Cost ($)', fontsize=11)
        ax1.grid(True, alpha=0.3, axis='y')
        ax1.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m'))
        plt.setp(ax1.xaxis.get_majorticklabels(), rotation=45, ha='right')
        
        # Components breakdown
        ax2 = fig.add_subplot(gs[1, 0])
        latest = monthly_data.iloc[-1]
        components = ['Gross\nSalary', 'Employer\nContrib', 'Benefits']
        values = [latest['salaire_brut'], latest['cot_patronale'], latest['avantages']]
        colors = ['#2E86AB', '#A23B72', '#F18F01']
        ax2.pie(values, labels=components, autopct='%1.1f%%', colors=colors, startangle=90)
        ax2.set_title(f'Cost Components ({monthly_data["month"].iloc[-1].strftime("%Y-%m")})', 
                     fontsize=12, fontweight='bold')
        
        # Average per employee
        ax3 = fig.add_subplot(gs[1, 1])
        monthly_data['avg_per_emp'] = monthly_data['total_cost'] / monthly_data['employee_id']
        ax3.plot(monthly_data['month'], monthly_data['avg_per_emp'], 
                marker='o', linewidth=2, color='#F18F01')
        ax3.set_title('Average Cost per Employee', fontsize=12, fontweight='bold')
        ax3.set_ylabel('Cost ($)', fontsize=10)
        ax3.grid(True, alpha=0.3)
        ax3.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m'))
        plt.setp(ax3.xaxis.get_majorticklabels(), rotation=45, ha='right')
        
        # Cumulative burn
        ax4 = fig.add_subplot(gs[2, :])
        monthly_data['cumulative'] = monthly_data['total_cost'].cumsum()
        ax4.fill_between(monthly_data['month'], monthly_data['cumulative'], 
                        alpha=0.5, color='#2E86AB')
        ax4.plot(monthly_data['month'], monthly_data['cumulative'], 
                marker='o', linewidth=2, color='#2E86AB')
        ax4.set_title('Cumulative Burn', fontsize=14, fontweight='bold')
        ax4.set_xlabel('Month', fontsize=11)
        ax4.set_ylabel('Cumulative Cost ($)', fontsize=11)
        ax4.grid(True, alpha=0.3)
        ax4.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m'))
        plt.setp(ax4.xaxis.get_majorticklabels(), rotation=45, ha='right')
        
        fig.suptitle('Monthly Burn Rate Dashboard', fontsize=16, fontweight='bold', y=0.995)
        
        canvas = FigureCanvasTkAgg(fig, self.plot_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
    
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
        ax.set_ylabel('Total Cost ($)', fontsize=12)
        ax.set_xticks(range(1, 13))
        ax.set_xticklabels(['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                           'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'])
        ax.legend(title='Year', fontsize=10)
        ax.grid(True, alpha=0.3)
        
        plt.tight_layout()
        
        canvas = FigureCanvasTkAgg(fig, self.plot_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
    
    def run(self):
        """Run the dashboard."""
        self.root.mainloop()


def main():
    """Main entry point."""
    dashboard = PayrollDashboard()
    dashboard.run()


if __name__ == "__main__":
    main()
