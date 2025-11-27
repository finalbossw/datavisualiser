

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from datetime import datetime
import os


class MiniExcelVisualizer:
    
    
    def __init__(self, root):
        self.root = root
        self.root.title("Mini Excel Data Visualizer")
        self.root.geometry("1200x700")
        self.root.configure(bg="#f5f7fa")
        
        # Data storage
        self.df = None
        self.filename = ""
        self.canvas_widget = None
        self.fig = None
        self.current_chart_info = None  
        self.setup_ui()
        
    def setup_ui(self):
        
        
        
        main_frame = tk.Frame(self.root, bg="#f5f7fa")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Top control panel
        control_frame = tk.Frame(main_frame, bg="#ffffff", relief=tk.RIDGE, bd=1)
        control_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Load CSV button
        btn_load = tk.Button(
            control_frame, 
            text="📁 Load CSV", 
            command=self.load_csv,
            bg="#4CAF50", 
            fg="white", 
            font=("Arial", 10, "bold"),
            padx=20, 
            pady=8,
            cursor="hand2"
        )
        btn_load.pack(side=tk.LEFT, padx=10, pady=10)
        
        # Chart controls frame
        chart_controls = tk.Frame(control_frame, bg="#ffffff")
        chart_controls.pack(side=tk.LEFT, padx=10, pady=10)
        
        # X Column
        tk.Label(chart_controls, text="X Column:", bg="#ffffff", font=("Arial", 9)).grid(row=0, column=0, padx=5, sticky="w")
        self.x_column_var = tk.StringVar()
        self.x_dropdown = ttk.Combobox(chart_controls, textvariable=self.x_column_var, state="readonly", width=15)
        self.x_dropdown.grid(row=0, column=1, padx=5)
        
        # Y Column
        tk.Label(chart_controls, text="Y Column:", bg="#ffffff", font=("Arial", 9)).grid(row=0, column=2, padx=5, sticky="w")
        self.y_column_var = tk.StringVar()
        self.y_dropdown = ttk.Combobox(chart_controls, textvariable=self.y_column_var, state="readonly", width=15)
        self.y_dropdown.grid(row=0, column=3, padx=5)
        
        # Aggregation
        tk.Label(chart_controls, text="Aggregation:", bg="#ffffff", font=("Arial", 9)).grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.agg_var = tk.StringVar(value="sum")
        self.agg_dropdown = ttk.Combobox(
            chart_controls, 
            textvariable=self.agg_var, 
            state="readonly", 
            width=15,
            values=["sum", "mean", "count", "min", "max"]
        )
        self.agg_dropdown.grid(row=1, column=1, padx=5, pady=5)
        
        # Chart Type
        tk.Label(chart_controls, text="Chart Type:", bg="#ffffff", font=("Arial", 9)).grid(row=1, column=2, padx=5, pady=5, sticky="w")
        self.chart_type_var = tk.StringVar(value="Bar")
        self.chart_dropdown = ttk.Combobox(
            chart_controls, 
            textvariable=self.chart_type_var, 
            state="readonly", 
            width=15,
            values=["Bar", "Line", "Pie", "Scatter", "Histogram"]
        )
        self.chart_dropdown.grid(row=1, column=3, padx=5, pady=5)
        
        # Chart action buttons
        btn_frame = tk.Frame(control_frame, bg="#ffffff")
        btn_frame.pack(side=tk.LEFT, padx=10)
        
        btn_generate = tk.Button(
            btn_frame, 
            text="📊 Generate Chart", 
            command=self.generate_chart,
            bg="#2196F3", 
            fg="white", 
            font=("Arial", 9, "bold"),
            padx=15, 
            pady=6,
            cursor="hand2"
        )
        btn_generate.pack(side=tk.TOP, pady=2)
        
        btn_clear = tk.Button(
            btn_frame, 
            text="🗑️ Clear Chart", 
            command=self.clear_chart,
            bg="#FF9800", 
            fg="white", 
            font=("Arial", 9, "bold"),
            padx=15, 
            pady=6,
            cursor="hand2"
        )
        btn_clear.pack(side=tk.TOP, pady=2)
        
        # PDF Report button
        btn_report = tk.Button(
            control_frame, 
            text="📄 Generate PDF Report", 
            command=self.generate_report,
            bg="#9C27B0", 
            fg="white", 
            font=("Arial", 10, "bold"),
            padx=20, 
            pady=8,
            cursor="hand2"
        )
        btn_report.pack(side=tk.RIGHT, padx=10, pady=10)
        
        # Content area (data table + chart) using PanedWindow for fixed split
        content_paned = tk.PanedWindow(main_frame, orient=tk.HORIZONTAL, bg="#f5f7fa", sashwidth=5, sashrelief=tk.RAISED)
        content_paned.pack(fill=tk.BOTH, expand=True)
        
        # Left panel - Data table
        table_frame = tk.Frame(content_paned, bg="#ffffff", relief=tk.RIDGE, bd=1, width=600)
        content_paned.add(table_frame, minsize=400)
        
        tk.Label(table_frame, text="Data Table", bg="#ffffff", font=("Arial", 11, "bold")).pack(pady=5)
        
        # Treeview with scrollbars
        tree_container = tk.Frame(table_frame, bg="#ffffff")
        tree_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        tree_scroll_y = ttk.Scrollbar(tree_container, orient=tk.VERTICAL)
        tree_scroll_x = ttk.Scrollbar(tree_container, orient=tk.HORIZONTAL)
        
        self.tree = ttk.Treeview(
            tree_container,
            yscrollcommand=tree_scroll_y.set,
            xscrollcommand=tree_scroll_x.set,
            show="tree headings"
        )
        
        tree_scroll_y.config(command=self.tree.yview)
        tree_scroll_x.config(command=self.tree.xview)
        
        tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        # Right panel - Chart area
        self.chart_frame = tk.Frame(content_paned, bg="#ffffff", relief=tk.RIDGE, bd=1, width=600)
        content_paned.add(self.chart_frame, minsize=400)
        
        chart_label = tk.Label(self.chart_frame, text="Chart View", bg="#ffffff", font=("Arial", 11, "bold"))
        chart_label.pack(pady=5)
        
        # Placeholder label for empty chart area
        self.chart_placeholder = tk.Label(
            self.chart_frame, 
            text="📊\n\nSelect columns and click\n'Generate Chart' to display visualization", 
            bg="#ffffff", 
            fg="#999999",
            font=("Arial", 11)
        )
        self.chart_placeholder.pack(expand=True)
        
        # Status bar
        self.status_bar = tk.Label(
            self.root, 
            text="Ready", 
            bd=1, 
            relief=tk.SUNKEN, 
            anchor=tk.W,
            bg="#e0e0e0",
            font=("Arial", 9)
        )
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
    def load_csv(self):
        """Load a CSV file and display in table"""
        try:
            filepath = filedialog.askopenfilename(
                title="Select CSV File",
                filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
            )
            
            if not filepath:
                return
            
            self.df = pd.read_csv(filepath)
            self.filename = os.path.basename(filepath)
            
            self.display_data()
            self.update_column_dropdowns()
            self.set_status(f"Loaded: {self.filename} ({len(self.df)} rows, {len(self.df.columns)} columns)")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load CSV:\n{str(e)}")
            self.set_status("Error loading file")
    
    def display_data(self):
        """Display DataFrame in Treeview"""
        # Clear existing data
        self.tree.delete(*self.tree.get_children())
        
        # Configure columns
        self.tree["columns"] = list(self.df.columns)
        self.tree["show"] = "headings"
        
        # Set column headings
        for col in self.df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor=tk.W)
        
        # Insert data rows
        for idx, row in self.df.iterrows():
            self.tree.insert("", tk.END, values=list(row))
    
    def update_column_dropdowns(self):
        """Update X and Y column dropdowns with DataFrame columns"""
        if self.df is not None:
            columns = list(self.df.columns)
            self.x_dropdown["values"] = columns
            self.y_dropdown["values"] = columns
            
            if len(columns) > 0:
                self.x_dropdown.current(0)
            if len(columns) > 1:
                self.y_dropdown.current(1)
            elif len(columns) > 0:
                self.y_dropdown.current(0)
    
    def generate_chart(self):
        """Generate chart based on selected parameters"""
        if self.df is None:
            messagebox.showwarning("Warning", "Please load a CSV file first")
            return
        
        x_col = self.x_column_var.get()
        y_col = self.y_column_var.get()
        agg_func = self.agg_var.get()
        chart_type = self.chart_type_var.get()
        
        if not x_col or not y_col:
            messagebox.showwarning("Warning", "Please select both X and Y columns")
            return
        
        try:
            # Clear previous chart
            self.clear_chart()
            
            # Remove placeholder if it exists
            if hasattr(self, 'chart_placeholder') and self.chart_placeholder.winfo_exists():
                self.chart_placeholder.pack_forget()
            
            # Create figure
            self.fig = Figure(figsize=(6, 5), dpi=100)
            ax = self.fig.add_subplot(111)
            
            # Prepare data with aggregation
            if chart_type == "Histogram":
                # Histogram only needs Y column
                if not pd.api.types.is_numeric_dtype(self.df[y_col]):
                    raise ValueError(f"Column '{y_col}' must be numeric for histogram")
                ax.hist(self.df[y_col].dropna(), bins=20, edgecolor='black', alpha=0.7)
                ax.set_xlabel(y_col)
                ax.set_ylabel("Frequency")
                ax.set_title(f"Histogram of {y_col}")
            else:
                # Apply aggregation
                grouped = self.df.groupby(x_col)[y_col].agg(agg_func).reset_index()
                x_data = grouped[x_col]
                y_data = grouped[y_col]
                
                # Generate appropriate chart
                if chart_type == "Bar":
                    ax.bar(x_data, y_data, color='#2196F3', alpha=0.8)
                    ax.set_xlabel(x_col)
                    ax.set_ylabel(f"{agg_func}({y_col})")
                    ax.set_title(f"{agg_func.capitalize()} of {y_col} by {x_col}")
                    plt.setp(ax.xaxis.get_majorticklabels(), rotation=45, ha='right')
                    
                elif chart_type == "Line":
                    ax.plot(x_data, y_data, marker='o', color='#4CAF50', linewidth=2)
                    ax.set_xlabel(x_col)
                    ax.set_ylabel(f"{agg_func}({y_col})")
                    ax.set_title(f"{agg_func.capitalize()} of {y_col} by {x_col}")
                    plt.setp(ax.xaxis.get_majorticklabels(), rotation=45, ha='right')
                    
                elif chart_type == "Pie":
                    ax.pie(y_data, labels=x_data, autopct='%1.1f%%', startangle=90)
                    ax.set_title(f"{agg_func.capitalize()} of {y_col} by {x_col}")
                    
                elif chart_type == "Scatter":
                    ax.scatter(x_data, y_data, color='#FF9800', alpha=0.7, s=100)
                    ax.set_xlabel(x_col)
                    ax.set_ylabel(f"{agg_func}({y_col})")
                    ax.set_title(f"{agg_func.capitalize()} of {y_col} by {x_col}")
            
            self.fig.tight_layout()
            
            # Embed chart in Tkinter
            self.canvas_widget = FigureCanvasTkAgg(self.fig, master=self.chart_frame)
            self.canvas_widget.draw()
            self.canvas_widget.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            # Store chart info for analysis report
            self.current_chart_info = {
                'x_col': x_col,
                'y_col': y_col,
                'agg_func': agg_func,
                'chart_type': chart_type,
                'data': grouped if chart_type != "Histogram" else self.df[[y_col]].dropna()
            }
            
            self.set_status(f"Chart generated: {chart_type} ({agg_func})")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate chart:\n{str(e)}")
            self.set_status("Chart generation failed")
    
    def clear_chart(self):
        """Clear the current chart"""
        if self.canvas_widget:
            self.canvas_widget.get_tk_widget().destroy()
            self.canvas_widget = None
        if self.fig:
            plt.close(self.fig)
            self.fig = None
        
        # Show placeholder again
        if hasattr(self, 'chart_placeholder'):
            self.chart_placeholder.pack(expand=True)
        
        self.set_status("Chart cleared")
    
    def generate_report(self):
        """Generate PDF report with data analysis"""
        if self.df is None:
            messagebox.showwarning("Warning", "Please load a CSV file first")
            return
        
        if self.current_chart_info is None:
            messagebox.showwarning("Warning", "Please generate a chart first to create an analysis report")
            return
        
        try:
            filepath = filedialog.asksaveasfilename(
                title="Save Analysis Report",
                defaultextension=".pdf",
                filetypes=[("PDF Files", "*.pdf")]
            )
            
            if not filepath:
                return
            
            # Create PDF
            doc = SimpleDocTemplate(filepath, pagesize=A4)
            elements = []
            styles = getSampleStyleSheet()
            
            # Title
            title = Paragraph(f"<b>Data Analysis Report</b>", styles['Title'])
            elements.append(title)
            elements.append(Spacer(1, 12))
            
            # Report metadata
            metadata_text = f"""
            <b>Dataset:</b> {self.filename}<br/>
            <b>Generated:</b> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}<br/>
            <b>Total Records:</b> {len(self.df)}<br/>
            """
            metadata = Paragraph(metadata_text, styles['Normal'])
            elements.append(metadata)
            elements.append(Spacer(1, 20))
            
            # Chart Configuration Section
            chart_config_title = Paragraph("<b>Chart Configuration</b>", styles['Heading2'])
            elements.append(chart_config_title)
            elements.append(Spacer(1, 8))
            
            chart_info = self.current_chart_info
            config_text = f"""
            <b>Chart Type:</b> {chart_info['chart_type']}<br/>
            <b>X-Axis:</b> {chart_info['x_col']}<br/>
            <b>Y-Axis:</b> {chart_info['y_col']}<br/>
            <b>Aggregation:</b> {chart_info['agg_func']}<br/>
            """
            config = Paragraph(config_text, styles['Normal'])
            elements.append(config)
            elements.append(Spacer(1, 20))
            
            # Data Analysis Section
            analysis_title = Paragraph("<b>Data Analysis & Insights</b>", styles['Heading2'])
            elements.append(analysis_title)
            elements.append(Spacer(1, 8))
            
            # Generate insights based on chart type and data
            insights = self.generate_insights()
            
            for insight in insights:
                insight_para = Paragraph(f"• {insight}", styles['Normal'])
                elements.append(insight_para)
                elements.append(Spacer(1, 6))
            
            elements.append(Spacer(1, 20))
            
            # Statistical Summary Section
            stats_title = Paragraph("<b>Statistical Summary</b>", styles['Heading2'])
            elements.append(stats_title)
            elements.append(Spacer(1, 8))
            
            # Generate statistics
            stats_data = self.generate_statistics()
            
            # Create statistics table
            stats_table = Table(stats_data, colWidths=[200, 150])
            stats_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 11),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 1, colors.black)
            ]))
            elements.append(stats_table)
            
            # Build PDF
            doc.build(elements)
            
            self.set_status(f"Analysis report saved: {os.path.basename(filepath)}")
            messagebox.showinfo("Success", f"Analysis report generated successfully:\n{filepath}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate report:\n{str(e)}")
            self.set_status("Report generation failed")
    
    def generate_insights(self):
        """Generate insights based on current chart and data"""
        insights = []
        chart_info = self.current_chart_info
        data = chart_info['data']
        
        try:
            if chart_info['chart_type'] == "Histogram":
                # Histogram insights
                y_col = chart_info['y_col']
                mean_val = data[y_col].mean()
                median_val = data[y_col].median()
                std_val = data[y_col].std()
                
                insights.append(f"The average {y_col} is {mean_val:.2f} with a standard deviation of {std_val:.2f}")
                insights.append(f"The median {y_col} is {median_val:.2f}, indicating the central tendency of the distribution")
                
                if mean_val > median_val * 1.1:
                    insights.append(f"The distribution is right-skewed, with higher values pulling the average up")
                elif median_val > mean_val * 1.1:
                    insights.append(f"The distribution is left-skewed, with lower values pulling the average down")
                else:
                    insights.append(f"The distribution appears relatively symmetric")
                    
            else:
                # Other chart types
                x_col = chart_info['x_col']
                y_col = chart_info['y_col']
                agg_func = chart_info['agg_func']
                
                # Find top and bottom performers
                if len(data) > 0:
                    sorted_data = data.sort_values(by=y_col, ascending=False)
                    top_item = sorted_data.iloc[0]
                    bottom_item = sorted_data.iloc[-1]
                    
                    insights.append(f"'{top_item[x_col]}' has the highest {agg_func} {y_col} of {top_item[y_col]:.2f}")
                    insights.append(f"'{bottom_item[x_col]}' has the lowest {agg_func} {y_col} of {bottom_item[y_col]:.2f}")
                    
                    # Calculate total and average
                    total = data[y_col].sum()
                    avg = data[y_col].mean()
                    
                    insights.append(f"Total {agg_func} across all categories: {total:.2f}")
                    insights.append(f"Average {agg_func} per category: {avg:.2f}")
                    
                    # Top contributor percentage
                    if total > 0:
                        top_percentage = (top_item[y_col] / total) * 100
                        insights.append(f"'{top_item[x_col]}' contributes {top_percentage:.1f}% of the total")
                    
                    # Comparison insight
                    if len(data) >= 2:
                        difference = top_item[y_col] - bottom_item[y_col]
                        insights.append(f"There is a difference of {difference:.2f} between the highest and lowest values")
                
        except Exception as e:
            insights.append("Analysis data unavailable for current configuration")
        
        return insights
    
    def generate_statistics(self):
        """Generate statistical summary table"""
        stats_data = [["Metric", "Value"]]
        chart_info = self.current_chart_info
        data = chart_info['data']
        
        try:
            if chart_info['chart_type'] == "Histogram":
                y_col = chart_info['y_col']
                stats_data.append(["Mean", f"{data[y_col].mean():.2f}"])
                stats_data.append(["Median", f"{data[y_col].median():.2f}"])
                stats_data.append(["Std Dev", f"{data[y_col].std():.2f}"])
                stats_data.append(["Min", f"{data[y_col].min():.2f}"])
                stats_data.append(["Max", f"{data[y_col].max():.2f}"])
                stats_data.append(["Count", f"{data[y_col].count()}"])
            else:
                y_col = chart_info['y_col']
                stats_data.append(["Total Sum", f"{data[y_col].sum():.2f}"])
                stats_data.append(["Average", f"{data[y_col].mean():.2f}"])
                stats_data.append(["Median", f"{data[y_col].median():.2f}"])
                stats_data.append(["Min Value", f"{data[y_col].min():.2f}"])
                stats_data.append(["Max Value", f"{data[y_col].max():.2f}"])
                stats_data.append(["Number of Categories", f"{len(data)}"])
                
        except Exception as e:
            stats_data.append(["Error", "Statistics unavailable"])
        
        return stats_data
    
    def set_status(self, message):
        """Update status bar message"""
        self.status_bar.config(text=message)


if __name__ == "__main__":
    root = tk.Tk()
    app = MiniExcelVisualizer(root)
    root.mainloop()