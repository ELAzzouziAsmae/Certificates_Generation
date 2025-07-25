import sys
import os
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QWidget, QTableWidget, QTableWidgetItem,
    QLabel, QComboBox, QHBoxLayout, QSpinBox, QPushButton, QGridLayout, QDateEdit
)
from PyQt5.QtWebEngineWidgets import QWebEngineView
from PyQt5.QtCore import QUrl
import plotly.express as px
from PyQt5.QtWidgets import QHeaderView

class KPIWindow(QMainWindow):
    def __init__(self, df, parent=None):
        super().__init__(parent)
        self.setWindowTitle("📊 KPI Dashboard")
        self.df_original = df.copy()

        # Mapping for regions based on SDSU units
        self.sdsu_to_region = {
            "SAM": "ERCIS", "AFD": "ERCIS", "AGS": "ERCIS", "AMR": "ERCIS",
            "BIO": "ERCIS", "CME": "ERCIS", "PLS": "ERCIS", "GLI": "ERCIS",
            "PSL": "ERCIS", "SAB": "ERCIS", "GWO7": "ERCIS", "GWV3": "ERCIS",
            "JSU1": "ERCIS",
            "GMM": "NAM", "PAM_NAM": "NAM", "SABAP": "MENAT",
            "ABS": "LAM", "ACS": "LAM", "AMS": "LAM", "PAM_LAM": "LAM",
            "QAR": "MENAT", "SAU": "MENAT", "UAE": "MENAT", "LFO": "MENAT",
            "ATT": "MENAT", "ATS": "MENAT", "MSC": "MENAT", "ETS": "MENAT",
            "TFO": "MENAT", "PAM_NAFT": "MENAT", "PAM_SABAP": "MENAT", "PAM__GULF": "MENAT",
            "ASG": "CEAP", "ITS": "CEAP", "TPA": "CEAP", "SCN": "CEAP", "PAM_CEAP": "CEAP",
            "PCP": "IND", "PAM_PCP": "IND"
        }

        self.prepare_data()
        self.init_ui()

    def prepare_data(self):
        df = self.df_original.copy()
        sdsu_col = next((col for col in df.columns if "SDSU Business Unit" in col), None)
        df["SDSU Unit"] = df[sdsu_col]
        df["Region"] = df["SDSU Unit"].map(self.sdsu_to_region).fillna("Unknown")

        date_col = next((col for col in df.columns if "Date" in col or "Timestamp" in col), None)
        if date_col:
            df["Date"] = pd.to_datetime(df[date_col], errors='coerce')  # Ensures conversion to datetime
            df["Year"] = df["Date"].dt.year
        else:
            df["Date"] = pd.NaT
            df["Year"] = "Unknown"

        df["Attempts"] = pd.to_numeric(df["No. of Attempts"], errors="coerce").fillna(0).astype(int) + 1

        self.df = df

    def init_ui(self):
        widget = QWidget()
        self.setCentralWidget(widget)

        main_layout = QVBoxLayout(widget)

        # Filters for date range (QDateEdit to show a calendar on click)
        filter_layout = QHBoxLayout()

        self.start_date = QDateEdit(self)
        self.start_date.setDisplayFormat("yyyy-MM-dd")
        self.start_date.setDate(self.df["Date"].min().date())  # Initialize with the earliest date
        self.start_date.setCalendarPopup(True)  # Allows showing a calendar when clicking the field
        self.start_date.dateChanged.connect(self.update_view)  # Connect to the date change

        self.end_date = QDateEdit(self)
        self.end_date.setDisplayFormat("yyyy-MM-dd")
        self.end_date.setDate(self.df["Date"].max().date())  # Initialize with the most recent date
        self.end_date.setCalendarPopup(True)  # Allows showing a calendar when clicking the field
        self.end_date.dateChanged.connect(self.update_view)  # Connect to the date change

        self.success_threshold = QSpinBox()
        self.success_threshold.setRange(0, 100)
        self.success_threshold.setValue(70)
        self.success_threshold.setSuffix(" %")
        self.success_threshold.valueChanged.connect(self.update_view)

        filter_layout.addWidget(QLabel("Filter by date:"))
        filter_layout.addWidget(self.start_date)
        filter_layout.addWidget(QLabel("to"))
        filter_layout.addWidget(self.end_date)
        filter_layout.addSpacing(20)
        filter_layout.addWidget(QLabel("Success Threshold:"))
        filter_layout.addWidget(self.success_threshold)

        self.export_btn = QPushButton("Export to CSV")
        self.export_btn.clicked.connect(self.export_csv)
        filter_layout.addSpacing(20)
        filter_layout.addWidget(self.export_btn)

        self.export_html_btn = QPushButton("Export to HTML")
        self.export_html_btn.clicked.connect(self.export_html)
        filter_layout.addWidget(self.export_html_btn)

        main_layout.addLayout(filter_layout)

        grid = QGridLayout()
        main_layout.addLayout(grid)

        self.table = QTableWidget()
        self.table.setStyleSheet(""" ... """)  # Style remains unchanged
        grid.addWidget(self.table, 0, 0)

        self.web_region = QWebEngineView()
        grid.addWidget(self.web_region, 0, 1)

        self.web_unit = QWebEngineView()
        grid.addWidget(self.web_unit, 1, 0)

        self.web_pie = QWebEngineView()
        grid.addWidget(self.web_pie, 1, 1)

        self.update_view()

    def export_csv(self):
        # Export data after filtering by dates
        if not hasattr(self, "last_table_df"):
            print("❌ No data to export. Ensure the view is updated.")
            return

        display_df = self.last_table_df  # Get the displayed DataFrame

        # Define the path for the CSV file
        filename_csv = r"C:\Users\260001889\Desktop\kpi_export.csv"  # Specify an absolute path

        try:
            # Save the data into the CSV file with the visible columns in the table
            display_df.to_csv(filename_csv, index=False)
            print(f"✅ CSV Export completed: {os.path.abspath(filename_csv)}")
        except Exception as e:
            print(f"❌ Error during CSV export: {e}")

    def export_html(self):
        try:
            # Check if graphs and table exist
            if not hasattr(self, "last_fig_region") or not hasattr(self, "last_table_df") or not hasattr(self, "last_fig_unit") or not hasattr(self, "last_fig_pie"):
                print("❌ The graphs are not yet generated. Please refresh the view first.")
                return

            # Generate the HTML for the charts with plotlyjs included only once (included in the first one)
            html_region = self.last_fig_region.to_html(include_plotlyjs=True, full_html=False)
            html_unit = self.last_fig_unit.to_html(include_plotlyjs=False, full_html=False)
            html_pie = self.last_fig_pie.to_html(include_plotlyjs=False, full_html=False)

            # HTML for the pandas table
            table_html = self.last_table_df.to_html(index=False, classes="data-table")

            # Style for the HTML
            css = """
            <style>
                body { font-family: Arial; margin: 20px; background: #f9f9f9; }
                h1 { text-align: center; }
                h2 { text-align: center; margin-bottom: 20px; }  /* Center the title */
                .grid { display: grid; grid-template-columns: 1fr 1fr; grid-gap: 20px; }
                .cell { background: white; padding: 12px; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
                .data-table { width: 100%; border-collapse: collapse; font-size: 12.5px; }
                .data-table th, .data-table td { border: 1px solid #ccc; padding: 4px 6px; text-align: center; }
                .data-table th { background-color: #eee; }
            </style>
            """

            # Getting the date range and success rate for the header
            start_date = self.start_date.date().toPyDate().strftime("%Y-%m-%d")
            end_date = self.end_date.date().toPyDate().strftime("%Y-%m-%d")
            success_threshold = self.success_threshold.value()

            # Creating the header with the date range and success rate threshold
            header = f"""
            <h2>KPI Dashboard</h2>  <!-- Title now centered -->
            <p>Filtering by date: {start_date} to {end_date}</p>
            <p>Success Threshold: {success_threshold}%</p>
            """

            # Final HTML page assembly
            html = f"""
            <html>
            <head>
                <meta charset="utf-8">
                <title>Dashboard KPI Export</title>
                <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
                {css}
            </head>
            <body>
                {header}  <!-- Date range and success threshold -->
                <div class="grid">
                    <div class="cell">{table_html}</div>
                    <div class="cell">{html_region}</div>
                    <div class="cell">{html_unit}</div>
                    <div class="cell">{html_pie}</div>
                </div>
            </body>
            </html>
            """

            # Save the HTML to a file
            with open("dashboard_export.html", "w", encoding="utf-8") as f:
                f.write(html)

            print("✅ HTML Export completed: dashboard_export.html")

        except Exception as e:
            print("❌ Error in export_html:", e)

    def update_view(self):
        success_threshold = self.success_threshold.value()
        df = self.df.copy()

        # Filter data by selected date range
        start_date = self.start_date.date().toPyDate()  # Use `date()` to get the date
        end_date = self.end_date.date().toPyDate()  # Use `date()` to get the date

        # Convert dates to datetime64 for comparison with the "Date" column
        start_date = pd.to_datetime(start_date)
        end_date = pd.to_datetime(end_date)

        # Apply the date filter
        df = df[(df["Date"] >= start_date) & (df["Date"] <= end_date)]

        # Group data by region and SDSU Unit
        grouped = df.groupby(["Region", "SDSU Unit"]).agg(
            Attempts=('Attempts', 'sum'),
            Successes=('Percent Score', lambda x: (x >= success_threshold).sum())
        ).reset_index()

        # Calculate success rate per SDSU Unit (local)
        grouped['Success Rate per Unit (%)'] = (grouped['Successes'] / grouped['Attempts'] * 100).round(1)

        # Group data by region to get total attempts and successes
        region_summary = grouped.groupby("Region").agg({
            "Attempts": "sum",
            "Successes": "sum"
        }).reset_index()

        # Calculate success rate per region
        region_summary['Success Rate by Region (%)'] = (
            region_summary['Successes'] / region_summary['Attempts'] * 100
        ).round(1)

        # Merge to add success rate by region to the original DataFrame
        grouped = grouped.merge(region_summary[['Region', 'Success Rate by Region (%)']], on='Region', how='left')

        # Prepare the table for display
        display_df = grouped.rename(columns={"SDSU Unit": " Unit"})[[ 
            "Region", " Unit", "Attempts", "Successes", 
            "Success Rate per Unit (%)", "Success Rate by Region (%)"
        ]]

        # Update the table
        self.table.setRowCount(len(display_df))
        self.table.setColumnCount(len(display_df.columns))
        self.table.setHorizontalHeaderLabels(display_df.columns.tolist())

        for i, row in display_df.iterrows():
            for j, val in enumerate(row):
                self.table.setItem(i, j, QTableWidgetItem(str(val)))

        self.table.resizeColumnsToContents()

        # Success rate by region graph
        fig_region = px.bar(
            region_summary,
            x="Region",
            y="Success Rate by Region (%)",
            title="Success Rate by Region (%)",
            color="Success Rate by Region (%)",
            color_continuous_scale=px.colors.sequential.Viridis
        )
        # Center and bold the title
        fig_region.update_layout(
            title=dict(
                text="Success Rate by Region (%)",
                x=0.5,  # Center horizontally
                xanchor='center',  # Center relative to the X axis
                font=dict(size=18, family='Arial', weight='bold')  # Bold font
            )
        )
        # Directly load the graph in QWebEngineView
        self.web_region.setHtml(fig_region.to_html(include_plotlyjs="cdn"))

        # Success rate by SDSU Unit graph
        fig_unit = px.bar(
            grouped,
            x="SDSU Unit",
            y="Success Rate per Unit (%)",
            title="Success Rate by SDSU Unit (%)",
            color="Success Rate per Unit (%)",
            color_continuous_scale=px.colors.sequential.Plasma
        )
        fig_unit.update_yaxes(range=[0, 100])
        # Center and bold the title
        fig_unit.update_layout(
            title=dict(
                text="Success Rate by SDSU Unit (%)",
                x=0.5,  # Center horizontally
                xanchor='center',  # Center relative to the X axis
                font=dict(size=18, family='Arial', weight='bold')  # Bold font
            )
        )
        # Directly load the graph in QWebEngineView
        self.web_unit.setHtml(fig_unit.to_html(include_plotlyjs="cdn"))

        # Pie chart based on the overall success rate
        region_summary_for_pie = region_summary.copy()

        # If the overall success rate is 0, replace with None to avoid showing it
        region_summary_for_pie.loc[region_summary_for_pie['Success Rate by Region (%)'] == 0, 'Success Rate by Region (%)'] = None

        fig_pie = px.pie(
            region_summary_for_pie,
            values="Success Rate by Region (%)",  # Uses the overall success rate for each region
            names="Region",  # Region names
            title="Distribution of Success Rates by Region",
            color_discrete_sequence=px.colors.qualitative.Bold  # Choose color palette
        )
        # Center and bold the title
        fig_pie.update_layout(
            title=dict(
                text="Distribution of Success Rates by Region",
                x=0.5,  # Center horizontally
                xanchor='center',  # Center relative to the X axis
                font=dict(size=18, family='Arial', weight='bold')  # Bold font
            )
        )

        # Directly load the graph in QWebEngineView
        self.web_pie.setHtml(fig_pie.to_html(include_plotlyjs="cdn"))

        # Save graphs for export
        self.last_fig_region = fig_region
        self.last_fig_unit = fig_unit
        self.last_fig_pie = fig_pie
        self.last_table_df = grouped.rename(columns={"SDSU Unit": " Unit"})[[ 
            "Region", " Unit", "Attempts", "Successes", "Success Rate per Unit (%)", "Success Rate by Region (%)"
        ]] 

if __name__ == "__main__":
    excel_path = r"C:\Users\260001889\Downloads\FORM_2188484_1753176334450_07-22-2025_09_25_34f38_SDSU.xlsx"
    if not os.path.exists(excel_path):
        raise FileNotFoundError("⚠️ Excel file not found. Please check the path.")

    df = pd.read_excel(excel_path)

    app = QApplication(sys.argv)
    window = KPIWindow(df)
    window.resize(1200, 800)
    window.show()

    sys.exit(app.exec_())
