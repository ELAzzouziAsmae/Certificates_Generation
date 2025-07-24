import sys
import os
import webbrowser
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QWidget, QTableWidget, QTableWidgetItem,
    QLabel, QComboBox, QHBoxLayout, QSpinBox, QPushButton, QGridLayout
)
from PyQt5.QtWebEngineWidgets import QWebEngineView
from PyQt5.QtCore import QUrl
import plotly.express as px

class KPIWindow(QMainWindow):
    def __init__(self, df, parent=None):
        super().__init__(parent)
        self.setWindowTitle("📊 Tableau de bord des KPI")
        self.df_original = df.copy()

        self.sdsu_to_region = {
            "SAM": "ERCIS", "AFD": "ERCIS", "AGS": "ERCIS", "AMR": "ERCIS",
            "BIO": "ERCIS", "CME": "ERCIS", "PLS": "ERCIS", "GLI": "ERCIS",
            "PSL": "ERCIS", "SAB": "ERCIS", "GWO7": "ERCIS", "GWV3": "ERCIS",
            "JSU1": "ERCIS", "GMM": "NAM", "PAM_NAM": "NAM", "SABAP": "MENAT",
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
        if not sdsu_col:
            raise ValueError("❌ Colonne SDSU manquante")
        df["SDSU Unit"] = df[sdsu_col]
        df["Région"] = df["SDSU Unit"].map(self.sdsu_to_region).fillna("Inconnue")

        date_col = next((col for col in df.columns if "Date" in col or "Timestamp" in col), None)
        if date_col:
            df["Date"] = pd.to_datetime(df[date_col], errors='coerce')
            df["Année"] = df["Date"].dt.year
        else:
            df["Date"] = pd.NaT
            df["Année"] = "Inconnue"

        self.df = df

    def init_ui(self):
        widget = QWidget()
        self.setCentralWidget(widget)
        main_layout = QVBoxLayout(widget)

        filter_layout = QHBoxLayout()
        self.annee_combo = QComboBox()
        self.annee_combo.addItem("Toutes les années")
        for an in sorted(self.df['Année'].dropna().unique()):
            self.annee_combo.addItem(str(int(an)))
        self.annee_combo.currentIndexChanged.connect(self.update_view)

        self.seuil_spin = QSpinBox()
        self.seuil_spin.setRange(0, 100)
        self.seuil_spin.setValue(70)
        self.seuil_spin.setSuffix(" %")
        self.seuil_spin.valueChanged.connect(self.update_view)

        filter_layout.addWidget(QLabel("Filtrer par année :"))
        filter_layout.addWidget(self.annee_combo)
        filter_layout.addSpacing(20)
        filter_layout.addWidget(QLabel("Seuil de réussite :"))
        filter_layout.addWidget(self.seuil_spin)

        self.export_btn = QPushButton("Exporter CSV")
        self.export_btn.clicked.connect(self.export_csv)
        filter_layout.addSpacing(20)
        filter_layout.addWidget(self.export_btn)

        # ✅ Nouveau bouton pour exporter HTML
        self.export_html_btn = QPushButton("Exporter interface HTML")
        self.export_html_btn.clicked.connect(self.export_html)
        filter_layout.addWidget(self.export_html_btn)

        main_layout.addLayout(filter_layout)

        grid = QGridLayout()
        main_layout.addLayout(grid)

        self.table = QTableWidget()
        self.table.setMinimumHeight(150)
        grid.addWidget(self.table, 0, 0)

        self.web_region = QWebEngineView()
        grid.addWidget(self.web_region, 0, 1)

        self.web_unit = QWebEngineView()
        grid.addWidget(self.web_unit, 1, 0)

        self.web_pie = QWebEngineView()
        grid.addWidget(self.web_pie, 1, 1)

        self.update_view()

    def export_csv(self):
        seuil = self.seuil_spin.value()
        df = self.get_filtered_data(seuil)
        filename = "kpi_export.csv"
        df.to_csv(filename, index=False)
        print(f"✅ Export CSV réalisé : {os.path.abspath(filename)}")

    def export_html(self):
        try:
            with open("graph_region.html", "r", encoding="utf-8") as f1, \
                 open("graph_unit.html", "r", encoding="utf-8") as f2, \
                 open("graph_pie.html", "r", encoding="utf-8") as f3:

                content = f"""
                <html>
                <head><title>📊 Tableau de bord Exporté</title></head>
                <body>
                    <h1 style="text-align:center;">📊 Tableau de bord des KPI</h1>
                    <h2>1️⃣ Taux de réussite par Région</h2>
                    {f1.read()}
                    <h2>2️⃣ Taux de réussite par SDSU Unit</h2>
                    {f2.read()}
                    <h2>3️⃣ Taux (%) de réussite par Région (Camembert)</h2>
                    {f3.read()}
                </body>
                </html>
                """
            with open("dashboard_export.html", "w", encoding="utf-8") as f_out:
                f_out.write(content)

            print("✅ Interface exportée dans : dashboard_export.html")
            webbrowser.open("dashboard_export.html")
        except Exception as e:
            print(f"❌ Erreur export HTML : {e}")

    def get_filtered_data(self, seuil):
        df = self.df.copy()
        selected_year = self.annee_combo.currentText()
        if selected_year != "Toutes les années":
            df = df[df["Année"] == int(selected_year)]

        grouped = df.groupby(["Région", "SDSU Unit"]).agg(
            Tentatives=('Percent Score', 'count'),
            Réussites=('Percent Score', lambda x: (x >= seuil).sum())
        ).reset_index()
        grouped['Taux Réussite (%)'] = (
            grouped['Réussites'] / grouped['Tentatives'] * 100
        ).clip(upper=100).round(1)
        return grouped

    def update_view(self):
        seuil = self.seuil_spin.value()
        grouped = self.get_filtered_data(seuil)

        self.table.setRowCount(len(grouped))
        self.table.setColumnCount(len(grouped.columns))
        self.table.setHorizontalHeaderLabels(grouped.columns.tolist())
        for i, row in grouped.iterrows():
            for j, val in enumerate(row):
                self.table.setItem(i, j, QTableWidgetItem(str(val)))

        region_summary = grouped.groupby("Région").agg({
            "Tentatives": "sum",
            "Réussites": "sum"
        }).reset_index()
        region_summary['Taux Réussite (%)'] = (
            region_summary['Réussites'] / region_summary['Tentatives'] * 100
        ).clip(upper=100).round(1)

        fig_region = px.bar(
            region_summary,
            x="Région", y="Taux Réussite (%)",
            title="Taux de réussite par Région",
            color="Taux Réussite (%)",
            color_continuous_scale=px.colors.sequential.Viridis,
            range_y=[0, 100]
        )
        fig_region.write_html("graph_region.html")
        self.web_region.load(QUrl.fromLocalFile(os.path.abspath("graph_region.html")))

        fig_unit = px.bar(
            grouped,
            x="SDSU Unit", y="Taux Réussite (%)",
            title="Taux de réussite par SDSU Unit",
            color="Taux Réussite (%)",
            color_continuous_scale=px.colors.sequential.Plasma,
        )
        fig_unit.update_yaxes(range=[0, 100])
        fig_unit.write_html("graph_unit.html")
        self.web_unit.load(QUrl.fromLocalFile(os.path.abspath("graph_unit.html")))

        fig_pie = px.pie(
            region_summary,
            values="Taux Réussite (%)",
            names="Région",
            title="Taux (%) de réussite par Région",
            color_discrete_sequence=px.colors.qualitative.Bold
        )
        fig_pie.write_html("graph_pie.html")
        self.web_pie.load(QUrl.fromLocalFile(os.path.abspath("graph_pie.html")))


if __name__ == "__main__":
    excel_path = r"C:\Users\260001889\Downloads\FORM_2188484_1753176334450_07-22-2025_09_25_34f38_SDSU.xlsx"
    if not os.path.exists(excel_path):
        raise FileNotFoundError("⚠️ Fichier Excel introuvable. Vérifie le chemin.")
    df = pd.read_excel(excel_path)

    app = QApplication(sys.argv)
    window = KPIWindow(df)
    window.resize(1200, 800)
    window.show()
    sys.exit(app.exec_())
