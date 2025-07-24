import sys
import os
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QWidget, QTableWidget, QTableWidgetItem,
    QLabel, QComboBox, QHBoxLayout, QSpinBox, QPushButton, QGridLayout
)
from PyQt5.QtWebEngineWidgets import QWebEngineView
from PyQt5.QtCore import QUrl
import plotly.express as px
from PyQt5.QtWidgets import QHeaderView

class KPIWindow(QMainWindow):
    def __init__(self, df, parent=None):
        super().__init__(parent)
        self.setWindowTitle("📊 Tableau de bord des KPI")
        self.df_original = df.copy()

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

        # Filtres
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

        self.export_html_btn = QPushButton("Exporter HTML")
        self.export_html_btn.clicked.connect(self.export_html)
        filter_layout.addWidget(self.export_html_btn)

        main_layout.addLayout(filter_layout)

        grid = QGridLayout()
        main_layout.addLayout(grid)

        self.table = QTableWidget()
        self.table.setStyleSheet("""
            QTableWidget {
                font-size: 11px;
                padding: 0px;
                margin: 0px;
                border: 1px solid #ccc;
                gridline-color: #ccc;
            }
            QTableWidget::item {
                padding: 2px;
                margin: 0px;
            }
            QHeaderView::section {
                padding: 2px;
                margin: 0px;
                font-size: 11px;
            }
        """)
        self.table.setWordWrap(False)
        self.table.setShowGrid(True)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table.setHorizontalScrollMode(QTableWidget.ScrollPerPixel)
        self.table.setMaximumHeight(400)

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
        df = self.df.copy()
        selected_year = self.annee_combo.currentText()
        if selected_year != "Toutes les années":
            df = df[df["Année"] == int(selected_year)]

        grouped = df.groupby(["Région", "SDSU Unit"]).agg(
            Tentatives=('Percent Score', 'count'),
            Réussites=('Percent Score', lambda x: (x >= seuil).sum())
        ).reset_index()

        # Calcul du taux réussite en % par région selon ta demande
        total_reussites = grouped['Réussites'].sum()
        region_summary = grouped.groupby("Région").agg({
            "Tentatives": "sum",
            "Réussites": "sum"
        }).reset_index()

        region_summary['Taux Réussite (%)'] = (
            region_summary['Réussites'] / total_reussites * 100
        ).round(1)

        # Merge pour avoir ce taux dans grouped
        grouped = grouped.merge(region_summary[['Région', 'Taux Réussite (%)']], on='Région', how='left')

        # Calcul du taux local par unité
        grouped['Taux Réussite par Unit (%)'] = (
            grouped['Réussites'] / grouped['Tentatives'] * 100
        ).round(1)

        # Ajout du taux par région dans le tableau unité
        grouped = grouped.merge(
            region_summary[["Région", "Taux Réussite (%)"]],
            on="Région",
            how="left"
        )
        display_df = grouped.rename(columns={"SDSU Unit": "De Unit"})[
            ["Région", "De Unit", "Tentatives", "Réussites", 
            "Taux Réussite par Unit (%)", "Taux Réussite par Région (%)"]
        ]


        display_df = grouped  # contient maintenant les colonnes nécessaires

        # Affichage du tableau
        self.table.setRowCount(len(display_df))
        self.table.setColumnCount(len(display_df.columns))
        self.table.setHorizontalHeaderLabels(display_df.columns.tolist())

        for i, row in display_df.iterrows():
            for j, val in enumerate(row):
                self.table.setItem(i, j, QTableWidgetItem(str(val)))


                filename = "kpi_export.csv"
                region_summary.to_csv(filename, index=False)
                print(f"✅ Export CSV réalisé: {os.path.abspath(filename)}")

    def export_html(self):
        try:
            # Vérifie que les graphiques et tableau existent
            if not hasattr(self, "last_fig_region") or not hasattr(self, "last_table_df") or not hasattr(self, "last_fig_unit") or not hasattr(self, "last_fig_pie"):
                print("❌ Les graphiques ne sont pas encore générés. Actualise d'abord la vue.")
                return

            # Génère le HTML des graphiques avec plotlyjs inclus uniquement une fois (ici inclus dans le premier)
            html_region = self.last_fig_region.to_html(include_plotlyjs=True, full_html=False)
            html_unit = self.last_fig_unit.to_html(include_plotlyjs=False, full_html=False)
            html_pie = self.last_fig_pie.to_html(include_plotlyjs=False, full_html=False)

            # HTML du tableau pandas
            table_html = self.last_table_df.to_html(index=False, classes="data-table")

            # Style CSS
            css = """
            <style>
                body { font-family: Arial; margin: 20px; background: #f9f9f9; }
                h1 { text-align: center; }
                .grid { display: grid; grid-template-columns: 1fr 1fr; grid-gap: 20px; }
                .cell { background: white; padding: 12px; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
                .data-table { width: 100%; border-collapse: collapse; font-size: 12.5px; }
                .data-table th, .data-table td { border: 1px solid #ccc; padding: 4px 6px; text-align: center; }
                .data-table th { background-color: #eee; }
            </style>
            """

            # Montage final de la page HTML
            html = f"""
            <html>
            <head>
                <meta charset="utf-8">
                <title>Dashboard KPI Export</title>
                <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
                {css}
            </head>
            <body>
                <h1>📊 Tableau de bord des KPI (Export HTML)</h1>
                <div class="grid">
                    <div class="cell">{table_html}</div>
                    <div class="cell">{html_region}</div>
                    <div class="cell">{html_unit}</div>
                    <div class="cell">{html_pie}</div>
                </div>
            </body>
            </html>
            """

            # Enregistre dans un fichier
            with open("dashboard_export.html", "w", encoding="utf-8") as f:
                f.write(html)

            print("✅ Export HTML généré : dashboard_export.html")

        except Exception as e:
            print("❌ Erreur dans export_html :", e)


    def update_view(self):
        seuil = self.seuil_spin.value()
        df = self.df.copy()
        selected_year = self.annee_combo.currentText()
        if selected_year != "Toutes les années":
            df = df[df["Année"] == int(selected_year)]

        # Group by région & unit
        grouped = df.groupby(["Région", "SDSU Unit"]).agg(
            Tentatives=('Percent Score', 'count'),
            Réussites=('Percent Score', lambda x: (x >= seuil).sum())
        ).reset_index()

        total_reussites = grouped['Réussites'].sum()

        # Taux par région = part de réussite région / réussite totale
        region_summary = grouped.groupby("Région").agg({
            "Tentatives": "sum",
            "Réussites": "sum"
        }).reset_index()

        region_summary['Taux Réussite (%)'] = (
            region_summary['Réussites'] / total_reussites * 100
        ).round(1)

        # Taux local par unité (réussite / tentative)
        grouped['Taux Réussite par Unit (%)'] = (
            grouped['Réussites'] / grouped['Tentatives'] * 100
        ).round(1)

        # Ajout de la colonne "Taux Réussite par Région (%)" à grouped
        grouped = grouped.merge(
            region_summary[["Région", "Taux Réussite (%)"]],
            on="Région",
            how="left"
        ).rename(columns={"Taux Réussite (%)": "Taux Réussite par Région (%)"})

        # Renommage pour affichage
        display_df = grouped.rename(columns={"SDSU Unit": "De Unit"})[
            ["Région", "De Unit", "Tentatives", "Réussites",
            "Taux Réussite par Unit (%)", "Taux Réussite par Région (%)"]
        ]

        # Affichage du tableau PyQt
        self.table.setRowCount(len(display_df))
        self.table.setColumnCount(len(display_df.columns))
        self.table.setHorizontalHeaderLabels(display_df.columns.tolist())
        for i, row in display_df.iterrows():
            for j, val in enumerate(row):
                self.table.setItem(i, j, QTableWidgetItem(str(val)))

        self.table.resizeColumnsToContents()


        # Graphique région : taux par région (part par rapport au total réussites global)
        fig_region = px.bar(
            region_summary,
            x="Région",
            y="Taux Réussite (%)",
            title="Taux de réussite par Région",
            color="Taux Réussite (%)",
            color_continuous_scale=px.colors.sequential.Viridis
        )
        tmp_file_region = "tmp_region.html"
        fig_region.write_html(tmp_file_region)
        self.web_region.load(QUrl.fromLocalFile(os.path.abspath(tmp_file_region)))

        # Graphique unité : taux local par unité (réussite / tentative * 100)
        fig_unit = px.bar(
            grouped,
            x="SDSU Unit",
            y="Taux Réussite par Unit (%)",
            title="Taux de réussite par SDSU Unit",
            color="Taux Réussite par Unit (%)",
            color_continuous_scale=px.colors.sequential.Plasma
        )
        fig_unit.update_yaxes(range=[0, 100])
        tmp_file_unit = "tmp_unit.html"
        fig_unit.write_html(tmp_file_unit)
        self.web_unit.load(QUrl.fromLocalFile(os.path.abspath(tmp_file_unit)))

        # Graphique pie région (taux par région)
        region_summary_for_pie = region_summary.copy()
        region_summary_for_pie.loc[region_summary_for_pie['Taux Réussite (%)'] == 0, 'Taux Réussite (%)'] = None

        fig_pie = px.pie(
            region_summary_for_pie,
            values="Taux Réussite (%)",
            names="Région",
            title="Taux de réussite (%) par Région",
            color_discrete_sequence=px.colors.qualitative.Bold
        )

        tmp_file_pie = "tmp_pie.html"
        fig_pie.write_html(tmp_file_pie)
        self.web_pie.load(QUrl.fromLocalFile(os.path.abspath(tmp_file_pie)))

        self.last_fig_region = fig_region
        self.last_fig_unit = fig_unit
        self.last_fig_pie = fig_pie
        self.last_table_df = grouped.rename(columns={
            "SDSU Unit": "De Unit"
        })[
            ["Région", "De Unit", "Tentatives", "Réussites",
            "Taux Réussite par Unit (%)", "Taux Réussite par Région (%)"]
        ]

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
