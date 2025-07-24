from graphviz import Digraph

# Crée un nouveau graph
dot = Digraph("CertificatesAutomationApp", format="svg")
dot.attr(rankdir="LR", size="12")

# Définir les styles
ui_style = {'style': 'filled', 'fillcolor': '#AED6F1', 'shape': 'ellipse'}
controller_style = {'style': 'filled', 'fillcolor': '#D5F5E3', 'shape': 'box'}
worker_style = {'style': 'filled', 'fillcolor': '#FAD7A0', 'shape': 'box'}
data_style = {'style': 'filled', 'fillcolor': '#D6DBDF', 'shape': 'cylinder'}
template_style = {'style': 'filled', 'fillcolor': '#F9E79F', 'shape': 'note'}
external_style = {'style': 'filled', 'fillcolor': '#F5B7B1', 'shape': 'box'}  # 'cloud' provoque une erreur
log_style = {'style': 'filled', 'fillcolor': '#D2B4DE', 'shape': 'folder'}
deployment_style = {'style': 'filled', 'fillcolor': '#F5CBA7', 'shape': 'component'}

# UI Layer
dot.node("LoginView", "LoginView\n(login_view.py)", **ui_style)
dot.node("CertView", "CertView\n(cert_view.py)", **ui_style)
dot.node("MainWindow", "MainWindow\n(main_window.py)", **ui_style)

# Controller
dot.node("Controller", "Main Controller\n(main.py)", **controller_style)

# Worker / Business Logic
dot.node("CertWorker", "CertWorker\n(cert_worker.py)", **worker_style)

# Data Source and Templates
dot.node("ExcelData", "Excel Input\nFORM_*.xlsx", **data_style)
dot.node("PPTTemplate", "PPTX Template\n(template_certificat.pptx)", **template_style)

# External Services
dot.node("PowerPointCOM", "PowerPoint COM\n(pptx → pdf)", **external_style)
dot.node("OutlookCOM", "Outlook COM\n(send email)", **external_style)

# Logging
dot.node("LogFile", "Log File\ncertificates_generation.log", **log_style)

# Deployment
dot.node("PyInstaller", "PyInstaller\n(admin.manifest, version.txt)", **deployment_style)

# Arrows / Connections
dot.edge("LoginView", "Controller", label="User Input")
dot.edge("CertView", "Controller", label="Generation Params")
dot.edge("MainWindow", "Controller", label="App Init")
dot.edge("Controller", "CertWorker", label="Start Thread")
dot.edge("CertWorker", "ExcelData", label="Read Data")
dot.edge("CertWorker", "PPTTemplate", label="Generate PPT")
dot.edge("CertWorker", "PowerPointCOM", label="Export to PDF")
dot.edge("CertWorker", "OutlookCOM", label="Send Emails")
dot.edge("CertWorker", "LogFile", label="Write Logs")
dot.edge("PyInstaller", "Controller", label="Packaged Entry Point")

# Groupes / Swimlanes
with dot.subgraph(name="cluster_ui") as c:
    c.attr(label="GUI Layer (PyQt5)", style='dashed')
    c.node("LoginView")
    c.node("CertView")
    c.node("MainWindow")

with dot.subgraph(name="cluster_logic") as c:
    c.attr(label="Business Logic", style='dashed')
    c.node("CertWorker")

with dot.subgraph(name="cluster_resources") as c:
    c.attr(label="Resources & Services", style='dashed')
    c.node("ExcelData")
    c.node("PPTTemplate")
    c.node("PowerPointCOM")
    c.node("OutlookCOM")
    c.node("LogFile")

# 📍 Spécifie le chemin de sortie (modifie si nécessaire)
output_path = r"C:\Users\260001889\Desktop\certificates_automation_architecture"

# Génère le diagramme
dot.render(output_path, format="svg", cleanup=True)

print(f"✅ Diagramme généré : {output_path}.svg")
