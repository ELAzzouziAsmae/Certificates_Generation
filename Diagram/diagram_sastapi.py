from graphviz import Digraph

# Create a directed graph
dot = Digraph("FullStackAutomationAPI", format="svg")
dot.attr(rankdir="LR", size="10")

# Define styles
frontend_style = {'style': 'filled', 'fillcolor': '#AED6F1', 'shape': 'box'}
api_style = {'style': 'filled', 'fillcolor': '#D5F5E3', 'shape': 'box'}
service_style = {'style': 'filled', 'fillcolor': '#A9DFBF', 'shape': 'box'}
async_style = {'style': 'filled', 'fillcolor': '#F9E79F', 'shape': 'box'}
external_style = {'style': 'filled', 'fillcolor': '#D7DBDD', 'shape': 'box'}
deployment_style = {'style': 'dashed', 'color': 'gray'}

# Frontend
dot.node("Frontend", "Frontend (index.html, app.js)", **frontend_style)

# FastAPI Layer
dot.node("API", "FastAPI App (main.py)", **api_style)
dot.node("Controllers", "Controllers\n(main_controller, customizer_controller, mpc_controller)", **api_style)

# Service Layer
dot.node("JiraService", "JiraService", **service_style)
dot.node("FileService", "FileService", **service_style)
dot.node("DownloadSceConfigService", "DownloadSceConfigService", **service_style)
dot.node("AutoSceService", "AutoSceService", **service_style)
dot.node("CustomizerService", "CustomizerService", **service_style)
dot.node("SceService", "SceService", **service_style)

# Async Task System
dot.node("Redis", "Redis Broker", **async_style)
dot.node("Worker", "Celery Worker\n(worker.py)", **async_style)
dot.node("Tasks", "Celery Tasks\n(tasks.py)", **async_style)

# External Systems
dot.node("Jira", "Jira REST API", **external_style)
dot.node("HyperV", "Hyper-V Batch\n(autosce.exe)", **external_style)
dot.node("FileSystem", "Local Filesystem\n(.mpc, logs, conf)", **external_style)

# Connections
dot.edge("Frontend", "API", label="HTTP")
dot.edge("API", "Controllers", label="Route")
dot.edge("Controllers", "Redis", label="Enqueue Task")
dot.edge("Redis", "Worker", label="Celery Queue")
dot.edge("Worker", "Tasks", label="Run Task Logic")

# Worker to services
dot.edge("Tasks", "JiraService")
dot.edge("Tasks", "FileService")
dot.edge("Tasks", "DownloadSceConfigService")
dot.edge("Tasks", "AutoSceService")
dot.edge("Tasks", "CustomizerService")
dot.edge("Tasks", "SceService")

# Service to external
dot.edge("JiraService", "Jira", label="REST Call")
dot.edge("AutoSceService", "HyperV", label="Batch")
dot.edge("FileService", "FileSystem", label="Write File")
dot.edge("DownloadSceConfigService", "FileSystem", label="Store Config")
dot.edge("Worker", "FileSystem", label="Logs/Artifacts")

# Docker Compose Box
with dot.subgraph(name="cluster_docker") as c:
    c.attr(label="Docker Compose", **deployment_style)
    c.node("API")
    c.node("Redis")
    c.node("Worker")

# 📍 Spécifie le chemin de sortie (modifie si nécessaire)
output_path = r"C:\Users\260001889\Desktop\FastAPI_architecture"

# Génère le diagramme
dot.render(output_path, format="svg", cleanup=True)

print(f"✅ Diagramme généré : {output_path}.svg")
