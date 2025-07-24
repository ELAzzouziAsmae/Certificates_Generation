
## 📄 `README.md` – *Certificates Generation*

```markdown
# 📜 Certificates Generation - GE Vernova

Application Windows pour générer automatiquement des certificats de formation à partir d’un modèle PowerPoint et d’un fichier Excel, avec envoi automatique par Outlook.

---

## 🖼️ Aperçu

- ✅ Lecture des données depuis Excel
- ✅ Filtrage des participants selon score et période
- ✅ Génération automatique de certificats PowerPoint + conversion en PDF
- ✅ Envoi d'emails personnalisés via Outlook avec signature HTML et pièces jointes
- ✅ Interface graphique (GUI) moderne avec PyQt5
- ✅ Journalisation complète (`.log`)

---

## 📁 Arborescence du projet

project_root/
│
├── main.py                        # Fichier principal de l'application
├── Worker/
│   └── cert_worker.py            # Logique de génération de certificat (threadé)
├── view/
│   ├── login_view.py             # Interface login
│   └── cert_view.py              # Interface de génération
│
├── GEVernova_logo.jpg            # Logo affiché dans la fenêtre
├── app_icon.ico                  # Icône de la fenêtre principale
├── version.txt                   # Informations de version (PyInstaller)
├── admin.manifest                # Manifest Windows pour exécuter en admin
├── requirements.txt              # Dépendances Python
└── README.md                     

````

---

## ⚙️ Installation

### Prérequis

- Python 3.8 – 3.11 (recommandé : Python 3.10)
- Windows uniquement (COM + Outlook requis)
- Microsoft Office installé (PowerPoint + Outlook)

### Installer les dépendances

```bash
pip install -r requirements.txt
````

> ⚠️ Assurez-vous d'utiliser un environnement virtuel (`venv`) pour isoler les dépendances.

---

## 🚀 Utilisation

### Lancement (mode développement)

```bash
python main.py
```

---

## 🧪 Compilation en exécutable Windows

Utilisez `PyInstaller` pour transformer le script en `.exe` autoportant.

### 1. Assurez-vous que `pyinstaller` est installé

```bash
pip install pyinstaller
```

### 2. Commande de compilation complète

```bash
pyinstaller --onefile --windowed ^
  --add-data "GEVernova_logo.jpg;." ^
  --add-data "app_icon.ico;." ^
  --version-file="version.txt" ^
  --manifest="admin.manifest" ^
  main.py
```

### Résultat

* Un exécutable `dist/Certificates Generation.exe` sera généré.
* Il inclura :

  * Logo
  * Icône
  * Droits administrateurs via `admin.manifest`
  * Informations de version via `version.txt`

---

## 🧾 Fichier `version.txt`

```python
# UTF-8
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=(1,1,0,3),
    prodvers=(1,1,0,3),
    mask=0x3f,
    flags=0x0,
    OS=0x4,
    fileType=0x1,
    subtype=0x0,
    date=(0,0)
  ),
  kids=[
    StringFileInfo([
      StringTable(
        '040904B0',
        [StringStruct('CompanyName', 'GE VERNOVA - Grid Solutions'),
         StringStruct('FileDescription', 'Certificates Generation'),
         StringStruct('FileVersion', '1.1.1'),
         StringStruct('InternalName', 'Certificates Generation'),
         StringStruct('LegalCopyright', '© 2025 GE Vernova - Grid Solutions. All rights reserved.'),
         StringStruct('OriginalFilename', 'Certificates Generation.exe'),
         StringStruct('ProductName', 'Certificates Generation'),
         StringStruct('ProductVersion', '1.1.0')]
      )
    ]),
    VarFileInfo([VarStruct('Translation', [1033, 1200])])
  ]
)
```

---

## 🔒 Manifest `admin.manifest`

Permet d’exécuter l'application avec les droits administrateur.

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
  <trustInfo xmlns="urn:schemas-microsoft-com:asm.v3">
    <security>
      <requestedPrivileges>
        <requestedExecutionLevel level="requireAdministrator" uiAccess="false"/>
      </requestedPrivileges>
    </security>
  </trustInfo>
</assembly>
```

---

## 🧩 Dépendances

Liste des packages utilisés :

```txt
PyQt5>=5.15.0         # Interface graphique
pandas>=1.1.0         # Lecture et traitement Excel
openpyxl>=3.0.0       # Backend .xlsx pour pandas
pywin32>=227          # Contrôle Outlook & COM
comtypes>=1.1.7       # COM automation pour PowerPoint
python-pptx>=0.6.21   # PowerPoint slide edition
```

---

## 📧 Fonctionnalités Email

* Utilise Outlook via `win32com.client`
* Ajoute automatiquement la **signature utilisateur Outlook**
* Gère les **images de signature** intégrées (`cid:`)
* Ajoute le certificat PDF en pièce jointe
* CC automatique vers `SAM.L3SSystem@gevernova.com`

---

## 🔐 Authentification

* Page de login prévue
* Intégration future possible avec les credentials Windows (`win32security`) pour validation SSO

---

## 📜 Licence

© 2025 GE Vernova – Grid Solutions. All rights reserved.

---

## 🙋 Support

Pour toute question, contactez l'équipe IT Formation GE Vernova.

```

---

Souhaites-tu également que je t’ajoute :

- ✅ Un `build.bat` pour tout compiler en un double-clic ?
- ✅ Un fichier `.spec` personnalisé pour PyInstaller ?
- ✅ Un modèle de mail professionnel (Outlook) plus riche en HTML ?

Je peux aussi générer un installateur `.exe` via **Inno Setup** ou te faire le fichier `.iss`.
```
