@echo off
chcp 65001 >nul
title DPGF Comparator — Installation
color 1F
cls

echo.
echo  ╔══════════════════════════════════════════════════════╗
echo  ║         DPGF COMPARATOR — Installation               ║
echo  ║         Economie de la construction                   ║
echo  ╚══════════════════════════════════════════════════════╝
echo.

:: Vérifier Python
echo  [1/4] Vérification de Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo  ╔══════════════════════════════════════════════╗
    echo  ║  Python n'est pas installe sur ce poste.    ║
    echo  ║                                              ║
    echo  ║  1. Allez sur : python.org/downloads         ║
    echo  ║  2. Téléchargez Python 3.11 ou 3.12          ║
    echo  ║  3. COCHEZ "Add Python to PATH"              ║
    echo  ║  4. Relancez ce fichier apres installation   ║
    echo  ╚══════════════════════════════════════════════╝
    echo.
    echo  Ouverture du site Python dans votre navigateur...
    start https://www.python.org/downloads/
    pause
    exit /b 1
)
for /f "tokens=*" %%i in ('python --version') do echo  ✓ %%i detecte

:: Installer les dépendances
echo.
echo  [2/4] Installation des dépendances Python...
echo  (Cette étape peut prendre 2 à 5 minutes, ne fermez pas cette fenêtre)
echo.
python -m pip install --upgrade pip --quiet
python -m pip install pandas openpyxl pdfplumber anthropic streamlit python-dotenv numpy --quiet
if errorlevel 1 (
    echo.
    echo  ERREUR lors de l'installation des dépendances.
    echo  Vérifiez votre connexion internet et relancez.
    pause
    exit /b 1
)
echo  ✓ Dépendances installées

:: Créer le fichier .env si absent
echo.
echo  [3/4] Configuration...
if not exist ".env" (
    copy ".env.example" ".env" >nul
    echo  ✓ Fichier de configuration créé (.env)
) else (
    echo  ✓ Configuration existante conservée
)

:: Créer le lanceur
echo.
echo  [4/4] Création du raccourci de lancement...
echo @echo off > "LANCER_DPGF_COMPARATOR.bat"
echo chcp 65001 ^>nul >> "LANCER_DPGF_COMPARATOR.bat"
echo title DPGF Comparator >> "LANCER_DPGF_COMPARATOR.bat"
echo echo. >> "LANCER_DPGF_COMPARATOR.bat"
echo echo  Démarrage de DPGF Comparator... >> "LANCER_DPGF_COMPARATOR.bat"
echo echo  Votre navigateur va s'ouvrir dans quelques secondes. >> "LANCER_DPGF_COMPARATOR.bat"
echo echo  Pour arrêter l'application : fermez cette fenêtre. >> "LANCER_DPGF_COMPARATOR.bat"
echo echo. >> "LANCER_DPGF_COMPARATOR.bat"
echo streamlit run app.py --server.headless false >> "LANCER_DPGF_COMPARATOR.bat"
echo  ✓ Raccourci créé : LANCER_DPGF_COMPARATOR.bat

echo.
echo  ╔══════════════════════════════════════════════════════╗
echo  ║              INSTALLATION TERMINÉE !                 ║
echo  ║                                                      ║
echo  ║  Double-cliquez sur :                                ║
echo  ║  >> LANCER_DPGF_COMPARATOR.bat                      ║
echo  ║                                                      ║
echo  ║  Pour le matching IA : éditez le fichier .env        ║
echo  ║  et renseignez votre clé ANTHROPIC_API_KEY           ║
echo  ╚══════════════════════════════════════════════════════╝
echo.
pause
