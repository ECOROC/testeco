@echo off
chcp 65001 >nul
title DPGF Comparator — Configuration clé API
cls

echo.
echo  ╔══════════════════════════════════════════════════════╗
echo  ║       Configuration de la clé API Anthropic          ║
echo  ╚══════════════════════════════════════════════════════╝
echo.
echo  Votre clé API permet le matching sémantique intelligent.
echo  Coût : quelques centimes par analyse d'appel d'offres.
echo.
echo  Où trouver votre clé : https://console.anthropic.com
echo  (Rubrique "API Keys" dans votre compte)
echo.

set /p CLE="  Collez votre clé API ici et appuyez sur Entrée : "

if "%CLE%"=="" (
    echo.
    echo  Aucune clé saisie. Annulation.
    pause
    exit /b 0
)

:: Vérifier que ça ressemble à une clé Anthropic
echo %CLE% | findstr /b "sk-" >nul
if errorlevel 1 (
    echo.
    echo  ⚠ Cette clé ne ressemble pas à une clé Anthropic (elle devrait commencer par sk-).
    echo  Vérifiez et relancez ce fichier.
    pause
    exit /b 1
)

:: Écrire dans .env
echo ANTHROPIC_API_KEY=%CLE% > .env
echo.
echo  ✓ Clé API enregistrée avec succès dans le fichier .env
echo.
echo  Vous pouvez maintenant lancer DPGF Comparator.
echo  Le matching sémantique Claude sera activé automatiquement.
echo.
pause
