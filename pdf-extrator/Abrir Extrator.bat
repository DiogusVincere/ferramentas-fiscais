@echo off
rem Abre a tela do Extrator de PDF sem janela de console.
cd /d "%~dp0"
where pythonw >nul 2>nul && (start "" pythonw "%~dp0extrator.pyw" & exit /b)
where pyw >nul 2>nul && (start "" pyw "%~dp0extrator.pyw" & exit /b)
echo Python nao encontrado neste computador. Instale o Python 3 para usar o Extrator.
pause
