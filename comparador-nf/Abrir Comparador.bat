@echo off
rem Abre a tela do Comparador de Notas Fiscais sem janela de console.
cd /d "%~dp0"
where pythonw >nul 2>nul && (start "" pythonw "%~dp0comparador.pyw" & exit /b)
where pyw >nul 2>nul && (start "" pyw "%~dp0comparador.pyw" & exit /b)
echo Python nao encontrado neste computador. Instale o Python 3 para usar o Comparador.
pause
