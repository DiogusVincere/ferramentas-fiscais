@echo off
rem Abre o hub das Ferramentas Fiscais sem janela de console.
cd /d "%~dp0"
where pythonw >nul 2>nul && (start "" pythonw "%~dp0hub\hub.pyw" & exit /b)
where pyw >nul 2>nul && (start "" pyw "%~dp0hub\hub.pyw" & exit /b)
echo Python nao encontrado neste computador. Instale o Python 3 para usar as ferramentas.
pause
