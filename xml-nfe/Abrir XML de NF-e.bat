@echo off
rem Abre a tela de XML de NF-e sem janela de console.
cd /d "%~dp0"
where pythonw >nul 2>nul && (start "" pythonw "%~dp0xml_nfe.pyw" & exit /b)
where pyw >nul 2>nul && (start "" pyw "%~dp0xml_nfe.pyw" & exit /b)
echo Python nao encontrado neste computador. Instale o Python 3 para usar esta ferramenta.
pause
