# Abre uma planilha pelo Excel (somente leitura, sem macros e sem atualizar vínculos)
# e salva uma cópia .xlsx em outro lugar. O arquivo de origem não é alterado.
param(
    [Parameter(Mandatory = $true)][string]$Origem,
    [Parameter(Mandatory = $true)][string]$Destino
)
$ErrorActionPreference = 'Stop'
$excel = $null
$livro = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 3          # desativa macros do arquivo aberto
    $livro = $excel.Workbooks.Open($Origem, 0, $true)   # UpdateLinks = 0, ReadOnly = $true
    $livro.SaveAs($Destino, 51)                          # 51 = pasta de trabalho .xlsx
    $livro.Close($false)
    $livro = $null
    exit 0
}
catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
}
finally {
    if ($livro) { try { $livro.Close($false) } catch {} }
    if ($excel) {
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
