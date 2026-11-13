# Gera Coleta.xlam (macro + guia "Assinaturas") a partir do Coleta.bas.
# Requer Python no PC de quem gera e, no Excel: Central de Confiabilidade > Configuracoes de Macro >
# "Confiar no acesso ao modelo de objeto do projeto do VBA".
$ErrorActionPreference = 'Stop'
$raiz = Split-Path $PSScriptRoot
$tmp = Join-Path $env:TEMP 'Coleta_sem_ui.xlam'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Add()
    [void]$wb.VBProject.VBComponents.Import((Join-Path $raiz 'Coleta.bas'))
    $wb.SaveAs($tmp, 55)   # 55 = suplemento do Excel (.xlam)
    $wb.Close($false)
} finally {
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
python (Join-Path $PSScriptRoot 'ribbon.py') $tmp (Join-Path $raiz 'Coleta.xlam')
Remove-Item $tmp
