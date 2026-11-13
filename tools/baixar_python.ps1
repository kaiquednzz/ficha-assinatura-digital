# Baixa o Python embeddable oficial para a pasta oculta python\ do projeto e remove o que o
# server.py nao usa (SSL, SQLite, ctypes, XML...): ~13 MB. Nenhum PC precisa instalar Python.
$ErrorActionPreference = 'Stop'
$versao = '3.12.10'
$dest = Join-Path (Split-Path $PSScriptRoot) 'python'
$zip = Join-Path $env:TEMP "python-$versao-embed-amd64.zip"

Invoke-WebRequest "https://www.python.org/ftp/python/$versao/python-$versao-embed-amd64.zip" -OutFile $zip -UseBasicParsing
if (Test-Path $dest) { Remove-Item $dest -Recurse -Force }
Expand-Archive $zip $dest
Remove-Item $zip

'_ssl.pyd', 'libssl-3.dll', 'libcrypto-3.dll', '_hashlib.pyd', '_sqlite3.pyd', 'sqlite3.dll', '_msi.pyd', '_wmi.pyd',
'_bz2.pyd', '_lzma.pyd', '_decimal.pyd', '_elementtree.pyd', 'pyexpat.pyd', '_ctypes.pyd', 'libffi-8.dll',
'_asyncio.pyd', '_overlapped.pyd', '_multiprocessing.pyd', 'winsound.pyd', '_zoneinfo.pyd', '_uuid.pyd', 'pythonw.exe' |
    ForEach-Object { Remove-Item (Join-Path $dest $_) -ErrorAction SilentlyContinue }

(Get-Item $dest).Attributes = 'Hidden,Directory'
& (Join-Path $dest 'python.exe') -c "import http.server, secrets, csv, json; print('python portatil ok')"
