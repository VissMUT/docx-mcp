# -*- coding: utf-8 -*-
# Деплой doc-editor на VM (v1smut@localhost):
#   .\deploy.ps1
# Правки: word_document_server/** → scp → рестарт сервиса (образ не пересобирать,
# каталог монтируется); Dockerfile → пересборка.
param([switch]$SkipTests)

$ErrorActionPreference = 'Stop'
$VM = 'v1smut@localhost'
$VM_DIR = '/home/v1smut/doc-editor'
$SSH = "ssh -o BatchMode=yes -o StrictHostKeyChecking=no $VM"
$SCP = "scp -o BatchMode=yes -o StrictHostKeyChecking=no"

function Base64-Bash([string]$script) {
    $b64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($script))
    return "echo $b64 | base64 -d | bash"
}

if (-not $SkipTests) {
    python -m py_compile word_mcp_server.py word_document_server\main.py word_document_server\tools\storage_tools.py
    if (-not $?) { throw 'py_compile failed' }
}

# Файлы на VM (весь каталог кода)
$s = @"
mkdir -p $VM_DIR
"@
Invoke-Expression ($SSH + ' "' + (Base64-Bash $s) + '"')
foreach ($f in @('word_mcp_server.py', 'Dockerfile', 'doc-editor.service',
                 'requirements.txt', 'pyproject.toml', '__init__.py',
                 'server.json', 'manifest.json')) {
    if (Test-Path $f) {
        & cmd /c "$SCP `"$f`" ${VM}:$VM_DIR/$f 2>&1" | Out-Null
    }
}
foreach ($d in @('word_document_server', 'office_word_mcp_server')) {
    $s = @"
mkdir -p $VM_DIR/$d
"@
    Invoke-Expression ($SSH + ' "' + (Base64-Bash $s) + '"') | Out-Null
    & cmd /c "scp -r -o BatchMode=yes -o StrictHostKeyChecking=no `"$d\*`" ${VM}:$VM_DIR/$d/ 2>&1" | Out-Null
}

# Сборка + сервис
$s = @"
set -e
cd $VM_DIR
docker build -t doc-editor:latest . 2>&1 | tail -2
echo 'anDUBaft87@' | sudo -S cp doc-editor.service /etc/systemd/system/doc-editor.service
echo 'anDUBaft87@' | sudo -S systemctl daemon-reload
echo 'anDUBaft87@' | sudo -S systemctl enable doc-editor.service
echo 'anDUBaft87@' | sudo -S systemctl restart doc-editor.service
sleep 12
systemctl is-active doc-editor.service
curl -s -o /dev/null -w 'http %{http_code}\n' http://127.0.0.1:9113/sse -m 10
"@
Invoke-Expression ($SSH + ' "' + (Base64-Bash $s) + '"')
Write-Output 'DEPLOY OK'
