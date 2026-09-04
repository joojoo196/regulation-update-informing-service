$ErrorActionPreference = 'Stop'

if (-not $env:LAW_API_OC -and -not (Test-Path -LiteralPath "$PSScriptRoot\.env")) {
    $env:LAW_API_OC = Read-Host '국가법령정보 API 인증키(OC)를 입력하세요'
}

$bundledNode = 'C:\Users\jojow\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe'
$node = Get-Command node -ErrorAction SilentlyContinue

if ($node) {
    & $node.Source "$PSScriptRoot\server.js"
} elseif (Test-Path -LiteralPath $bundledNode) {
    & $bundledNode "$PSScriptRoot\server.js"
} else {
    throw 'Node.js를 찾을 수 없습니다.'
}
