# Mermaid on PowerPoint を Windows の PowerPoint に登録する。
#   右クリック →「PowerShell で実行」、または  powershell -ExecutionPolicy Bypass -File install.ps1
#   取り消すときは  install.ps1 -Uninstall
#
# 行うこと（Microsoft の「共有フォルダーからの読み込み」と同じ設定）:
#   1. manifest.xml を %LOCALAPPDATA%\MermaidOnPptx\manifest に保存する
#   2. そのフォルダーを、実行した本人だけが読める共有フォルダーにする（ここだけ管理者の確認が出る）
#   3. HKCU の Office\16.0\WEF\TrustedCatalogs に共有フォルダーを登録する
param([switch]$Uninstall)

$ErrorActionPreference = "Stop"

$BaseUrl   = "__BASE_URL__"
$Dir       = Join-Path $env:LOCALAPPDATA "MermaidOnPptx\manifest"
$ShareName = "MermaidOnPptx"
$Unc       = "\\$env:COMPUTERNAME\$ShareName"
# 何度実行しても同じ登録を上書きするよう、固定の GUID を使う
$CatalogId = "{7b0f3c52-1d4e-4a8b-9e63-5c2a7d1f0b84}"
$CatalogKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$CatalogId"

function Invoke-Elevated([string]$Command) {
  # 共有の作成・削除だけ管理者権限で行う
  $p = Start-Process powershell -Verb RunAs -Wait -PassThru -ArgumentList "-NoProfile", "-Command", $Command
  if ($p.ExitCode -ne 0) { throw "管理者権限での実行に失敗しました（終了コード $($p.ExitCode)）" }
}

if ($Uninstall) {
  if (Test-Path $CatalogKey) { Remove-Item $CatalogKey -Recurse }
  if (Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue) {
    Invoke-Elevated "Remove-SmbShare -Name '$ShareName' -Force"
  }
  if (Test-Path $Dir) { Remove-Item (Split-Path $Dir) -Recurse -Force }
  Write-Host "取り消しました。PowerPoint を開き直してください。"
  return
}

New-Item -ItemType Directory -Force -Path $Dir | Out-Null
Invoke-WebRequest -UseBasicParsing -Uri "$BaseUrl/manifest.xml" -OutFile (Join-Path $Dir "manifest.xml")
Write-Host "manifest.xml を保存しました: $Dir"

if (-not (Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue)) {
  $user = "$env:USERDOMAIN\$env:USERNAME"
  Invoke-Elevated "New-SmbShare -Name '$ShareName' -Path '$Dir' -ReadAccess '$user' | Out-Null"
}
Write-Host "共有フォルダー: $Unc"

New-Item -Force -Path $CatalogKey | Out-Null
Set-ItemProperty -Path $CatalogKey -Name Id -Value $CatalogId
Set-ItemProperty -Path $CatalogKey -Name Url -Value $Unc
Set-ItemProperty -Path $CatalogKey -Name Flags -Value 1 -Type DWord
Write-Host "PowerPoint の信頼できるアドイン カタログに登録しました。"
Write-Host ""
Write-Host "PowerPoint を開き直し、ホーム → アドイン → 詳細 →「共有フォルダー」から Mermaid on PowerPoint を追加してください。"
