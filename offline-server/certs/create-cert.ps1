# Tạo chứng chỉ SSL cục bộ cho localhost (port 8443) và tin cậy cho User hiện tại
$ErrorActionPreference = "Stop"

$certDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$pfxPath = Join-Path $certDir "localhost.pfx"
$cerPath = Join-Path $certDir "localhost.cer"
$password = ConvertTo-SecureString -String "pcdt-offline" -Force -AsPlainText

Write-Host "1. Đang tạo chứng chỉ tự ký cho localhost..." -ForegroundColor Cyan
$cert = New-SelfSignedCertificate `
  -DnsName "localhost", "127.0.0.1" `
  -CertStoreLocation "Cert:\CurrentUser\My" `
  -FriendlyName "PCDT Word Add-in Offline Localhost" `
  -NotAfter (Get-Date).AddYears(5)

Write-Host "2. Xuất file PFX cho Node.js server..." -ForegroundColor Cyan
Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $password | Out-Null

Write-Host "3. Xuất file CER và cài vào Trusted Root (CurrentUser)..." -ForegroundColor Cyan
Export-Certificate -Cert $cert -FilePath $cerPath | Out-Null
Import-Certificate -FilePath $cerPath -CertStoreLocation "Cert:\CurrentUser\Root" | Out-Null

Write-Host "✅ Hoàn tất! Đã tạo $pfxPath và tin cậy chứng chỉ localhost." -ForegroundColor Green
