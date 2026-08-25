param(
    [Parameter(Mandatory=$true)][string]$ProjectId,
    [string]$DriveFileId = "12UDMaTEeqvNDaFH8sUDJy9E6YYLehI_J",
    [string]$Region = "southamerica-west1",
    [string]$ServiceName = "eai-bolivia-economic-monitor",
    [string]$ServiceAccountName = "eai-monitor"
)

$ErrorActionPreference = "Stop"
$ServiceAccountEmail = "$ServiceAccountName@$ProjectId.iam.gserviceaccount.com"

Write-Host "== EAI Bolivia Economic Monitor | Cloud Run ==" -ForegroundColor Cyan
Write-Host "Proyecto: $ProjectId"
Write-Host "Region:   $Region"
Write-Host "Servicio: $ServiceName"
Write-Host "Drive ID: $DriveFileId"

gcloud config set project $ProjectId | Out-Null
gcloud services enable run.googleapis.com cloudbuild.googleapis.com artifactregistry.googleapis.com iam.googleapis.com drive.googleapis.com

$saExists = $true
try {
    gcloud iam service-accounts describe $ServiceAccountEmail | Out-Null
} catch {
    $saExists = $false
}
if (-not $saExists) {
    gcloud iam service-accounts create $ServiceAccountName --display-name="EAI Bolivia Economic Monitor"
}

Write-Host ""
Write-Host "Comparte el Excel/Google Sheet con este correo como Lector/Viewer:" -ForegroundColor Yellow
Write-Host "  $ServiceAccountEmail" -ForegroundColor Green
Write-Host "No generes una llave JSON. Cloud Run usará Application Default Credentials."
Read-Host "Cuando el archivo ya esté compartido, presiona Enter para desplegar"

gcloud run deploy $ServiceName `
  --source . `
  --region $Region `
  --platform managed `
  --allow-unauthenticated `
  --service-account $ServiceAccountEmail `
  --set-env-vars "GOOGLE_DRIVE_FILE_ID=$DriveFileId,EAI_DRIVE_AUTH=adc" `
  --min 1 `
  --max 1 `
  --cpu 1 `
  --memory 1Gi `
  --timeout 3600

Write-Host ""
Write-Host "URL pública:" -ForegroundColor Cyan
gcloud run services describe $ServiceName --region $Region --format="value(status.url)"
