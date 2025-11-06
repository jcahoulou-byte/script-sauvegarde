cd "$PSScriptRoot"
Write-Host "Dossier de travail : $PSScriptRoot"

git pull origin master

$gitStatus = git status --porcelain

$claspConfigPath = Join-Path $PSScriptRoot ".clasp.json"
if (Test-Path $claspConfigPath) {
    $claspJson = Get-Content $claspConfigPath | ConvertFrom-Json
    $scriptId = $claspJson.scriptId
    Write-Host "Projet Apps Script cible : $scriptId"
} else {
    Write-Host "Fichier .clasp.json introuvable. Impossible d'afficher le scriptId."
}

if ($gitStatus) {
    Write-Host "Modifications detectees. Poussee vers Apps Script en cours..."
    clasp push
} else {
    Write-Host "Aucun changement detecte. Rien a pousser."
}