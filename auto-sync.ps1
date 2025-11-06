# 🧠 Encodage UTF-8 pour affichage correct
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# 📁 Se placer dans le dossier du script
cd "$PSScriptRoot"

# 🕒 Horodatage
$timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
Write-Host "$timestamp - Synchronisation automatique lancée"

# 📂 Afficher le dossier de travail
Write-Host "Dossier de travail : $PSScriptRoot"

# 🔄 Synchroniser avec GitHub distant
git pull origin master

# 📜 Lire le fichier .clasp.json pour afficher le scriptId
$claspConfigPath = Join-Path $PSScriptRoot ".clasp.json"
if (Test-Path $claspConfigPath) {
    $claspJson = Get-Content $claspConfigPath | ConvertFrom-Json
    $scriptId = $claspJson.scriptId
    Write-Host "Projet Apps Script cible : $scriptId"
} else {
    Write-Host "❌ Fichier .clasp.json introuvable. Impossible d'afficher le scriptId."
}

# 🔍 Vérifier les modifications locales
$gitStatus = git status --porcelain

if ($gitStatus) {
    # 🧹 Ajouter et commit les modifications
    git add .
    git commit -m "Commit automatique à $timestamp"
    git push origin master
    Write-Host "$timestamp - Modifications poussées vers GitHub"

    # 🚀 Pousser vers Apps Script
    clasp push
    Write-Host "$timestamp - Modifications poussées vers Apps Script"

    $syncResult = "$timestamp - Modifications poussées / Changes pushed"
} else {
    Write-Host "$timestamp - Aucun changement à pousser"
    $syncResult = "$timestamp - Aucun changement / No changes"
}

# 📝 Enregistrer dans le journal
$logPath = "$env:USERPROFILE\sync-log.txt"
Add-Content $logPath $syncResult