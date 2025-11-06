# 🧠 Forcer l'encodage UTF-8 pour afficher correctement les accents et emojis
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# 📁 Se placer dans le dossier du script
cd "$PSScriptRoot"

# 🕒 Afficher l'heure d'exécution
Write-Host "🕒 Script lancé à $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

# 📂 Afficher le dossier de travail
Write-Host "Dossier de travail : $PSScriptRoot"

# 🔄 Synchroniser avec GitHub
git pull origin master

# 🔍 Vérifier s'il y a des modifications locales
$gitStatus = git status --porcelain

# 📜 Lire le fichier .clasp.json pour afficher le scriptId
$claspConfigPath = Join-Path $PSScriptRoot ".clasp.json"
if (Test-Path $claspConfigPath) {
    $claspJson = Get-Content $claspConfigPath | ConvertFrom-Json
    $scriptId = $claspJson.scriptId
    Write-Host "Projet Apps Script cible : $scriptId"
} else {
    Write-Host "❌ Fichier .clasp.json introuvable. Impossible d'afficher le scriptId."
}

# 🚀 Pousser vers Apps Script si des modifications sont détectées
if ($gitStatus) {
    Write-Host "Modifications détectées. Poussée vers Apps Script en cours..."
    clasp push
    $syncResult = "✅ Modifications poussées"
} else {
    Write-Host "Aucun changement détecté. Rien à pousser."
    $syncResult = "🟡 Aucun changement"
}

# 📝 Enregistrer l'exécution dans le fichier log
$logPath = "$env:USERPROFILE\sync-log.txt"
$timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
Add-Content $logPath "$timestamp — $syncResult"