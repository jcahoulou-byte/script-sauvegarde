# 🧠 Encodage UTF-8 pour affichage correct
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# 📁 Se placer dans le dossier du script
cd "$PSScriptRoot"

# 🕒 Horodatage
$timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
Write-Host "$timestamp - Synchronisation automatique lancée"
Write-Host "Dossier de travail : $PSScriptRoot"

# 📦 Initialiser le journal
$logPath = "$env:USERPROFILE\sync-log.txt"
$syncResult = "$timestamp - "

# 🔄 Git pull
try {
    git pull origin master
    Write-Host "$timestamp - Git pull terminé"
} catch {
    Write-Host "$timestamp - ❌ Échec du git pull : $($_.Exception.Message)"
    Add-Content $logPath "$timestamp - Git pull FAILED"
}

# 📜 Lire le scriptId depuis .clasp.json
$claspConfigPath = Join-Path $PSScriptRoot ".clasp.json"
if (Test-Path $claspConfigPath) {
    try {
        $claspJson = Get-Content $claspConfigPath | ConvertFrom-Json
        $scriptId = $claspJson.scriptId
        Write-Host "Projet Apps Script cible : $scriptId"
    } catch {
        Write-Host "❌ Erreur lecture .clasp.json : $($_.Exception.Message)"
    }
} else {
    Write-Host "❌ Fichier .clasp.json introuvable"
}

# 🔍 Vérifier les modifications locales
$gitStatus = git status --porcelain

if ($gitStatus) {
    try {
        git add .
        git commit -m "Commit automatique à $timestamp"
        git push origin master
        Write-Host "$timestamp - Modifications poussées vers GitHub"
        $syncResult += "GitHub OK / "

    } catch {
        Write-Host "$timestamp - ❌ Échec du git push : $($_.Exception.Message)"
        $syncResult += "GitHub FAILED / "
    }

    try {
        clasp push
        Write-Host "$timestamp - Modifications poussées vers Apps Script"
        $syncResult += "Apps Script OK"
    } catch {
        Write-Host "$timestamp - ❌ Échec du clasp push : $($_.Exception.Message)"
        $syncResult += "Apps Script FAILED"
    }
} else {
    Write-Host "$timestamp - Aucun changement à pousser"
    $syncResult += "No changes"
}

# 📝 Enregistrer dans le journal
Add-Content $logPath $syncResult