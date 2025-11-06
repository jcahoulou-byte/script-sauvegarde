# 🧠 Forcer l'encodage UTF-8 pour l'affichage
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# 📁 Se placer dans le dossier du script
cd "$PSScriptRoot"

# 🕒 Horodatage
$timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
Write-Host "$timestamp - Push automatique lancé"

# 🔍 Vérifier les modifications locales
$gitStatus = git status --porcelain

if ($gitStatus) {
    # 🧹 Ajouter et commit les modifications
    git add .
    git commit -m "Commit automatique à $timestamp"
    git push origin master

    # ✅ Confirmation console et log
    Write-Host "$timestamp - Modifications poussées vers GitHub"
    Add-Content "$env:USERPROFILE\push-log.txt" "$timestamp - Modifications poussées"
} else {
    # 🟡 Aucun changement
    Write-Host "$timestamp - Aucun changement à pousser"
    Add-Content "$env:USERPROFILE\push-log.txt" "$timestamp - Aucun changement"
}