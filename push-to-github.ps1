[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
cd "$PSScriptRoot"

$timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
Write-Host "🕒 Push automatique lancé à $timestamp"

$gitStatus = git status --porcelain

if ($gitStatus) {
    git add .
    git commit -m "🔄 Commit automatique à $timestamp"
    git push origin master
    Write-Host "✅ Modifications poussées vers GitHub"
    Add-Content "$env:USERPROFILE\push-log.txt" "$timestamp — ✅ Modifications poussées"
} else {
    Write-Host "🟡 Aucun changement à pousser"
    Add-Content "$env:USERPROFILE\push-log.txt" "$timestamp — 🟡 Aucun changement"
}