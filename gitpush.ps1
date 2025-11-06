[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
cd "$PSScriptRoot"
$timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
Write-Host "$timestamp - Push GitHub lancé"
$logPath = "$env:USERPROFILE\git-log.txt"

$gitStatus = git status --porcelain

if ($gitStatus) {
    try {
        git add .
        git commit -m "Commit automatique à $timestamp"
        git push origin master
        Write-Host "$timestamp - Modifications poussées vers GitHub"
        Add-Content $logPath "$timestamp - GitHub push OK"
    } catch {
        Write-Host "$timestamp - ❌ Échec du git push : $($_.Exception.Message)"
        Add-Content $logPath "$timestamp - GitHub push FAILED"
    }
} else {
    Write-Host "$timestamp - Aucun changement à pousser"
    Add-Content $logPath "$timestamp - Aucun changement"
}