$ErrorActionPreference = 'Stop'
Write-Host '----------------------------------------------------------------------------------------' -ForegroundColor Cyan
Write-Host 'Enabling Teams sideloading. Please keep this window open until the process is complete.' -ForegroundColor Cyan
try {
    Connect-MicrosoftTeams | Out-Null
    Set-CsTeamsAppSetupPolicy -Identity Global -AllowSideLoading $true | Out-Null

    $policy = Get-CsTeamsAppSetupPolicy -Identity Global
    if ($policy.AllowSideLoading -ne $true) {
        throw 'Sideloading could not be verified as enabled.'
    }

    Write-Host ''
    Write-Host '------------------------------------------------------------' -ForegroundColor Cyan
    Write-Host '  Sideloading is enabled in the Global app setup policy.' -ForegroundColor Green
    Write-Host '  You can close this window.' -ForegroundColor Green    
    Write-Host '------------------------------------------------------------' -ForegroundColor Cyan
}
catch {
    Write-Host ''
    Write-Host '------------------------------------------------------------' -ForegroundColor Red
    Write-Host '  Sideloading could not be enabled or verified.' -ForegroundColor Red
    Write-Host '  Please run the script again.' -ForegroundColor Yellow
    Write-Host "  Details: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host '------------------------------------------------------------' -ForegroundColor Red
}
Read-Host
