<#
.SYNOPSIS
    Tests des fonctions du projet

.DESCRIPTION
    Exemples de tests pour valider les fonctions

.EXAMPLE
    .\Tests\Test-Helper.ps1
#>

# Charger les fonctions à tester
. ../Functions/Helper.ps1

Write-Host "`n╔════════════════════════════════════════╗" -ForegroundColor Yellow
Write-Host "║          Tests - Projet Iveco         ║" -ForegroundColor Yellow
Write-Host "╚════════════════════════════════════════╝`n" -ForegroundColor Yellow

# Test 1: Get-WelcomeMessage
Write-Host "Test 1: Get-WelcomeMessage" -ForegroundColor Cyan
try {
    Get-WelcomeMessage -Message "Test message"
    Write-Host "✓ Test 1 réussi`n" -ForegroundColor Green
} catch {
    Write-Host "✗ Test 1 échoué: $_`n" -ForegroundColor Red
}

# Test 2: Get-SystemInfo
Write-Host "Test 2: Get-SystemInfo" -ForegroundColor Cyan
try {
    Get-SystemInfo
    Write-Host "✓ Test 2 réussi`n" -ForegroundColor Green
} catch {
    Write-Host "✗ Test 2 échoué: $_`n" -ForegroundColor Red
}

Write-Host "╔════════════════════════════════════════╗" -ForegroundColor Yellow
Write-Host "║      Tous les tests sont terminés      ║" -ForegroundColor Yellow
Write-Host "╚════════════════════════════════════════╝`n" -ForegroundColor Yellow
