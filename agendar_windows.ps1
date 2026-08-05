<#
    Registra o robo no Agendador de Tarefas do Windows.
    Alternativa local ao agendamento na nuvem (GitHub Actions).

    Uso:
        .\agendar_windows.ps1              # agenda 09:00 e 16:00, dias uteis
        .\agendar_windows.ps1 -Horarios 09:00
        .\agendar_windows.ps1 -Remover     # remove as tarefas criadas

    Nao precisa de administrador: as tarefas ficam no seu usuario.
#>
param(
    [string[]] $Horarios = @('09:00', '16:00'),
    [switch]   $Remover
)

$ErrorActionPreference = 'Stop'
$bat = Join-Path $PSScriptRoot 'executar_robo.bat'
$prefixo = 'Robo CEAGESP Pescado'

if ($Remover) {
    Get-ScheduledTask | Where-Object { $_.TaskName -like "$prefixo*" } | ForEach-Object {
        Unregister-ScheduledTask -TaskName $_.TaskName -Confirm:$false
        Write-Host "Removida: $($_.TaskName)"
    }
    return
}

if (-not (Test-Path $bat)) { throw "Nao encontrei $bat" }

foreach ($h in $Horarios) {
    $nome = "$prefixo - $h"
    $acao = New-ScheduledTaskAction -Execute $bat -WorkingDirectory $PSScriptRoot
    $gatilho = New-ScheduledTaskTrigger -Weekly -At $h `
        -DaysOfWeek Monday, Tuesday, Wednesday, Thursday, Friday
    # StartWhenAvailable: se o PC estiver desligado no horario, roda assim que ligar.
    $config = New-ScheduledTaskSettingsSet -StartWhenAvailable `
        -ExecutionTimeLimit (New-TimeSpan -Minutes 30) -MultipleInstances IgnoreNew

    Register-ScheduledTask -TaskName $nome -Action $acao -Trigger $gatilho `
        -Settings $config -Description 'Coleta as cotacoes de pescado no atacado da CEAGESP' `
        -Force | Out-Null
    Write-Host "Agendada: $nome (seg a sex)"
}

Write-Host ''
Write-Host 'Conferir:  Get-ScheduledTask -TaskName "Robo CEAGESP Pescado*"'
Write-Host 'Rodar ja:  Start-ScheduledTask -TaskName "Robo CEAGESP Pescado - 09:00"'
Write-Host 'Remover:   .\agendar_windows.ps1 -Remover'
