# HotspotOn.ps1

$LogFile = Join-Path $PSScriptRoot "HotspotOn.log"

function Write-Log {
    param([string]$Message)

    Add-Content -Path $LogFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $Message" -Encoding UTF8
}

Write-Log "========== 시작 =========="

try {

    Write-Log "프로필 검색"

    $profile = [Windows.Networking.Connectivity.NetworkInformation,Windows.Networking.Connectivity,ContentType=WindowsRuntime]::GetConnectionProfiles() |
        Where-Object { $_.ProfileName -eq "이더넷" }

    if ($null -eq $profile) {
        Write-Log "오류 : '이더넷' 프로필을 찾을 수 없습니다."
        exit 1
    }

    Write-Log "프로필 발견 : $($profile.ProfileName)"

    $tether =
        [Windows.Networking.NetworkOperators.NetworkOperatorTetheringManager,Windows.Networking.NetworkOperators,ContentType=WindowsRuntime]::CreateFromConnectionProfile($profile)

    Write-Log "StartTetheringAsync 호출"

    $null = $tether.StartTetheringAsync()

    Start-Sleep -Seconds 2

    Write-Log "현재 상태 : $($tether.TetheringOperationalState)"

}
catch {

    Write-Log "예외 발생 : $($_.Exception.Message)"

}

Write-Log "========== 종료 =========="