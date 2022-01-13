Function Set-MS365HealthReportLastRunTime {
    [CmdletBinding()]
    param (
        [parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$TenantID,

        [parameter()]
        [datetime]$LastRunTime
    )
    $now = Get-Date
    $RegPath = "HKCU:\Software\MS365HealthReport\$TenantID"

    $regSplat = @{
        Path  = $RegPath
        Value = $(
            if ($LastRunTime) {
                "{0:yyyy-MM-dd H:mm}" -f $LastRunTime
            }
            else {
                $now
            }
        )
    }
    $null = New-Item @regSplat -Force
}