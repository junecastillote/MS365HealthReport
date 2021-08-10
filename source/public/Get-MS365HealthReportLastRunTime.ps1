Function Get-MS365HealthReportLastRunTime {
    [CmdletBinding()]
    param (
        [parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$TenantID
    )
    $now = Get-Date
    $RegPath = "HKCU:\Software\MS365HealthReport\$TenantID"

    try {
        $value = Get-ItemPropertyValue -Path $RegPath -Name "(default)" -ErrorAction Stop
        return $(Get-Date $value)
    }
    catch {
        Set-MS365HealthReportLastRunTime -TenantID $TenantID -LastRunTime $now
        return $now
    }
}