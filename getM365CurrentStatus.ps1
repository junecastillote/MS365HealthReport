# Function to List all current Office 365 service status
Function Get-m365Status {
    [cmdletbinding()]
    param(
        [parameter(Mandatory, Position = 0)]
        [ValidateNotNullOrEmpty()]
        $Token
    )
    $tenantID = ($token | Get-JWTDetails).tid
    $header = @{'Authorization' = "Bearer $($Token)" }
    $result = (Invoke-RestMethod -Uri "https://manage.office.com/api/v1.0/$($tenantID)/ServiceComms/CurrentStatus" -Headers $header -Method Get -ContentType 'application/json')
    return ($result.value)
}