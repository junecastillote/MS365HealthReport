# Function to get Office 365 Messages
Function Get-M365Messages {
    [cmdletbinding()]
    param(
        [parameter(Mandatory, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Token,

        [parameter(Position = 2)]
        [ValidateSet('Incident', 'PlannedMaintenance', 'MessageCenter')]
        [string]$MessageType,

        [parameter()]
        [ValidateNotNullOrEmpty()]
        [string[]]$Workload,

        [parameter()]
        [datetime]$LastUpdatedTime
    )

    $tenantID = ($token | Get-JWTDetails).tid
    $header = @{'Authorization' = "Bearer $($Token)" }
    $uri = "https://manage.office.com/api/v1.0/$($tenantID)/ServiceComms/Messages"

    if ($MessageType) {
        $uri = $uri+"`?`$filter=MessageType eq `'$MessageType`'"
    }

    $messages = (Invoke-RestMethod -Uri $uri -Headers $header -Method Get -ContentType 'application/json')

    $result = @()
    # Filter by workload if $workload is specified
    if ($Workload) {
        foreach ($message in ($messages.value | Where-Object { $_.Workload -in $Workload })) {
            $result += $message
        }
    }
    else {
        $result = $messages.value
    }

    if ($LastUpdatedTime) {
        $result = $result | Where-Object {[datetime]$_.LastUpdatedTime -ge $LastUpdatedTime}
    }

    return $result
}