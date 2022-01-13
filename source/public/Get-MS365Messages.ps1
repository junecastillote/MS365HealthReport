# Function to get Office 365 Messages
Function Get-MS365Messages {
    [cmdletbinding()]
    param(
        [parameter(Mandatory, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Token,

        [parameter()]
        [ValidateNotNullOrEmpty()]
        [string[]]$Workload,

        [parameter()]
        [datetime]$LastUpdatedTime,

        [parameter()]
        [ValidateSet('Ongoing', 'Resolved')]
        [string]$Status
    )

    $header = @{'Authorization' = "Bearer $($Token)" }
    # $uri = "https://manage.office.com/api/v1.0/$($tenantID)/ServiceComms/Messages"
    $uri = "https://graph.microsoft.com/beta/admin/serviceAnnouncement/issues"

    if ($Status -eq 'Ongoing') {
        $uri = "$uri`?`$filter=isResolved ne true"
    }

    if ($Status -eq 'Resolved') {
        $uri = "$uri`?`$filter=isResolved eq true"
    }

    if ($LastUpdatedTime) {
        $lastModifiedDateTime = Get-Date ($LastUpdatedTime.ToUniversalTime()) -UFormat "%Y-%m-%dT%RZ"
        $uri = "$uri and lastModifiedDateTime ge $lastModifiedDateTime"
    }

    if ($Workload) {
        $uri = "$uri and (service eq '$($Workload[0])'"
        for ($i = 1; $i -lt $workload.Count; $i++) {
            $uri = "$uri or service eq '$($Workload[$i])'"
        }
        $uri = "$uri)"
    }

    SayInfo "Query = $uri"

    try {
        $result = @((Invoke-RestMethod -Uri $uri -Headers $header -Method Get -ErrorAction Stop).value)
        return $result
    }
    catch {
        SayError "$($_.Exception.Message)"
        return $null
    }
}