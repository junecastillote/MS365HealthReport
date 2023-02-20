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
    $uri = 'https://graph.microsoft.com/beta/admin/serviceAnnouncement/issues'

    # If any (Workload, LastUpdatedTime, Status)
    if ($Workload -or $LastUpdatedTime -or $Status) {
        $uri = "$($uri)?filter="
    }

    # If event status is not yet resolved
    if ($Status -eq 'Ongoing') {
        $uri = "$($uri)isResolved ne true"
    }

    # If event status is resolved
    if ($Status -eq 'Resolved') {
        $uri = "$($uri)isResolved eq true"
    }

    # If LastUpdatedTime is specified
    if ($LastUpdatedTime) {
        $lastModifiedDateTime = Get-Date ($LastUpdatedTime.ToUniversalTime()) -UFormat "%Y-%m-%dT%RZ"
        if ($uri.Split('=')[1] -ne '') {
            $uri = "$uri and lastModifiedDateTime ge $lastModifiedDateTime"
        }
        else {
            $uri = "$($uri)lastModifiedDateTime ge $lastModifiedDateTime"
        }
    }

    # If -Workload [workload names[]]
    if ($Workload) {
        if ($uri.Split('=')[1] -ne '') {
            $uri = "$($uri) and (service eq '$($Workload[0])'"
        }
        else {
            $uri = "$uri(service eq '$($Workload[0])'"
        }
        # $uri = "$uri and (service eq '$($Workload[0])'"
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