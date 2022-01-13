Function Get-MS365HealthOverview {
    [cmdletbinding(DefaultParameterSetName = 'All')]
    param (
        [Parameter(Mandatory)]
        [string]
        $Token,

        # List of services to retrieve
        [Parameter()]
        [string[]]$Workload
    )
    $headers = @{"Authorization" = "Bearer $($Token)" }
    $uri = "https://graph.microsoft.com/beta/admin/serviceAnnouncement/healthOverviews"

    if ($Workload) {
        $uri = "$uri`?`$filter=service eq '$($Workload[0])'"
        for ($i = 1; $i -lt $workload.Count; $i++) {
            $uri = "$uri or service eq '$($Workload[$i])'"
        }
    }

    SayInfo "Query = $uri"

    try {
        $response = @((Invoke-RestMethod -Uri $uri -Method GET -Headers $headers -ErrorAction STOP).value)
    }
    catch {
        SayError "$($_.Exception.Message)"
        return $null
    }

    return $($response | Sort-Object service)

}