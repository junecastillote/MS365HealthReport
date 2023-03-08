Function New-MS365IncidentReport {
    [cmdletbinding(DefaultParameterSetName = 'Client Secret')]
    param (
        [parameter()]
        [string]
        $OrganizationName,

        [parameter(Mandatory, ParameterSetName = 'Client Certificate')]
        [parameter(Mandatory, ParameterSetName = 'Certificate Thumbprint')]
        [parameter(Mandatory, ParameterSetName = 'Client Secret')]
        [guid]
        $ClientID,

        [parameter(Mandatory, ParameterSetName = 'Client Secret')]
        [string]
        $ClientSecret,

        [parameter(Mandatory, ParameterSetName = 'Client Certificate')]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]
        $ClientCertificate,

        [parameter(Mandatory, ParameterSetName = 'Certificate Thumbprint')]
        [string]
        $ClientCertificateThumbprint,

        [parameter(Mandatory, ParameterSetName = 'Client Certificate')]
        [parameter(Mandatory, ParameterSetName = 'Certificate Thumbprint')]
        [parameter(Mandatory, ParameterSetName = 'Client Secret')]
        [string]
        $TenantID,

        [Parameter()]
        [switch]
        $StartFromLastRun,

        [Parameter()]
        [datetime]
        $LastUpdatedTime,

        [Parameter()]
        [string[]]
        $Workload,

        [parameter()]
        [ValidateSet('Ongoing', 'Resolved')]
        [string]$Status,

        [Parameter()]
        [switch]
        $SendEmail,

        [Parameter()]
        [string]
        $From,

        [Parameter()]
        [string[]]
        $To,

        [Parameter()]
        [string[]]
        $CC,

        [Parameter()]
        [string[]]
        $Bcc,

        [Parameter()]
        [boolean]
        $WriteReportToDisk = $true,

        [Parameter()]
        [boolean]
        $WriteRawJSONToDisk = $false,

        [Parameter()]
        [boolean]
        $Consolidate = $true,

        [Parameter()]
        [boolean]
        $SendTeams = $false,

        [Parameter()]
        [string[]]
        $TeamsWebHookURL,

        [Parameter()]
        [string]
        $RunHistoryFile
    )

    if ($StartFromLastRun -and $LastUpdatedTime) {
        SayWarning "Do not use -StartFromLastRun and -LastUpdatedTime at the same time."
        return $null
    }

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $moduleInfo = Get-Module $($MyInvocation.MyCommand.ModuleName)

    $errorFlag = $false

    # $WriteReportToDisk = $true
    # $WriteRawJSONToDisk = $true

    $now = Get-Date

    #Region Prepare Output Directory
    $outputDir = ([System.IO.Path]::Combine([environment]::getfolderpath('userprofile'), $($moduleInfo.Name), $($TenantID)))

    if (!(Test-Path -Path $outputDir)) {
        $null = New-Item -ItemType Directory -Path $outputDir -Force
    }

    else {
        Remove-Item -Path $outputDir\* -Exclude runHistory.csv -Force -Confirm:$false
    }
    SayInfo "Output Directory: $outputDir"

    #EndRegion

    # Set the run times history file if it isn't specified.
    if (!($RunHistoryFile)) {
        $runHistoryFile = ([System.IO.Path]::Combine($outputDir, "runHistory.csv" ))
    }

    # $runHistoryFile = ([System.IO.Path]::Combine($outputDir, "runHistory.csv" ))
    # Create the history file if it doesn't exist.
    # if (!(Test-Path $RunHistoryFile) -or !(Get-Content $RunHistoryFile -Raw -ErrorAction SilentlyContinue)) {
    if ((Test-Path $RunHistoryFile) -eq $false) {
        SayInfo "Creating $RunHistoryFile"
        "RunTime,Status" | Set-Content -Path $RunHistoryFile -Force -Confirm:$false
        # Add initial entry 'OK' (which means successful) dated 7 days ago. This way there will always be a starting point.
        "$("{0:yyyy-MM-dd H:mm}" -f $now.AddDays(-7)),OK" | Add-Content -Path $RunHistoryFile -Force -Confirm:$false
    }

    # $RunHistoryFile = (Resolve-Path $RunHistoryFile).Path


    if (!$OrganizationName) { $OrganizationName = $TenantID }

    SayInfo "Authentication type: $($pscmdlet.ParameterSetName)"
    SayInfo "Client ID: $ClientID"
    SayInfo "Tenant ID: $TenantID"

    # Get Service Communications API Token
    if ($pscmdlet.ParameterSetName -eq 'Client Secret') {
        $SecureClientSecret = New-Object System.Security.SecureString
        $ClientSecret.toCharArray() | ForEach-Object { $SecureClientSecret.AppendChar($_) }
        $OAuth = Get-MsalToken -ClientId $ClientID -ClientSecret $SecureClientSecret -TenantId $tenantID -ErrorAction Stop
        Sayinfo $($ClientSecret -replace $($ClientSecret.Substring(0, $ClientSecret.Length - 8)), $('X' * $($ClientSecret.Substring(0, $ClientSecret.Length - 8)).Length))
    }
    elseif ($pscmdlet.ParameterSetName -eq 'Client Certificate') {
        $OAuth = Get-MsalToken -ClientId $ClientID -ClientCertificate $ClientCertificate -TenantId $tenantID -ErrorAction Stop
    }
    elseif ($pscmdlet.ParameterSetName -eq 'Certificate Thumbprint') {
        if (Test-Path Cert:\CurrentUser\My\$($ClientCertificateThumbprint)) {
            $ClientCertificate = Get-Item Cert:\CurrentUser\My\$($ClientCertificateThumbprint) -ErrorAction Stop
        }
        elseif (Test-Path Cert:\LocalMachine\My\$($ClientCertificateThumbprint)) {
            $ClientCertificate = Get-Item Cert:\LocalMachine\My\$($ClientCertificateThumbprint) -ErrorAction Stop
        }
        else {
            SayError $_.Exception.Message
            return $null
        }
        $OAuth = Get-MsalToken -ClientId $ClientID -ClientCertificate $ClientCertificate -TenantId $tenantID -ErrorAction Stop
    }

    $GraphAPIHeader = @{'Authorization' = "Bearer $($OAuth.AccessToken)" }

    # Get GraphAPI Token
    if ($SendEmail) {
        if (!$From) { SayWarning "You ask me to send an email report but you forgot to add the -From address."; return $null }
        if (!$To) { SayWarning "You ask me to send an email report but you forgot to add the -To address(es)."; return $null }
    }

    #Region Get Incidents
    $searchParam = @{
        Token = ($OAuth.AccessToken);
    }

    if ($Status) {
        $searchParam += (@{Status = $Status })
    }

    ## If -StartFromLastRun, this function will only get the incidents whose LastUpdatedTime is after the timestamp in "$outputDir\runHistory.csv"
    if ($StartFromLastRun) {
        SayInfo "Getting last successful run time from $RunHistoryFile."
        [datetime]$LastUpdatedTime = @(Import-Csv $RunHistoryFile | Where-Object { $_.Status -eq 'Ok' })[-1].RunTime
    }

    ## If -LastUpdatedTime, this function will only get the incidents whose LastUpdatedTime is after the $LastUpdatedTime datetime value.
    if ($LastUpdatedTime) {
        $searchParam += (@{LastUpdatedTime = $LastUpdatedTime })
        SayInfo "Getting incidents from the last successful run time: $LastUpdatedTime"
    }

    if ($Workload) {
        $searchParam += (@{Workload = $Workload })
        SayInfo "Workload: $($Workload -join ',')"
    }
    try {
        $events = @(Get-MS365Messages @searchParam -ErrorAction STOP)
        SayInfo "Total Incidents Retrieved: $($events.Count)"
    }
    catch {
        SayError "Failed to get data. $($_.Exception.Message)"
        $errorFlag = $true
        return $null
    }

    #EndRegion

    #Region Create Report
    ## Get the CSS style
    $css_string = Get-Content (($moduleInfo.ModuleBase.ToString()) + '\source\private\style.css') -Raw

    #Region Consolidate Email
    if ($Consolidate) {
        if ($events.Count -gt 0) {
            $mailSubject = "[$($organizationName)] Microsoft 365 Service Health Report"
            $event_id_file = "$outputDir\consolidated_report.html"
            $event_id_json_file = "$outputDir\consolidated_report.json"
            $htmlBody = [System.Collections.ArrayList]@()
            $null = $htmlBody.Add("<html><head><title>$($mailSubject)</title>")
            $null = $htmlBody.Add('<style type="text/css">')
            $null = $htmlBody.Add($css_string)
            $null = $htmlBody.Add("</style>")
            $null = $htmlBody.Add("</head><body>")
            $null = $htmlBody.Add("<hr>")
            $null = $htmlBody.Add('<table id="section"><tr><th><a name="summary">Summary</a></th></tr></table>')
            $null = $htmlBody.Add("<hr>")
            $null = $htmlBody.Add('<table id="data">')
            $null = $htmlBody.Add("<tr><th>Workload</th><th>Event ID</th><th>Classification</th><th>Status</th><th>Title</th></tr>")
            foreach ($event in ($events | Sort-Object Classification -Descending)) {
                $ticket_status = ($event.Status.substring(0, 1).toupper() + $event.Status.substring(1) -creplace '[^\p{Ll}\s]', ' $&').Trim();
                $null = $htmlBody.Add("<tr><td>$($event.Service)</td>
                <td>" + '<a href="#' + $($event.ID) + '">' + "$($event.ID)</a></td>
                <td>$($event.Classification.substring(0, 1).toupper() + $event.Classification.substring(1))</td>
                <td>$($ticket_status)</td>
                <td>$($event.Title)</td></tr>")
            }
            $null = $htmlBody.Add('</table>')

            foreach ($event in $events | Sort-Object Classification -Descending) {
                $ticket_status = ($event.Status.substring(0, 1).toupper() + $event.Status.substring(1) -creplace '[^\p{Ll}\s]', ' $&').Trim();
                $null = $htmlBody.Add("<hr>")
                $null = $htmlBody.Add('<table id="section"><tr><th><a name="' + $event.ID + '">' + $event.ID + '</a> | ' + $event.Service + ' | ' + $event.Title + '</th></tr></table>')
                $null = $htmlBody.Add("<hr>")
                $null = $htmlBody.Add('<table id="data">')
                $null = $htmlBody.Add('<tr><th>Status</th><td><b>' + $ticket_status + '</b></td></tr>')
                $null = $htmlBody.Add('<tr><th>Organization</th><td>' + $organizationName + '</td></tr>')
                $null = $htmlBody.Add('<tr><th>Classification</th><td>' + $($event.Classification.substring(0, 1).toupper() + $event.Classification.substring(1)) + '</td></tr>')
                $null = $htmlBody.Add('<tr><th>User Impact</th><td>' + $event.ImpactDescription + '</td></tr>')
                $null = $htmlBody.Add('<tr><th>Last Updated</th><td>' + "{0:yyyy-MM-dd H:mm}" -f [datetime]$event.lastModifiedDateTime + '</td></tr>')
                $null = $htmlBody.Add('<tr><th>Start Time</th><td>' + "{0:yyyy-MM-dd H:mm}" -f [datetime]$event.startDateTime + '</td></tr>')
                $null = $htmlBody.Add('<tr><th>End Time</th><td>' + $(
                        if ($event.endDateTime) {
                            "{0:yyyy-MM-dd H:mm}" -f [datetime]$event.endDateTime
                        }
                        else {
                            $null
                        }
                    ) + '</td></tr>')

                $latestMessage = ($event.posts[-1].description.content) -replace "`n", "<br />"

                $null = $htmlBody.Add('<tr><th>Latest Message</th><td>' + $latestMessage + '</td></tr>')
                $null = $htmlBody.Add('</table>')
                $null = $htmlBody.Add('<div style="font-family: Tahoma;font-size: 10px"><a href = "#summary">(back to summary)</a></div>')
            }

            $null = $htmlBody.Add('<p><font size="2" face="Segoe UI Light"><br />')
            $null = $htmlBody.Add('<br />')
            $null = $htmlBody.Add('<a href="' + $moduleInfo.ProjectURI.ToString() + '" target="_blank">' + $moduleInfo.Name.ToString() + ' v' + $moduleInfo.Version.ToString() + ' </a><br></p>')
            $null = $htmlBody.Add('</body>')
            $null = $htmlBody.Add('</html>')
            $htmlBody = $htmlBody -join "`n" #convert to multiline string

            $htmlBody = ReplaceSmartCharacter $htmlBody

            if ($WriteReportToDisk -eq $true) {
                $htmlBody | Out-File $event_id_file -Force
            }

            if ($SendEmail -eq $true) {
                # Recipients
                $toAddressJSON = @()
                $To | ForEach-Object {
                    $toAddressJSON += @{EmailAddress = @{Address = $_ } }
                }

                try {
                    #message
                    $mailBody = @{
                        message = @{
                            subject                = $mailSubject
                            body                   = @{
                                contentType = "HTML"
                                content     = $htmlBody
                            }
                            toRecipients           = @(
                                $ToAddressJSON
                            )
                            internetMessageHeaders = @(
                                @{
                                    name  = "X-Mailer"
                                    value = "MS365HealthReport (junecastillote)"
                                }
                            )
                        }
                    }

                    ## Add CC recipients if specified
                    if ($Cc) {
                        $ccAddressJSON = @()
                        $Cc | ForEach-Object {
                            $ccAddressJSON += @{EmailAddress = @{Address = $_ } }
                        }
                        $mailBody.Message += @{ccRecipients = $ccAddressJSON }
                    }

                    ## Add BCC recipients if specified
                    if ($Bcc) {
                        $BccAddressJSON = @()
                        $Bcc | ForEach-Object {
                            $BccAddressJSON += @{EmailAddress = @{Address = $_ } }
                        }
                        $mailBody.Message += @{BccRecipients = $BccAddressJSON }
                    }

                    $mailBody = $($mailBody | ConvertTo-Json -Depth 4)

                    if ($WriteRawJSONToDisk) {
                        $mailBody | Out-File $event_id_json_file -Force
                    }

                    ## Send email
                    # $ServicePoint = [System.Net.ServicePointManager]::FindServicePoint('https://graph.microsoft.com')
                    SayInfo "Sending Consolidated Alert for $($events.id -join ';')"
                    $null = Invoke-RestMethod -Method Post -Uri "https://graph.microsoft.com/beta/users/$($From)/sendmail" -Body $mailBody -Headers $GraphAPIHeader -ContentType 'application/json'
                    # $null = $ServicePoint.CloseConnectionGroup('')
                }
                catch {
                    SayInfo "Failed to send Alert for $($events.id -join ';') | $($_.Exception.Message)"
                    $errorFlag = $true
                    return $null
                }
            }
        }
    }
    #EndRegion Consolidate Email
    #Region NoConsolidate Email
    else {
        if ($events.Count -gt 0) {
            foreach ($event in ($events | Sort-Object Classification -Descending) ) {
                $ticket_status = ($event.Status.substring(0, 1).toupper() + $event.Status.substring(1) -creplace '[^\p{Ll}\s]', ' $&').Trim();
                $mailSubject = "[$($organizationName)] MS365 Service Health Report | $($event.id) | $($event.Service)"
                $event_id_file = "$outputDir\$($event.ID).html"
                $event_id_json_file = "$outputDir\$($event.ID).json"
                $htmlBody = [System.Collections.ArrayList]@()
                $null = $htmlBody.Add("<html><head><title>$($mailSubject)</title>")
                $null = $htmlBody.Add('<style type="text/css">')
                $null = $htmlBody.Add($css_string)
                $null = $htmlBody.Add("</style>")
                $null = $htmlBody.Add("</head><body>")
                $null = $htmlBody.Add("<hr>")
                $null = $htmlBody.Add('<table id="section"><tr><th>' + $event.ID + ' | ' + $event.Service + ' | ' + $event.Title + '</th></tr></table>')
                $null = $htmlBody.Add("<hr>")
                $null = $htmlBody.Add('<table id="data">')
                $null = $htmlBody.Add('<tr><th>Status</th><td><b>' + $ticket_status + '</b></td></tr>')
                $null = $htmlBody.Add('<tr><th>Organization</th><td>' + $organizationName + '</td></tr>')
                $null = $htmlBody.Add('<tr><th>Classification</th><td>' + $($event.Classification.substring(0, 1).toupper() + $event.Classification.substring(1)) + '</td></tr>')
                $null = $htmlBody.Add('<tr><th>User Impact</th><td>' + $event.ImpactDescription + '</td></tr>')
                $null = $htmlBody.Add('<tr><th>Last Updated</th><td>' + [datetime]$event.lastModifiedDateTime + '</td></tr>')
                $null = $htmlBody.Add('<tr><th>Start Time</th><td>' + [datetime]$event.startDateTime + '</td></tr>')
                $null = $htmlBody.Add('<tr><th>End Time</th><td>' + $(
                        if ($event.endDateTime) {
                            [datetime]$event.endDateTime
                        }
                        else {
                            $null
                        }
                    ) + '</td></tr>')

                $latestMessage = ($event.posts[-1].description.content) -replace "`n", "<br />"

                $null = $htmlBody.Add('<tr><th>Latest Message</th><td>' + $latestMessage + '</td></tr>')
                $null = $htmlBody.Add('</table>')

                $null = $htmlBody.Add('<p><font size="2" face="Segoe UI Light"><br />')
                $null = $htmlBody.Add('<br />')
                $null = $htmlBody.Add('<a href="' + $moduleInfo.ProjectURI.ToString() + '" target="_blank">' + $moduleInfo.Name.ToString() + ' v' + $moduleInfo.Version.ToString() + ' </a><br></p>')
                $null = $htmlBody.Add('</body>')
                $null = $htmlBody.Add('</html>')
                $htmlBody = $htmlBody -join "`n" #convert to multiline string

                $htmlBody = ReplaceSmartCharacter $htmlBody
                # $htmlBody | Out-File -FilePath $env:temp

                if ($WriteReportToDisk -eq $true) {
                    $htmlBody | Out-File $event_id_file -Force
                }

                if ($SendEmail -eq $true) {

                    # Recipients
                    $toAddressJSON = @()
                    $To | ForEach-Object {
                        $toAddressJSON += @{EmailAddress = @{Address = $_ } }
                    }

                    try {
                        #message
                        $mailBody = @{
                            message = @{
                                subject                = $mailSubject
                                body                   = @{
                                    contentType = "HTML"
                                    content     = $htmlBody
                                }
                                toRecipients           = @(
                                    $ToAddressJSON
                                )
                                internetMessageHeaders = @(
                                    @{
                                        name  = "X-Mailer"
                                        value = "MS365HealthReport (junecastillote)"
                                    }
                                )
                            }
                        }

                        ## Add CC recipients if specified
                        if ($Cc) {
                            $ccAddressJSON = @()
                            $Cc | ForEach-Object {
                                $ccAddressJSON += @{EmailAddress = @{Address = $_ } }
                            }
                            $mailBody.Message += @{ccRecipients = $ccAddressJSON }
                        }

                        ## Add BCC recipients if specified
                        if ($Bcc) {
                            $BccAddressJSON = @()
                            $Bcc | ForEach-Object {
                                $BccAddressJSON += @{EmailAddress = @{Address = $_ } }
                            }
                            $mailBody.Message += @{BccRecipients = $BccAddressJSON }
                        }

                        $mailBody = $($mailBody | ConvertTo-Json -Depth 4)

                        if ($WriteRawJSONToDisk) {
                            $mailBody | Out-File $event_id_json_file -Force
                        }

                        ## Send email
                        # $ServicePoint = [System.Net.ServicePointManager]::FindServicePoint('https://graph.microsoft.com')
                        SayInfo "Sending Alert for $($event.id)"
                        $null = Invoke-RestMethod -Method Post -Uri "https://graph.microsoft.com/v1.0/users/$($From)/sendmail" -Body $mailBody -Headers $GraphAPIHeader -ContentType 'application/json'
                        # $null = $ServicePoint.CloseConnectionGroup('')

                    }
                    catch {
                        SayInfo "Failed to send Alert for $($event.id) | $($_.Exception.Message)"
                        $errorFlag = $true
                        return $null
                    }
                }
            }
        }
    }
    #EndRegion NoConsolidate Email

    if ($events.Count -gt 0) {
        if ($SendTeams -eq $true -and $TeamsWebHookURL.Count -gt 0) {
            #Region Consolidate (Teams)
            # if ($Consolidate -eq $true) {
            $teamsAdaptiveCard = New-ConsolidatedCard -InputObject $events
            SayInfo "Posting Alert to Teams Channel for $($events.id -join ';')"
            foreach ($url in $TeamsWebHookURL) {
                $Params = @{
                    "URI"         = $url
                    "Method"      = 'POST'
                    "Body"        = $teamsAdaptiveCard | ConvertTo-Json -Depth 50
                    # "Body"        = $teamsAdaptiveCard
                    "ContentType" = 'application/json'
                }
                $result = Invoke-RestMethod @Params

                if ($result -eq 1) {
                    # SayInfo "OK. Posted to $url."
                }
                else {
                    SayError "Failed to post to channel. $result."
                    $errorFlag = $true
                }
            }
            # }
            #EndRegion Consolidate (Teams)
        }

        #Region Don't Consolidate (Teams)
        if ($SendTeams -eq $true -and $TeamsWebHookURL.Count -gt 0 -and $Consolidate -eq $false) {

            ## TODO: Add per event Teams notification.
        }
        #EndRegion Don't Consolidate (Teams)
    }

    #EndRegion Create Report

    if ($errorFlag) {
        SayInfo "Setting last run time (NotOK) in $($runHistoryFile) to $("{0:yyyy-MM-dd HH:mm}" -f $now)"
        "$("{0:yyyy-MM-dd H:mm}" -f $now),NotOK" | Add-Content -Path $RunHistoryFile -Force -Confirm:$false
    }
    elseif ($events.Count -lt 1) {
        SayInfo "Setting last run time (NoResult) in $($runHistoryFile) to $("{0:yyyy-MM-dd HH:mm}" -f $now)"
        "$("{0:yyyy-MM-dd H:mm}" -f $now),NoResult" | Add-Content -Path $RunHistoryFile -Force -Confirm:$false
    }
    else {
        SayInfo "Setting last run time (OK) in $($runHistoryFile) to $("{0:yyyy-MM-dd HH:mm}" -f $now)"
        "$("{0:yyyy-MM-dd H:mm}" -f $now),OK" | Add-Content -Path $RunHistoryFile -Force -Confirm:$false
    }
}