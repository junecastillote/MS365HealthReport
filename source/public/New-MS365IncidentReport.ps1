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
        $Consolidate = $true
    )
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $ModuleInfo = Get-Module $($MyInvocation.MyCommand.ModuleName)
    $Now = Get-Date
    if (!$OrganizationName) { $OrganizationName = $TenantID }

    SayInfo "Authentication type: $($pscmdlet.ParameterSetName)"
    SayInfo "Client ID: $ClientID"
    SayInfo "Tenant ID: $TenantID"

    # Get Service Communications API Token
    if ($pscmdlet.ParameterSetName -eq 'Client Secret') {
        $SecureClientSecret = New-Object System.Security.SecureString
        $ClientSecret.toCharArray() | ForEach-Object { $SecureClientSecret.AppendChar($_) }
        $OAuth = Get-MsalToken -ClientId $ClientID -ClientSecret $SecureClientSecret -TenantId $tenantID -ErrorAction Stop
        Sayinfo $($ClientSecret -replace $($ClientSecret.Substring(0,$ClientSecret.Length -8)),$('X'*$($ClientSecret.Substring(0,$ClientSecret.Length -8)).Length))
    }
    elseif ($pscmdlet.ParameterSetName -eq 'Client Certificate') {
        $OAuth = Get-MsalToken -ClientId $ClientID -ClientCertificate $ClientCertificate -TenantId $tenantID -ErrorAction Stop
    }
    elseif ($pscmdlet.ParameterSetName -eq 'Certificate Thumbprint') {
        $OAuth = Get-MsalToken -ClientId $ClientID -ClientCertificate (Get-Item Cert:\CurrentUser\My\$($ClientCertificateThumbprint)) -TenantId $tenantID -ErrorAction Stop
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

    ## If -StartFromLastRun, this function will only get the incidents whose LastUpdatedTime is after the timestamp in "HKCU:\Software\MS365HealthReport\$TenantID"
    if ($StartFromLastRun) {
        SayInfo "Getting last run time from the registry."
        [datetime]$LastUpdatedTime = Get-MS365HealthReportLastRunTime -TenantID $TenantID
    }

    ## If -LastUpdatedTime, this function will only get the incidents whose LastUpdatedTime is after the $LastUpdatedTime datetime value.
    if ($LastUpdatedTime) {
        $searchParam += (@{LastUpdatedTime = $LastUpdatedTime })
        SayInfo "Getting incident reports from: $LastUpdatedTime"
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
        return $null
    }

    #EndRegion

    #Region Prepare Output Directory
    if ($WriteReportToDisk -eq $true) {
        $outputDir = "$($env:USERPROFILE)\$($ModuleInfo.Name)\$($TenantID)"
        if (!(Test-Path -Path $outputDir)) {
            $null = New-Item -ItemType Directory -Path $outputDir -Force
        }
        else {
            Remove-Item -Path $outputDir\* -Recurse -Force -Confirm:$false
        }
        SayInfo "Output Directory: $outputDir"
    }

    #EndRegion

    #Region Create Report

    ## Get the CSS style
    $css_string = Get-Content (($ModuleInfo.ModuleBase.ToString()) + '\source\public\style.css') -Raw

    #Region Consolidate
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
                $null = $htmlBody.Add("<tr><td>$($event.Service)</td>
                <td>" + '<a href="#' + $($event.ID) + '">' + "$($event.ID)</a></td>
                <td>$($event.Classification)</td>
                <td>$($event.Status)</td>
                <td>$($event.Title)</td></tr>")
            }
            $null = $htmlBody.Add('</table>')

            foreach ($event in $events | Sort-Object Classification -Descending) {
                $null = $htmlBody.Add("<hr>")
                $null = $htmlBody.Add('<table id="section"><tr><th><a name="' + $event.ID + '">' + $event.ID + '</a> | ' + $event.Service + ' | ' + $event.Title + '</th></tr></table>')
                $null = $htmlBody.Add("<hr>")
                $null = $htmlBody.Add('<table id="data">')
                $null = $htmlBody.Add('<tr><th>Status</th><td><b>' + $event.Status + '</b></td></tr>')
                $null = $htmlBody.Add('<tr><th>Organization</th><td>' + $organizationName + '</td></tr>')
                $null = $htmlBody.Add('<tr><th>Classification</th><td>' + $event.Classification + '</td></tr>')
                $null = $htmlBody.Add('<tr><th>User Impact</th><td>' + $event.ImpactDescription + '</td></tr>')
                $null = $htmlBody.Add('<tr><th>Last Updated</th><td>' + "{0:yyyy-MM-dd H:mm}" -f [datetime]$event.lastModifiedDateTime + '</td></tr>')
                $null = $htmlBody.Add('<tr><th>Start Time</th><td>' + "{0:yyyy-MM-dd H:mm}" -f [datetime]$event.startDateTime + '</td></tr>')
                $null = $htmlBody.Add('<tr><th>End Time</th><td>' + $(
                        if ($event.endDateTime) {
                            "{0:yyyy-MM-dd H:mm}" -f [datetime]$event.endDateTime
                        }
                        else {
                            ""
                        }
                    ) + '</td></tr>')


                # https://4sysops.com/archives/dealing-with-smart-quotes-in-powershell/
                $latestMessage = ($event.posts[-1].description.content) -replace "`n", "<br />"
                $latestMessage = $latestMessage -replace '[\u2019\u2018]', "'"
                $latestMessage = $latestMessage -replace '[\u201C\u201D]', '"'

                $null = $htmlBody.Add('<tr><th>Latest Message</th><td>' + $latestMessage + '</td></tr>')
                $null = $htmlBody.Add('</table>')
                $null = $htmlBody.Add('<div style="font-family: Tahoma;font-size: 10px"><a href = "#summary">(back to summary)</a></div>')
            }

            $null = $htmlBody.Add('<p><font size="2" face="Segoe UI Light"><br />')
            $null = $htmlBody.Add('<br />')
            $null = $htmlBody.Add('<a href="' + $ModuleInfo.ProjectURI.ToString() + '" target="_blank">' + $ModuleInfo.Name.ToString() + ' v' + $ModuleInfo.Version.ToString() + ' </a><br></p>')
            $null = $htmlBody.Add('</body>')
            $null = $htmlBody.Add('</html>')
            $htmlBody = $htmlBody -join "`n" #convert to multiline string

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
                    $null = Invoke-RestMethod -Method Post -Uri "https://graph.microsoft.com/beta/users/$($From)/sendmail" -Body $mailBody -Headers $GraphAPIHeader -ContentType application/json
                    # $null = $ServicePoint.CloseConnectionGroup('')

                }
                catch {
                    SayInfo "Failed to send Alert for $($event.id) | $($_.Exception.Message)"
                    return $null
                }
            }
        }
    }
    #EndRegion Consolidate
    #Region NoConsolidate
    else {
        if ($events.Count -gt 0) {
            foreach ($event in ($events | Sort-Object Classification -Descending) ) {
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
                $null = $htmlBody.Add('<tr><th>Status</th><td><b>' + $event.Status + '</b></td></tr>')
                $null = $htmlBody.Add('<tr><th>Organization</th><td>' + $organizationName + '</td></tr>')
                $null = $htmlBody.Add('<tr><th>Classification</th><td>' + $event.Classification + '</td></tr>')
                $null = $htmlBody.Add('<tr><th>User Impact</th><td>' + $event.ImpactDescription + '</td></tr>')
                $null = $htmlBody.Add('<tr><th>Last Updated</th><td>' + [datetime]$event.lastModifiedDateTime + '</td></tr>')
                $null = $htmlBody.Add('<tr><th>Start Time</th><td>' + [datetime]$event.startDateTime + '</td></tr>')
                $null = $htmlBody.Add('<tr><th>End Time</th><td>' + $(
                        if ($event.endDateTime) {
                            [datetime]$event.endDateTime
                        }
                        else {
                            ""
                        }
                    ) + '</td></tr>')


                # https://4sysops.com/archives/dealing-with-smart-quotes-in-powershell/
                $latestMessage = ($event.posts[-1].description.content) -replace "`n", "<br />"
                $latestMessage = $latestMessage -replace '[\u2019\u2018]', "'"
                $latestMessage = $latestMessage -replace '[\u201C\u201D]', '"'

                $null = $htmlBody.Add('<tr><th>Latest Message</th><td>' + $latestMessage + '</td></tr>')
                $null = $htmlBody.Add('</table>')

                $null = $htmlBody.Add('<p><font size="2" face="Segoe UI Light"><br />')
                $null = $htmlBody.Add('<br />')
                $null = $htmlBody.Add('<a href="' + $ModuleInfo.ProjectURI.ToString() + '" target="_blank">' + $ModuleInfo.Name.ToString() + ' v' + $ModuleInfo.Version.ToString() + ' </a><br></p>')
                $null = $htmlBody.Add('</body>')
                $null = $htmlBody.Add('</html>')
                $htmlBody = $htmlBody -join "`n" #convert to multiline string

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
                        $null = Invoke-RestMethod -Method Post -Uri "https://graph.microsoft.com/v1.0/users/$($From)/sendmail" -Body $mailBody -Headers $GraphAPIHeader -ContentType application/json
                        # $null = $ServicePoint.CloseConnectionGroup('')

                    }
                    catch {
                        SayInfo "Failed to send Alert for $($event.id) | $($_.Exception.Message)"
                        return $null
                    }
                }
            }
        }
    }
    #EndRegion NoConsolidate

    #EndRegion Create Report

    # if ($StartFromLastRun) {
        SayInfo "Setting last run time in the registry to $Now"
        Set-MS365HealthReportLastRunTime -TenantID $tenantID -LastRunTime $Now
    # }
}