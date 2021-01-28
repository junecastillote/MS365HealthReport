Function New-MS365IncidentReport {
    [cmdletbinding(DefaultParameterSetName = 'bySecret')]
    param (
        [parameter()]
        [string]
        $OrganizationName,

        [parameter(Mandatory, ParameterSetName = 'byCertificate')]
        [parameter(Mandatory, ParameterSetName = 'byThumbprint')]
        [parameter(Mandatory, ParameterSetName = 'bySecret')]
        [guid]
        $ClientID,

        [parameter(Mandatory, ParameterSetName = 'bySecret')]
        [string]
        $ClientSecret,

        [parameter(Mandatory, ParameterSetName = 'byCertificate')]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]
        $ClientCertificate,

        [parameter(Mandatory, ParameterSetName = 'byThumbprint')]
        [string]
        $ClientCertificateThumbprint,

        [parameter(Mandatory, ParameterSetName = 'byCertificate')]
        [parameter(Mandatory, ParameterSetName = 'byThumbprint')]
        [parameter(Mandatory, ParameterSetName = 'bySecret')]
        [string]
        $TenantID,

        [Parameter()]
        [datetime]
        $LastUpdatedTime,

        [Parameter()]
        [string[]]
        $Workload,

        [Parameter()]
        [switch]
        $SendEmail,

        [Parameter()]
        [mailaddress]
        $From,

        [Parameter()]
        [mailaddress[]]
        $To,

        [Parameter()]
        [mailaddress[]]
        $CC,

        [Parameter()]
        [mailaddress[]]
        $Bcc
    )

    $ModuleInfo = Get-Module MS365HealthReport
    if (!$OrganizationName) { $OrganizationName = $TenantID }

    Write-Verbose $($pscmdlet.ParameterSetName)

    if ($pscmdlet.ParameterSetName -eq 'bySecret') {
        $SecureClientSecret = New-Object System.Security.SecureString
        $ClientSecret.toCharArray() | ForEach-Object { $SecureClientSecret.AppendChar($_) }
        $ServiceCommsApiOAuth2 = Get-MsalToken -ClientId $ClientID -ClientSecret $SecureClientSecret -TenantId $tenantID -Scopes 'https://manage.office.com/.default' -ErrorAction Stop
    }
    elseif ($pscmdlet.ParameterSetName -eq 'byCertificate') {
        $ServiceCommsApiOAuth2 = Get-MsalToken -ClientId $ClientID -ClientCertificate $ClientCertificate -TenantId $tenantID -Scopes 'https://manage.office.com/.default' -ErrorAction Stop
    }
    elseif ($pscmdlet.ParameterSetName -eq 'byThumbprint') {
        $ServiceCommsApiOAuth2 = Get-MsalToken -ClientId $ClientID -ClientCertificate (Get-Item Cert:\CurrentUser\My\$($ClientCertificateThumbprint)) -TenantId $tenantID -Scopes 'https://manage.office.com/.default' -ErrorAction Stop
    }

    $searchParam = @{
        Token       = ($ServiceCommsApiOAuth2.AccessToken);
        MessageType = 'Incident';
    }

    if ($LastUpdatedTime) {
        $searchParam += (@{LastUpdatedTime = $LastUpdatedTime })
    }

    if ($Workload) {
        $searchParam += (@{Workload = $Workload })
    }

    $events = @(Get-MS365Messages @searchParam)
    Write-Verbose "Events=$($events.Count)"
    #$events

    $css_string = Get-Content (($ModuleInfo.ModuleBase.ToString()) + '\source\public\style.css') -Raw
    $outputDir = "$($env:TMP)\$($ModuleInfo.Name)\$($TenantID)"

    if (!(Test-Path -Path $outputDir)) {
        $null = New-Item -ItemType Directory -Path $outputDir -Force
    }
    else {
        Remove-Item -Path $outputDir\* -Recurse -Force -Confirm:$false
    }

    if ($events.Count -gt 0) {
        foreach ($event in $events) {
            $event_id_file = "$outputDir\$($event.ID).html"
            $htmlBody = [System.Collections.ArrayList]@()
            $mailSubject = '[' + $event.Status + '] ' + $event.ID + ' | ' + $event.WorkloadDisplayName + ' | ' + $event.Title
            $null = $htmlBody.Add("<html><head><title>$($mailSubject)</title>")
            $null = $htmlBody.Add('<style type="text/css">')
            $null = $htmlBody.Add($css_string)
            $null = $htmlBody.Add("</style>")
            $null = $htmlBody.Add("</head><body>")
            $null = $htmlBody.Add("<hr>")
            $null = $htmlBody.Add('<table id="section"><tr><th>' + $event.ID + ' | ' + $event.WorkloadDisplayName + ' | ' + $event.Title + '</th></tr></table>')
            $null = $htmlBody.Add("<hr>")
            $null = $htmlBody.Add('<table id="data">')
            $null = $htmlBody.Add('<tr><th>Status</th><td><b>' + $event.Status + '</b></td></tr>')
            $null = $htmlBody.Add('<tr><th>Organization</th><td>' + $organizationName + '</td></tr>')
            $null = $htmlBody.Add('<tr><th>Classification</th><td>' + $event.Classification + '</td></tr>')
            $null = $htmlBody.Add('<tr><th>User Impact</th><td>' + $event.ImpactDescription + '</td></tr>')

            $null = $htmlBody.Add('<tr><th>Last Updated</th><td>' + [datetime]$event.LastUpdatedTime + '</td></tr>')

            $null = $htmlBody.Add('<tr><th>Start Time</th><td>' + [datetime]$event.StartTime + '</td></tr>')
            if ($event.EndTime) {
                $null = $htmlBody.Add('<tr><th>End Time</th><td>' + [datetime]$event.EndTime + '</td></tr>')
            }
            else {
                $null = $htmlBody.Add('<tr><th>End Time</th><td></td></tr>')
            }

            $null = $htmlBody.Add('<tr><th>Latest Message</th><td>' + ($event.Messages[-1].MessageText).Replace("`n", "<br />").Replace('�','') + '</td></tr>')
            $null = $htmlBody.Add('</table>')

            $null = $htmlBody.Add('<p><table id="section">')
            $null = $htmlBody.Add('<p><font size="2" face="Segoe UI Light"><br />')
            $null = $htmlBody.Add('<br />')
            $null = $htmlBody.Add('<a href="' + $ModuleInfo.ProjectURI.ToString() + '" target="_blank">' + $ModuleInfo.Name.ToString() + ' v' + $ModuleInfo.Version.ToString() + ' </a><br>')
            $null = $htmlBody.Add('</body>')
            $null = $htmlBody.Add('</html>')
            $htmlBody = $htmlBody -join "`n" #convert to multiline string
            $htmlBody | Out-File $event_id_file -Force

            if ($SendEmail -eq $true) {
                if (!$From) { Write-Warning "You ask me to send an email report but you forgot to add the -From address."; return $null }
                if (!$To) { Write-Warning "You ask me to send an email report but you forgot to add the -To address(es)."; return $null }

                #recipients
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

                    $ServicePoint = [System.Net.ServicePointManager]::FindServicePoint('https://graph.microsoft.com')

                    # Get GraphAPI Token
                    if ($pscmdlet.ParameterSetName -eq 'bySecret') {
                        $SecureClientSecret = New-Object System.Security.SecureString
                        $ClientSecret.toCharArray() | ForEach-Object { $SecureClientSecret.AppendChar($_) }
                        $GraphApiOAuth2 = Get-MsalToken -ClientId $ClientID -ClientSecret $SecureClientSecret -TenantId $tenantID -Scopes @('https://graph.microsoft.com/Mail.Send') -ErrorAction Stop
                    }
                    elseif ($pscmdlet.ParameterSetName -eq 'byCertificate') {
                        $GraphApiOAuth2 = Get-MsalToken -ClientId $ClientID -ClientCertificate $ClientCertificate -TenantId $tenantID -Scopes @('https://graph.microsoft.com/Mail.Send') -ErrorAction Stop
                    }
                    elseif ($pscmdlet.ParameterSetName -eq 'byThumbprint') {
                        $GraphApiOAuth2 = Get-MsalToken -ClientId $ClientID -ClientCertificate (Get-Item Cert:\CurrentUser\My\$($ClientCertificateThumbprint)) -TenantId $tenantID -Scopes @('https://graph.microsoft.com/Mail.Send') -ErrorAction Stop
                    }

                    $header = @{'Authorization' = "Bearer $($GraphApiOAuth2.AccessToken)" }

                    $null = Invoke-RestMethod -Method Post -Uri "https://graph.microsoft.com/v1.0/users/$($From)/sendmail" -Body ($mailBody | ConvertTo-Json -Depth 4) -Headers $header -ContentType application/json

                    $null = $ServicePoint.CloseConnectionGroup('')
                }
                catch {
                    Write-Error "Failed to send Alert for $($event.id). $($_.Exception.Message)"
                }
            }
        }
    }
}