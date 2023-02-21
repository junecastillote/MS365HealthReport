Function New-ConsolidatedCard {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        $InputObject
    )

    $moduleInfo = Get-Module $($MyInvocation.MyCommand.ModuleName)

    Function New-FactItem {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            $InputObject
        )

        $factHeader = [pscustomobject][ordered]@{
            type  = "Container"
            style = "emphasis"
            bleed = $true
            items = @(
                $([pscustomobject][ordered]@{
                        type      = 'TextBlock'
                        wrap      = $true
                        separator = $true
                        weight    = 'Bolder'
                        text      = "$($InputObject.id) | $($InputObject.Service) | $($InputObject.Title)"
                    } )
            )
        }

        $factSet = [pscustomobject][ordered]@{
            type      = 'FactSet'
            separator = $true
            facts     = @(
                $([pscustomobject][ordered]@{Title = 'Impact'; Value = $($InputObject.impactDescription) } ),
                $([pscustomobject][ordered]@{Title = 'Type'; Value = ($InputObject.Classification.substring(0, 1).toupper() + $InputObject.Classification.substring(1)) } ),
                $([pscustomobject][ordered]@{Title = 'Status'; Value = ($InputObject.Status.substring(0, 1).toupper() + $InputObject.Status.substring(1) -creplace '[^\p{Ll}\s]', ' $&').Trim(); } ),
                $([pscustomobject][ordered]@{Title = 'Update'; Value = ("{0:MMMM dd, yyyy hh:mm tt}" -f [datetime]$InputObject.lastModifiedDateTime) }),
                $([pscustomobject][ordered]@{Title = 'Start'; Value = ("{0:MMMM dd, yyyy hh:mm tt}" -f [datetime]$InputObject.startDateTime) }),
                $([pscustomobject][ordered]@{Title = 'End'; Value = $(if ($InputObject.endDateTime) { ("{0:MMMM dd, yyyy hh:mm tt}" -f [datetime]$InputObject.startDateTime) }) })
            )
        }
        return @($factHeader, $factSet)
    }

    $teamsAdaptiveCard = (Get-Content (($moduleInfo.ModuleBase.ToString()) + '\source\private\TeamsConsolidated.json') -Raw | ConvertFrom-Json)
    foreach ($item in $InputObject) {
        $teamsAdaptiveCard.attachments[0].content.body += (New-FactItem -InputObject $item)
    }
    return $teamsAdaptiveCard
}