[CmdletBinding()]
param (
    [string]$ModulePath
)
$moduleManifest = Get-ChildItem -Path $PSScriptRoot -Filter *.psd1
$Moduleinfo = Test-ModuleManifest -Path ($moduleManifest.FullName)

Remove-Module ($Moduleinfo.Name) -ErrorAction SilentlyContinue

if (!$ModulePath) {
    # Get all PSModulePath
    $paths = ($env:PSModulePath -split ";")

    do {
        Clear-Host
        # Display selection menu
        Say "====== Module Install Location ======"
        Say ""
        $i = 1
        $paths | ForEach-Object {
            Say "$($i): $_"
            $i = $i + 1
        }
        Say "Q: QUIT"
        Say ""
        # AS for input
        $userInput = Read-Host "Select the installation path"
    }
    until ($userInput -eq 'Q' -or ($userInput -lt ($paths.count + 1) -and $userInput -gt 0))

    if ($userInput -eq 'Q') {
        Say ""
        Say "QUIT"
        Say ""
        return $null
    }
    $ModulePath = $paths[($userInput - 1)]
}
$ModulePath = $ModulePath + "\$($Moduleinfo.Name.ToString())\$($Moduleinfo.Version.ToString())"

if (!(Test-Path $ModulePath)) {
    New-Item -Path $ModulePath -ItemType Directory | Out-Null
}

try {
    Copy-Item -Path $PSScriptRoot\* -Include *.psd1, *.psm1 -Destination $ModulePath -Force -Confirm:$false -ErrorAction Stop
    Copy-Item -Path $PSScriptRoot\source -recurse -Destination $ModulePath -Force -Confirm:$false -ErrorAction Stop
    Say ""
    Say "Success. Installed to $ModulePath"
    Say ""
    #Import-Module ExCmdReport
    Get-ChildItem -Recurse $ModulePath | Unblock-File -Confirm:$false
}
catch {
    Say ""
    Say "Failed"
    Say $_.Exception.Message
    Say ""
}

