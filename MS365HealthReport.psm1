$Path = [System.IO.Path]::Combine($PSScriptRoot, 'source')
Get-Childitem $Path -Filter *.ps1 -Recurse | Foreach-Object {
    . $_.Fullname
}