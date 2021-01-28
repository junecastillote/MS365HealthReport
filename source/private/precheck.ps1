# Confirm the host os is Windows
if ($psversiontable.os -notlike "*windows*") {
    Write-Error "This module is compatible only with Windows."
    break
}