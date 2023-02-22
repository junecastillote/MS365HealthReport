Function ReplaceSmartCharacter {
    #https://4sysops.com/archives/dealing-with-smart-quotes-in-powershell/
    param(
        [parameter(Mandatory)]
        [string]$String
    )

    # Unicode Quote Characters
    $unicodePattern = @{
        '[\u2019\u2018]'                                                                                                                       = "'" # Single quote
        '[\u201C\u201D]'                                                                                                                       = '"' # Double quote
        '\u00A0|\u1680|\u180E|\u2000|\u2001|\u2002|\u2003|\u2004|\u2005|\u2006|\u2007|\u2008|\u2009|\u200A|\u200B|\u202F|\u205F|\u3000|\uFEFF' = " " # Space
        '\u0027'                                                                                                                               = "'" # Apostrophe
    }

    $unicodePattern.Keys | ForEach-Object {
        $stringToReplace = $_
        $String = $String -replace $stringToReplace, $unicodePattern[$stringToReplace]
    }

    return $String
}