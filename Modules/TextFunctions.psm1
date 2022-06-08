
function correctToTitle ($textToCorrect)
{
    $TextInfo = (Get-Culture).TextInfo
    return $TextInfo.ToTitleCase($textToCorrect.ToLower())
}