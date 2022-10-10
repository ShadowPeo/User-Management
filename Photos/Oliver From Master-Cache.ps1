$folderPhotos = "<<PATH TO CENTRAL STORE>>>"
$folderCache = "./Cache"

#Get-ChildItem -Path $folderCache -Include *.* -File -Recurse | foreach { $_.Delete()}

#Copy-Item -Path "$folderPhotos/*" -Destination $folderCache -PassThru | Out-Null

$photosToProcess = Get-ChildItem -Path $folderCache

foreach ($photo in $photosToProcess)
{
    if (($photo.BaseName.ToString()) -match '[01s]{1}[0-9t]{1}[0-9]{5,6}')
    {
        
        Rename-Item -Path $photo.FullName -NewName "$($matches[0])$($photo.Extension)"
    }
}

