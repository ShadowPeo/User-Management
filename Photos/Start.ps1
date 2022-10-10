$folderPhotos = "<<PATH TO PHOTO DOWNLOAD>>"

$photosToProcess = Get-ChildItem -Path $folderPhotos

foreach ($photo in $photosToProcess)
{
    $adUser = $null
    $adSearchUser = $null
    try 
    {
        $adSearchUser = $photo.BaseName.ToString() 
        $adUser =Get-ADUser -Properties * -Filter {(employeeID -eq $adSearchUser)}
        #Write-Host "$($adUser.employeeID) - $($adUser.samAccountName) - $($adUser.Surname), $($adUser.givenName)$($photo.Extension)"
        Rename-Item -Path $photo.FullName -NewName "$($adUser.employeeID) - $($adUser.samAccountName) - $($adUser.Surname), $($adUser.givenName)$($photo.Extension)"
    }
    catch {
        Write-Host "Error on $($photo.BaseName.ToString())"
    }
}

