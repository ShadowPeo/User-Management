

function Set-AttributeFiles ()
{
    $customAttributeNames = @()
    foreach ($attributeFile in $attributeFiles)
    {
        $customAuthChangeAttribs = @()
        $AttributeName = $attributeFile.Name.Substring($attributeFilePrefix.Length+1,(($attributeFile.Name.Length-1) - ($attributeFilePrefix.Length + $attributeFile.Extension.Length)))
        $global:customAttributeNames += $AttributeName
        $customSettings[$AttributeName] = Import-CSV $attributeFile.FullName
        foreach ($validAttribute in $customSettings[$AttributeName])
        {
            if ($customAuthChangeAttribs -notcontains $validAttribute.licencingValue)
            {
                $customAuthChangeAttribs += $validAttribute.licencingValue
            }
        }
        $customAuthChanges[$AttributeName] =  $customAuthChangeAttribs
    }
}

function Set-Licencing ()
{
    foreach ($customAttribute in $customAttributeNames) 
    {
        $customAuthChangeAttribs = $customSettings[$customAttribute].licencingValue
            
        foreach ($accessControl in $customSettings[$customAttribute]) 
        {
                
            $attributeRow = 0
            if ($accessControl.attributeValue -eq "FALSE")
            {
                $accessControl.attributeValue = $false
            }
            elseif ($accessControl.attributeValue -eq "TRUE")
            {
                $accessControl.attributeValue = $true
            }
                
            if ($accessControl.attributeName -eq "PermChange")
            {
                continue
            }

            if ($accessControl.attributeName -ne "Default")
            {
            
                #LogWrite "Checking $(($AD_User.$($accessControl.attributeName))) -eq $($accessControl.attributeValue))"
                if (($AD_User.$($accessControl.attributeName) -eq $accessControl.attributeValue) -and ($AD_User.$customAttribute -ne $accessControl.licencingValue))
                {
                    if (($customAuthChangeAttribs -contains $AD_User.$customAttribute) -or (($AD_User.$customAttribute -eq "") -or ($AD_User.$customAttribute -eq $null)))
                    {
                        updateADValue $customAttribute $customAttribute $accessControl.licencingValue
                    }
                    else
                    {
                        LogWrite "Cannot update $userCommonName's $customAttribute to $($accessControl.licencingValue) as it has been manually overridden" -foregroundColour:"Magenta"
                    }

                    break
                }
                elseif (($AD_User.$($accessControl.attributeName) -eq $accessControl.attributeValue) -and ($AD_User.$customAttribute -eq $accessControl.licencingValue))
                {
                    break
                }
                    
                $attributeRow++

            }
            elseif (($accessControl.attributeName -eq "Default") -and ($AD_User.$customAttribute -ne $accessControl.licencingValue))
            {
                if (($customAuthChangeAttribs -contains $AD_User.$customAttribute) -or ($AD_User.$customAttribute -eq $null) -or ($AD_User.$customAttribute -eq ""))
                {
                    updateADValue $customAttribute $customAttribute $accessControl.licencingValue
                }
                else
                {
                    LogWrite "Cannot update $userCommonName's $customAttribute to $($accessControl.licencingValue) as it has been manually overridden" -foregroundColour:"Magenta"

                }
                break
            }
        }
    }
}