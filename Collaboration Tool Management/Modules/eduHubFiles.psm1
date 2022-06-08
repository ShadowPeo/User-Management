function Join-eduHubDelta
{
    Param 
    (
        [Parameter(Mandatory=$true)][string]$file1, 
        [Parameter(Mandatory=$true)][string]$file2, 
        [Parameter(Mandatory=$true)][string]$outputPath,
        [Parameter(Mandatory=$true)][string]$matchAttribute,
        [switch]$force
    )
    
    if(!(test-path $file1))
    {
            LogWrite ("$file1 does not exist, breaking")
            break
    }

    if(!(test-path $file2))
    {
            LogWrite ("$file2 does not exist, returning File 1")
            return $file1
    }


    ## Merte the two files if the second file is newer than the second
    if (((Get-Item $file1).LastWriteTime -lt (Get-Item $file2).LastWriteTime) -or $force -eq $true)
    {
        LogWrite "Delta File is newer than Base File, Merging" -noOutput:$true
        
        ### Set the output file location
        
        ### Test to ensure output path is vaild, if not create it
        if(!(test-path $outputPath))
        {
              New-Item -ItemType Directory -Force -Path $outputPath | Out-Null
        }
        
        $outputFile = "$outputPath\$((Get-Item $file1).Name)"

        $file1Import = Import-Csv -Path $file1
        $file2Import = Import-Csv -Path $file2

        

        foreach ($record in $file2Import)
        {
            if ($file1Import.$matchAttribute -contains $record.$matchAttribute)
            {
                LogWrite "Record ($($record.$matchAttribute)) Matches Existing Record, Merging" -noOutput:$true
                
                foreach ($row in $file1Import)
                {
                    if ($row.$matchAttribute -eq $record.$matchAttribute)
                    {
                        $row = $record
                    }
                }
            }
            else
            {
                LogWrite "New Record Found, Inserting" -noOutput:$true
                
                return $file1
                $file1Import += $record
            }
        }

        $file1Import | Export-CSV $outputFile -Encoding ASCII  -NoTypeInformation

        return $outputFile
    }
    else
    { 
        LogWrite "Newer file not detected, skipping" -noOutput:$true
        return $file1
    }
}