function Get-edvalsClasses
{
    foreach ($subject in $TempTest)
    {
        if( $subject.'Course Subject'.Substring(0,1) -gt 1 -and ($subject.'Course Subject'.Substring(0,1) -ge 7 -and $subject.'Course Subject'.Substring(0,1) -le 9))
        {
            $subject.'CorrectedSubject' = "0$($subject.'Course Subject')"
        }
        else
        {
            $subject.'CorrectedSubject' = $subject.'Course Subject'
        }

        if ($subject.'CorrectedSubject'.Substring(0,2) -le 12)
        {
            $subject.'YearLevel' = $subject.'CorrectedSubject'.Substring(0,2)
        }

        $subject.'Course Name' = $subject.'Course Name' -replace '-and-',' & '
        $subject.'Course Name' = $subject.'Course Name' -replace '  ',' '
    
        if((($subject.'SIS ID').Substring(0,1) -match "^\d+$"))
        {
            if(($subject.'SIS ID').Substring(($subject.'SIS ID').IndexOf('.')+1,(($subject.'SIS ID').IndexOf('-') - ($subject.'SIS ID').IndexOf('.') - 1)) -match "^\d+$")
            {
                # Do something
                Write-Host "Digit"
                $tempCourseName = ("$($subject.YearLevel).")+(($subject.'SIS ID').Substring(($subject.'SIS ID').IndexOf('.')+1))
                $tempSubject = $subject.'Course Name'
                $subject.'Course Name' = $tempCourseName -replace "-"," $tempSubject - "
            }
            else
            {
                $tempCourseName = $subject.YearLevel+(($subject.'SIS ID').Substring(($subject.'SIS ID').IndexOf('.')+1))
                $tempSubject = $subject.'Course Name'
                $subject.'Course Name' = $tempCourseName -replace "-"," $tempSubject - "
            }
        }
        else
        {
                $tempCourseName = ($subject.'Course Name')+" "+(($subject.'SIS ID').Substring(($subject.'SIS ID').IndexOf('.')+1))
                $tempSubject = $subject.'Course Name'
                $subject.'Course Name' = $tempCourseName -replace "-"," - "
        }

    }
}