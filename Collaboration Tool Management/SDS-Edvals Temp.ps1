$TempTest = Import-Csv '.\SDS\Sem 1,2\Section.csv' | SELECT 'SIS ID','Course Subject','Course Name','YearLevel','CorrectedSubject'
$TempTest += Import-Csv '.\SDS\Sem 2\Section.csv' | SELECT 'SIS ID','Course Subject','Course Name','YearLevel','CorrectedSubject'

foreach ($subject in $TempTest)
{
    #$subject.'Course Subject' = $subject.'Course Subject' -replace '[^a-zA-Z]',''
    if( $subject.'Course Subject'.Substring(0,1) -gt 1)
    {
        $subject.'CorrectedSubject' = "0$($subject.'Course Subject')"
    }
    else
    {
        $subject.'CorrectedSubject' = $subject.'Course Subject'
    }

    $subject.'YearLevel' = $subject.'CorrectedSubject'.Substring(0,2)

    $subject.'Course Name' = $subject.'Course Name' -replace '-and-','&'
}

$TempTest | Sort-Object -Unique -Property "CorrectedSubject","Course Name" | FT
$TempTest | Sort-Object -Property "CorrectedSubject","Course Name" | FT