# 
# Sort-MTS.ps1 - This script cleans up the exported MTS file from Tuberacker and sorts it by tube location
# Author             - Larkin O'quinn
# Date               - 071321
#
####--Changelog--####
# v1.0  - Initial version
# v1.1  - Added dynamic column header finder, as evidently the columns can shift around during report export from Tuberacker - LO 071421
# v1.2  - Fixed issue with regex in position column finder. Changed $outputfilepath to target whichever directory the csv ws imported from. 
#           Added total sample count as title of output document. Added -or "EDTA" to tube type column finder, and changed variable name. LO 071521
# v1.3  - Added every-other shading and page numbers to output document. LO 071921
# v1.4  - Added duplicate sample position identifier and only retains the newest position in the final output. 
#       - In $outputfile, changed .Trim() to .TrimEnd() to resolve filepath issue.
#       - Added expression to scoot "Date Racked" property up a line in the $import array. LO 072021
# v1.41 - Fixed column 5 width issue. Added retention of Q rack samples from duplicate checker, and removal of Q EDTA positions. LO 072321

#OpenFileDialog form for retrieval of csv
Write-host "Please select MTS export" -BackgroundColor Black
Add-Type -AssemblyName System.Windows.Forms
$MTS = New-Object System.Windows.Forms.OpenFileDialog
$MTS.filter = "csv (*.csv)| *.csv"
$MTS.Title = "Please select MTS export csv"
[void]$MTS.ShowDialog()
$MTS.FileName

#--------Csv file work-----------#
[pscustomobject[]]$samples = $null
$outputfile = "$($MTS.FileName.TrimEnd($($MTS.Safefilename)))MTSexport$($MTS.SafeFileName.Trim(".csv")).xlsx"
$import = Get-Content -Path $MTS.FileName | Select-Object -Skip 4 | convertfrom-Csv -WarningAction SilentlyContinue

#Moves 'Date Racked' cells up one to be with the appropriate sample
$i = 0
foreach ($line in $import) {
    if ($line.'Date Racked' -ne '') {
        $import[($i-1)].'Date Racked' += $line.'Date Racked'
        }
    $i++
    }

#Removes blanks
[System.Collections.Generic.List[pscustomobject]]$samples = @()
foreach ($sample in $import) {
    if ($sample.'Tube Cart' -ne '') {
        $null = $samples.Add($sample)
        #$samples += $sample
        }
    }

#Checks for samples with multiple positions and finds the associated sample ID and applies it to that row
$i = 0
foreach ($sample in $samples) {
    if ($sample.'sample ID' -eq '' -and $sample.'Tube Cart' -ne '') {
        $ID = $sample.'sample ID'
        $j = $i
        while ($ID -eq '') {
            $j --
            $ID = $($samples[($j)].'Sample ID')
            }
        $samples[$i].'Sample ID' = $ID
        }
    $i ++
    }

#Find data column headers
foreach ($prop in $($samples | gm -MemberType NoteProperty)) {
    if ($prop.Definition -match "CLOT" -or $prop.Definition -match "EDTA") {
        $typecol = $prop.Name
        }
    if ($prop.Definition -match "=\d") {
        $poscol = $prop.Name
        }
    }

#Retypes Position and Date Racked column into ints and DateTimes for sorting
$samples | % {$_.$poscol = [int]$_.$poscol}
$samples | % {$_.'Date Racked' = [System.DateTime]$_.'Date Racked'} 

#----------------------------#

#Identifies and removes duplicates, sorts samples, creates Excel object for formatting
$duplicate = ($samples | where -FilterScript {$_.'Tube Cart' -notlike "Q0*"}) | Group-Object -Property "Sample ID" | Where-Object -FilterScript {$_.Count -gt 1}
[pscustomobject[]]$removal = $null
foreach ($item in $duplicate) {
    $removal += $item.group | where -FilterScript {$_.'Date Racked' -ne ($item.Group | Measure-Object -Property 'Date Racked' -Maximum).Maximum}
    }

foreach ($thing in $removal) {
    $samples.Remove($thing) | Out-Null
    }

#Removes any EDTA postions in Q racks
[pscustomobject[]]$Qremoval = $null
foreach ($item in $samples) {
    $Qremoval += $item | where -FilterScript {$_.'Tube Cart' -like "Q0*" -and $_.$typecol -eq "EDTA"}
    }
foreach ($thing in $Qremoval) {
    $samples.Remove($thing) | Out-Null
    }
  

$excel = ($samples | select -Property 'Sample ID', @{N='Tube Type';E={$_.$typecol}}, 'Tube Rack', @{N='Position';E={$_.$poscol}}, @{N='Date Racked';E={$_.'Date Racked'}} | sort -Property 'Tube Rack', Position) | Export-Excel -Path $outputfile -Title "Total samples: $($samples.Count)" -WorksheetName LocatedSamples -PassThru -TitleSize 16 -TitleBold -AutoSize


#Formats, exports, and opens final Excel file
Add-ConditionalFormatting -Worksheet $excel.Workbook.Worksheets[1] -Address "a2:h100000" -ConditionValue "=MOD(Row(),2)=1" -BackgroundColor LightGray -RuleType Expression
$excel.Workbook.Worksheets[1].HeaderFooter.FirstFooter.CenteredText = "Page &P of &N"
$excel.Workbook.Worksheets[1].HeaderFooter.EvenFooter.CenteredText = "Page &P of &N"
$excel.Workbook.Worksheets[1].HeaderFooter.OddFooter.CenteredText = "Page &P of &N"
$excel.Workbook.Worksheets[1].Cells["D3:D100000"].Style.HorizontalAlignment = "Center"
$excel.Workbook.Worksheets[1].Column(5).Width = 18
Close-ExcelPackage -Show $excel