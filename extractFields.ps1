
#   Extracts fields from a PDF file using iTextSharp
#
#   Copyright (C) 2018, Cyber Moxie, LLC
#
#   This program is free software: you can redistribute it and/or modify
#   it under the terms of the GNU General Public License as published by
#   the Free Software Foundation, either version 3 of the License, or
#   (at your option) any later version.
#
#   This program is distributed in the hope that it will be useful,
#   but WITHOUT ANY WARRANTY; without even the implied warranty of
#   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#   GNU General Public License for more details.
#
#   You should have received a copy of the GNU General Public License
#   along with this program.  If not, see <https://www.gnu.org/licenses/>.
#
# 
#  This script extracts fields or values from a PDF file where values have the format:
#         
#           <Label>: <Value>
#
#  and prints the data to the screen in CSV format.  The process allows you to map the values in the file to a column 
#  in the CSV file by specifying a column name in the field definition field.  The enable you to extract values and combine
#  them into a single field.  This is helpful for address values were address, address2, city, state, zip are separate fields in the file, but
#  should be combined into a single column field in the output.  First name, last name is another example.  THIS ONLY WORKS FOR consectutive 
#  values in the file.  A future enhancement could sort the tuple below to combine fields that are dispersed.
# 
#  Parameters:
#      fieldLabelsFile - file containing field labels of values to search for and extract
#      inputFile - file that is target of extraction
#
param([string]$fieldLabelsFile="", [string]$inputFile="", [int]$maxPages=2)

# using the iTextSharp library that needs to be in the current folder
$itextsharpdll = ".\itextsharp.dll"
Add-Type -Path $itextsharpdll

$regex = "\:[ \t]+(.*)"
$identifyingStringNew = "Appraisal Quality Rating"

# create a tuple that contains the file label and the output fields column.
$fieldList = New-Object 'Collections.Generic.List[Tuple[string, string]]'
foreach($line in Get-Content $fieldLabelsFile) {
        $sourceFieldName, $fieldLabel = $line.Split(',')

        # the $ line.Split function returns arrays
        if ($fieldLabel[0].Length -gt 0)
        {
            $seperator = ","
        }
        else
        {
            $seperator = ""
        }
        $tuple = New-Object 'Tuple[string,string]'($sourceFieldName.Trim(), $fieldLabel.Trim())
        $fieldList.add($tuple)
}

# use the iTextSharp pdfreader object
$reader = New-Object iTextSharp.text.pdf.pdfreader -ArgumentList "$inputFile"
if ($reader.NumberOfPages -lt 1)
{
    Write-Host "Input file is not a valid file" + $orderFile.FullName;
}
else
{    
    $test = ""
    for ($i = 0; $i -lt $maxPages -and $i -lt $reader.NumberOfPages; ++$i)
    {

        $text += [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($reader, $i + 1).Split([char]0x000A)
    } 


    # remove the tabs and non-breaking space characters
    $normalizedText = $text -replace "`t"," "
    $normalizedText = $normalizedText -replace [char]0x00a0,' '

    $fieldValues = New-Object 'Collections.Generic.List[Tuple[string, string, string]]'

    # obtain all the values
    $i = 0
    foreach ($field in $fieldlist)
    {
        $lineIndex = [array]::FindIndex($normalizedText, [Predicate[String]]{param($s)$s.Contains($field.Item1) -eq 1})
        if ($lineIndex -ge 0)
        {
            $a = Select-String -Pattern $regex -InputObject $normalizedText[$lineIndex]
            if ($a.Matches.Groups.Count -gt 1)
            {
                # save the field name, field separator, and field value to a tuple
                $tuple = New-Object 'Tuple[string, string, string]'($field.Item2, $field.Item3, $a.Matches.Groups[1].Value)
                $fieldValues.add($tuple)
            }
        }
        $i++
    }
                                    
    # print the values to the output file
    $newColumn = $true
    $i = 0

    # start with the opening quotes
    $valuesString = ""
    $openingQuotes = "`"" 
    $separator = ""
    $lastFieldLabel = ""
    foreach ($tuple in $fieldValues)
    {                                        
        $fieldValue = $tuple.item3
        $fieldSeparator = $tuple.Item2
        $fieldLabel = $tuple.Item1

        if ($valuesString.Length -eq 0)
        {
            $valuesString = "`""
        }
        else
        {
            # if this value has the same label as the previous, then don't start a new column
            if ($fieldLabel.Trim() -ne $lastFieldLabel.Trim())
            {
                $valuesString += "`",`""
            }
            else
            {
                $valuesString += " "
            }
        }
        $valuesString += $fieldValue.Trim()
        $lastFieldLabel = $fieldLabel
    }
    $valuesString += "`""

    Write-Host $valuesString
}
$reader.Close()
 