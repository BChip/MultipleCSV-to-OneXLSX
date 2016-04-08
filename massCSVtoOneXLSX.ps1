function convertToExcel {
    $rootFolder = Read-Host "Folder?"
    $filename = Read-Host "Output Filename?"
    $files = dir -Path $rootFolder *.csv

    $excel = New-Object -ComObject Excel.Application
    $excel.visible = $true
    $workbook = $excel.workbooks.add()
    $sheets = $workbook.sheets
    $sheetCount = $Sheets.count
    $mySheet = 1
    $mySheetName = "Sheet" + $mySheet
    $s1 = $sheets | where{$_.name -eq $mySheetName}
    $s1.Activate()

    If ($sheetCount -gt 1) {
        #Delete other Sheets
        $Sheets | ForEach {
            $tmpSheetName = $_.Name
            $tmpSheet = $_
            If ($tmpSheetName -ne "Sheet1") {
                $tmpSheet.Delete()
                }
            }
        }

        ForEach ($file in $files) {
            If ($mySheet -gt 1){
                $s1 = $workbook.sheets.add()
                }
            $s1.Name = $file.BaseName
            $s1.Activate()
            $s1Data = Import-Csv $file.FullName
            $s1Data | ConvertTo-Csv -Delimiter "`t" -NoTypeInformation | Clip
            $s1.cells.item(1,1).Select()
            $s1.Paste()
            $mySheet ++
            }

        $workbook.SaveAs("$rootFolder\$filename")
        $excel.Quit()
    }

    convertToExcel
    rm "$rootFolder\*.csv"