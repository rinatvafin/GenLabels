#Set-Location "C:\Users\rinat_2\Desktop\LEGOR"
Write-Host "Скрипт должен запускаться в директории с xlsx файлами."
Write-Host "Файлы в текущей директории: "(Get-ChildItem).Name
$SetName = Read-Host -Prompt "Введите номер набора: "
$excel = New-Object -ComObject excel.application
$WorkFile = "$($pwd)\$setname.xlsx"
$WorkBook = $Excel.Workbooks.Open($WorkFile)
$SheetMainTable = $WorkBook.Worksheets.Item(1)
#$SheetColors = $WorkBook.Worksheets.Item(2)
$SheetReport = $WorkBook.Worksheets.Item(3)
$SheetReport.name = "Вкладыши"

$excelColor = New-Object -ComObject excel.application
$colorFile = "$($pwd)\colors.xlsx"
$colorBook = $ExcelColor.Workbooks.Open($ColorFile)
$SheetColors = $ColorBook.Worksheets.Item(1)

#VerticalAlignment:
$xlTop = -4160
$xlCenter = -4108
$xlBottom = -4107
#HorizontalAlignment:
$xlLeft = -4131
$xlCenter = -4108
$xlRight = -4152
#Borders
$xlContinuous = 1
$xlThin=2
$xlMedium = -4138
$xlThick = 4

$EndFile = $True
$MainObject = @()
[int]$Column = 3
[int]$row = 5
While($EndFile -eq $True)
{
    $obj = New-Object PSObject
    $obj | Add-Member -MemberType NoteProperty -Name "ID" -Value ($SheetMainTable.cells.Item($row,$Column).text).Substring(1)
    $obj | Add-Member -MemberType NoteProperty -Name "Color" -Value $SheetColors.cells.item((([int](($SheetMainTable.cells.Item($row,$Column).hyperlinks[1] | select-object Address).address.split("=")[2]))+1), 3).text
    $Column = 4
    $obj | Add-Member -MemberType NoteProperty -Name "Name" -Value (($SheetMainTable.cells.Item($row,$Column).text).replace($obj.Color,"")).Substring(1)
    $Column = 7
    $obj | Add-Member -MemberType NoteProperty -Name "Count" -Value $SheetMainTable.cells.Item($row,$Column).text
    $Column = 4
    $row++
    $obj | Add-Member -MemberType NoteProperty -Name "Type" -Value (($SheetMainTable.cells.Item($row,$Column).text).Split(":")[1]).Substring(1)
    $obj | Add-Member -MemberType NoteProperty -Name "Model" -Value (($SheetMainTable.cells.Item($row,$Column).text).Split(":")[2]).Substring(1)
    "Деталь: "+$MainObject.count+" Название: "+$obj.Name
    $MainObject += $obj
    $row+=3
    $column = 3

        if(($SheetMainTable.cells.Item($row,$Column).text) -eq "")
        {$row++
         if(($SheetMainTable.cells.Item($row,$Column).text) -eq "")
            {$row++
             if(($SheetMainTable.cells.Item($row,$Column).text) -eq "")
             {$row++
             if(($SheetMainTable.cells.Item($row,$Column).text) -eq ""){$EndFile = $False}
             }
            }
        }
}


[int]$row = 1
[int]$columMax = 6
$NumberOfDetails = 0
$columnWidth = 16
$SheetReport.columns.item('A').columnWidth = $columnWidth 
$SheetReport.columns.item('B').columnWidth = $columnWidth 
$SheetReport.columns.item('C').columnWidth = $columnWidth 
$SheetReport.columns.item('D').columnWidth = $columnWidth 
$SheetReport.columns.item('E').columnWidth = $columnWidth 
while ($NumberOfDetails -lt $MainObject.Count)
{
    $column = 1
    While(($column -lt $columMax) -and ($NumberOfDetails -lt $MainObject.Count))
    {
        if ($MainObject[$NumberOfDetails].count -gt 0)
        {
        $SheetReport.Cells.Item($row,$Column) = $MainObject[$NumberOfDetails].Count
        $SheetReport.Cells.Item($row,$Column).font.size = 12
        $SheetReport.Cells.Item($row,$Column).borders.LineStyle = 1
        $SheetReport.Rows.item($row).RowHeight  = 90
        $SheetReport.Rows.Item($row).HorizontalAlignment = $xlLeft
        $SheetReport.Rows.Item($row).VerticalAlignment = $xlTop
        $row++

        $SheetReport.Cells.Item($row,$Column) = $MainObject[$NumberOfDetails].ID
        $SheetReport.Cells.Item($row,$Column).font.size = 12
        $SheetReport.Cells.Item($row,$Column).borders.LineStyle = 1
        $SheetReport.Rows.Item($row).HorizontalAlignment = $xlCenter 
        $SheetReport.Rows.Item($row).VerticalAlignment = $xlCenter
        $row++

        $SheetReport.Cells.Item($row,$Column) = $MainObject[$NumberOfDetails].Name
        $SheetReport.Cells.Item($row,$Column).font.size = 10
        $SheetReport.Cells.Item($row,$Column).borders.LineStyle = 1
        $SheetReport.Rows.item($row).RowHeight  = 50       
        $SheetReport.Rows.item($row).WrapText = $True
        $SheetReport.Rows.Item($row).HorizontalAlignment = $xlLeft 
        $SheetReport.Rows.Item($row).VerticalAlignment = $xlCenter
        $row++

        $SheetReport.Cells.Item($row,$Column) = $MainObject[$NumberOfDetails].Color
        $SheetReport.Cells.Item($row,$Column).font.size = 12
        $SheetReport.Cells.Item($row,$Column).borders.LineStyle = 1
        $SheetReport.Rows.Item($row).HorizontalAlignment = $xlLeft
        $SheetReport.Rows.Item($row).VerticalAlignment = $xlCenter
        $row++

#        $SheetReport.Cells.Item($row,$Column) = $MainObject[$NumberOfDetails].Type
#        $row++
        $SheetReport.Cells.Item($row,$Column) = $MainObject[$NumberOfDetails].Model
        $SheetReport.Cells.Item($row,$Column).font.size = 12
        $SheetReport.Cells.Item($row,$Column).borders.LineStyle = 1
        $SheetReport.Rows.Item($row).HorizontalAlignment = $xlLeft
        $SheetReport.Rows.Item($row).VerticalAlignment = $xlCenter
        $row-=4
        $NumberOfDetails++
        $column++
        "$NumberOfDetails"+" - отрисован, Column $column"
        
       }
       else { 
        $NumberOfDetails++
        "$NumberOfDetails"+" - пропущен, Column $column"
             }
    }

$row+=5
"Отрисовано "+$NumberOfDetails+" из "+$MainObject.Count
}

#$Start = "A1"
#$End = "E"+"$row"
#$RangeCoordinates = "$Start`:$End"
#$range1 = $SheetReport.Range($RangeCoordinates)
#$range1.Borders.LineStyle = 1
#$SheetReport.Rows.Item($row).entirerow.rowheight

#Сохраняем книгу
$WorkBook.Saveas($WorkFile)
#Закрываем книгу набора
$WorkBook.Close()
$Excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
Remove-Variable excel

#Закрываем книгу цветов
$ColorBook.Close()
$ExcelColor.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelColor)
Remove-Variable ExcelColor

Read-Host -Prompt "Генерация списка табличек закончена. Файл $WorkFile обновлен. Нажмите Enter чтобы закрыть."