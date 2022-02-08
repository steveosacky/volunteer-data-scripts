#require version 5
class Javaupdate{    
    [string]$Site
    [string]$Address    
    [string]$City
    [string]$State
    [string]$Zip
    [string]$Color
    [string]$Type
}


$global:Rows=1000

#Absolute File path location of the excel file
$Global:FilePath= 'C:\users\SteveOsacky\Downloads\Steve Spreadsheet II.xlsx'

$outputArray=@()
function ReadDataInSpreadSheets{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=0)]$WorkSheet
    )
    begin{
        $output=@()
        $WSname= $WorkSheet.Name;
        $WorkSheet = $WorkBook.sheets.item($WorkSheet)
        #Activate the Wanted sheet
		$WorkSheet.Activate() | Out-Null
    }
    process{
        #read each row and if it's empty stop the process
        for($row=2; $row -le $global:Rows; $row++){
            $complete = $row/$global:Rows * 100;
            Write-Progress -Activity "Loading Data in $WSname" -Id 1 -PercentComplete $complete
            
            #if the value of action is empty, assume that there's no more data and break the collection of the information from file
			if([string]::IsNullOrEmpty($WorkSheet.Cells.Item($row,1).Text)){
                Write-Progress -Activity "Loading Data in $WSname" -Id 1 -PercentComplete 100 -Completed
                break;
			}

            $newobject = New-Object -TypeName Javaupdate
                                                  
            #column 1 (A2 Site)		              
            $newobject.Site = $WorkSheet.Cells.Item($row,1).Text;
          
		    #column 2 (B2 Address)            
            $newobject.Address = $WorkSheet.Cells.Item($row,2).Text;
            
            #column 3 (A2 City)            
		    $newobject.City = $WorkSheet.Cells.Item($row,3).Text;
            
		    #column 4 (B2 State)      
            $newobject.State = $WorkSheet.Cells.Item($row,4).Text;
            
            # #column 5 (A2 Zip)            		    
            $newobject.Zip = $WorkSheet.Cells.Item($row,5).Text;
           
            $newobject.Color = $WorkSheet.Cells.Item($row,5).Interior.ColorIndex;	            
            $newobject.Type = $WorkSheet.Name;
            
           
            $output+=$newobject
        }
    }
    end{
        return $output
    }

}

# Create an Object Excel.Application using Com interface
$readExcel = New-Object -ComObject Excel.Application
# Disable the 'visible' property so the document won't open in excel
$readExcel.Visible = $true

# Open the Excel file and save it in $WorkBook
$WorkBook = $readExcel.Workbooks.Open($Global:FilePath)

#Get All WorkSheets in the Book
$WorkSheetsName=@()
write-host -ForegroundColor Cyan "Loading Worksheets names"
foreach($item in $workBook.Worksheets){
    write-host -ForegroundColor Gray -BackgroundColor DarkBlue "Working with $($item.Name)"
	$WorkSheetsName+=$item.Name
}

#Read data in SpreadSheets
foreach($wb in $WorkSheetsName){
    $outputArray+= ReadDataInSpreadSheets $wb
}
#$outputArray | select -ExpandProperty ProductCode

Write-Host -BackgroundColor DarkGray -ForegroundColor Black "Use `$outputArray variable, to query the whole object"
$WorkBook.close()
$readExcel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($readExcel) | Out-Null


Write-Host "Original doc address count: " $outputArray.Count
$valid = @()
#$valid = $outputArray | Where-Object {$_.Color -ne 3 -and $_.Address -ne '' -and $_.Zip -ne ''}
$valid = $outputArray | Where-Object {$_.Address -ne '' -and (($_.City -ne '' -and $_.State -ne '') -or ($_.Zip -ne ''))}
Write-Host "Filtered address count: " $valid.Count

$results = @()
foreach ($v in $valid) {

    $newresult = New-Object System.Object                                                    	                  
    $newresult | Add-Member -MemberType NoteProperty -Name "Site" -Value $v.Site;  
    $newresult | Add-Member -MemberType NoteProperty -Name "Address" -Value $v.Address;
    $newresult | Add-Member -MemberType NoteProperty -Name "City" -Value $v.City;    
    $newresult | Add-Member -MemberType NoteProperty -Name "State" -Value $v.State;
    $newresult | Add-Member -MemberType NoteProperty -Name "Zip" -Value $v.Zip;
    $newresult | Add-Member -MemberType NoteProperty -Name "Color" -Value $v.Color;	            
    $newresult | Add-Member -MemberType NoteProperty -Name "Type" -Value $v.Type;       
    $results += $newresult
}
$date = Get-Date -Format "MM-dd-yyyy_HH_mm"
$Path = $Global:FilePath.Replace(".xlsx", "_$date.xlsx")
$results | Export-Excel $Path


Remove-Variable -Scope Global Rows,FilePath
Remove-Variable ReadExcel,workbook,WorksheetsName