<# 
.SYNOPSIS 
	Exports a Datatable and Headers to an Excel Spreadsheet
.DESCRIPTION 
	Exports a Datatable and Headers to an Excel Spreadsheet
		
		By Default:
			This Module will always freeze the top row and first column
			This Module will always write Column Names as Headers
			This Module will shut down any running copies of excel
		
		Options:
			-strSort Index is used to A-Z Sort on a Column
			-blnFreezePanes will Freeze the top and first column
			-blnHeader is used to control wether or not to use the datatable column names as a header or not
			-debug will output module progress to the Debug channel
		
.INPUTS 
	Object 
	Any PowerShell System.Data.DataTable 
.OUTPUTS 
   	An Excel Spreadheet
.EXAMPLE 
	Out-DTtoExcel -DataTable $datatable -strSheet "SheetName" -strPath "Path to File" -strSortIndex "Column to Sort" -blnFreezePanes  "$True or $False"  -blnHeader "$True or $False" -debug "To see Verbose Output"
.NOTES 
	Author: Jace Jenkins
	Version: 1.3
		1.3 - Added the blnHeader Argument
		1.2 - Used a copy paste method to write spreadsheet data
		1.0 - Created Script
	Dependancies:
	
#>


function Out-DTtoExcel{

	[CmdletBinding()]
	
	param (
		
		# Function Parameters
		[Parameter(Position=0,Mandatory = $true)]
		[Object]$datatable = '',
		
		[Parameter(Position=1,Mandatory = $false)]
		[string]$strSheet = '',
		
		[Parameter(Position=2,Mandatory = $true)]
		[string]$strPath = '',
		
		[Parameter(Position=3,Mandatory = $false)]
		[string]$strSortIndex = '',
		
		[Parameter(Position=4,Mandatory = $false)]
		[boolean]$blnFreezeTopRow = $true,
		
		[Parameter(Position=5,Mandatory = $false)]
		[boolean]$blnHeaderRow = $true
	)
	
	# If Debug Mode is Turned on via -debug, Enable "Continue" vs the default of "Inquire", the Else stops cascading -debugs from called scripts
	if ($PSCmdlet.MyInvocation.BoundParameters["Debug"].IsPresent) {$DebugPreference = "Continue"} else {$DebugPreference = "SilentlyContinue"}
	
	# Test to ensure datatable has data
	if ($datatable.DefaultView.Count -eq 0){ Write-Debug "No Data Found in Datatable"; return}
	
	# See if Excel is Running for this user and Kill it
	$objAllProcesses = Get-WmiObject -Class Win32_Process | where {$_.ProcessName -match "Excel"}
	foreach ($objProcess in $objAllProcesses){
		if (($objProcess | Measure-Object).Count -ne 0){
			
			try{$objProcessOwner = ($objProcess.GetOwner()).User}
			catch {Write-Debug "Could Not detect process owner, skipping excel shutdown" ; $objProcessOwner = "None"}
			
			if ($objProcessOwner -eq $env:username){
				Stop-Process -Id $objProcess.ProcessID
			}
		}
	}
	$objAllProcesses = $null
	
	# Excel Constants
	$xlAscending = 1
	$xlDescending = 2
	$xlSortRows = 2
	$xlTopToBottom = 1
	$xlSortOnValues = 0
	$xlSortNormal = 0
	$xlYes = 1
	$xlNo = 2
	
	# Open Excel in Ram
	$objExcel = New-Object -ComObject Excel.Application
	$objExcel.Visible = $false
	$objExcel.DisplayAlerts = $False
	
	# If the Spreadsheet does not Exist Create it and Delete all But 1 Tab
	If (!(Test-Path $strPath)){
	
		Write-Debug "File does not exists $strPath"
		
		# Add a New Workbook
		Write-Debug "Create a new workbook"
		$objWorkbook = $objExcel.Workbooks.Add()
		
		#Delete the Default Extra Sheets
		Write-Debug "Deleting Sheet 2 and Sheet 3"
		$objWorksheet = $objWorkbook.Sheets | Where {$_.Name -eq "sheet2"} 
		$objWorksheet.Delete() 
		$objWorksheet = $objWorkbook.Sheets | Where {$_.Name -eq "sheet3"} 
		$objWorksheet.Delete() 
		
		# Attached to Sheet1
		$objWorksheet = $objWorkbook.Sheets | Where {$_.Name -eq "sheet1"}  
		$objWorksheet.Name = $strSheet
		
	} else {
		
		Write-Debug "File exists $strPath"
		
		# Open the Existing WorkBook
		Write-Debug "Opening existing file"
		$objWorkbook = $objExcel.Workbooks.Open($strPath) 
		
		# See if the Tab Already Exists Clear the Datra 
		$objWorksheet = $objWorkbook.Sheets | Where {$_.Name -eq $strSheet} 
		If ($objWorksheet.Name -eq $strSheet){
			Write-Debug "Current Sheet $strSheet Exists, renaming it to $($strSheet)_old "
			$objWorksheet.Name = $strSheet + "_old"
			$blnDeleteOld = $true
		} else {$blnDeleteOld = $false}
	
		# Create a new sheet and rename it
		Write-Debug "Creating new sheet named $strSheet"
		$objWorksheet = $objWorkbook.Sheets.Add()
		$objWorksheet = $objWorkbook.Sheets | Where {$_.Name -eq 'Sheet1'} 
		$objWorksheet.Name = $strSheet
		
		# See if there is an old Tab to delete
		$objWorksheet = $objWorkbook.Sheets | Where {$_.Name -eq $strSheet + "_old"} 
		If ($blnDeleteOld){
			
			Write-Debug "Deleting $($strSheet)_old"
			$objWorksheet.Delete()
		} 
		
		# Connect back to the new one just created
		$objWorksheet = $objWorkbook.Sheets | Where {$_.Name -eq $strSheet}
	}

	# Write Header Based on Columns Names
	if ($blnHeaderRow){Write-Debug "Writing Header Line"; $intHeader = 1} else {Write-Debug "Not Writing Header Line"; $intHeader = 0}
	$intColumnCtr = 1
	foreach ($dtColumn in $datatable.Columns){
		if ($blnHeaderRow){$objWorksheet.Cells.Item($intHeader,$intColumnCtr) = (($dtColumn.ColumnName).Trim())}
		$intColumnCtr++
	}
			
	# Convert DataTable to an Mutidimentional Array[Row,Column]
	Write-Debug "Converting Datatable to a $($datatable.DefaultView.Count) by $intColumnCtr Multidimentional Array"
	$RowData = New-Object 'object[,]' $datatable.DefaultView.Count,$intColumnCtr
	$intCurrentRecord = 0
	foreach ($dtRecord in $datatable.DefaultView){
		
		#Write Array Left to Right
		for ($intField=0; $intField -lt ($intColumnCtr - 1) ; $intField++){
			$RowData[$intCurrentRecord,$intField] = $dtRecord[$intField] 
		}
		
		# Increment Line Counter after every Row
		$intCurrentRecord++
	}
	
	# Excel uses Alpha's as Columns Coordinates, Convert Column Count to an Alpha
	$intStartA = ([int][char]("A")) - 1
	if ($intColumnCtr -gt 26) {
    	$strColumnAplha = [char]([int][math]::Floor($intColumnCtr/26) + $intStartA) + [char](($intColumnCtr%26) + $intStartA)
	} else {
    	$strColumnAplha = [char]($intColumnCtr + $intStartA)
	}
	
	# Write Row Data on A1 and add $IntHeader if a Header was written
	Write-Debug "Writing Row Data"
	$Range = $objWorksheet.Range("A$(1+$intHeader)","$strColumnAplha$($intCurrentRecord+$intHeader)")
	$Range.Value2 = $RowData
		
	#Sort Option, Pass the Column or Row you want to sort on
	if($strSortIndex.Length -ne 0){
		
		Write-Debug "Sorting Sheet on $strSortIndex"
		
		# Range of Data used in the spreadsheet
		$objRange = $objWorksheet.UsedRange
		
		# Column or Row to Sort
		$objSortRange = $objWorksheet.Range($strSortIndex)

		# Apply Sort
		[void]$objWorksheet.Sort.SortFields.Add($objSortRange, $xlSortOnValues, $xlDescending, $xlSortNormal)
		$objWorksheet.Sort.SetRange($objRange)
		$objWorksheet.Sort.Header = $xlYes
		$objWorksheet.Sort.Orientation = $xlTopToBottom
		$objWorksheet.Sort.apply()
		
		# Remove COM Object
		[Void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($objRange)
		[Void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($objSortRange)
	}
	
	#Freeze top Row and Column
	if($blnFreezeTopRow -eq $true){

		Write-Debug "Freezing top pane of spreadsheet"
		$objWorksheet.Application.ActiveWindow.SplitColumn = 1
		$objWorksheet.Application.ActiveWindow.SplitRow = 1
		$objWorksheet.Application.ActiveWindow.FreezePanes = $true
	}
	
	#AutoSize Columns
	Write-Debug "Autofitting Columns"	
	[Void]$objWorksheet.Columns.AutoFit()
	
	#Save and Quit
	Write-Debug "Save and Quit $strPath"
	$objWorksheet.SaveAs($strPath)
	$objExcel.Quit()
	
	# Remove COM Object
	[Void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel)
	[Void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($objWorkbook)
	[Void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($objWorksheet)

}
