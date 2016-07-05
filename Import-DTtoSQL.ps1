<# 
.SYNOPSIS 
	Imports all Columns and Rows of a Datatable into a SQL Table
.DESCRIPTION 
	Imports all Columns and Rows of a Datatable intor a SQL Table. 
	
	This module does not use the DefaultView of a datatable so commit all changes or make a new table from DefaultView before passing to this module
		
		By Default:
			This Module will always drop and create the SQL Table if the Table Exists
			This Module will always create the SQL Table if it does not Exist
			
			The default column size for VarChar is Max
			The default column size for Decimal is 18,2
					
		Optional:
			Setting -blnKeepExistingTable to $true will basically append the data to the existing table, but will create the table if it does not exist
			
.INPUTS 
	Object 
   	Any PowerShell System.Data.DataTable 
.OUTPUTS 
   A SQL Table with Datatypes Matching the Columns in the Datatable
.EXAMPLE 
	Import-DTtoSQL -datatable "DataTableObject" -strSQLTableName "TableName" -strSQLDatabase "DatabaseName" -strSQLServer "ServerName" [ -blnKeepExistingTable "$true" ]
.NOTES 
	Author: Jace Jenkins
	Version: 1.0
#>

function Import-DTtoSQL{

	[CmdletBinding()]
		
	param (
		[Parameter(Position=0,Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
  		[object]$datatable = $(throw "-a datatable is required."),
		
		[Parameter(Position=1,Mandatory = $true, ValueFromPipelineByPropertyName=$true)]
		[string]$strSQLTableName = $(throw "-a SQL Table Name is required."),
		
		[Parameter(Position=2,Mandatory = $true, ValueFromPipelineByPropertyName=$true)]
		[string]$strSQLDatabase= $(throw "-a SQL Database Name is required."),
		
		[Parameter(Position=3,Mandatory = $true, ValueFromPipelineByPropertyName=$true)]
		[string]$strSQLServer= $(throw "-a SQL Server Name is required."),
		
		[Parameter(Position=4,Mandatory = $false, ValueFromPipelineByPropertyName=$true)]
		[boolean]$blnKeepExistingTable = $false		
	)
		
	# Create Connection String, Results Datatable, SQL Adapater
	$strSQLConnectionString = "Data Source=$strSQLServer;Initial Catalog=$strSQLDatabase;Integrated Security=True"
	$dtResults = new-object "System.Data.Datatable"
	$objSQLDataAdapter = New-Object System.Data.SqlClient.SqlDataAdapter   
	
	# Test for existance of the Table
	$strSQLQuery = "select count(*) from $strSQLTableName"
	$objSQLDataAdapter.SelectCommand = new-object System.Data.SqlClient.SqlCommand ($strSQLQuery,$strSQLConnectionString)                        
	
	# See if the table exists
	try {$objSQLDataAdapter.Fill($dtResults); $blnTableExists = $true; Write-Debug "Table Exists: $strSQLTableName"} catch {$blnTableExists = $false; Write-Debug "Table Does Not Exist: $strSQLTableName"}
					
	# Drop Existing Table only if $blnKeepExistingTable = $false
	if ($blnTableExists -and !($blnKeepExistingTable)){
		
		$strSQLQuery = "DROP TABLE $strSQLTableName"
		Write-Debug $strSQLQuery
		
		$objSQLDataAdapter.SelectCommand = new-object System.Data.SqlClient.SqlCommand ($strSQLQuery,$strSQLConnectionString)        
		
		try {
			$objSQLDataAdapter.Fill($dtResults)
			Write-Debug "Drop Successful"
		} catch {
			$strError = ""
			$strError = $Error[0].ToString()
			Write-Debug $strError
		}
		$blnTableExists = $false
	} elseif ($blnTableExists -and $blnKeepExistingTable){
		Write-Debug "Keeping Existing Table: $strSQLTableName"
	}
	
	# Create the Table if it Does not Exists
	If (-not($blnTableExists)){
	
		# Build Create Table Syntax
		$strColumns = ""
		
		foreach ($dtcolumn in $datatable.Columns){
				
			# Transpose DataTable Types to SQL Column Types 
			$strSQLDataType = switch ($dtcolumn.DataType.Name)
		   		{
		        "Boolean" {"Bit"}
		        "Byte[]" {"VarBinary"}
		        "Byte"  {"VarBinary"}
		        "Datetime"  {"DateTime"}
		        "Decimal" {"Decimal(18,2)"}
		        "Double" {"Float"}
		        "Guid" {"UniqueIdentifier"}
		        "Int16"  {"SmallInt"}
		        "Int32"  {"Int"}
		        "Int64" {"BigInt"}
		        default {"VarChar(Max)"}
		    	}
				
			#Construct the SQL Syntax
			$strColumns += ([char]34 + $dtcolumn.ColumnName + [char]34 + " " + "$strSQLDataType,")
		}
		
		# Trim the Common off the End of the Sting
		$strColumns = $strColumns.Substring(0,($strColumns.Length -1 ))
		
		# Create SQL Table
		$strSQLQuery = "CREATE TABLE $strSQLTableName ($strColumns)"
		Write-Debug $strSQLQuery
		
		$objSQLDataAdapter.SelectCommand = new-object System.Data.SqlClient.SqlCommand ($strSQLQuery,$strSQLConnectionString)        
		
		try {
			$objSQLDataAdapter.Fill($dtResults)
			Write-Debug "Create Successful"
		} catch {
			$strError = ""
			$strError = $Error[0].ToString()
			Write-Debug $strError
		}
	}
	
	#Upload Data to SQL
	$bulkCopy = new-object ("Data.SqlClient.SqlBulkCopy") $strSQLConnectionString
	$bulkCopy.DestinationTableName = $strSQLTableName
	$bulkCopy.BatchSize = 5000
    	$bulkCopy.BulkCopyTimeout = 0
   
	Write-Debug "Upload Data to: $strSQLTableName"
	
	try {
		$bulkCopy.WriteToServer($datatable)
		Write-Debug "Upload Data Successful"
	} catch {
		$strError = ""
		$strError = $Error[0].ToString()
		Write-Debug $strError
	}	
}