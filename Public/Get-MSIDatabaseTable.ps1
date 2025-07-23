Function Get-MSIDatabaseTable {
    <#
    .DESCRIPTION
    Opens an MSI file and retrieves the specified database table. The Filter parameter can be used to narrow the data returned.
    See https://learn.microsoft.com/en-us/windows/win32/msi/sql-syntax for information about proper syntax when querying the MSI database.


    .PARAMETER MSIPath
    The path to the MSI file to be queried.

    .PARAMETER MSITableName
    The name of the table to query. If no table is specified, the Property table is queried by default.

    .PARAMETER Filter
    Adds a "Where" clause to the query.

    .PARAMETER TestQueryOnly
    Returns the query that will be executed instead of opening the MSI database.

    .LINK
    https://github.com/zebulonsmith/NativePSMSI

    .EXAMPLE
    Query the Property table.
    Get-msitable -MSIPath .\7z2201-x64.msi

    .EXAMPLE
    Query the Feature table
    Get-MSIDatabaseTable -MSIPath .\7z2201-x64.msi -MSITableName "Feature"

    .EXAMPLE
    Query the Feature table, but only return rows where the Feature_Parent column's value is "Complete"
    Get-MSIDatabaseTable -MSIPath .\7z2201-x64.msi -MSITableName "Feature" -Filter "Feature_Parent='Complete'"

    .EXAMPLE
    Query the Feature table with a filter and return the query that will be executed without opening the MSI.
    Get-MSIDatabaseTable -MSIPath .\7z2201-x64.msi -MSITableName "Feature" -Filter "Feature_Parent='Complete'" -TestQueryOnly
    #>
    param (
        [Parameter(Mandatory=$true)]
        [String]$MSIPath,

        [Parameter(Mandatory=$false)]
        [string]$MSITableName = "Property",

        [Parameter(Mandatory=$false)]
        [string]
        $Filter,

        [Parameter(Mandatory=$false)]
        [switch]
        $TestQueryOnly
    )

    if ([string]::IsNullOrEmpty($filter)) {
        $MSIQuery = "Select * from $($MSITableName)"
    } else {
        $MSIQuery = "Select * from $($MSITableName) WHERE $($Filter)"
    }

    if ($TestQueryOnly) {
        Return $MSIQuery
    }

    #validate that the file exists
    if (!(test-path -path $MSIPath -PathType Leaf)) {
        Throw [System.IO.FileNotFoundException]::New("File $($MSIPath) not found.")
    } else {
        $MSIFile = Get-Item -Path $MSIPath
    }

    #Load the WindowsInstaller
    Try {
        $WindowsInstaller = New-Object -ComObject WindowsInstaller.Installer
    } Catch {
        Write-Error "Failed to create WindowsInstaller.Installer com object."
        Throw $_
    }

     #Open the MSI file as Read Only
     Try {
        $MSIDBObject = $WindowsInstaller.Gettype().InvokeMember("OpenDatabase", "InvokeMethod", $null, $WindowsInstaller, @($MSIFile.fullname, 0))
    } Catch {
        Write-Error "Failed to open MSI Database $($MSIFile.FullName)."
        Throw $_
    }

    #See if the specified table exists.
    $TablePersist = $MSIDBObject.GetType().InvokeMember("TablePersistent","GetProperty",$null,$MSIDBObject,$MSITableName)
    If ($TablePersist -ne 1) {
        Throw "The specified table doesn't exist."

    }

    #Create the View Object
    Try {
        $ThisView = $MSIDBObject.GetType().InvokeMember("OpenView","InvokeMethod",$null, $MSIDBObject, $MSIQuery)
    }
    Catch {
        Write-Error "Query $($MSIQuery) is not valid."
        #Close the database
        $MSIDBObject.GetType().InvokeMember("Commit", "InvokeMethod", $null, $MSIDBObject, $null)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($WindowsInstaller) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($MSIDBObject) | Out-Null
        $MSIDBObject = $null

        Throw $_
    }


    #Get the column count for the query results in the view.
    $columnNameRecord = $thisview.GetType().InvokeMember("ColumnInfo", "GetProperty", $null, $thisview, 0)

    #Get an array of psobjects representing the column index and name for the query results.
    $ColumnsCount = $columnNameRecord.gettype().InvokeMember("FieldCount", "GetProperty", $null, $columnNameRecord, $null)

    $TableColumns = @{}
    For ($i=1; $i -le $ColumnsCount; $i++) {
        $thisColumn = $columnNameRecord.GetType().InvokeMember("StringData", "GetProperty", $null, $columnNameRecord, $i)
        if ([string]::IsNullOrEmpty($thisColumn)) {
            $thisColumn = "NULL"
        }
        $TableColumns.add($thisColumn,$i)
    }

    #Execute the query
    $ThisView.GetType().InvokeMember("Execute", "InvokeMethod", $null, $ThisView, $null)


    $QueryResults = @()
    #Fetch the first row of the query result, then loop through each row and fetch the next until nothing is returned.
    $QueryRecord = $ThisView.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $ThisView, $null)
    While ($null -ne $QueryRecord) {

        $thisRow = [PSCustomObject]@{}
        #Enumerate the table column list and generate one row of the return results for each column.
        Foreach ($thisTableColumn in $TableColumns.GetEnumerator()) {
            $columnValue = $queryRecord.Gettype().InvokeMember("StringData", "GetProperty", $null, $queryRecord, $thisTableColumn.value)
            $thisRow | Add-Member -MemberType NoteProperty -Name $thisTableColumn.name -Value $columnValue
        }
            $QueryResults += $thisRow
        #Fetch the next record
        $QueryRecord = $ThisView.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $ThisView, $null)
    }

    #Close out the View
    $thisview.GetType().InvokeMember("Close", "InvokeMethod", $null, $thisview, $null)
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ThisView) | Out-Null
    $thisview = $null

    #Close the database
    $MSIDBObject.GetType().InvokeMember("Commit", "InvokeMethod", $null, $MSIDBObject, $null)


    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($WindowsInstaller) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($MSIDBObject) | Out-Null
    $MSIDBObject = $null



    Return $QueryResults

}