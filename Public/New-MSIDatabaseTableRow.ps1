Function New-MSIDatabaseTableRow {
    <#
    .DESCRIPTION
    Inserts a row into an MSI Database. See https://learn.microsoft.com/en-us/windows/win32/msi/sql-syntax for info about proper syntax.

    .PARAMETER MSIPath
    The path to the MSI file to modify.

    .PARAMETER MSITableName
    The MSI database table to modify.

    .PARAMETER ColumnList
    A comma separated list of column names for the table that is to be changed, in order. This should be a single string, not an array. Items in the list should NOT be wrapped in single quotes.

    .PARAMETER ValueList
    A comma separated list of column values in order matching ColumnList. This should be a single string, not an array. Values should be wrapped in single quotes.

    .LINK
    https://github.com/zebulonsmith/NativePSMSI

    .EXAMPLE
    Insert Directory='Stink' Directory_Parent='Stank' DefaultDir='Stunk' into the Directory table.
    New-MSIDatabaseTableRow -MSIPath "C:\Users\zebul\OneDrive\Got Bored and Wrote Code\PackageManagerDrafts\7z2201-x64 - Modified.msi" -MSITableName Directory -ColumnList "Directory, Directory_Parent, DefaultDir" -ValueList "'Stink', 'Stank', 'Stunk'"

    .EXAMPLE
    Same as above, but returns the query that will be executed instead of making changes.
    New-MSIDatabaseTableRow -MSIPath "C:\Users\zebul\OneDrive\Got Bored and Wrote Code\PackageManagerDrafts\7z2201-x64 - Modified.msi" -MSITableName Directory -ColumnList "Directory, Directory_Parent, DefaultDir" -ValueList "'Stink', 'Stank', 'Stunk'" -TestQueryOnly

    #>
    param (
        [Parameter(Mandatory=$true)]
        [System.IO.FileInfo]$MSIPath,

        [Parameter(Mandatory=$true)]
        [string]
        $MSITableName,

        [Parameter(Mandatory=$true)]
        [string]
        $ColumnList,

        [Parameter(Mandatory=$true)]
        [string]
        $ValueList,

        [Parameter(Mandatory=$false)]
        [Switch]
        $TestQueryOnly
    )
    #Build the query
    $InsertQuery = "INSERT INTO ``$($MSITableName)`` ($($ColumnList)) VALUES ($($ValueList))"

    if ($TestQueryOnly) {
        Return $InsertQuery
    }

    #validate that the file exists
    if (!(test-path -path $MSIPath -PathType Leaf)) {
        Throw [System.IO.FileNotFoundException]::New("File $($MSIPath.FullName) not found.")
    }

    #Load the WindowsInstaller
    Try {
        $WindowsInstaller = New-Object -ComObject WindowsInstaller.Installer
    } Catch {
        Write-Error "Failed to create WindowsInstaller.Installer com object."
        Throw $_
    }

    #Open the MSI file in Transact mode
     Try {
        $MSIDBObject = $WindowsInstaller.Gettype().InvokeMember("OpenDatabase", "InvokeMethod", $null, $WindowsInstaller, @($MSIPath.fullname, 1))
    } Catch {
        Write-Error "Failed to open MSI Database $($MSIPath.FullName)."
        Throw $_
    }

    #Create a view object to insert the data
    Write-Verbose "Executing query '$($InsertQuery)' on the MSI database."

    Try {
        $UpdateView = $MSIDBObject.Gettype().InvokeMember("OpenView","InvokeMethod",$null,$MSIDBObject,($InsertQuery))
    } Catch {
        Write-Error "Query $($InsertQuery) is invalid."
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($MSIDBObject) | out-null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($WindowsInstaller) | Out-Null
        Throw $_
    }

    Try {
        $UpdateView.Gettype().InvokeMember("Execute","InvokeMethod", $null,$UpdateView,$null) | Out-Null
        $UpdateView.Gettype().InvokeMember("Close","InvokeMethod",$null,$UpdateView,$null) | Out-Null

        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($UpdateView) | Out-Null
    } Catch {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($UpdateView) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($MSIDBObject) | out-null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($WindowsInstaller) | Out-Null

        Write-Error "Failed to commit changes to the MSI Database."
        Throw $_
    }

    #Commit the changes
    $MSIDBObject.Gettype().InvokeMember("Commit","InvokeMethod",$null,$MSIDBObject,$null)


    #Release objects and stuff
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($MSIDBObject) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($WindowsInstaller) | Out-null
}