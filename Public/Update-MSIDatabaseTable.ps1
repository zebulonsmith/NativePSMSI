Function Update-MSIDatabaseTable {
    <#
    .DESCRIPTION
    Changes the specified row in an MSI Database. It will be necessary to find the Primary Key for the table that is being changed.
    The primary key can be found easily by using Orca's "Adjust Schema" function or by referencing https://learn.microsoft.com/en-us/windows/win32/msi/database-tables.

    .PARAMETER MSIPath
    The path to the MSI file to modify.

    .PARAMETER MSITableName
    The MSI database table to modify.

    .PARAMETER PrimaryKeyName
    The name of the Primary Key for the table.

    .PARAMETER PrimaryKeyValue
    The value of the Primary Key for the row that will be modified.

    .PARAMETER ValueToChange
    The name of the column that will be changed.

    .PARAMETER NewValue
    The new value for the column that will be changed.

    .LINK
    https://github.com/zebulonsmith/NativePSMSI

    .EXAMPLE
    Change the DefaultDir column in the Directory table where Directory = INSTALLDIR
    Update-MSIDatabaseTable -MSIPath '.\7z2201-x64 - Modified.msi' -MSITableName "Directory" -PrimaryKeyName "Directory" -PrimaryKeyValue "INSTALLDIR" -ValueToChange "DefaultDir" -NewValue "7-Zoop"

    .EXAMPLE
    Same as above, but return the query without making changes.
    Update-MSIDatabaseTable -MSIPath '.\7z2201-x64 - Modified.msi' -MSITableName "Directory" -PrimaryKeyName "Directory" -PrimaryKeyValue "INSTALLDIR" -ValueToChange "DefaultDir" -NewValue "7-Zoop" -TestQueryOnly

    .EXAMPLE Change the Value column in the Property table where Property = TestProperty.
    Update-MSIDatabaseTable -MSIPath '.\7z2201-x64 - Modified.msi' -MSITableName Property -PrimaryKeyName Property -PrimaryKeyValue "TestProperty" -ValueToChange "Value" -NewValue "NewTestValue"
    #>
    param (
        [Parameter(Mandatory=$true)]
        [String]$MSIPath,

        [Parameter(Mandatory=$true)]
        [string]
        $MSITableName,

        [Parameter(Mandatory=$true)]
        $PrimaryKeyName,

        [parameter(Mandatory=$true)]
        [String]
        $PrimaryKeyValue,

        [parameter(Mandatory=$true)]
        [string]
        $ValueToChange,

        [parameter(Mandatory=$true)]
        [string]
        $NewValue,

        [Parameter(Mandatory=$false)]
        [Switch]
        $TestQueryOnly
    )

    #Build the query
    $UpdateQuery = "UPDATE $MSITableName SET $ValueToChange='$($NewValue)' WHERE $PrimaryKeyName='$PrimaryKeyValue'"

    if ($TestQueryOnly) {
        Return $UpdateQuery
    }

    #validate that the file exists
    if (!(test-path -path $MSIPath -PathType Leaf)) {
        Throw [System.IO.FileNotFoundException]::New("File $($MSIPath.FullName) not found.")
    }else {
        $MSIFile = Get-Item -Path $MSIPath
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
        $MSIDBObject = $WindowsInstaller.Gettype().InvokeMember("OpenDatabase", "InvokeMethod", $null, $WindowsInstaller, @($MSIFile.fullname, 1))
    } Catch {
        Write-Error "Failed to open MSI Database $($MSIFile.FullName)."
        Throw $_
    }

    #Create a view object to insert the data
    Write-Verbose "Executing query '$($UpdateQuery)' on the MSI database."

    Try {
        $UpdateView = $MSIDBObject.Gettype().InvokeMember("OpenView","InvokeMethod",$null,$MSIDBObject,($UpdateQuery))
    } Catch {
        Write-Error "Query $($UpdateQuery) is invalid."
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