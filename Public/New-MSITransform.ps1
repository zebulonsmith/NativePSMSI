Function New-MSITransform {

    <#
    .DESCRIPTION
    Compares two MSI files and outputs an MSI Transform (.mst) file.

    To create a transform, create a copy of the MSI file and make the desired changes using Update-MSIDatabaseTable (or another tool) then use
    this cmdlet to compare it to the original MSI and output an MST.

    .PARAMETER ReferenceMSIPath
    The path to the original, unmodified MSI file.

    .PARAMETER DifferenceMSIPath
    The path to the copy of the MSI that has been modified with the changes that should be added to the MST.

    .PARAMETER OutputMSTPath
    Optional output path for the MST file. By default, the MST will be created in the same directory as the original MSI.

    .LINK
    https://github.com/zebulonsmith/NativePSMSI

    .EXAMPLE
    Create a transfrom based on the differences between '7z2201-x64.msi' and '7z2201-x64 - Modified.msi'
    New-MSITransform -ReferenceMSIPath .\7z2201-x64.msi -DifferenceMSIPath '.\7z2201-x64 - Modified.msi'

    .EXAMPLE
    Same as above, but the output file name is specified.
    New-MSITransform -ReferenceMSIPath .\7z2201-x64.msi -DifferenceMSIPath '.\7z2201-x64 - Modified.msi' -OutputMSTPath .\7zTransform.mst
    #>

    param(
        [Parameter(Mandatory=$true)]
        [String]$ReferenceMSIPath,

        [Parameter(Mandatory=$true)]
        [String]$DifferenceMSIPath,

        [parameter(Mandatory=$false)]
        [ValidatePattern('.*\.mst$')]
        [String]$OutputMSTPath
    )

        #validate that the files exist
        Write-Verbose "Validating file paths"
        if (!(test-path -path $ReferenceMSIPath -PathType Leaf)) {
            Throw [System.IO.FileNotFoundException]::New("File $($ReferenceMSIPath) not found.")
        }

        if (!(test-path -path $DifferenceMSIPath -PathType Leaf)) {
            Throw [System.IO.FileNotFoundException]::New("File $($DifferenceMSIPath) not found.")
        }

        #Figure out the output path
        if (!$OutputMSTPath) {
            [System.IO.FileInfo]$OutputMSTPathObject = "$([System.IO.Path]::GetFileNameWithoutExtension($ReferenceMSIPath)).mst"
        } else {
            [System.IO.FileInfo]$OutputMSTPathObject = $OutputMSTPath
        }

        #Make sure that the destination directory exists
        if (!(Test-path -path $OutputMSTPathObject.Directory)) {
            throw [System.IO.DirectoryNotFoundException]::New("Output directory $($OutputMSTPathObject.Directory) does not exist.")
        }

        #Make sure that the output mst doesn't exist already
        if (Test-path -path $OutputMSTPathObject.FullName -PathType Leaf) {
            throw [System.IO.IOException]::new("Output MST file $($OutputMSTPathObject.fullname) already exists.")
        }

        #Load the WindowsInstaller
        Write-Verbose "Loading Windows Installer object"
        Try {
            $WindowsInstaller = New-Object -ComObject WindowsInstaller.Installer
        } Catch {
            Write-Error "Failed to create WindowsInstaller.Installer com object."
            Throw $_
        }

        #open the diff MSI database read only
        Write-Verbose "Opening diff MSI"
        Try {
            $DiffMSIDBObject = $WindowsInstaller.Gettype().InvokeMember("OpenDatabase", "InvokeMethod", $null, $WindowsInstaller, @($DifferenceMSIPath, 0))
        } Catch {
            Write-Error "Failed to open MSI Database $($DifferenceMSIPath.fullname)."
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($WindowsInstaller) | Out-null
            Throw $_
        }

        #Open the reference MSI database read only
        Write-Verbose "Opening Reference MSI"
        Try {
            $ReferenceMSIDBObject = $WindowsInstaller.Gettype().InvokeMember("OpenDatabase", "InvokeMethod", $null, $WindowsInstaller, @($ReferenceMSIPath, 0))
        } Catch {
            Write-Error "Failed to open MSI Database $($ReferenceMSIPath.fullname)."
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($DiffMSIDBObject) | Out-Null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ReferenceMSIDBObject) | Out-Null
            Throw $_
        }


        #Generate a transform using the diff of the two MSIs
        Write-verbose "Generating transform"

        Try {
            $Transform = $DiffMSIDBObject.Gettype().InvokeMember("GenerateTransform","InvokeMethod",$null,$DiffMSIDBObject,@($ReferenceMSIDBObject,$OutputMSTPathObject.FullName))
        } Catch {
            Write-Error "Could not create MST file.`n$($Transform)"
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($DiffMSIDBObject) | Out-Null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ReferenceMSIDBObject) | Out-Null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($WindowsInstaller) | Out-null
            Throw $_
        }

        #Add the MSI summary data to the MSI
        Try {
            $Summary = $DiffMSIDBObject.Gettype().InvokeMember("CreateTransformSummaryInfo","InvokeMethod",$null,$DiffMSIDBObject,@($ReferenceMSIDBObject,$OutputMSTPathObject.FullName,0,0))
        } Catch {
            Write-Error "Failed to add summary info to $($OutputMSTPath.FullName)`n$($Summary)"
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($DiffMSIDBObject) | Out-Null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ReferenceMSIDBObject) | Out-Null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($WindowsInstaller) | Out-null
            Throw $_
        }


        #Close the databases
        Write-Verbose "Closing resources"
        $DiffMSIDBObject.GetType().InvokeMember("Commit", "InvokeMethod", $null, $DiffMSIDBObject, $null)
        $ReferenceMSIDBObject.GetType().InvokeMember("Commit", "InvokeMethod", $null, $ReferenceMSIDBObject, $null)

        #Release files and objects
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($DiffMSIDBObject) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ReferenceMSIDBObject) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($WindowsInstaller) | Out-null

        if ($transform -eq $false) {
            Throw "Difference MSI didn't contain any changes."
        }

}