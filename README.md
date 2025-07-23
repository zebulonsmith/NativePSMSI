# NativePSMSI
A few PowerShell functions that can read and edit MSI files as well as generate an MSI Transform (.mst).

## Why?
It's Monday night, I'm part of the way through a $6 bottle of wine from Trader Joe's and I've needed a way to read and manipulate MSI files using PowerShell without any external dependencies for a while.
There are other modules out there that will do a lot more with an MSI, but they require the WIX toolkit binaries to do it. This module has less functionality but also has no external dependencies.

It can read and edit an MSI Database's contents and also spit out a Transform (.mst) file.

## How does it work?
The module uses the WindowsInstaller.Installer com object to interface with MSI files. See https://learn.microsoft.com/en-us/windows/win32/msi/automation-interface-reference for Microsoft's documentation on the subject.

The goal was to be able to automate common software packaging tasks that are often done manually with tools like Orca, etc.

For example, you could write a script that uses Get-MSIDatabaseTable to retrieve the ProductCode from an MSI file and then create an Application with the Microsoft Configuration Manager module, using the retrieved ProductCode value to create the detection rule.
If necessary, you could create a transform to go with the original MSI that contains any necessary customizations.

## Notes
The individual functions in the 'Public' directory are written so that they can function independently of the rest of the module. This means that the .ps1 files themselves, or the code inside, can be copied and re-used elsewhere if needed without the module being present. I ask that if the code is re-used proper credit is attributed.

## Functions in the Module
Import the module and use Get-Help for examples and full documentation.

### Get-MSIDatabaseTable
Reads string data from the MSI Database. Can't read binary data.

### Update-MSIDatabaseTable
Changes the MSI Database.

### New-MSIDatabaseTableRow
Insert a new row into an MSI Database table.

### New-MSITransform
Generates a transform by comparing an edited copy of an MSI to the original.
