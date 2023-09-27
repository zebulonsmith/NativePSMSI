# NativePSMSI
A few PowerShell functions that can read and edit MSI files as well as generate an MSI Transform (.mst).

## Why?
It's Monday night, I'm part of the way through a $6 bottle of wine from Trader Joe's and I've needed a way to read and manipulate MSI files using PowerShell without any external dependencies for a while.
There are other modules out there that will do a lot more with an MSI, but they require the WIX toolkit binaries to do it. This module has less functionality but also has no external dependencies. 

It can read and edit an MSI Database's contents and also spit out a Transform (.mst) file.

## How does it work?
The module uses the WindowsInstaller.Installer com object to interface with MSI files. See https://learn.microsoft.com/en-us/windows/win32/msi/automation-interface-reference for Microsoft's documentation on the subject.
Functions within the module are written without any interdependencies to each other so if you want to use them in your own script, they can be copied from this repo and pasted in directly if the module can't be loaded for some reason.

The goal was to be able to automate software packaging tasks that are often done manually with tools like Orca, etc.

For example, you could write a script that uses Get-MSIDatabaseTable to retrieve the ProductCode from an MSI file and then create an Application with the Microsoft Configuration Manager module, using the retrieved ProductCode value to create the detection rule.
If necessary, you could create a transform to go with the original MSI that contains any necessary customizations.

## Future Plans
~I'd like to add an Add-MSIDatabaseTableRow function at some point, but it's late and I'm out of wine.~
Add DELETE FROM, CREATE TABLE, DROP TABLE, etc. Maybe also a function that allows manual execution of a provided query. 

## Functions in the Module
Import the module and use Get-Help for examples and full documentation.

### Get-MSIDatabaseTable
Reads data from the MSI Database

### Update-MSIDatabaseTable
Changes the MSI Database

### New-MSITransform
Generates a transform by comparing an edited copy of an MSI to the original.
