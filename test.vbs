Set fso = CreateObject("Scripting.FileSystemObject")

destination = "C:\"
zipFilename = "C:\testZip.zip"

WScript.Echo "destination:", destination
WScript.Echo "zipFilename:", zipFilename

If NOT fso.FileExists(zipFilename) Then
 WScript.Echo "zipFilename doesn't exist!"
End If

WScript.Echo "Creating Shell.Application Object"
set objShell=CreateObject("Shell.Application")

WScript.Echo "Creating objShell.NameSpace of testZip.zip"
set zipFiles=objShell.NameSpace(zipFilename).items

WScript.Echo "Outputting all items in testZip.zip"
For Each Item in zipFiles
 WScript.Echo Item.Name
Next

WScript.Echo "Copying files to destination"
objShell.NameSpace(destination).CopyHere(zipFilename)

Set fso = Nothing
Set objShell = Nothing
