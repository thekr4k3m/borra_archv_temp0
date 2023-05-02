On Error Resume Next

' Carpeta de archivos temporales de Windows
strTempFolder = "C:\Windows\Temp"

' Carpeta de archivos temporales de usuario
Set objShell = CreateObject("WScript.Shell")
strUserName = objShell.ExpandEnvironmentStrings("%USERNAME%")
strUserTempFolder = "C:\Users\" & strUserName & "\AppData\Local\Temp"

' Eliminar archivos temporales de la carpeta de Windows
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(strTempFolder)
For Each objFile In objFolder.Files
    objFSO.DeleteFile objFile.Path, True
Next

' Eliminar archivos temporales de la carpeta de usuario
Set objFolder = objFSO.GetFolder(strUserTempFolder)
For Each objFile In objFolder.Files
    objFSO.DeleteFile objFile.Path, True
Next

' Eliminar archivos temporales de Internet Explorer
Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.NameSpace(0x20).Self
Set objFolderItem = objFolder.ParseName("Temporary Internet Files")
Set objFolder = objFolderItem.GetFolder
For Each objFile In objFolder.Files
    objFSO.DeleteFile objFile.Path, True
Next

MsgBox "Se han eliminado los archivos temporales."
