# borra_archv_temp0
VBScript (.vbs) que te permite borrar los archivos temporales de tu PC:

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



Este script elimina los archivos temporales de Windows, los archivos temporales 
de usuario y los archivos temporales de Internet Explorer. Ten en cuenta que, 
una vez ejecutado, el proceso de eliminación puede tardar un tiempo en función
de la cantidad de archivos temporales que haya en el sistema.

Para ejecutar este script, guarda el código en un archivo con 
extensión ".vbs" siguiendo los mismos pasos que te mencioné anteriormente y 
luego haz doble clic en el archivo para ejecutarlo. Asegúrate de ejecutar 
el script con permisos de administrador para que pueda tener acceso 
a las carpetas y archivos que necesita eliminar.
