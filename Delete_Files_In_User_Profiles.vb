'This script running on my terminal servers srv01 to srv03 to remove settings file for a application, And last on the file server where roaming profile is stored.

Dim fso
Dim objFSO
Dim objFolder
Dim objSubfolders
Dim objSubfolder
Dim strFolder
Dim strDocuments
Set objFSO = CreateObject("Scripting.FileSystemObject")

'File that will be removed in user Profile(Path has user profile folder as root)
strDocuments = "\system\Servers.pcf"


strFolder = "\\srv01\c$\Users"
Set objFolder = objFSO.GetFolder(strFolder)
Set objSubfolders = objFolder.Subfolders
For Each objSubfolder In objSubfolders
	if objFSO.FileExists(strFolder & "\" & objSubfolder.Name & strDocuments) then
		objFSO.DeleteFile strFolder & "\" & objSubfolder.Name & strDocuments, True
	end if
Next

strFolder = "\\srv03\c$\Users"
Set objFolder = objFSO.GetFolder(strFolder)
Set objSubfolders = objFolder.Subfolders
For Each objSubfolder In objSubfolders
	if objFSO.FileExists(strFolder & "\" & objSubfolder.Name & strDocuments) then
		objFSO.DeleteFile strFolder & "\" & objSubfolder.Name & strDocuments, True
	end if
Next

strFolder = "\\srv03\c$\Users"
Set objFolder = objFSO.GetFolder(strFolder)
Set objSubfolders = objFolder.Subfolders
For Each objSubfolder In objSubfolders
	if objFSO.FileExists(strFolder & "\" & objSubfolder.Name & strDocuments) then
		objFSO.DeleteFile strFolder & "\" & objSubfolder.Name & strDocuments, True
	end if
Next



strFolder = "\\FileSRV01\e$\profiles"
Set objFolder = objFSO.GetFolder(strFolder)
Set objSubfolders = objFolder.Subfolders
For Each objSubfolder In objSubfolders
	if objFSO.FileExists(strFolder & "\" & objSubfolder.Name & strDocuments) then
		objFSO.DeleteFile strFolder & "\" & objSubfolder.Name & strDocuments, True
	end if
Next

WScript.Echo "Finish"