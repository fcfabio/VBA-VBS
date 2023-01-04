Function checkFolder(path)
	Dim ObjFSO, objFolder
	Set ObjFSO = CreateObject("Scripting.FileSystemObject")
	if not (ObjFSO.FolderExists(path)) then
		Set objFolder = objFSO.CreateFolder(path)	
	End If
	Set ObjFSO = Nothing
End Function

Function createTXT(path)
	Dim ObjFSO
	Set ObjFSO = CreateObject("Scripting.FileSystemObject")
	If Not(ObjFSO.FileExists(path)) Then
		ObjFSO.CreateTextFile(path)
	End If
	Set ObjFSO = Nothing
End Function

Function CleanTXT(filePath)
	Dim ObjFSO
	Set ObjFSO = CreateObject("Scripting.FileSystemObject")
	ObjFSO.DeleteFile filePath
	createTXT(filePath)
	Set ObjFSO = Nothing
End Function

Function OpenFolder(folder_path) 
	Dim folder
	Set folder = CreateObject("WSCript.Shell")
	folder.Run folder_path
	Set folder = Nothing
End function


Function SearchFolder(file)
	Dim path, ObjFSO, searchSubFolder, searchSubFolder2, file, searchFolder1
	
	path = "C:\Temp"
	
	Set ObjFSO = CreateObject("Scripting.FileSystemObject")
	Set searchFolder1 = ObjFSO.GetFolder(path)
	
	for each searchSubFolder in searchFolder1.subfolders
		for each searchSubFolder2 in searchSubFolder.subfolders
			for each file in searchSubFolder2.Files
				if instr(file.name, file) then
					SearchFolder = file.path
				end if
			next
		Next
	Next
	
	Set ObjFSO = Nothing
	Set searchFolder1 = Nothing
	
End Function
