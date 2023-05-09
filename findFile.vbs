' ' Specify the directory to search in
' Dim searchDirectory
' searchDirectory = "D:\"

' ' Specify the target file name to search for
' Dim targetFileName
' targetFileName = "findme.txt"

' ' Call the recursive function to search for the file
' Dim filePath
' filePath = RecursiveFileSearch(searchDirectory, targetFileName)

' ' Display the result
' If Not IsEmpty(filePath) Then
'     MsgBox "File found at: " & filePath
' Else
'     MsgBox "File not found."
' End If

' ' Recursive file search function
' Function RecursiveFileSearch(directory, fileName)
'     ' Create a file system object
'     Dim fso
'     Set fso = CreateObject("Scripting.FileSystemObject")
    
'     ' Get the folder object for the current directory
'     Dim folder
'     Set folder = fso.GetFolder(directory)
    
'     ' Iterate through each file in the folder
'     Dim file
'     For Each file In folder.Files
'         ' Check if the file name matches the target
'         If LCase(file.Name) = LCase(fileName) Then
'             ' File found, return the file path
'             RecursiveFileSearch = file.Path
'             Exit Function
'         End If
'     Next
    
'     ' Recursively search subfolders
'     Dim subfolder
'     For Each subfolder In folder.Subfolders
'         ' Call the function recursively on each subfolder
'         RecursiveFileSearch = RecursiveFileSearch(subfolder.Path, fileName)
        
'         ' Check if the file was found in the subfolder
'         If Not IsEmpty(RecursiveFileSearch) Then
'             Exit Function
'         End If
'     Next
    
'     ' File not found
'     RecursiveFileSearch = Empty
' End Function


function recursiveTest(num)
    if(num = 1 or num = 0) then
        recursiveTest = 1
    else 
        recursiveTest = num * recursiveTest(num - 1)
    end if

end function

WScript.Echo recursiveTest(4)