'
' filename : fileNameOutput.vbs
'
Dim objFSO, objFolder, objFile, strPath, outputFile
Set objFSO = CreateObject("Scripting.FileSystemObject")

' 現在のディレクトリを取得
strPath = objFSO.GetAbsolutePathName(".")

' アウトプットするテキストファイルを指定 (ここでは "filelist.txt" としていますが、必要に応じて変更してください)
outputFile = strPath & "\filelist.txt"

' ファイルを作成/上書きするためのテキストストリームオブジェクトを開く
Set objTextFile = objFSO.CreateTextFile(outputFile, True)

' 現在のディレクトリのファイルをリストアップ
Set objFolder = objFSO.GetFolder(strPath)
For Each objFile In objFolder.Files
    objTextFile.WriteLine(objFile.Name)
Next

' オブジェクトを閉じてリリース
objTextFile.Close
Set objTextFile = Nothing
Set objFolder = Nothing
Set objFSO = Nothing

WScript.Echo "File list has been saved to " & outputFile
