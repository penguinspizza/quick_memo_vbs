' ファイル名を入力するダイアログを表示
Set objShell = CreateObject("WScript.Shell")
strFileName = InputBox("Please input the file name:", "Input file name")

' キャンセルが押されたらプログラムを終了
If IsEmpty(strFileName) Then
    WScript.Quit
End If

' ファイル名に現在日時を結合
datetimeNow = Now()
now_YYYYMMDDhhmmss= Year(datetimeNow)
now_YYYYMMDDhhmmss= now_YYYYMMDDhhmmss & Right("0" & Month(datetimeNow) , 2)
now_YYYYMMDDhhmmss= now_YYYYMMDDhhmmss & Right("0" & Day(datetimeNow) , 2)
now_YYYYMMDDhhmmss= now_YYYYMMDDhhmmss & Right("0" & Hour(datetimeNow) , 2)
now_YYYYMMDDhhmmss= now_YYYYMMDDhhmmss & Right("0" & Minute(datetimeNow) , 2)
now_YYYYMMDDhhmmss= now_YYYYMMDDhhmmss & Right("0" & Second(datetimeNow) , 2)

If strFileName = "" Then
    strFileName = now_YYYYMMDDhhmmss
Else
    strFileName = strFileName & "_" & now_YYYYMMDDhhmmss
End If

' 保存先ディレクトリを指定する
strFileName = "./created/" & strFileName

' テキストファイルを作成して書き込む
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile(strFileName & ".txt", True)
objFile.WriteLine "Created " & datetimeNow
objFile.Close

' メモ帳を開いてファイルを表示
Set objShell = CreateObject("WScript.Shell")
objShell.Run "notepad.exe " & strFileName
