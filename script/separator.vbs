Option Explicit

' 定義
Const PictureExtensionFileName = ".\..\extensionPicture.txt"
Const MovieExtensionFileName = ".\..\extensionMovie.txt"
Const ToMovePictureFolderName = ".\..\Picture"
Const ToMoveMovieFolderName = ".\..\Movie"

' 識別子
Const Picture = 0
Const Movie = 1

Dim isDelete : isDelete = True


'=======================================================
Class Progressbar
	Private total '総数
	Private treated '処理済み
	Private isLess '100未満かどうか
	Private nProgress '現在の進捗
	Private pProgress '100%で置き換えた時の1%当たりの数

	Private Sub Class_Initialize
		pProgress = 1'0除算回避
		nProgress = 0
		total = 0
		treated = 0
		' 処理開始時は一度コールする
		update
	End Sub

	Public Sub setTotalProgress(Byval num)
		total = num
		if total < 100 then
			pProgress = total
			isLess = True
		else
			pProgress = total \ 100
			isLess = False
		end if
	End Sub

	Public Sub notify()
		if 0 = total then
			wscript.echo "総数の設定が漏れています。"
		end if

		treated = treated+1
		' 進捗に変化があったらバーを更新
		Dim bef : bef = nProgress
		if True = isLess Then
			nProgress = nProgress + 100 \ pProgress
		else
			if treated mod pProgress = 0 then
				nProgress = nProgress + 1
			end if
		end if
		if nProgress <> bef then
			update
		end if
	End Sub

	Public Sub finish()
		nProgress = 100
		update
	End Sub

	Private Sub update()
		Dim count,bar
		bar = ""

		' 進捗バーの表示
		wscript.echo cstr(nProgress) & "/" & cstr(100)
		for count = 0 To nProgress
			 bar = bar & "■"
		Next
		'wscript.echo bar
	End Sub

	Private Sub clear()
		Dim objShell
		Set objShell = CreateObject("WScript.Shell")
		objShell.Run "cmd cls", 0, false
		set objShell = nothing
	End Sub
	
End Class
'=======================================================

Dim progressObj

set progressObj = New Progressbar

main

'  Main
Sub main()

	if 0 = Wscript.Arguments.Count then
		WScript.Echo "識別したいフォルダのパスを入力して下さい。"
		Wscript.Quit
	end if

	Dim inPass
	Dim targetPictureExtension
	Dim targetMovieExtension

	inPass = Wscript.Arguments(0)

	Set targetPictureExtension = CreateObject("System.Collections.ArrayList")
	Set targetMovieExtension = CreateObject("System.Collections.ArrayList")

	' 分類対象の拡張子を取得
	call GetExtension(Picture, targetPictureExtension)
	call GetExtension(Movie, targetMovieExtension)

	' 指定元から対象拡張子が存在すれば、分類分けを行う
	' 動画
	call CreateMoveToFolder(ToMoveMovieFolderName)
	call Calssfier(inPass, ToMoveMovieFolderName, targetMovieExtension)
	' 静止画
	call CreateMoveToFolder(ToMovePictureFolderName)
	call Calssfier(inPass, ToMovePictureFolderName, targetPictureExtension)

	' 動画と静止画のファイル総数を進捗総数とする
	Dim fso, total
	Set fso = CreateObject("Scripting.FileSystemObject")
	total = fso.GetFolder(ToMoveMovieFolderName).Files.Count
	total = total + fso.GetFolder(ToMovePictureFolderName).Files.Count
	call progressObj.setTotalProgress(total)

	'移動元にファイルが無ければ削除する
	call DeleteFromMoveFolder(inPass)

	' 撮影年月のフォルダに移動する
	call MoveShootingMonth(ToMovePictureFolderName, Picture)
	call MoveShootingMonth(ToMoveMovieFolderName, Movie)

	progressObj.finish

End Sub

Sub DeleteFromMoveFolder(byval inputPass)
	Dim fso
	set fso = CreateObject("Scripting.FileSystemObject")

	UpdateIsDeleteFromMoveFolder(inputPass)

	' フォルダ内にファイルが無ければ削除する
	if true = isDelete then
		fso.DeleteFolder inputPass
	end if

	set fso = nothing

End Sub


Sub UpdateIsDeleteFromMoveFolder(byval inputPass)
	Dim fso
	Dim folder, Subfolder
	set fso = CreateObject("Scripting.FileSystemObject")
	set folder = fso.GetFolder(inputPass)

	if  0 <> folder.Files.Count then
		isDelete = False
	End if

	For Each Subfolder in folder.SubFolders
		call DeleteFromMoveFolder(Subfolder.Path)
	Next

	set fso = nothing

End Sub

Sub CreateMoveToFolder(byval folderName)
	Dim fso
	set fso = CreateObject("Scripting.FileSystemObject")

	if  false = fso.FolderExists( folderName ) then
		fso.CreateFolder(folderName)
	End if

	set fso = nothing
End Sub

' 分類わけを実施
Function Calssfier(ByVal inputPass, ByVal targetPass,ByRef targetExtension)
	Dim fso
	Dim index
	Set fso = CreateObject("Scripting.FileSystemObject")

	index = targetExtension.Count - 1
	while index >= 0
		' 移動を実施
		call ClassifierSub(fso.GetFolder(inputPass),targetPass,targetExtension(index))
		index = index - 1
	wend

	set fso = nothing
End Function

' 分類わけ（サブ）
Sub ClassifierSub(folder, ByVal targetPass,ByVal targetExtension)
	Dim file
	Dim Subfolder
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")

	For Each file in folder.Files
	' 対象の拡張子であれば、移動
		If LCase(fso.GetExtensionName(file.name))=targetExtension Or _
			UCase(fso.GetExtensionName(file.name))=targetExtension Then
			Call fso.MoveFile(folder.Path & "\" & file.name, targetPass& "\" & file.name)
		End If
	Next

	' 階層構造のフォルダの場合は、再帰する
	For Each Subfolder in folder.SubFolders
		call ClassifierSub(Subfolder, targetPass,targetExtension)
	Next

	Set fso = Nothing
End Sub


' 拡張子を取得
Sub GetExtension(ByVal kind, Byref extensionArray)
	Dim objFileSys
	Dim objReadStream

	Set objFileSys = CreateObject("Scripting.FileSystemObject")

	if Picture = kind then
		Set objReadStream  = objFileSys.OpenTextFile(PictureExtensionFileName, 1)
	else
		Set objReadStream  = objFileSys.OpenTextFile(MovieExtensionFileName, 1)
	End if

	while objReadStream.AtEndOfLine = false
		extensionArray.add objReadStream.ReadLine
	wend

End Sub

' 指定フォルダ内のファイルを撮影日時月のフォルダに移動する
Sub MoveShootingMonth(byval inPath, byval mode)
	Dim fso, folder, file
	Dim shell, objFolder, objFolderItem
	Dim targetPass
	Dim shootingTime
	Dim year, month, day, hour, minute
	Dim originalFile, targetFile
	Dim count, extension

	Set fso = CreateObject("Scripting.FileSystemObject")
	Set shell = CreateObject("shell.application")

	Set folder = fso.GetFolder(inPath)
	For Each file in folder.Files
		count = 1
		' 静止画の時は編集される事で、更新日時が変わるので、撮影日時を取得する。
		if Picture = mode then
			Set objFolder = shell.NameSpace(fso.GetAbsolutePathName(inPath))
			Set objFolderItem = objFolder.ParseName(file.name)
			shootingTime = objFolder.GetDetailsOf(objFolderItem, 12)
			year  = Mid(shootingTime, 2, 4)
			month = Mid(shootingTime, 8, 2)
			day = Mid(shootingTime, 12,2)
			hour = Mid(shootingTime,17,2)
			minute = Mid(shootingTime, 20,2)
		' 動画の時は編集される事が無い前提で、更新日時＝撮影日時で取得する。
		else
			shootingTime = file.DateLastModified
			year = Left(shootingTime,4)
			month = Mid(shootingTime,6,2)
			day = Mid(shootingTime,9,2)
			hour = Mid(shootingTime,12,2)
			minute = Mid(shootingTime, 15,2)
		end if
		'撮影記録の年、月のフォルダに移動する
		targetPass = inPath & "\" & year & "_" & month
		originalFile = file.Name
		CreateMoveToFolder(targetPass)
		extension = fso.GetExtensionName(file.name)
		targetFile = year&"_"&month&"_"&day&"_"&hour&"_"&minute
		targetFile = Replace(targetFile, ":", "")

		'動画、写真撮影において、1分間に何度も撮影されている場合は同名のファイル名となるので、リネームする。
		if True = fso.FileExists(targetPass& "\" & targetFile&"."&extension) Then
			while True = fso.FileExists(targetPass& "\" & targetFile&"_"&count&"."&extension)
				count = count + 1
			wend
			targetFile = targetFile &"_"&count
		end if
		Call fso.MoveFile(inPath & "\" & originalFile, targetPass& "\" & targetFile&"."&extension)
		' 進捗バーの更新
		progressObj.notify
	Next

	Set objFolderItem = Nothing
	Set objFolder = Nothing
	Set fso = Nothing
	Set shell = Nothing
End Sub