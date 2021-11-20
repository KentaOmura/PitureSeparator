Option Explicit

' ��`
Const PictureExtensionFileName = ".\..\extensionPicture.txt"
Const MovieExtensionFileName = ".\..\extensionMovie.txt"
Const ToMovePictureFolderName = ".\..\Picture"
Const ToMoveMovieFolderName = ".\..\Movie"

' ���ʎq
Const Picture = 0
Const Movie = 1

Dim isDelete : isDelete = True


'=======================================================
Class Progressbar
	Private total '����
	Private treated '�����ς�
	Private isLess '100�������ǂ���
	Private nProgress '���݂̐i��
	Private pProgress '100%�Œu������������1%������̐�

	Private Sub Class_Initialize
		pProgress = 1'0���Z���
		nProgress = 0
		total = 0
		treated = 0
		' �����J�n���͈�x�R�[������
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
			wscript.echo "�����̐ݒ肪�R��Ă��܂��B"
		end if

		treated = treated+1
		' �i���ɕω�����������o�[���X�V
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

		' �i���o�[�̕\��
		wscript.echo cstr(nProgress) & "/" & cstr(100)
		for count = 0 To nProgress
			 bar = bar & "��"
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
		WScript.Echo "���ʂ������t�H���_�̃p�X����͂��ĉ������B"
		Wscript.Quit
	end if

	Dim inPass
	Dim targetPictureExtension
	Dim targetMovieExtension

	inPass = Wscript.Arguments(0)

	Set targetPictureExtension = CreateObject("System.Collections.ArrayList")
	Set targetMovieExtension = CreateObject("System.Collections.ArrayList")

	' ���ޑΏۂ̊g���q���擾
	call GetExtension(Picture, targetPictureExtension)
	call GetExtension(Movie, targetMovieExtension)

	' �w�茳����Ώۊg���q�����݂���΁A���ޕ������s��
	' ����
	call CreateMoveToFolder(ToMoveMovieFolderName)
	call Calssfier(inPass, ToMoveMovieFolderName, targetMovieExtension)
	' �Î~��
	call CreateMoveToFolder(ToMovePictureFolderName)
	call Calssfier(inPass, ToMovePictureFolderName, targetPictureExtension)

	' ����ƐÎ~��̃t�@�C��������i�������Ƃ���
	Dim fso, total
	Set fso = CreateObject("Scripting.FileSystemObject")
	total = fso.GetFolder(ToMoveMovieFolderName).Files.Count
	total = total + fso.GetFolder(ToMovePictureFolderName).Files.Count
	call progressObj.setTotalProgress(total)

	'�ړ����Ƀt�@�C����������΍폜����
	call DeleteFromMoveFolder(inPass)

	' �B�e�N���̃t�H���_�Ɉړ�����
	call MoveShootingMonth(ToMovePictureFolderName, Picture)
	call MoveShootingMonth(ToMoveMovieFolderName, Movie)

	progressObj.finish

End Sub

Sub DeleteFromMoveFolder(byval inputPass)
	Dim fso
	set fso = CreateObject("Scripting.FileSystemObject")

	UpdateIsDeleteFromMoveFolder(inputPass)

	' �t�H���_���Ƀt�@�C����������΍폜����
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

' ���ނ킯�����{
Function Calssfier(ByVal inputPass, ByVal targetPass,ByRef targetExtension)
	Dim fso
	Dim index
	Set fso = CreateObject("Scripting.FileSystemObject")

	index = targetExtension.Count - 1
	while index >= 0
		' �ړ������{
		call ClassifierSub(fso.GetFolder(inputPass),targetPass,targetExtension(index))
		index = index - 1
	wend

	set fso = nothing
End Function

' ���ނ킯�i�T�u�j
Sub ClassifierSub(folder, ByVal targetPass,ByVal targetExtension)
	Dim file
	Dim Subfolder
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")

	For Each file in folder.Files
	' �Ώۂ̊g���q�ł���΁A�ړ�
		If LCase(fso.GetExtensionName(file.name))=targetExtension Or _
			UCase(fso.GetExtensionName(file.name))=targetExtension Then
			Call fso.MoveFile(folder.Path & "\" & file.name, targetPass& "\" & file.name)
		End If
	Next

	' �K�w�\���̃t�H���_�̏ꍇ�́A�ċA����
	For Each Subfolder in folder.SubFolders
		call ClassifierSub(Subfolder, targetPass,targetExtension)
	Next

	Set fso = Nothing
End Sub


' �g���q���擾
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

' �w��t�H���_���̃t�@�C�����B�e�������̃t�H���_�Ɉړ�����
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
		' �Î~��̎��͕ҏW����鎖�ŁA�X�V�������ς��̂ŁA�B�e�������擾����B
		if Picture = mode then
			Set objFolder = shell.NameSpace(fso.GetAbsolutePathName(inPath))
			Set objFolderItem = objFolder.ParseName(file.name)
			shootingTime = objFolder.GetDetailsOf(objFolderItem, 12)
			year  = Mid(shootingTime, 2, 4)
			month = Mid(shootingTime, 8, 2)
			day = Mid(shootingTime, 12,2)
			hour = Mid(shootingTime,17,2)
			minute = Mid(shootingTime, 20,2)
		' ����̎��͕ҏW����鎖�������O��ŁA�X�V�������B�e�����Ŏ擾����B
		else
			shootingTime = file.DateLastModified
			year = Left(shootingTime,4)
			month = Mid(shootingTime,6,2)
			day = Mid(shootingTime,9,2)
			hour = Mid(shootingTime,12,2)
			minute = Mid(shootingTime, 15,2)
		end if
		'�B�e�L�^�̔N�A���̃t�H���_�Ɉړ�����
		targetPass = inPath & "\" & year & "_" & month
		originalFile = file.Name
		CreateMoveToFolder(targetPass)
		extension = fso.GetExtensionName(file.name)
		targetFile = year&"_"&month&"_"&day&"_"&hour&"_"&minute
		targetFile = Replace(targetFile, ":", "")

		'����A�ʐ^�B�e�ɂ����āA1���Ԃɉ��x���B�e����Ă���ꍇ�͓����̃t�@�C�����ƂȂ�̂ŁA���l�[������B
		if True = fso.FileExists(targetPass& "\" & targetFile&"."&extension) Then
			while True = fso.FileExists(targetPass& "\" & targetFile&"_"&count&"."&extension)
				count = count + 1
			wend
			targetFile = targetFile &"_"&count
		end if
		Call fso.MoveFile(inPath & "\" & originalFile, targetPass& "\" & targetFile&"."&extension)
		' �i���o�[�̍X�V
		progressObj.notify
	Next

	Set objFolderItem = Nothing
	Set objFolder = Nothing
	Set fso = Nothing
	Set shell = Nothing
End Sub