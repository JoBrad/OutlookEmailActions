Option Explicit
Option Base 0
Option Compare Text

'******************************************************************************
'
'	md_CategoryBackup
'	Backup or restore all of your categories, with their colors
'
'******************************************************************************

'---------------------------------------------------------------------------------------
' Procedure :	BackupCategories
' Author	:	jobrad
' Date		:	2015-01-21
' Purpose	:	Backs up all categories, with their color information, to a file
'				Suggested file layout: CSV, categoryName, CategoryColor
'---------------------------------------------------------------------------------------
'
Public Sub BackupCategories()
	Dim strBackupPath As String, appNamespace As NameSpace, strFileName As String, intFileNum As Integer, catThisCategory As Category, strThisBackupRecord As String

	strBackupPath = JoBrad.FolderPicker

	If strBackupPath <> "-1" Then
		Set appNamespace = Application.GetNamespace("MAPI")
		If appNamespace.Categories.Count > 0 Then
			strFileName = strBackupPath & "\" & Format(Now(), "yyyymmdd_hhmmss") & "_outlook_category_backup.csv"

			intFileNum = FreeFile
			Open strFileName For Output As #intFileNum

			For Each catThisCategory In appNamespace.Categories
				strThisBackupRecord = catThisCategory.Name & ", " & catThisCategory.Color
				Print #intFileNum, strThisBackupRecord
			Next catThisCategory

			Close #intFileNum
		Else
			MsgBox Prompt:="There are no categories to export.", Buttons:=vbOKOnly + vbInformation, Title:="No Backup Created"
		End If

	End If

End Sub

'---------------------------------------------------------------------------------------
' Procedure :	RestoreCategories
' Author	:	jobrad
' Date		:	2015-01-28
' Purpose	:	Reads a list of categories from a backup file, and adds them to Outlook
'---------------------------------------------------------------------------------------
'
Public Sub RestoreCategories()
	On Error GoTo Err_Handler
	Dim colBackupFiles As Collection, vrtThisFileName As Variant, intFileNum As Integer, strThisRecord As String, aryThisRecord() As String, strThisCategoryName As String, intThisCategoryColor As Integer, dicCategories As Dictionary, vrtThisCategory As Variant

	Set colBackupFiles = JoBrad.FilePicker

	Set dicCategories = New Dictionary

	For Each vrtThisFileName In colBackupFiles
		intFileNum = FreeFile()
		Open CStr(vrtThisFileName) For Input As #intFileNum
			Do While Not EOF(intFileNum)
				Line Input #intFileNum, strThisRecord
				aryThisRecord = Split(strThisRecord, ",")
				strThisCategoryName = Trim(aryThisRecord(0))
				intThisCategoryColor = CInt(Trim(aryThisRecord(1)))
				dicCategories.Add Key:=strThisCategoryName, item:=intThisCategoryColor
			Loop
		Close #intFileNum
	Next vrtThisFileName
	
	For Each vrtThisCategory In dicCategories.Keys
		strThisCategoryName = CStr(vrtThisCategory)
		intThisCategoryColor = dicCategories(strThisCategoryName)
		AddCategory strCategoryName:=strThisCategoryName, intColor:=intThisCategoryColor
	Next vrtThisCategory

	GoTo The_End

Err_Handler:
	If err.Number = 457 Then
		Resume Next
	Else
		MsgBox Prompt:=err.Description & " (" & err.Number & "). The restore process will be cancelled.", Buttons:=vbOKOnly + vbCritical, Title:="Cannot Continue Restore Process"
		Close
		Exit Sub
	End If
The_End:

End Sub

'******************************************************************************
'	Supporting Functions
'******************************************************************************
'---------------------------------------------------------------------------------------
' Procedure :	AddCategory
' Author	:	jobrad
' Date		:	2010-03-04
' Purpose	:	Adds a category to Outlook
'---------------------------------------------------------------------------------------
'
Private Sub AddCategory(strCategoryName As String, intColor As Integer)
	On Error GoTo Err_Handler
	Dim objNS As NameSpace

	Set objNS = Application.GetNamespace("MAPI")

	objNS.Categories.Add strCategoryName, intColor
	'Sleep 500
	Set objNS = Nothing

	GoTo The_End

Err_Handler:
	Select Case err.Number
		Case -2147024809
			' The category already exists
			Resume Next
		Case Else
			MsgBox Prompt:=err.Description & " (" & err.Number & ").", Buttons:=vbCritical + vbOKOnly, Title:="Could Not Add Category!"
	End Select

The_End:
End Sub

'---------------------------------------------------------------------------------------
' Procedure :	FolderPicker
' Author	:	jobrad
' Date		:	2015-01-28
' Purpose	:	Uses MS Word to show a FolderPicker. Returns the path, or a -1 if cancelled
'
' Parameters
'	strDefaultDirectory	:	The directory to start looking in
'---------------------------------------------------------------------------------------
'
Private Function FolderPicker(Optional strDefaultDirectory As String = "") As String
	Dim objWordApp As Word.Application, dlgFileDialog As Office.FileDialog, strChosenPath As String

	Set objWordApp = New Word.Application
	Set dlgFileDialog = objWordApp.Application.FileDialog(msoFileDialogFolderPicker)
	strChosenPath = -1

	With dlgFileDialog
		.AllowMultiSelect = False
		.InitialView = msoFileDialogViewLargeIcons

		If strDefaultDirectory <> "" Then
			.InitialFileName = strDefaultDirectory
		Else
			.InitialFileName = Environ$("USERPROFILE") & "\Documents"
		End If

		.Title = "Choose a Directory"

		If .Show = -1 Then
			strChosenPath = .selectedItems(1)
		Else
			strChosenPath = -1
		End If
	End With

	objWordApp.Quit

	Set objWordApp = Nothing

	FolderPicker = strChosenPath

End Function

'---------------------------------------------------------------------------------------
' Procedure :	FilePicker
' Author	:	jobrad
' Date		:	2015-01-28
' Purpose	:	Uses MS Word to show a FileDialog. Returns a collection of file names,
'				or a -1 if cancelled
' Parameters
'	strDefaultDirectory	:	The directory to start looking in
'
' Todo		:	* Add a fileTypes option
'					fileTypes:	A list of friendly names and masks for allowed
'								file types. The name and mask should be comma-separated,
'								and each name-mask pair should be separated by semicolons.
'								Example: Images,*.png; Videos,*.mov
'				* Allow multiple masks for a single mask name, using pipe-delimited values
'---------------------------------------------------------------------------------------
'
Private Function FilePicker(Optional strDefaultDirectory As String = "") As Collection
	Dim objWordApp As Word.Application, dlgFileDialog As Office.FileDialog, vrtSelectedItems As Variant, colSelectedItems As Collection

	Set objWordApp = New Word.Application
	Set dlgFileDialog = objWordApp.Application.FileDialog(msoFileDialogFilePicker)
	Set colSelectedItems = New Collection

	With dlgFileDialog
		.AllowMultiSelect = False
		.InitialView = msoFileDialogViewLargeIcons

		If strDefaultDirectory <> "" Then
			.InitialFileName = strDefaultDirectory
		Else
			.InitialFileName = Environ$("USERPROFILE") & "\Documents"
		End If

		.Title = "Choose a File"

		If .Show = -1 Then
			For Each vrtSelectedItems In .selectedItems
				colSelectedItems.Add item:=vrtSelectedItems
			Next vrtSelectedItems
		Else
			colSelectedItems.Add item:=-1
		End If
	End With

	objWordApp.Quit
	Set objWordApp = Nothing

	Set FilePicker = colSelectedItems

End Function
