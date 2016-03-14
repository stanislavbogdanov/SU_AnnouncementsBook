VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ЭтаКнига"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
  Dim BackupDir As String: BackupDir = ThisWorkbook.Path & "\backup"
  Dim BackupName As String
  Dim FileExt As String
  FileExt = Mid$(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, "."))
  BackupName = Left$(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1) & _
               " " & Format(Date, "yyyy-mm-dd") & " " & Format(Time, "hhmmss") & _
               FileExt
  Dim fs As FileSystemObject
  
  On Error GoTo HandleErrors
  
  Set fs = New FileSystemObject
  Dim ThisDoc As String
  ThisDoc = fs.BuildPath(ThisWorkbook.Path, ThisWorkbook.Name)
  Dim BackupDoc As String
  BackupDoc = fs.BuildPath(BackupDir, BackupName)
  If Not fs.FolderExists(BackupDir) Then
    Call fs.CreateFolder(BackupDir)
  End If
  Call fs.CopyFile(ThisDoc, BackupDoc)
  
ExitHere:

  Set fs = Nothing
  Exit Sub
  
HandleErrors:
  
  MsgBox "Невозможно создать резервную копию." & vbNewLine & _
         "Рекомендуется сохранить документ под другим именем." & vbNewLine & _
         "Ошибка: " & Err.Description & " (" & Err.Number & " )", _
         vbExclamation
  SaveAsUI = True
  Resume ExitHere
End Sub

Private Sub Workbook_Open()
End Sub
