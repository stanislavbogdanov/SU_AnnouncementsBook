VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExportChapterForm 
   Caption         =   "Экспорт"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7680
   OleObjectBlob   =   "ExportChapterForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExportChapterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" _
(ByVal hInstance As Long, ByVal lpCursorName As Long) As Long

Private Declare Function SetCursor Lib "user32" _
(ByVal hCursor As Long) As Long

Const IDC_HAND = 32649&

Const CHAPTER1_NAME = "Основные"
Const CHAPTER2_NAME = "Дополнительные"
Const CHAPTER3_NAME = "Аварийные"

Const SBOGDANOV_HYPERLINK = "http://ya.ru"

'Символ-заполнитель пустых ячеек для экспорта в XML
Const FILL_CHAR = "@"

'Номер столбца для первого языка
Dim nFirstLangCol As Long

'Номер столбца для последнего языка
Dim nLastLangCol As Long

'Флаг остановки
Dim bStopFlag As Boolean

'Построить имя файла (без расширения) в соответствии с выбранными опциями
Private Function BuildFileName() As String
  Dim sLangSet As String
  sLangSet = ""
  Dim langRec As ListRow
  For Each langRec In ThisWorkbook.Worksheets("TxtStyles").ListObjects("colorstab").ListRows
    If Me.Controls.Item("L" & langRec.Range(1, 2)).Value = True Then
      sLangSet = sLangSet & langRec.Range(1, 3)
    End If
  Next
  BuildFileName = GetChapterName() & "_" & sLangSet
End Function

'Заменить пустую строку на символ-заполнитель
'(требуется из-за особенностей верстки в Adobe InDesign)
Private Function PrepareEmptyCell(text As String) As String
  If text = "" Then
    PrepareEmptyCell = FILL_CHAR
  Else
    PrepareEmptyCell = text
  End If
End Function

'Вернуть имя выбранной главы
Private Function GetChapterName() As String
  If Me.ChapterButton1.Value = True Then GetChapterName = CHAPTER1_NAME
  If Me.ChapterButton2.Value = True Then GetChapterName = CHAPTER2_NAME
  If Me.ChapterButton3.Value = True Then GetChapterName = CHAPTER3_NAME
End Function

'Сохранить выбранное содержимое как XML
Private Sub BuildXMLDoc(sFullFileName As Variant)
  Dim Chapter As ListObject
  Dim sCText As String
  Dim oRec As ListRow
  Dim n As Long
  Dim DelSeparatorFlag As Boolean
  Dim bEmptyLine As Boolean
  
  Dim xmlDoc As DOMDocument
  Dim xmlFields As IXMLDOMElement
  Dim xmlField As IXMLDOMElement
  Dim Handbook As IXMLDOMElement
  Dim Announcement As IXMLDOMElement
  Dim xmlSeparator As IXMLDOMElement
  Dim tempElm As IXMLDOMElement
  
  'Создать новый DOMDocument
  Set xmlDoc = CreateObject("Microsoft.XMLDOM")
  xmlDoc.LoadXML ("<Root/>")
  Set Handbook = xmlDoc.DocumentElement.appendChild(xmlDoc.createElement("Handbook"))
  Set Chapter = ThisWorkbook.Worksheets(GetChapterName()).ListObjects(1)
  
  Dim nFirstLangCol As Long
  nFirstLangCol = Chapter.ListColumns("LstNum").Index + 1
  Dim nLastLangCol As Long
  nLastLangCol = Chapter.ListColumns("Separator").Index - 1
  Dim sLangName As String
  
  For Each oRec In Chapter.ListRows
    Set Announcement = xmlDoc.createElement("Announcement")
    Set xmlFields = Handbook.appendChild(Announcement)
    bEmptyLine = True
    With oRec.Range
      ' Заголовочные записи
      If (.Cells(2) = 1) Or (.Cells(2) = 2) Or (.Cells(2) = 3) Or (.Cells(2) = 4) Or (.Cells(2) = 5) Or (.Cells(2) = 7) Or (.Cells(2) = 8) Then
        Set xmlField = xmlFields.appendChild(xmlDoc.createElement("TxtStyle" & .Cells(2)))
        sCText = Trim(.Cells(7))
        If sCText <> "" Then
          xmlField.text = sCText
          bEmptyLine = False
        End If
      ' Обычное сообщение или список
      ElseIf (.Cells(2) = 0) Or (.Cells(2) = 6) Then
        For n = nFirstLangCol To nLastLangCol
          sLangName = ThisWorkbook.Worksheets("TxtStyles").ListObjects("colorstab").DataBodyRange(n - nFirstLangCol + 1, 2)
          If Me.Controls.Item("L" & sLangName).Value = True Then
            'Добавить поле в XML, если соответствующий флажок установлен
            Set xmlField = xmlFields.appendChild(xmlDoc.createElement("TxtStyle" & .Cells(2) & "_" & sLangName))
            sCText = Trim(.Cells(n))
            bEmptyLine = bEmptyLine And (sCText = "")
            xmlField.text = PrepareEmptyCell(sCText)
          End If
        Next
        Set xmlSeparator = xmlDoc.createElement("Separator")
        Set xmlField = xmlFields.appendChild(xmlSeparator)
        xmlField.text = FILL_CHAR
      Else
      End If
    End With
    
    'Не сохранять сообщение (Announcement), если на всех выбранных языках оно не произносится
    If bEmptyLine And Not EmptyLineOut.Value Then
      Call Handbook.RemoveChild(Announcement)
    End If
  Next
  Call xmlDoc.Save(sFullFileName)
End Sub

'Получить окончательное имя файла через системный диалог Application.GetSaveAsFilename и подтверждение перезаписи
Private Function GetSaveAsFilename(sFileExt As String, sFileFilter As String, sCaption As String) As Variant
  Dim fd As Variant
  Dim confirm As Boolean
  Static WorkDir As String
  
  If WorkDir = "" Then WorkDir = ThisWorkbook.Path
  fd = WorkDir & "\" & BuildFileName() & sFileExt
  confirm = False
  Do While Not confirm
    fd = Application.GetSaveAsFilename(fd, sFileFilter, 1, sCaption, "Сохранить")
    If fd = False Then Exit Do
      
    Dim fs As New FileSystemObject
    If fs.FileExists(fd) Then
      If MsgBox(fd & " уже существует. " & vbNewLine & "Перезаписать?", vbYesNo + vbQuestion, "Подтвердить сохранение") = vbYes Then
        confirm = True
        WorkDir = Left$(fd, InStrRev(fd, "\") - 1)
      End If
    Else
      confirm = True
      WorkDir = Left$(fd, InStrRev(fd, "\") - 1)
    End If
  Loop
  GetSaveAsFilename = fd
End Function

'Вычислить количество выбранных языков numLang
Private Function GetNumSelectedLangs() As Long
  Dim langRec As ListRow
  GetNumSelectedLangs = 0
  For Each langRec In ThisWorkbook.Worksheets("TxtStyles").ListObjects("colorstab").ListRows
    If Me.Controls.Item("L" & langRec.Range(1, 2)).Value = True Then
      GetNumSelectedLangs = GetNumSelectedLangs + 1
    End If
  Next
End Function

'Сохранить выбранное содержимое как MS Word Document
Private Sub BuildWordDoc(sFullFileName As Variant)
  Dim Chapter As ListObject
  Set Chapter = ThisWorkbook.Worksheets(GetChapterName()).ListObjects(1)
  
  'Параметры таблицы вывода
  nFirstLangCol = Chapter.ListColumns("LstNum").Index + 1
  nLastLangCol = Chapter.ListColumns("Separator").Index - 1
  
  'Приложение и новый документ Word
  Dim WA As New Word.Application
  WA.DisplayAlerts = wdAlertsNone
  WA.Options.CheckGrammarAsYouType = False
  WA.Options.CheckSpellingAsYouType = False
  Dim WD As Word.Document
  Set WD = WA.Documents.Add
  Dim wr As Word.Range
  'Dim wt As Word.Table
  
  'Вычислить номер выбранной главы numChapter
  Dim numChapter As Long
  numChapter = Switch(Me.ChapterButton1.Value, 1, Me.ChapterButton2.Value, 2, Me.ChapterButton3.Value, 3)
  
  'Создать список для нумерации заголовков
  Dim InternalListTemp As Word.ListTemplate
  Set InternalListTemp = HeadingList(WD, numChapter)
  
  Dim oRec As ListRow
  Dim styleID As Long
  Dim nRow As Long: nRow = 0 'Указатель на строку в таблице документа Word
  Dim n As Long: n = 0
  
  For Each oRec In Chapter.ListRows
    n = n + 1
    styleID = oRec.Range.Cells(2)
    Select Case styleID
      Case 0 'Обычный
        Call PlaceSimpleText(nRow, WD, oRec.Range, False)
              
      Case 1 'Название главы
        Call PlaceHeading(nRow, WD, oRec.Range, WD.Styles(wdStyleHeading1))
      
      Case 2 'Заголовок2
        Call PlaceHeading(nRow, WD, oRec.Range, WD.Styles(wdStyleHeading2))
      
      Case 3 'Заголовок3
        Call PlaceHeading(nRow, WD, oRec.Range, WD.Styles(wdStyleHeading3))
      
      Case 4 'Уточнение1
        Call PlaceHeading(nRow, WD, oRec.Range, WD.Styles(wdStyleHeading4))
        
      Case 5 'Уточнение2
        Call PlaceHeading(nRow, WD, oRec.Range, WD.Styles(wdStyleHeading5))
        
      Case 6 'Список
        Call PlaceSimpleText(nRow, WD, oRec.Range, True)
      
      Case 7 'Режим полета
        Call PlaceHeading(nRow, WD, oRec.Range, WD.Styles(wdStyleMessageHeader))
        
      Case 8 'Врезка
        Call PlaceHeading(nRow, WD, oRec.Range, WD.Styles(wdStyleCaption))
    End Select
    If n Mod 16 = 0 Then
      LogLabel.Caption = "Обработано " & CStr(Fix(n / Chapter.ListRows.Count * 100)) & "%"
      DoEvents
      If bStopFlag Then GoTo Interrupt_By_User
    End If
  Next
  Set wr = WD.Content
  wr.Collapse Direction:=wdCollapseStart
  wr.ListFormat.ApplyListTemplate ListTemplate:=InternalListTemp

  LogLabel.Caption = "Сохранение документа…"
  DoEvents
  If Not bStopFlag Then _
    WD.SaveAs2 Filename:=sFullFileName
  
Interrupt_By_User:

  LogLabel.Caption = ""
  WD.Close SaveChanges:=wdDoNotSaveChanges
  WA.Quit
  Set WD = Nothing: Set WA = Nothing
  If bStopFlag Then Unload Me
End Sub

'Поместить в документ Word обычный текст или список
Private Sub PlaceSimpleText(nRow As Long, WD As Word.Document, rngLine As Range, bIsList As Boolean)
  Dim wr As Word.Range
  Static wt As Word.Table
  Set wr = WD.Content
  wr.Collapse Direction:=wdCollapseEnd
  
  Dim i As Long
  Dim n As Long
  Dim bEmptyLine As Boolean
  
  Dim arrLang() As String: ReDim arrLang(GetNumSelectedLangs())
  Dim arrColor() As Long: ReDim arrColor(GetNumSelectedLangs())
  
  n = 1: bEmptyLine = True
  
  'Перебрать все языки и сохранить в arrLang публикуемые тексты, в arrColor соответствующие им цвета
  For i = nFirstLangCol To nLastLangCol
    With ThisWorkbook.Worksheets("TxtStyles").ListObjects("colorstab")
      If Me.Controls.Item("L" & .DataBodyRange(i - nFirstLangCol + 1, 2)).Value = True Then
        arrLang(n) = rngLine.Cells(i)
        arrColor(n) = RGB(.DataBodyRange(i - nFirstLangCol + 1, 4), _
                          .DataBodyRange(i - nFirstLangCol + 1, 5), _
                          .DataBodyRange(i - nFirstLangCol + 1, 6))
        bEmptyLine = bEmptyLine And (arrLang(n) = "")
        n = n + 1
      End If
    End With
  Next
  
  'Вывести строку, если она непустая или установлена принудительная опция
  If (EmptyLineOut.Value = True) Or Not bEmptyLine Then
    'Новая таблица или добавление строки в таблицу
    If nRow = 0 Then
      Set wt = wr.Tables.Add(wr, 1, GetNumSelectedLangs())
      wt.Style = -155 'Table Grid
      wt.Rows.AllowBreakAcrossPages = False
      nRow = 1
    Else
      wt.Rows.Add
      nRow = nRow + 1
    End If
    'Вывести и раскрасить
    For n = 1 To GetNumSelectedLangs()
      Call wt.cell(nRow, n).Range.InsertAfter(arrLang(n))
      If bIsList Then
        wt.cell(nRow, n).Range.ListFormat.ApplyListTemplate _
            ListTemplate:=WD.Application.ListGalleries(wdBulletGallery).ListTemplates(1), _
            ContinuePreviousList:=False, _
            ApplyTo:=wdListApplyToSelection
      Else
        wt.cell(nRow, n).Range.Style = WD.Styles(wdStyleBodyText)
      End If
      wt.cell(nRow, n).Range.Font.TextColor.RGB = arrColor(n)
    Next
  End If
End Sub

'Поместить в документ Word заголовок, режим полета или врезку
Private Sub PlaceHeading(nRow As Long, WD As Word.Document, rngLine As Range, parStyle As Word.Style)
  Dim wr As Word.Range
  Set wr = WD.Content
  wr.Collapse Direction:=wdCollapseEnd
        
  nRow = 0
  Call wr.InsertAfter(rngLine.Cells(nFirstLangCol))
  wr.InsertParagraphAfter
  wr.Collapse Direction:=wdCollapseStart
  wr.Style = parStyle
End Sub

'Шаблон списка для заголовков
Private Function HeadingList(WD As Word.Document, nStartChapter As Long) As Word.ListTemplate
  Set HeadingList = WD.Application.ListGalleries(wdOutlineNumberGallery).ListTemplates(1)
    With HeadingList.ListLevels(1)
        .NumberFormat = "%1."
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(0.63)
        .TabPosition = wdUndefined
        .ResetOnHigher = 0
        .StartAt = nStartChapter
        .LinkedStyle = WD.Styles(wdStyleHeading1)
    End With
    With HeadingList.ListLevels(2)
        .NumberFormat = "%1.%2."
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0.63)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(1.4)
        .TabPosition = wdUndefined
        .ResetOnHigher = 1
        .StartAt = 1
        .LinkedStyle = WD.Styles(wdStyleHeading2)
    End With
    With HeadingList.ListLevels(3)
        .NumberFormat = "%1.%2.%3."
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(1.27)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(2.16)
        .TabPosition = wdUndefined
        .ResetOnHigher = 2
        .StartAt = 1
        .LinkedStyle = WD.Styles(wdStyleHeading3)
    End With
End Function

Private Sub CancelButton_Click()
  Unload Me
End Sub

Private Sub Label2_Click()
  ActiveWorkbook.FollowHyperlink SBOGDANOV_HYPERLINK
End Sub

Private Sub Label2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call SetHandCursor
  Me.Label2.ForeColor = vbWindowText
  Label3.ForeColor = vbWindowText
End Sub

Private Sub SaveAsWordButton_Click()
  Dim fd As Variant
  fd = GetSaveAsFilename(".docx", "Документ Word (*.docx), *.xml, Все файлы (*.*), *.*", "Сохранение в формате MS Word")
  If fd <> False Then
    bStopFlag = False
    Me.SaveAsXMLButton.Visible = False
    Me.SaveAsWordButton.Visible = False
    Me.CancelButton.Visible = False
    Me.StopButton.Visible = True
    Me.Frame1.Enabled = False
    Me.Frame2.Enabled = False
    Me.EmptyLineOut.Enabled = False
    Me.LogLabel.Caption = "Подготовка данных…"
    
    Call BuildWordDoc(fd)
    
    Me.SaveAsXMLButton.Visible = True
    Me.SaveAsWordButton.Visible = True
    Me.CancelButton.Visible = True
    Me.StopButton.Visible = False
    Me.Frame1.Enabled = True
    Me.Frame2.Enabled = True
    Me.EmptyLineOut.Enabled = True
  End If
End Sub

Private Sub SaveAsXMLButton_Click()
  Dim fd As Variant
  fd = GetSaveAsFilename(".xml", "XML (*.xml), *.xml, Все файлы (*.*), *.*", "Сохранение в XML")
  If fd <> False Then Call BuildXMLDoc(fd)
End Sub

Private Sub StopButton_Click()
  bStopFlag = True
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Label2.ForeColor = vbButtonShadow
  Label3.ForeColor = vbButtonShadow
End Sub

Private Sub SetHandCursor()
  SetCursor LoadCursor(0, IDC_HAND)
End Sub
