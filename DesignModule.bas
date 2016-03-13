Attribute VB_Name = "DesignModule"
Option Explicit

'������ �������� ����� ��� ���� ������, ���������������� � ��������� InitColors
'������ ������������� �������� �� ������� ListNum
Dim LangColor() As Long

'����� ������� LstNum
Dim nLstNumCol As Long

'����� ������� ��� ������� �����
Dim nFirstLangCol As Long

'����� ������� ��� ���������� �����
Dim nLastLangCol As Long

'�������� ������ �� "��������" ��������
Private Function Sanitate(sDirty As String) As String
  Sanitate = Trim(sDirty)
  '������� ���������� �������
  Dim n As Long
  For n = 0 To 31
    Sanitate = Replace(Sanitate, Chr(n), " ")
  Next
  '������� ������� �������
  Do
    n = Len(Sanitate)
    Sanitate = Replace(Sanitate, "  ", " ")
  Loop While n > Len(Sanitate)
End Function

'������������ ���������� ������ ��� ������ �������
Public Sub DesignTabLine(trg As Range)
  Dim InteriorRange As Range
  Dim cell As Range
  Dim ThisTable As ListObject
  Dim nRow As Long
  Dim nCol As Long
  'Dim i As Long
  Dim styleID As Long
  Dim PureText As String
  
  Set ThisTable = trg.ListObject
  '���������, ���� ������ ��� �������
  If ThisTable Is Nothing Then Exit Sub
  
  Application.ScreenUpdating = False
  
  '������������� ���������� ����������
  InitColors
  nLstNumCol = ThisTable.ListColumns("LstNum").Index
  nFirstLangCol = nLstNumCol + 1
  nLastLangCol = ThisTable.ListColumns("Separator").Index - 1
  
  For Each cell In trg.Cells
    '����������, ���� ������ �� ��������� ������� ������ �������
    If IsRangeInside(ThisTable.DataBodyRange, cell) = True Then
      Call GetRelCoordinates(cell, nRow, nCol)
    
      '���������� �������� InteriorRange ��� ���������� �����:
      '������� �������� ������ ��� ��� ������ ������� �������
      Set InteriorRange = ThisTable.ListRows(nRow).Range
      styleID = InteriorRange.Cells(1, 2)
      If nCol = 0 Then
        '��� ������� ����� ������ � ������� ���������� ����� �� ���������
        If cell.Value = "" Then
          cell.Value = ThisWorkbook _
                       .Worksheets("TxtStyles") _
                       .ListObjects("txtstylestab") _
                       .ListColumns("style_Name") _
                       .DataBodyRange _
                       .Cells(1) _
                       .Value
          '������������ ������ � ���������
          Call RestoreAutofill(ThisTable)
        End If
        '������� ��� ������, ���� �������� �����
        With InteriorRange
          Set InteriorRange = Range(.Cells(1, nFirstLangCol), .Cells(1, nLastLangCol))
        End With
      ElseIf nCol >= nFirstLangCol And nCol <= nLastLangCol Then
        '������� ������� ������, ���� �������� ����� �� �����-���� �����
        Set InteriorRange = cell
        '�������� ����� �� ������
        PureText = Sanitate(cell.Value)
        If Len(PureText) < Len(cell.Value) Then cell.Value = PureText
      Else
        '��� ��������� �����
        If cell.Errors.Item(xlEvaluateToError).Value = True Then
          '������������ ������ � ���������
          Call RestoreAutofill(ThisTable)
        End If
        Set InteriorRange = Nothing
      End If
        
      '��������� ����� � ��������� InteriorRange
      If Not (InteriorRange Is Nothing) Then
        Call SetStyle(styleID, InteriorRange)
      End If
    
    End If
  Next

  Application.ScreenUpdating = True
End Sub

'������������ �������������� ���� ����� � ���������.
'� �������� ������� ������� ������ ������ � �������
Private Sub RestoreAutofill(Table As ListObject)
  '����� ������� LstNum, �������� "$F$1"
  Dim sLstNumAddr As String
  sLstNumAddr = Table.ListColumns("LstNum").Range(1).Address()
  
  Dim DstRange As Range
  Dim SrcRange As Range
  Set DstRange = Table.DataBodyRange.Range("B2:" & Mid(sLstNumAddr, 2, (InStr(2, sLstNumAddr, "$") - 2)) & Table.DataBodyRange.Rows.Count)
  Set SrcRange = DstRange.Rows(1)
  SrcRange.AutoFill Destination:=DstRange, Type:=xlFillValues
End Sub

' ��������� ������������� ���������� ������ � ������� ������ ������� � ��������� �� � row � col.
' ������ ���������� � 1. ������� ������������ �������� �������� ������.
Private Sub GetRelCoordinates(cell As Range, row As Long, col As Long)
  Dim regEx As New RegExp
  Dim Matches As MatchCollection
  Dim sAddr As String
  
  sAddr = UCase( _
                cell.Address(RowAbsolute:=False, _
                             ColumnAbsolute:=False, _
                             ReferenceStyle:=xlR1C1, _
                             RelativeTo:=cell.ListObject.DataBodyRange(1, 1)) _
                )
  regEx.IgnoreCase = True
  regEx.Global = True
  regEx.Pattern = "[RC]\[?(\d*)"
  Set Matches = regEx.Execute(sAddr)
  row = 0
  col = 0
  If Matches(0).SubMatches(0) <> "" Then row = CLng(Matches(0).SubMatches(0)) + 1
  If Matches(1).SubMatches(0) <> "" Then col = CLng(Matches(1).SubMatches(0)) + 1
End Sub

'��������� ����� ��������� area
Private Sub SetStyle(style_ID As Long, area As Range)
  Dim cell As Range
  
  Select Case style_ID
    Case 0
    '�������
    area.Style = "Normal"
    area.Font.Bold = False
    area.Font.Italic = False
    area.Font.Underline = False
    area.WrapText = True
    Call SetStyleInLstNum(area)
    For Each cell In area
      Call ColorizeLangCell(cell)
    Next
          
    Case 1
    '�������� �����
    area.Style = "��������� 1"
    area.WrapText = False
    Call SetStyleInLstNum(area, "��������� 1")
          
    Case 2
    '���������2
    area.Style = "��������� 2"
    area.WrapText = False
    Call SetStyleInLstNum(area, "��������� 2")
          
    Case 3
    '���������3
    area.Style = "��������� 3"
    area.WrapText = False
    Call SetStyleInLstNum(area, "��������� 3")
          
    Case 4
    '���������1
    area.Style = "���������1"
    Call SetStyleInLstNum(area)
    area.WrapText = False
          
    Case 5
    '���������2
    area.Style = "���������2"
    Call SetStyleInLstNum(area)
    area.WrapText = False
          
    Case 6
    '������
    area.Style = "Normal"
    area.Font.Bold = False
    area.Font.Italic = False
    area.WrapText = True
    Call SetStyleInLstNum(area)
    For Each cell In area
      Call ColorizeLangCell(cell)
    Next
          
    Case 7
    '����� ������
    area.Style = "Normal"
    area.Font.Bold = True
    area.Font.Italic = False
    area.WrapText = False
    Call SetStyleInLstNum(area)
          
    Case 8
    '������
    area.Style = "Normal"
    area.Font.Bold = False
    area.Font.Italic = False
    area.WrapText = False
    Call SetStyleInLstNum(area)
  End Select
End Sub

'��������� ����� ���� LstNum
Private Sub SetStyleInLstNum(area As Range, Optional styleName As String = "Normal")
  Dim r As Long
  Dim c As Long
  Call GetRelCoordinates(area.Cells(1, 1), r, c)
  area.ListObject.DataBodyRange(r, area.ListObject.ListColumns("LstNum").Index).Style = styleName
End Sub

'���������� �������� ������
Private Sub ColorizeLangCell(cell As Range)
  Dim r As Long
  Dim c As Long
  Call GetRelCoordinates(cell, r, c)
  If c >= nFirstLangCol And c <= nLastLangCol Then
    cell.Font.Color = LangColor(c - nFirstLangCol)
  End If
End Sub

'��������� � ������ LangColor �������� ������ �� ������� colorstab
Public Sub InitColors()
  Dim i As Integer
  Dim fc As Long
  With ThisWorkbook.Worksheets("TxtStyles").ListObjects("colorstab")
    ReDim LangColor(.ListRows.Count - 1)
    For i = 1 To .ListRows.Count
      fc = RGB(.DataBodyRange(i, 4), .DataBodyRange(i, 5), .DataBodyRange(i, 6))
      LangColor(i - 1) = fc
      .DataBodyRange(i, 2).Font.Color = fc
      .DataBodyRange(i, 3).Font.Color = fc
    Next
  End With
End Sub

'���������� true, ���� �������� ExternRange ��������� �������� � ���� InternRange
Public Function IsRangeInside(ExternRange As Range, InternRange As Range) As Boolean
  Dim isec As Range
  IsRangeInside = False
  If Not (ExternRange Is Nothing Or InternRange Is Nothing) Then
    Set isec = Application.Intersect(ExternRange, InternRange)
    If Not (isec Is Nothing) Then
      If isec.Count <= InternRange.Count Then
        IsRangeInside = True
      End If
    End If
  End If
End Function

