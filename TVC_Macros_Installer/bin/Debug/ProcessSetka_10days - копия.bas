'Описание суперфайла
Attribute VB_Name = "ProcessSetka"
Function RemoveNulls(ByVal DurStr As String) As String
  '0:XX:XX->_XX:XX
  DurStr = replace("_" + DurStr, "_0:", "_")
  '_0X:XX->_X:XX
  DurStr = replace(DurStr, "_0", "_")
  'X:00:XX->X:_00_:XX
  DurStr = replace(DurStr, ":00:", ":_00_:")
  'X:XX:00->X:XX
  DurStr = replace(DurStr, ":00", "")
  '_0->
  If Left(DurStr, 2) = "_0" And Mid(DurStr, 3, 1) <> ":" Then
    DurStr = ""
  End If
  RemoveNulls = replace(DurStr, "_", "")
End Function

Sub MacroChangeNullsAndFont()
'
' ChangeNullsAndFont
'
Dim AZoom, i, j, p, p1 As Integer
Dim ACharacters As Characters
Dim StrMinutes As String
Dim Skobka1Pos, Skobka2Pos As Integer
Dim s, ClearDur, AdvertDur, PromoDur, Durs As String
Dim BegTime As String
Dim Delta As Double
'infinity
Delta = 1E+100

Application.ScreenUpdating = False

'���������� ����� �� ������������ �������
For j = 1 To ActiveSheet.UsedRange.Columns.Count
  If (Len(ActiveSheet.Cells(2, j).Value) = 5) And (Mid(ActiveSheet.Cells(2, j).Value, 3, 1) = ":") Then
    For i = 2 To ActiveSheet.UsedRange.Rows.Count
      If (Len(ActiveSheet.Cells(i, j).Value) = 5) And (Mid(ActiveSheet.Cells(i, j).Value, 3, 1) = ":") Then
        StrMinutes = Right(ActiveSheet.Cells(i, j).Value, 1) '�������� ��������� �����
        If CByte(StrMinutes) >= 5 Then
          StrMinutes = "5"
        Else
          StrMinutes = "0"
        End If
        ActiveSheet.Cells(i, j).Value = Left(ActiveSheet.Cells(i, j).Value, 4) + StrMinutes
      End If
    Next i
  End If
Next j

For i = 1 To ActiveSheet.Shapes.Count
  Set ACharacters = ActiveSheet.Shapes(i).TextFrame.Characters
  ' ������ �������
  ACharacters.Font.Bold = True
  ACharacters.Font.Size = ACharacters.Font.Size + 1.5
  ' ����� ������ �������
  ' ������ ����� �� 20, ���� �� ������ 20
  If ACharacters.Font.Size > 22 Then
    ACharacters.Font.Size = 22
  End If
  s = ACharacters.Text
  '���� ��������� ����������� ������
  p = 0
  Skobka1Pos = 0
  Skobka2Pos = 0
  Do
    p = InStr(p + 1, s, "(")
     ' ������� And (Mid(s, p + 1, 1) <> "�")
    If (p > 0) And (Mid(s, p + 1, 1) <> "�") Then
    Skobka1Pos = p
    End If
  Loop While p > 0
  '� ��������� �����������
  If Skobka1Pos > 0 Then
    Skobka2Pos = InStr(Skobka1Pos + 1, s, ")")
  End If
  If Skobka1Pos > 0 And Skobka2Pos > 0 Then
    '���� �� ��� ����� �� ������ (X:XX:XX+X:XX:XX+X:XX:XX)
    ClearDur = ""
    AdvertDur = ""
    PromoDur = ""
    p = InStr(Skobka1Pos + 1, s, "+")
    If p > 0 And p < Skobka2Pos Then
      ClearDur = Mid(s, Skobka1Pos + 1, p - Skobka1Pos - 1)
      p1 = InStr(p + 1, s, "+")
      If p1 > 0 And p1 < Skobka2Pos Then
        AdvertDur = Mid(s, p + 1, p1 - p - 1)
        PromoDur = Mid(s, p1 + 1, Skobka2Pos - p1 - 1)
      End If
    End If
    '��������� ������ ������, ���������� ���� � �������� � ����� �������
    If ClearDur <> "" And AdvertDur <> "" And PromoDur <> "" Then
      ClearDur = RemoveNulls(ClearDur)
      AdvertDur = RemoveNulls(AdvertDur)
      '���
      f = 0
      If Right(AdvertDur, 1) = "]" Then
      f = InStr(f + 1, AdvertDur, "[")
      AdvertDur = Left(AdvertDur, f - 1)
      End If
      f = 0
      If Right(PromoDur, 1) = "]" Then
      f = InStr(f + 1, PromoDur, "[")
      PromoDur = Left(PromoDur, f - 1)
      End If
      '����� ���
      PromoDur = RemoveNulls(PromoDur)
      '��� ������� ����� �, ��� ����� (�����) - ����� �
      Durs = ""
      If ClearDur <> "" Then
        Durs = ClearDur + "+"
      End If
      If AdvertDur <> "" Then
        Durs = Durs + AdvertDur + "�+"
      End If
      If PromoDur <> "" Then
        Durs = Durs + PromoDur + "�+"
      End If
      Durs = Left(Durs, Len(Durs) - 1)
      If Durs = "" Then
        Durs = "0"
      End If
      
      '������� ��������� ��-�� ������������ �������
      If Delta = 1E+100 Then
        BegTime = Mid(s, Skobka2Pos + 2, 5)
        For j = 2 To ActiveSheet.UsedRange.Rows.Count
          If ActiveSheet.Cells(j, 1).Value = BegTime Then
            Delta = ActiveSheet.Cells(j + 1, 1).Top - ActiveSheet.Shapes(i).Top
            Exit For
          End If
        Next j
      End If
      
      '�������� ��-�� � ������� ����� �����
      ACharacters.Text = Left(s, Skobka1Pos) + Durs + ")" + Mid(s, Skobka2Pos + 19)
    End If
  End If
Next i

'��������� ��-�� ������������ �������
'If Delta <> 0 And Delta <> 1E+100 Then
' For i = 1 To ActiveSheet.Shapes.Count
'  ActiveSheet.Shapes(i).Top = ActiveSheet.Shapes(i).Top + Delta
'Next i
'End If

' ������ �������
 f = 0
If Range("A5").Value = Range("L5").Value Then
f = 1
End If
'�����������
If f = 0 Then
Columns("A:A").Select
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    Selection.replace What:=":**", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByColumns, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
Columns("B:B").Select
    Selection.replace What:="**:", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByColumns, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
k = 0
s = 0
Range("A1:B1").Select
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
For i = 3 To 351
If Range("A" & i).Value <> Range("A" & i + 1).Value Then
Range("A" & i & ":A" & (i - k)).Merge
Range("A" & i & ":B" & i).Select
  With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With

k = -1
Else
Range("A" & i).Clear
End If
k = k + 1
Next i

Columns("B:B").ColumnWidth = 4
Columns("A:A").Font.Size = 22
   Columns("A:A").Select
    Selection.Columns.AutoFit
     With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
Columns("B:B").Select
Selection.ClearContents
Columns("B:B").Copy
Columns("AE:AE").Insert
Columns("A:A").Copy
Columns("AF:AF").Insert
Columns("AG:AG").Delete

Else

'�������

If CInt(Right(Left(Range("A3").Value, 5), 2)) - (Right(Left(Range("A2").Value, 5), 2)) > 0 Then
rn = CInt(Right(Left(Range("A3").Value, 5), 2)) - (Right(Left(Range("A2").Value, 5), 2))
Else
rn = CInt(Right(Left(Range("A4").Value, 5), 2)) - (Right(Left(Range("A3").Value, 5), 2))
End If

For i = 2 To 13
If (CInt(Left(Range("A" & i + 1).Value, 2)) - CInt(Left(Range("A" & i).Value, 2)) = 1) And ((CInt(Right(Range("A" & i + 1).Value, 2)) / rn) < 0.5) Then
st = CInt(Right(Range("A" & i + 1).Value, 2)) / rn
    Range("A2:A350").Select
    Selection.Cut Destination:=Range("A3:A351")
    Range("A2").Value = Range("A3").Value
    i = 15
End If
Next i

Columns("A:A").Select
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    Selection.replace What:=":**", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByColumns, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
Columns("B:B").Select
    Selection.replace What:="**:", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByColumns, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
k = 0
s = 0
Range("A1:B1").Select
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
For i = 2 To 351
If Range("A" & i).Value <> Range("A" & i + 1).Value Then
Range("A" & i & ":A" & (i - k)).Merge
Range("A" & i & ":B" & i).Select
  With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With

k = -1
Else
Range("A" & i).Clear
End If
k = k + 1
Next i

Columns("B:B").ColumnWidth = 4
Columns("A:A").Font.Size = 22
   Columns("A:A").Select
    Selection.Columns.AutoFit
     With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
Columns("B:B").Select
Selection.ClearContents
Columns("B:B").Copy
Columns("M:M").Insert
Columns("A:A").Copy
Columns("N:N").Insert
Columns("O:O").Delete
'Columns("AF:AF").Delete
'Columns("AG:AG").Delete

End If
ActiveSheet.PageSetup.CenterHeader = replace(ActiveSheet.PageSetup.CenterHeader, " - ������", "")
ActiveSheet.PageSetup.PrintArea = ""
' ����� ������ �������

AZoom = ActiveWindow.Zoom
ActiveWindow.Zoom = 10
ActiveWindow.Zoom = AZoom
Application.ScreenUpdating = True

End Sub 'MacroChangeNullsAndFont



