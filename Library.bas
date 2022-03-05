Attribute VB_Name = "Library"
Public inv As String
Public kq_File_duoc_chon As String
Public tongsheet As Integer



Function Edward(ed2 As String, ed3 As String)
Edward = Sheets(ed2).Range(ed3)

End Function
Function UniConvert(Text As String, InputMethod As String) As String
  Dim VNI_Type, Telex_Type, CharCode, Temp, i As Long
  UniConvert = Text
  VNI_Type = Array("a81", "a82", "a83", "a84", "a85", "a61", "a62", "a63", "a64", "a65", "e61", _
      "e62", "e63", "e64", "e65", "o61", "o62", "o63", "o64", "o65", "o71", "o72", "o73", "o74", _
      "o75", "u71", "u72", "u73", "u74", "u75", "a1", "a2", "a3", "a4", "a5", "a8", "a6", "d9", _
      "e1", "e2", "e3", "e4", "e5", "e6", "i1", "i2", "i3", "i4", "i5", "o1", "o2", "o3", "o4", _
      "o5", "o6", "o7", "u1", "u2", "u3", "u4", "u5", "u7", "y1", "y2", "y3", "y4", "y5")
  Telex_Type = Array("aws", "awf", "awr", "awx", "awj", "aas", "aaf", "aar", "aax", "aaj", "ees", _
      "eef", "eer", "eex", "eej", "oos", "oof", "oor", "oox", "ooj", "ows", "owf", "owr", "owx", _
      "owj", "uws", "uwf", "uwr", "uwx", "uwj", "as", "af", "ar", "ax", "aj", "aw", "aa", "dd", _
      "es", "ef", "er", "ex", "ej", "ee", "is", "if", "ir", "ix", "ij", "os", "of", "or", "ox", _
      "oj", "oo", "ow", "us", "uf", "ur", "ux", "uj", "uw", "ys", "yf", "yr", "yx", "yj")
  CharCode = Array(ChrW(7855), ChrW(7857), ChrW(7859), ChrW(7861), ChrW(7863), ChrW(7845), ChrW(7847), _
      ChrW(7849), ChrW(7851), ChrW(7853), ChrW(7871), ChrW(7873), ChrW(7875), ChrW(7877), ChrW(7879), _
      ChrW(7889), ChrW(7891), ChrW(7893), ChrW(7895), ChrW(7897), ChrW(7899), ChrW(7901), ChrW(7903), _
      ChrW(7905), ChrW(7907), ChrW(7913), ChrW(7915), ChrW(7917), ChrW(7919), ChrW(7921), ChrW(225), _
      ChrW(224), ChrW(7843), ChrW(227), ChrW(7841), ChrW(259), ChrW(226), ChrW(273), ChrW(233), ChrW(232), _
      ChrW(7867), ChrW(7869), ChrW(7865), ChrW(234), ChrW(237), ChrW(236), ChrW(7881), ChrW(297), ChrW(7883), _
      ChrW(243), ChrW(242), ChrW(7887), ChrW(245), ChrW(7885), ChrW(244), ChrW(417), ChrW(250), ChrW(249), _
      ChrW(7911), ChrW(361), ChrW(7909), ChrW(432), ChrW(253), ChrW(7923), ChrW(7927), ChrW(7929), ChrW(7925))
  Select Case InputMethod
    Case Is = "VNI": Temp = VNI_Type
    Case Is = "Telex": Temp = Telex_Type
  End Select
  For i = 0 To UBound(CharCode)
    UniConvert = Replace(UniConvert, Temp(i), CharCode(i))
    UniConvert = Replace(UniConvert, UCase(Temp(i)), UCase(CharCode(i)))
  Next i
End Function
Sub Open_Single_File()
     'Khai báo các bi?n s? d?ng
     Dim dk_Ten_tieu_de As String     'Tên tiêu d? c?a s? Workbook Open
     Dim dk_Loc_LoaiFile As String    'L?c các lo?i file có trong c?a s? Workbook Open
     Dim dk_Loc_Index As Integer      'Th? t? l?c m?c d?nh
        'File du?c ch?n
     
          dk_Loc_LoaiFile = "Excel Files (*.xls*),*.xls," & "CSV Files (*.csv),*.csv,"
          dk_Loc_Index = 1
          dk_Ten_tieu_de = "Select Your Input File of Choice"
     'M? c?a s? Workbook Open
     kq_File_duoc_chon = Application.GetOpenFilename _
          (FileFilter:=dk_Loc_LoaiFile, _
           FilterIndex:=dk_Loc_Index, _
           Title:=dk_Ten_tieu_de)
     'Các tru?ng h?p không thành công
     'Tru?ng h?p 1: Không có file d?oc ch?n
     
     If kq_File_duoc_chon = "" Then
          MsgBox ("Khong co file duoc chon")
          Exit Sub
     'Tru?ng h?p 2. B?m vào nút Cancel
     ElseIf kq_File_duoc_chon = "False" Then
          MsgBox ("Ban da bam lenh Huy thao tac")
          Exit Sub
     End If
     'Tru?ng h?p thành công: M? file du?c ch?n
     Workbooks.Open kq_File_duoc_chon
     inv = Right(kq_File_duoc_chon, Len(kq_File_duoc_chon) - InStr(kq_File_duoc_chon, "Invoice") + 1)
     w = ThisWorkbook.Name
    Application.Workbooks("" & inv).Activate
    Range("F26:G55").Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    Windows("" & w). _
        Activate
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

End Sub



Function Eval(Ref As String)
Application.Volatile
Eval = Evaluate(Ref)
End Function

Sub UnhideAllSheets()
  
 
 Range("N2") = "PSED-10-500"
 Range("N3") = "PE01-10-121A"
 Range("N4") = "FZA1-19-4JYA"
 Range("N5") = "FZA5-19-4JYA"
 Range("N6") = "FZA4-19-4JYA"
 Range("N7") = "S550-10-121"
 Range("N8") = "S550-10-131"
 Range("N9") = "S550-10-141B"
 Range("N10") = "PSED-10-190"
 Range("C1") = "=CONCATENATE(ROUND(IF(OR(A1=$N$3,A1=$N$4,A1=$N$5,A1=$N$6,A1=$N$7,A1=$N$8,A1=$N$9),""" & "SSS" & """,IF(A1=$N$10,Edward(A1,""" & "L18" & """),IF(A1=$N$2,Edward(A1,""" & "L23" & """),Edward(A1,""" & "L19" & """))))*100,2),""" & "%" & """)"
 
   Dim wsSheet As Worksheet
    Sheets("Sheet1").Tab.ColorIndex = 4
   For Each wsSheet In ActiveWorkbook.Worksheets

       wsSheet.Visible = xlSheetVisible

   Next wsSheet

End Sub
Sub SheetAscending()
 n = WorksheetFunction.CountA(Range("A1:A20"))
 tongsheet = ThisWorkbook.Sheets.Count
For i = 1 To Application.Sheets.Count
    For j = 1 To Application.Sheets.Count
    If UCase$(Sheets(j).Name) = UCase$(Sheets("sheet1").Cells(i, 1)) Then
    Sheets(j).Tab.ColorIndex = 4
    Sheets(j).Move before:=Sheets(i)
    End If
    Next
    
Next

Sheets("Sheet1").Activate
 
End Sub
Sub SheetAshiding()
For i = 1 To tongsheet
If Sheets(i).Tab.ColorIndex <> 4 Then
Sheets(i).Visible = False
End If
Next
For j = 1 To Application.Sheets.Count
    Sheets(j).Tab.ColorIndex = 6
Next
Sheets("" & UniConvert("Toorng", "Telex")).Visible = True
Sheets("" & UniConvert("Toorng", "Telex")).Move before:=Sheets(1)
End Sub
Sub onehit()
    Call Open_Single_File
    If kq_File_duoc_chon <> "" Then
    Call UnhideAllSheets
    Call SheetAscending
    Call SheetAshiding
    End If
    Sheets("Sheet1").Activate
End Sub

