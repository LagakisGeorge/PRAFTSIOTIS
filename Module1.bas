Attribute VB_Name = "Module1"
Global gConnect As String
Global gdb As Connection
Global g_Stop
Declare Function GetCurrentTime Lib "kernel32" Alias "GetTickCount" () As Long



Dim f_kodik As Recordset, F_T As Recordset, f_tab(40)
Dim f_db As Database

Function print7_excel(ByVal sql As String, SUgm_str, ByVal EPIKEF As String, ByVal GROUPN As Integer)
'"*********************** P R I N T _ X A R ************************
'"**************      ����� ���������� ������� ��� �����
'"*** n=������� fields ��� ��������������
'"*** synuhkh2 � ������� ��� ��  IF, sugm_str ������ ��� �� ����� ������ ,�.�. "00100" ����� ������ ��� 3� field
'"*** sum_pic  �� picture ��� ��� ������
'"*** Ei,Fi,Pi  :����������� �����������,Fields ��� ��������������,Picture �����������
'"** synt_eject:=0  ������� ������ ���� ������������ � �����
Dim mfields(120), mSYN
Dim synt_eject
Dim SUMES, CC, pp, ar_Print(4), k, m_sthl_ektyp(120), f(120)
Dim mBSEIRA
Dim scr2, dhdr(1), dfld(1), marxeio2, mPal, mPAL22, MOLIS_ALLAJE
Dim PrinSeir1, PrinSeir2, PrinSeir3, PrinSeir4
Dim aaP, aaP2, ektypoths
Dim EPIK, xeirisths, PPF, epik3
Dim sthles, kw, PPD, SELIS
Dim LSYN, aa, epik2, I As Integer
Dim AYJ, end_page, Typose, aaF, aaF2, mpal2, eject
Dim m_sumes(120), SUMES0(120)
Dim ss

Dim m_ekt As Integer
Dim DUM
Dim returnValue

Dim f_excelPath




Dim Excel As Excel.Application
Dim workbook As Excel.workbook
Dim myXL As Excel.Worksheet

  Set Excel = New Excel.Application
  ' Excel.Visible = True
  Set workbook = Excel.Workbooks.Add

 On Error Resume Next


'  If (MenuShow.Caption = "&Show") Then
'    MenuShow.Caption = "&Hide"
    workbook.Activate

Set myXL = workbook.ActiveSheet

Dim FF As New UDialog
FF.Show
FF.OKButton.Visible = False
FF.List1.Visible = False

FF.CancelButton.Caption = "�������"
FF.CancelButton.Top = 120
FF.CancelButton.Left = 120
FF.CancelButton.Width = 2895
FF.CancelButton.Height = 495

'FF.Top = 3000
'FF.Left = 3000

FF.Width = 3210
FF.Height = 810


    FF.Left = Screen.Width / 2 - FF.Width / 2
    FF.Top = Screen.Height / 2 - FF.Height / 2




'FF.Top = MDIForm1.Top + (MDIForm1.height) / 2 ' Command2.Top
'FF.Left = MDIForm1.Left + (MDIForm1.width) / 2 ' Command2.Left





FF.Caption = "���������� �����������"



''------------------------������ �������� �� ������ ���� ������ ���� [����.����]------------------------------
'Dim db1 As Database
'Dim FROM As Long
'On Error Resume Next
'  Kill "C:\EKTYP.XLS"
'  DoEvents
'  myXL.SaveAs "C:\EKTYP.XLS"
'
'  Call workbook.Close(False)
'  excel.Quit
'  Set excel = Nothing
'
'FROM = InStr(1, UCase(sql), "FROM")
'Set db1 = OpenDatabase("", False, False, gConnect)
'db1.Execute Left(sql, FROM - 1) + " " + " into 32 in 'c:\EKTYP.xls' 'Excel 8.0;' " + Trim(Mid(sql, FROM, 500))
''------------------------------------------------------







MDIForm1.MousePointer = vbHourglass
'excel.Visible = True

'f_excelPath = FindParametroi("MDIFORM1", "f_excelPath", "C:\Program Files\Microsoft Office\OFFICE11", "������� Excel")
''C:\Program Files\Microsoft Office\OFFICE11
'returnValue = Shell(f_excelPath + "\EXCEL.EXE", vbMaximizedFocus) ' vbMinimizedNoFocus) ' Run Microsoft Excel.
'
'Set myXL = GetObject("", "Excel.Sheet")




For k = 1 To 120: m_sumes(k) = 0: SUMES0(k) = 0: Next



Dim n
Dim F_T As New ADODB.Recordset

Dim TT As Long



'Dim recs As Long, fp As Long
'fp = InStr(UCase(sql), "FROM")


'F_T.Open "select count(*) " + Trim(Mid(sql, fp, 100)), Gdb, adOpenForwardOnly, adLockReadOnly
'recs = F_T(0)
'F_T.Close


F_T.Open sql, gdb, adOpenForwardOnly, adLockReadOnly




n = F_T.Fields.Count



For k = 1 To n
    If F_T(k - 1).Type = 8 Or F_T(k - 1).Type = 129 Then 'DATE
      f_tab(k) = 2 + f_tab(k - 1) + 8
    Else
      If F_T.Fields(k - 1).DefinedSize > 200 Then
         f_tab(k) = 2 + f_tab(k - 1) + 30
      Else
         f_tab(k) = 3 + f_tab(k - 1) + F_T.Fields(k - 1).DefinedSize
      End If
    End If
Next


LSYN = f_tab(k - 1)





mPal = "    "
mPAL22 = "     "
MOLIS_ALLAJE = 0


'If IsNull(EPIK) Then
 '   EPIK = Format(Date, "dd/mm/yyyy")
'End If

marxeio2 = "EKT" + xeirisths + ".TXT"
'On Error Resume Next

'   m_sthl_ektyp(1) = 0  ' int ( IF(type("STHLES")="U",40,STHLES/2) - lsyn / 2 )
'   For k = 1 To n
'          aa = LTrim(Str(k))
'          m_sthl_ektyp(k + 1) = m_sthl_ektyp(k) + Len(macro("p", aa)) + 1
'   Next

Dim R As New ADODB.Recordset
R.Open "SELECT *FROM MEM", gdb, adOpenDynamic, adLockOptimistic
        k = 1
        With myXL
            .Cells(1, 3) = R("pelono")
            .Cells(2, 3) = R("pelepa")
            .Cells(3, 3) = Now
            .Cells(4, 3) = EPIKEF
        End With

  On Error GoTo 0


    '----------------------- ����������� ----------------------------
    For k = 0 To n - 1
        myXL.Cells(5, k + 1) = F_T(k).Name
    Next

myXL.Rows(5).Font.Size = 14
myXL.Rows(5).Font.FontStyle = 12

'  With myXL
'     On Error GoTo 0
'       rows("5:5").Select
'       With selection.Font
'         .Name = "Arial"
'         .fontStyle = "������"
'         .Size = 12
'       End With
'  End With


Dim LAST_TIMH
Dim synt1
synt1 = IIf(IsNull("SYNT1"), "true", synt1)   ' ���� ������� ��� �n� ������� �������� �� synt1
AYJ = 5


Typose = 0
Dim row


Dim synola_SELIDOS
synola_SELIDOS = False

'--------------------------------------------------------
g_Stop = 1 'entos loop
Do While Not F_T.EOF
    AYJ = AYJ + 1
    If FF.CancelButton.Enabled = False Then
       Exit Do
    End If
        
    
    'If g_Stop = 2 Then
    '   Exit Do
    'End If
    
    'MDIForm1.Caption = AYJ
     '----------------- ������ ��� �� ����� ---------------------
    For k = 0 To n - 1

      If Left(F_T(k), 3) = "@@@" Then '  A/A
            myXL.Cells(AYJ, k + 1) = Right(Space(30) + Format(AYJ - 5, "######"), F_T(k).DefinedSize)
      ElseIf F_T(k).Type = 7 Or F_T(k).Type = 4 Or F_T(k).Type = 5 Or F_T(k).Type = 3 Or F_T(k).Type = 131 Then   'IsNumeric(f_t(K)) ����������
         If IsNull(F_T(k)) Then
             myXL.Cells(AYJ, k + 1) = Right(Space(30) + Format(0, "######,##0.00"), 13)  '  F_T(K).DefinedSize)
         Else
             myXL.Cells(AYJ, k + 1) = Right(Space(30) + Format(F_T(k), "######,##0.00"), 13) '  F_T(K).DefinedSize)
         End If
      ElseIf F_T(k).Type = 8 Or F_T(k).Type = 135 Then   'DATE
         myXL.Cells(AYJ, k + 1) = Right(Space(30) + Format(F_T(k), "DD/MM/YYYY"), 10)
      Else ' 10 STRING

           If IsNull(F_T(k)) Then

           Else
             ' Print #1, Tab(f_tab(K)); to928(F_T(K));
             If k < n - 1 Then
                ' ��� �� ��� ������ ���� ������� �����
                myXL.Cells(AYJ, k + 1) = "'" + Left(F_T(k), f_tab(k + 1) - f_tab(k) - 1)
             Else
                On Error Resume Next
                myXL.Cells(AYJ, k + 1) = "'" + F_T(k) 'to928(F_T(K))
             End If
           End If

           If m_sthl_ektyp(k) > 2 Then
            ' If K = 1 Then Print #1,
           End If  'm_sthl_ektyp(K) > 2

       End If
       ' soymes---------------------
       If Mid$(SUgm_str, k + 1, 1) = "1" Or Mid$(SUgm_str, k + 1, 1) = "2" Then
               If IsNull(F_T(k)) Then

               Else
                  m_sumes(k) = m_sumes(k) + Val(F_T(k))
               End If
       End If ' mid$(sugm_str,k,1)
    Next

     ' Print #1,
      If GROUPN > 0 Then
         LAST_TIMH = F_T(GROUPN - 1)
      End If
    F_T.MoveNext
    If Not F_T.EOF Then
       If GROUPN > 0 Then
           If LAST_TIMH <> F_T(GROUPN - 1) Then
          '   AYJ = AYJ + 1
           End If
       End If
    End If

    If AYJ Mod 100 = 0 Then
      DoEvents
      FF.Caption = "������� " + Format(AYJ, "######")    ' + "/" + Format(F_T.RecordCount, "######")
    End If
    

Loop
Unload FF



g_Stop = 0 'adiaforo
'"
AYJ = AYJ + 1
'myXL.cells(AYJ, 1) = String$(LSYN - 2, "-")
'"
  ' Print #1, Chr(13)  ' 6/12/2007
'"
'   aa = f_kodik("sum_seltxt")
   Dim PARAM '
   'PARAM = IIf(aa = " ", "  ", macro(aa, 0))
   'pr_SUMselidas param
   'PRINT2_Xsumes "������ ������"
AYJ = AYJ + 1
   GoSub printSUM

'AppActivate returnValue

' On Error Resume Next


  myXL.Columns("A:K").Select
  myXL.Columns.AutoFit

MDIForm1.MousePointer = vbNormal
'excel.Visible = True


Excel.Visible = True
On Error Resume Next
Kill "C:\EKTYP.XLS"
  DoEvents


  myXL.SaveAs "C:\EKTYP.XLS"


Dim ans3 As Long


ans3 = MsgBox("������ �� EXCEL", vbYesNo)
If ans3 = vbYes Then
 Call workbook.Close(False)
  Excel.Quit
  Set Excel = Nothing
End If






Exit Function


printSUM:

myXL.Rows(AYJ).Font.Size = 14
myXL.Rows(AYJ).Font.FontStyle = 12
For k = 0 To n - 1

    If m_sumes(k) > 0 Then
       If Mid$(SUgm_str, k + 1, 1) = "1" Then ' SUMA
           myXL.Cells(AYJ, k + 1) = Right(Space(30) + Format(m_sumes(k), "######,##0.00"), F_T(k).DefinedSize + 2)
       Else
           myXL.Cells(AYJ, k + 1) = Right(Space(30) + Format(m_sumes(k) / (AYJ - 7), "######,##0.00"), F_T(k).DefinedSize + 2)
       End If
       
    Else
         myXL.Cells(AYJ, k + 1) = "" ' Right(Space(50), F_T(K).DefinedSize);
    End If
Next

Return

printEpik:

For k = 0 To n - 1
    If m_sumes(k + 1) > 0 Then
         myXL.Cells(AYJ, k + 1) = Right(Space(30) + Format(m_sumes(k + 1), "######,##0.00"), F_T(k).DefinedSize + 2)
    Else
         myXL.Cells(AYJ, k + 1) = Right(Space(50), F_T(k).DefinedSize)
    End If
Next

Return







End Function



Function to437(string_ As String) As String
Dim a$, k As Integer, S As String, t As Integer, s928 As String, s437 As String
'metatrepei eggrafo apo 437->928
s928 = "������������������������-������������������������-��������"
s437 = "������������������������-������������������������-��������" ' saehioyv
's437 = "������������������������-������������������������" '�������"
'������������������������������������������������

a$ = string_
'                                                        saehioyv
'GoTo 11
'Open Text2.Text For Output As #2
'Open Text1.Text For Input As #1
'Do While Not EOF(1)
'  Line Input #1, a$
  For k = 1 To Len(a$)
     S = Mid(a$, k, 1)
     t = InStr(s928, S)
     If t > 0 Then
        Mid$(a$, k, 1) = Mid$(s437, t, 1)
     End If
  Next
11
  to437 = a$


End Function





Public Function Olografos(J)
'*******************************************************************
'* ������ ������ ��� ��� ���� ���������
Dim ekat, xil, MON, S, nn
 S = ""

 ekat = Int(J / 1000000)
If ekat > 0 Then
    If ekat = 1 Then
       S = "��� EKATOMY�I� "
    Else
       S = Tria_Olografos(ekat) + "EKATOMY�IA "
    End If
End If

 nn = J - ekat * 1000000
 xil = Int(nn / 1000)

If xil > 0 Then
   If xil = 1 Then
      S = S + "����A "
   Else
      S = S + Tria_Olografos(xil) + "�������� "
   End If
End If

 nn = nn - xil * 1000
 MON = Int(nn)

 nn = nn - MON

 If nn = 0 Then
    S = S + Tria_Olografos(MON) + " ����"
 Else
    S = S + Tria_Olografos(MON) + " ���� & " + LTrim(Str(Round(nn * 100, 0))) + " �����"
 End If
 
 Olografos = S
End Function  'return s

Public Function Tria_Olografos(n)
'"*******************************************************************
'"* ������ �������� ��� �� ���� ���������
Dim ek, dek, MON, strOL, nn
ek = Int(n / 100)
strOL = ""
Select Case ek
   Case 1
        strOL = "EKATO "
   Case 2
        strOL = "�������A "
   Case 3
        strOL = "��������A "
   Case 4
        strOL = "���������A "
   Case 5
        strOL = "���������A "
   Case 6
        strOL = "E������A "
   Case 7
        strOL = "��������A "
   Case 8
        strOL = "��������A "
   Case 9
        strOL = "���������A "
End Select
nn = n - ek * 100   '&& px  578 - 500 = 78
dek = Int(nn / 10)

Select Case dek
   Case 1
        If nn = 11 Then
            strOL = strOL + "������ "
            Tria_Olografos = strOL
        ElseIf nn = 12 Then
            strOL = strOL + "������ "
            Tria_Olografos = strOL
        Else
            strOL = strOL + "���� "
        End If
        
   Case 2
        strOL = strOL + "������ "
   Case 3
        strOL = strOL + "������� "
   Case 4
        strOL = strOL + "������� "
   Case 5
        strOL = strOL + "������� "
   Case 6
        strOL = strOL + "������ "
   Case 7
        strOL = strOL + "��������� "
   Case 8
        strOL = strOL + "������� "
   Case 9
        strOL = strOL + "�������� "
End Select

nn = nn - dek * 10  ' && px  78 - 70 = 8
MON = Int(nn)

Select Case MON
   Case 1
        strOL = strOL + "ENA "
   Case 2
        strOL = strOL + "��� "
   Case 3
        strOL = strOL + "��IA "
   Case 4
        strOL = strOL + "����E�A "
   Case 5
        strOL = strOL + "����� "
   Case 6
        strOL = strOL + "��� "
   Case 7
        strOL = strOL + "���� "
   Case 8
        strOL = strOL + "���� "
   Case 9
        strOL = strOL + "����� "
End Select

Tria_Olografos = strOL
End Function  '"return strOL
