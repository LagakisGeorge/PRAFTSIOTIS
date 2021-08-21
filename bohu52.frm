VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form bohu53 
   BackColor       =   &H00FF0000&
   Caption         =   "IMPORT ΕΙΔΩΝ"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13635
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8910
   ScaleWidth      =   13635
   WindowState     =   2  'Maximized
   Begin VB.CommandButton pel 
      Caption         =   "εισαγωγη εγγραφων απο ΕΛΤΑ"
      Height          =   495
      Left            =   195
      TabIndex        =   15
      Top             =   5595
      Width           =   5850
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ΚΑΤΑΧΩΡΗΣΗ ΧΩΡΙΩΝ"
      Height          =   555
      Left            =   180
      TabIndex        =   14
      Top             =   4620
      Width           =   5865
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Εξοδος"
      Height          =   555
      Left            =   7680
      TabIndex        =   10
      Top             =   6840
      Width           =   1470
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   4320
      TabIndex        =   9
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CheckBox Check2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      Caption         =   "Πρόσθεση νέων εγγραφών σε Κύριο αρχείο"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   150
      TabIndex        =   7
      Top             =   1005
      Value           =   1  'Checked
      Width           =   4815
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      Caption         =   "Ενημέρωση τιμών,περιγραφών από Excel"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   150
      TabIndex        =   6
      Top             =   645
      Value           =   1  'Checked
      Width           =   4815
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3780
      TabIndex        =   4
      Text            =   "2"
      Top             =   210
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Left            =   9960
      TabIndex        =   3
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Text            =   " "
      Top             =   480
      Width           =   4335
   End
   Begin MSComDlg.CommonDialog Cd1 
      Left            =   9600
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Καταχώρηση ΑΝΤΙΣΤΟΙΧΊΣΕΩΝ"
      Enabled         =   0   'False
      Height          =   375
      Left            =   135
      TabIndex        =   0
      Top             =   2520
      Width           =   5925
   End
   Begin VB.Label Label4 
      BackColor       =   &H000080FF&
      Height          =   450
      Left            =   150
      TabIndex        =   13
      Top             =   3105
      Width           =   2835
   End
   Begin VB.Label Label5 
      BackColor       =   &H000080FF&
      Height          =   450
      Left            =   3225
      TabIndex        =   12
      Top             =   3105
      Width           =   2835
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Πίνακας με πεδία"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   150
      TabIndex        =   11
      Top             =   3645
      Width           =   1965
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF0000&
      Caption         =   "Αν δεν αναφέρεται άλλο βάζω προεπιλεγμένη κατηγορία ΦΠΑ  την :"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   150
      TabIndex        =   8
      Top             =   1365
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF0000&
      Caption         =   "Εκκίνηση από την σειρά του Excel :"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   150
      TabIndex        =   5
      Top             =   165
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "Αρχείο Excel που περιέχει τον τιμοκατάλογο"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5640
      TabIndex        =   2
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "bohu53"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xl As New Excel.Application
Dim xlsheet As Excel.Worksheet
Dim xlsheet3 As Excel.Worksheet
Dim xlwbook As Excel.workbook
Private Sub Command1_Click()

Dim R As New ADODB.Recordset
Dim R0 As New ADODB.Recordset
Dim mok As Integer

Dim PIN(30, 2)

Dim COUNTER As Integer

Dim KOD_OK  As Boolean ' ΕΧΕΙ ΔΗΛΩΣΕΙ ΤΟΝ ΚΩΔΙΚΟ
Dim BARCODE_OK  As Boolean ' ΕΧΕΙ ΔΗΛΩΣΕΙ TA BARCODES
Dim sql As String
Dim NON_STOP As Boolean
NON_STOP = False
KOD_OK = False
BARCODE_OK = False
Dim k As Long

COUNTER = 1  ' ARIUMOS PEDION POY THA METAFERTHOYN
Dim KOD_COLUMN  As Integer
Dim BARCODE_COLUMN  As Integer
R0.Open "SELECT TOP 1 * FROM tk3 ", gdb, adOpenDynamic, adLockOptimistic
gdb.Execute "DELETE FROM tk3"
Set xlwbook = xl.Workbooks.Open(Text1.Text)
Set xlsheet = xlwbook.Sheets.Item(1)
Dim ko As String
Dim mNew As Long, mUpd As Long
mNew = 0
mUpd = 0
        Label4.Caption = "Nέες εγγραφές 0"
        Label5.Caption = "Ενημέρωση εγγραφές 0"
Dim mRow As Long
' data1.Recordset.MoveFirst
mRow = Val(Text2.Text)    '  data1.Recordset.Move
 On Error GoTo error_name
        
Do While True  ' Not xlsheet.cells(mRow, 1) = Null ' Not data1.Recordset.EOF
     If IsNull(xlsheet.Cells(mRow, 1)) Then
         Exit Do
     End If
     
     If IsEmpty(xlsheet.Cells(mRow, 1)) Then
         Exit Do
     End If
     
     
'     If R.EOF Then 'DEN BRHKA TON KODIKO
       'R.AddNew
            R0.AddNew
            R0("ARX") = Val(xlsheet.Cells(mRow, 1))
            R0("TEL") = Val(xlsheet.Cells(mRow, 2))
            R0("PRAKTOREIO") = xlsheet.Cells(mRow, 3)
            mNew = mNew + 1
            Label4.Caption = "Nέες εγγραφές " + Format(mNew, "######")
        R0.Update
        'R.Close
        DoEvents
        Me.Caption = ko
        mRow = mRow + 1 'data1.Recordset.MoveNext
Loop

xl.Quit
Set xlwbook = Nothing
Set xl = Nothing




MsgBox "τέλος ενημέρωσης"


Exit Sub

error_name:
If NON_STOP = False Then
   MsgBox "λάθος στην σειρά " + Format(mRow, "#######")
   If MsgBox("ΤΕΡΜΑΤΙΣΜΟΣ ;", vbYesNo) = vbYes Then
      Exit Sub
   End If
   If MsgBox("ΣΥΝΕΧΕΙΑ ΧΩΡΙΣ ΕΡΩΤΗΣΗ ;", vbYesNo) = vbYes Then
      NON_STOP = True
   End If
 End If





Resume Next
End Sub

Private Sub Command2_Click()
  
If Len(Trim(Text1.Text)) = 0 Then
  Cd1.ShowOpen
  Text1.Text = Cd1.FileName
Else
  If Len(Dir(LTrim(Text1.Text), vbNormal)) < 2 Then
     MsgBox "δεν υπάρχει το αρχείο " + Text1.Text
     Exit Sub
  End If
End If

On Error GoTo open_error
Me.MousePointer = vbHourglass
'Set xlwbook = xl.Workbooks.Open(Text1.Text)
'Set xlsheet = xlwbook.Sheets.Item(1)

On Error GoTo 0
Me.MousePointer = vbNormal
Command1.Enabled = True

Exit Sub


open_error:
MsgBox "λάθος στo  αρχείο " + Text1.Text
Me.MousePointer = vbNormal

Exit Sub

  
  
  
  
  
  
  
  
  
  
  
  
 ' data1.RecordSource = xlwbook.Sheets.Item(1).Name '   "Φύλλο1$"
  
' On Error Resume Next
'data1.Refresh

  
  
 ' data1.Recordset.Move 6
 ' Me.Caption = data1.Recordset(2)
End Sub

Private Sub Command3_Click()
  Unload Me
  
End Sub

Private Sub Command4_Click()

Dim R As New ADODB.Recordset
Dim R0 As New ADODB.Recordset
Dim mok As Integer

Dim PIN(30, 2)

Dim COUNTER As Integer

Dim KOD_OK  As Boolean ' ΕΧΕΙ ΔΗΛΩΣΕΙ ΤΟΝ ΚΩΔΙΚΟ
Dim BARCODE_OK  As Boolean ' ΕΧΕΙ ΔΗΛΩΣΕΙ TA BARCODES
Dim sql As String
Dim NON_STOP As Boolean
NON_STOP = False
KOD_OK = False
BARCODE_OK = False
Dim k As Long

COUNTER = 1  ' ARIUMOS PEDION POY THA METAFERTHOYN
Dim KOD_COLUMN  As Integer
Dim BARCODE_COLUMN  As Integer


R0.Open "SELECT TOP 1 * FROM tk ", gdb, adOpenDynamic, adLockOptimistic
gdb.Execute "DELETE FROM tk"
Set xlwbook = xl.Workbooks.Open(Text1.Text)
Set xlsheet = xlwbook.Sheets.Item(1)
Dim ko As String
Dim mNew As Long, mUpd As Long
mNew = 0
mUpd = 0
        Label4.Caption = "Nέες εγγραφές 0"
        Label5.Caption = "Ενημέρωση εγγραφές 0"
Dim mRow As Long
' data1.Recordset.MoveFirst
mRow = Val(Text2.Text)    '  data1.Recordset.Move
 On Error GoTo error_name
        
Do While True  ' Not xlsheet.cells(mRow, 1) = Null ' Not data1.Recordset.EOF
     If IsNull(xlsheet.Cells(mRow, 1)) Then
         Exit Do
     End If
     
     If IsEmpty(xlsheet.Cells(mRow, 1)) Then
         Exit Do
     End If
            R0.AddNew
            R0("xorio") = xlsheet.Cells(mRow, 1)
            R0("tk") = Replace(xlsheet.Cells(mRow, 2), " ", "")
            R0("nomos") = xlsheet.Cells(mRow, 3)
            'R0("") = Val(xlsheet.Cells(mRow, 2))
            
            mNew = mNew + 1
            Label4.Caption = "Nέες εγγραφές " + Format(mNew, "######")
        R0.Update
        'R.Close
        DoEvents
        Me.Caption = ko
        mRow = mRow + 1 'data1.Recordset.MoveNext
Loop

xl.Quit
Set xlwbook = Nothing
Set xl = Nothing




MsgBox "τέλος ενημέρωσης"


Exit Sub

error_name:
If NON_STOP = False Then
   MsgBox "λάθος στην σειρά " + Format(mRow, "#######")
   If MsgBox("ΤΕΡΜΑΤΙΣΜΟΣ ;", vbYesNo) = vbYes Then
      Exit Sub
   End If
   If MsgBox("ΣΥΝΕΧΕΙΑ ΧΩΡΙΣ ΕΡΩΤΗΣΗ ;", vbYesNo) = vbYes Then
      NON_STOP = True
   End If
 End If





Resume Next

End Sub

Private Sub Command5_Click()
'ΛΗΨΗ ΑΠΟ EXCEL
' ============================


Dim xl As New Excel.Application
Dim xlsheet As Excel.Worksheet
Dim xlsheet3 As Excel.Worksheet
Dim xlwbook As Excel.workbook









Dim DBF As Database
Dim sql As New ADODB.Connection
Dim rDBF As Recordset
Dim rSQL As New ADODB.Recordset
Dim rSQL2 As New ADODB.Recordset
Dim rYPOL As New ADODB.Recordset
Dim mYPOL As Single


Dim conDBF As String
Dim conSQL As String
Dim k As Long
Dim Fname As String

Dim db As DAO.Database

Dim arxeio As String, arxeio2 As String
Dim SynStoYpoloipo As Integer

Dim M_ATIM
M_ATIM = "Σ" + Format(Time(), "hhmmss")



'SynStoYpoloipo = MsgBox("Να προστεθεί στο ήδη υπάρχον υπόλοιπο;" + Chr(13) + "Σε περίπτωση που απαντήσετε όχι η ποσότητα που απογράψατε θα είναι και το υπόλοιπο", vbYesNoCancel)
'If SynStoYpoloipo = vbCancel Then
'    MsgBox "Η εργασία ακυρώθηκε"
'    Exit Sub
'End If

Dim BARC1STILI As Integer

' BARC1STILI = MsgBox("Η πρώτη στήλη είναι το BARCODE; " + Chr(13) + " Oxι σημαίνει ότι είναι ο κωδικός", vbYesNo)




'====================================================

'On Error GoTo open_error
Me.MousePointer = vbHourglass


If Len(Dir(Text1.Text, vbNormal)) > 0 Then
  ' OK
Else
   MsgBox " ΔΕΝ ΥΠΑΡΧΕΙ ΤΟ ΑΡΧΕΙΟ " + Text1.Text
   Exit Sub
End If



Set xlwbook = xl.Workbooks.Open(Text1.Text)
Set xlsheet = xlwbook.Sheets.Item(1)


  

'Set db = OpenDatabase(Dir1.Path, False, False, "dBase III;")

'Set rDBF = db.OpenRecordset("SELECT * FROM " + File1.TEXT1)





rSQL.Open "SELECT * FROM EGG WHERE YEAR(HME)=1999", gdb, adOpenDynamic, adLockOptimistic


Dim Z
Z = 0
k = 0
On Error GoTo WRITEERROR ' Resume Next

Dim OK As Boolean
Dim M_CODE
Dim RR As New ADODB.Recordset



Dim RR2 As New ADODB.Recordset

Dim TREX_YP As Single
Dim nn As Integer


' If IsNull(xlsheet.cells(mRow, KOD_COLUMN)) Then

Dim FOUND As Boolean




' mRow = 1
mRow = Val(Text2.Text)


Do While True  '   Not IsNull(xlsheet.cells(mRow, 1))
  
  
  
     If IsNull(xlsheet.Cells(mRow, 1)) Then
         Exit Do
     End If
     
     If IsEmpty(xlsheet.Cells(mRow, 1)) Then
         Exit Do
     End If
       
  
  OK = True
  
  M_CODE = xlsheet.Cells(mRow, 1)
        '  On Error Resume Next
        
       FOUND = True
        
       
         RR.Open "select * FROM PEL WHERE EIDOS='" + Left(Combo1.Text, 1) + "' and KOD='" + M_CODE + "'", gdb, adOpenDynamic, adLockOptimistic
         If RR.EOF Then
             FOUND = False
             MsgBox "ΔΕΝ ΒΡΕΘΗΚΕ O ΠΕΛΑΤΗΣ ΜΕ ΚΩΔΙΚΟ  " + M_CODE
             List2.AddItem "ΔΕΝ ΒΡΕΘΗΚΕ : " + M_CODE
             xlsheet.Cells(mRow, 4) = "ΔΕΝ ΒΡΕΘΗΚΕ"
         End If
         M_CODE = RR("KOD")
         RR.Close
         
       
       
If FOUND And Abs(xlsheet.Cells(mRow, 2)) > 0 Then
        
         rSQL.AddNew
          rSQL("eidos") = Left(Combo1.Text, 1)
          
          rSQL("ATIM") = "Σ00001" ' M_ATIM '"λ00002"
          rSQL("hme") = APOT.Value
          
          
          rSQL("ait") = xlsheet.Cells(mRow, 3)
          rSQL("KOD") = M_CODE
          
          rSQL("XRE") = xlsheet.Cells(mRow, 2)
       If xlsheet.Cells(mRow, 2) > 0 Then
          rSQL("XREOSI") = xlsheet.Cells(mRow, 2)
       Else
          rSQL("PISTOSI") = -xlsheet.Cells(mRow, 2)
       End If
       
          rSQL.Update
End If
         
         mRow = mRow + 1
         If mRow Mod 10 = 0 Then
            Me.Caption = mRow
         End If


         DoEvents
Loop

rSQL.Close
          

If List2.ListCount > 0 Then
    MsgBox "ΤΑ ΛΑΘΗ ΑΠΟΘΗΚΕΥΤΗΚΑΝ ΣΤΟ EXCEL  "
    xlwbook.Save  '  "c:\ektyp2.xls"
End If
  
  
 

  
  
  Set xlsheet = Nothing
  Set xlwbook = Nothing
' excel.Quit

xl.Quit





MsgBox "Ενημερώθηκαν " + Format(mRow, "###0") + " εγγραφές"


Exit Sub


WRITEERROR:
Resume Next


End Sub

Private Sub Form_Load()
Dim R As New ADODB.Recordset

'   R.Open "SELECT *FROM PINAKES WHERE TYPOS=1 ORDER BY AYJON", gdb, adOpenDynamic, adLockOptimistic
   

'FPA
'Do While Not R.EOF
'   If R("typos") = 1 Then
'      Combo2.AddItem Str(R("AYJON")) + " -> " + Str(R("TIMH"))
'   End If
'   R.MoveNext
'Loop
' R.Close

'Combo2.Text = Combo2.List(1)

'Combo1.Text = Combo1.List(0)

End Sub

Private Sub pel_Click()


Dim R As New ADODB.Recordset
Dim R0 As New ADODB.Recordset
Dim mok As Integer

Dim PIN(30, 2)

Dim COUNTER As Integer

Dim KOD_OK  As Boolean ' ΕΧΕΙ ΔΗΛΩΣΕΙ ΤΟΝ ΚΩΔΙΚΟ
Dim BARCODE_OK  As Boolean ' ΕΧΕΙ ΔΗΛΩΣΕΙ TA BARCODES
Dim sql As String
Dim NON_STOP As Boolean
NON_STOP = False
KOD_OK = False
BARCODE_OK = False
Dim k As Long

COUNTER = 1  ' ARIUMOS PEDION POY THA METAFERTHOYN
Dim KOD_COLUMN  As Integer
Dim BARCODE_COLUMN  As Integer
R0.Open "SELECT TOP 1 * FROM pel ", gdb, adOpenDynamic, adLockOptimistic

Set xlwbook = xl.Workbooks.Open(Text1.Text)
Set xlsheet = xlwbook.Sheets.Item(1)
Dim ko As String
Dim mNew As Long, mUpd As Long
mNew = 0
mUpd = 0
        Label4.Caption = "Nέες εγγραφές 0"
        Label5.Caption = "Ενημέρωση εγγραφές 0"
Dim mRow As Long
' data1.Recordset.MoveFirst
mRow = Val(Text2.Text)    '  data1.Recordset.Move
 On Error GoTo error_name
        
        Dim MC1 As String
        
        
Do While True  ' Not xlsheet.cells(mRow, 1) = Null ' Not data1.Recordset.EOF
     If IsNull(xlsheet.Cells(mRow, 1)) Then
         Exit Do
     End If
     
     If IsEmpty(xlsheet.Cells(mRow, 1)) Then
         Exit Do
     End If
     
     
'     If R.EOF Then 'DEN BRHKA TON KODIKO
       'R.AddNew
'            R0.AddNew
            MC1 = Trim(xlsheet.Cells(mRow, 2))
            gdb.Execute "insert into pel (c1) values ('" + MC1 + "')"
            gdb.Execute "update pel set d1='" + Format(xlsheet.Cells(mRow, 1), "MM/DD/YYYY") + "' WHERE c1='" + MC1 + "'"
            gdb.Execute "update pel set epo='" + Left(xlsheet.Cells(mRow, 6), 35) + "' WHERE c1='" + MC1 + "'"
            
            gdb.Execute "update pel set n2=" + Str(xlsheet.Cells(mRow, 9)) + "  WHERE c1='" + MC1 + "'"
            gdb.Execute "update pel set thl='" + LTrim(Str(xlsheet.Cells(mRow, 10))) + "'  WHERE c1='" + MC1 + "'"
            gdb.Execute "update pel set n1=" + Str(xlsheet.Cells(mRow, 12)) + "  WHERE c1='" + MC1 + "'"
            gdb.Execute "update pel set die='" + xlsheet.Cells(mRow, 14) + "'  WHERE c1='" + MC1 + "'"
              gdb.Execute "update pel set praktoreio='" + xlsheet.Cells(mRow, 15) + "'  WHERE c1='" + MC1 + "'"
           gdb.Execute "update pel set tk='" + LTrim(Str(xlsheet.Cells(mRow, 7))) + "' WHERE c1='" + MC1 + "'"
              
              
              
              
           ' R0("d1") = xlsheet.Cells(mRow, 1)
           ' R0("c1") = xlsheet.Cells(mRow, 2)
           ' R0("epo") = Left(xlsheet.Cells(mRow, 6), 35)
           ' R0("tk") = Val(xlsheet.Cells(mRow, 7))
           ' R0("n2") = Val(xlsheet.Cells(mRow, 9))
           '  R0("thl") = xlsheet.Cells(mRow, 10)
          ' R0("n1") = Val(xlsheet.Cells(mRow, 12))
            ' R0("c1") = Val(xlsheet.Cells(mRow, 2))
            ' R0("epo") = xlsheet.Cells(mRow, 6)
            
            
            
            
            
            
            
            mNew = mNew + 1
            Label4.Caption = "Nέες εγγραφές " + Format(mNew, "######")
'        R0.Update
        'R.Close
        DoEvents
        Me.Caption = ko
        mRow = mRow + 1 'data1.Recordset.MoveNext
Loop

xl.Quit
Set xlwbook = Nothing
Set xl = Nothing




MsgBox "τέλος ενημέρωσης"


Exit Sub

error_name:
If NON_STOP = False Then
   MsgBox "λάθος στην σειρά " + Format(mRow, "#######")
   If MsgBox("ΤΕΡΜΑΤΙΣΜΟΣ ;", vbYesNo) = vbYes Then
      Exit Sub
   End If
   If MsgBox("ΣΥΝΕΧΕΙΑ ΧΩΡΙΣ ΕΡΩΤΗΣΗ ;", vbYesNo) = vbYes Then
      NON_STOP = True
   End If
 End If





Resume Next

























End Sub
