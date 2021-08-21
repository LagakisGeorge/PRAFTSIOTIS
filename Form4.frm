VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   10305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15360
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form4"
   Picture         =   "Form4.frx":0000
   ScaleHeight     =   10305
   ScaleWidth      =   15360
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FF0000&
      Caption         =   "M¸ÌÔ email "
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   12285
      TabIndex        =   18
      Top             =   2745
      Width           =   2520
   End
   Begin RichTextLib.RichTextBox TEXT2 
      Height          =   2145
      Left            =   6465
      TabIndex        =   17
      Top             =   90
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   3784
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form4.frx":57E442
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Word"
      Height          =   330
      Left            =   10380
      TabIndex        =   16
      Top             =   2715
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Excel"
      Height          =   315
      Left            =   10380
      TabIndex        =   15
      Top             =   3180
      Value           =   1  'Checked
      Width           =   1590
   End
   Begin VB.TextBox c2 
      Height          =   315
      Left            =   8490
      TabIndex        =   13
      ToolTipText     =   "c2"
      Top             =   3150
      Width           =   1605
   End
   Begin VB.TextBox c1 
      Height          =   315
      Left            =   8490
      TabIndex        =   11
      ToolTipText     =   "c1"
      Top             =   2685
      Width           =   1605
   End
   Begin MSComCtl2.DTPicker d1 
      Height          =   285
      Left            =   8490
      TabIndex        =   7
      Top             =   3615
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      Format          =   16449537
      CurrentDate     =   40198
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   5220
      Top             =   10200
      Visible         =   0   'False
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form4.frx":57E4C6
      Height          =   4995
      Left            =   195
      TabIndex        =   6
      Top             =   4710
      Width           =   14940
      _ExtentX        =   26353
      _ExtentY        =   8811
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1032
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1032
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "External DB"
      Height          =   360
      Left            =   60
      TabIndex        =   5
      Top             =   9930
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.CommandButton Command1 
      Caption         =   "≈ ‘≈À≈”«"
      Height          =   675
      Left            =   10365
      TabIndex        =   4
      Top             =   3630
      Width           =   1590
   End
   Begin VB.CommandButton Command3 
      Caption         =   "EXECUTE SQL"
      Height          =   315
      Left            =   3600
      TabIndex        =   3
      Top             =   9945
      Width           =   1695
   End
   Begin VB.FileListBox File1 
      Height          =   4380
      Left            =   195
      TabIndex        =   2
      Top             =   120
      Width           =   5955
   End
   Begin VB.CommandButton Command11 
      Caption         =   "¡ÔËﬁÍÂıÛÁ"
      Height          =   255
      Left            =   11145
      TabIndex        =   1
      Top             =   2340
      Width           =   3105
   End
   Begin VB.CheckBox ODBC 
      Caption         =   "ODBC"
      Height          =   240
      Left            =   1965
      TabIndex        =   0
      Top             =   9975
      Visible         =   0   'False
      Width           =   1440
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker d2 
      Height          =   285
      Left            =   8490
      TabIndex        =   8
      Top             =   4035
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      Format          =   16449537
      CurrentDate     =   40198
   End
   Begin VB.Label Label4 
      Caption         =   "–·Ò·ÏÂÙÒÔÚ 2"
      Height          =   330
      Left            =   6885
      TabIndex        =   14
      Top             =   3165
      Width           =   1530
   End
   Begin VB.Label Label3 
      Caption         =   "–·Ò·ÏÂÙÒÔÚ 1"
      Height          =   330
      Left            =   6885
      TabIndex        =   12
      Top             =   2700
      Width           =   1530
   End
   Begin VB.Label Label2 
      Caption         =   "≈˘Ú"
      Height          =   285
      Left            =   6885
      TabIndex        =   10
      Top             =   4065
      Width           =   1530
   End
   Begin VB.Label Label1 
      Caption         =   "¡¸"
      Height          =   315
      Left            =   6885
      TabIndex        =   9
      Top             =   3645
      Width           =   1530
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fSQL As String


Private Sub Command1_Click()



Dim DBF As New ADODB.Connection
Dim sql As New ADODB.Connection
Dim rDBF As New ADODB.Recordset
Dim rSQL As New ADODB.Recordset




Dim conDBF As String
Dim conSQL As String
On Error Resume Next

Me.MousePointer = vbHourglass
Dim n As Long
n = 0  'GetCurrentTime()
'Dim sql As String


fSQL = Replace(Text2.Text, "@d1", Format(d1, "mm/dd/yyyy"))
fSQL = Replace(fSQL, "@d2", Format(DateAdd("d", 1, d2), "mm/dd/yyyy"))
fSQL = Replace(fSQL, "@c1", c1.Text)











If ODBC.Value = vbChecked Then
  ' DBGrid1.Visible = True
   Data2.Connect = "ODBC;" + gConnect
   Data2.RecordSource = fSQL
   Data2.Refresh
Else
  ' DBGrid1.Visible = False
   Adodc2.ConnectionString = gConnect
   Adodc2.RecordSource = fSQL
   Adodc2.Refresh
End If



DataGrid1.Refresh


'Me.Caption = (GetCurrentTime() - n) / 1000

'TDBGrid.AlternatingRowStyle = True

'TDBGrid.OddRowStyle.BackColor = &H8000000F   ' GRI   vbCyan
'TDBGrid.EvenRowStyle.BackColor = &HFFFFC0   'OYRANI     &H8000000F  ' GRI


If Check1.Value = vbChecked Then
   print7_excel fSQL, "011111111", "", 1
End If





Me.MousePointer = vbNormal

End Sub

Private Sub Command11_Click()
 CD1.InitDir = "c:\salonika\queries"
  
  CD1.ShowSave
  
  Dim f
  f = CD1.FileName
  
Open f For Output As #5
    Print #5, Text2.Text
Close #5
  
  
  '+ ".txt"
  
End Sub

Private Sub Command12_Click()
  print7_excel fSQL, "011111111", "", 1
End Sub

Private Sub Command2_Click()
On Error Resume Next
gdb.Close
gConnect = InputBox("ƒŸ”≈ Õ≈œ CONNECTION STRING ")




  gdb.Open gConnect

End Sub

Private Sub Command3_Click()
Dim DBF As New ADODB.Connection
Dim sql As New ADODB.Connection
Dim rDBF As New ADODB.Recordset
Dim rSQL As New ADODB.Recordset



Dim conDBF As String
Dim conSQL As String
Dim LO As Long
Dim db As Database

Dim SEIRES(30)

Dim MLINE As String, k

Dim n As Long

  For k = 1 To 30: SEIRES(k) = "": Next



LO = 0
Me.MousePointer = vbHourglass
   On Error GoTo LATOS  'On Error Resume Next


If ODBC.Value = Checked Then

   Set db = OpenDatabase("", False, False, gConnect)
   db.Execute Text2.Text
   LO = db.RecordsAffected
Else
MLINE = Text2.Text

 '  DUM = FETES2_DELIM(MLINE, SEIRES)
   

' For k = 1 To 30
'   If Len(SEIRES(k)) > 2 Then
'       gdb.Execute Trim(SEIRES(k)), LO
'   End If
' Next
 
 gdb.Execute MLINE, LO
 
 
 'Me.Caption = (GetCurrentTime() - n) / 1000
End If
   MsgBox Str(LO) + " ≈√√—¡÷≈” ≈Õ«Ã≈—Ÿ»« ¡Õ"
   Me.MousePointer = vbNormal


Exit Sub

LATOS:
MsgBox Err.Description
Resume Next

End Sub

Private Sub Command4_Click()

Adodc2.Recordset("ANAZHTHSH") = 1
Adodc2.Recordset.Update







Dim appWord As Word.Application
Dim wrdDoc As Word.Document
Dim strFileName As String
strFileName = "c:\salonika\ANAZHTHSHA.doc"
Set appWord = New Word.Application
Set wrdDoc = appWord.Documents.Open(strFileName)
'MsgBox wrdDoc.Path & "\" & wrdDoc.Name


 'Insert a paragraph at the beginning of the document.
        ' wrdDoc.Paragraphs(1).Range.Text = "LAGAKIS"

wrdDoc.Tables(1).Cell(1, 1).Range.Text = wrdDoc.Tables(1).Cell(1, 1).Range.Text + Format(Now, "DD/MM/YYYY")

wrdDoc.Tables(1).Cell(2, 2).Range.Text = Chr(13) + "PP" + Adodc2.Recordset(2).Value + "GR" 'pp
wrdDoc.Tables(1).Cell(2, 2).Range.Font.Bold = True


wrdDoc.Tables(1).Cell(2, 3).Range.Text = Chr(13) + Format(Adodc2.Recordset("n2").Value, "###0")


wrdDoc.Tables(1).Cell(2, 4).Range.Text = Chr(13) + Format(Adodc2.Recordset(8).Value, "###0.00") + "Ä"
wrdDoc.Tables(1).Cell(2, 4).Range.Font.Bold = True

wrdDoc.Tables(1).Cell(4, 2).Range.Text = wrdDoc.Tables(1).Cell(4, 2).Range.Text + Chr(13) + Adodc2.Recordset(1).Value

wrdDoc.Tables(1).Cell(5, 2).Range.Text = wrdDoc.Tables(1).Cell(5, 2).Range.Text + Chr(13) + Adodc2.Recordset(3).Value + Chr(13) + Adodc2.Recordset(4).Value + Chr(13) + Adodc2.Recordset(5).Value + "  " + Adodc2.Recordset("tk").Value + Chr(13) + Adodc2.Recordset(6).Value







'Dim mobjWORD As Word.Document

'Start a new document in Word
 '   Set mWordA = CreateObject("Word.Application")
  '  Set mobjWORD = mWordA.Documents.Add

'wrdDoc.Paragraphs (0)





'wrdDoc.Save

If Check2.Value = vbChecked Then
  wrdDoc.SaveAs "C:\salonika\email-ANAZHTHSH\" + Adodc2.Recordset(2)
Else
  wrdDoc.PrintOut
End If



wrdDoc.Close False
appWord.Quit
Set wrdDoc = Nothing
Set appWord = Nothing


Adodc2.Refresh




End Sub

Sub print_word_form()

Dim appWord As Word.Application
Dim wrdDoc As Word.Document
Dim strFileName As String
strFileName = "c:\salonika\ANAZHTHSHA.doc"
Set appWord = New Word.Application
Set wrdDoc = appWord.Documents.Open(strFileName)
'MsgBox wrdDoc.Path & "\" & wrdDoc.Name


 'Insert a paragraph at the beginning of the document.
        ' wrdDoc.Paragraphs(1).Range.Text = "LAGAKIS"
        
wrdDoc.Tables(1).Cell(1, 1).Range.Text = wrdDoc.Tables(1).Cell(1, 1).Range.Text + Format(Now, "DD/MM/YYYY")
wrdDoc.Tables(1).Cell(1, 2).Range.Text = wrdDoc.Tables(1).Cell(1, 2).Range.Text + "2,2"
wrdDoc.Tables(1).Cell(1, 3).Range.Text = wrdDoc.Tables(1).Cell(1, 3).Range.Text + "2,2"
wrdDoc.Tables(1).Cell(1, 4).Range.Text = wrdDoc.Tables(1).Cell(1, 4).Range.Text + "2,2"
wrdDoc.Tables(1).Cell(1, 5).Range.Text = wrdDoc.Tables(1).Cell(1, 5).Range.Text + "2,2"

wrdDoc.Tables(1).Cell(2, 3).Range.Text = wrdDoc.Tables(1).Cell(2, 1).Range.Text + "2,1"
wrdDoc.Tables(1).Cell(3, 1).Range.Text = "3,1"
wrdDoc.Tables(1).Cell(4, 1).Range.Text = "3,1"
wrdDoc.Tables(1).Cell(5, 1).Range.Text = "3,1"
wrdDoc.Tables(1).Cell(6, 1).Range.Text = "3,1"


'wrdDoc.Save
wrdDoc.PrintOut



wrdDoc.Close False
appWord.Quit
Set wrdDoc = Nothing
Set appWord = Nothing







































End Sub



Private Sub File1_Click()

  Dim a
  a = 0

Dim f As String

f = File1.FileName


Dim ss As String
Dim b As String

ss = ""

Open File1.Path + "\" + f For Input As #1
Do While Not EOF(1)
    
    Line Input #1, b
    ss = ss + b + Chr(13)
Loop
Close #1

'Text2.width = 4800
Text2.Text = ss






End Sub

Private Sub Form_Load()
 File1.Path = "C:\salonika\queries"
 d1.Value = Now
 d2.Value = Now
 

End Sub
