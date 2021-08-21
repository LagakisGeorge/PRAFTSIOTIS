VERSION 5.00
Begin VB.Form epist 
   BackColor       =   &H00FF0000&
   Caption         =   "Form5"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1380
      Left            =   6870
      TabIndex        =   6
      Top             =   450
      Width           =   3225
      Begin VB.OptionButton Option2 
         Caption         =   "Επιστροφές"
         Height          =   345
         Left            =   300
         TabIndex        =   8
         Top             =   765
         Width           =   2745
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Πληρωμές"
         Height          =   300
         Left            =   285
         TabIndex        =   7
         Top             =   315
         Value           =   -1  'True
         Width           =   2535
      End
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3075
      TabIndex        =   0
      Top             =   1305
      Width           =   1875
   End
   Begin VB.ListBox List1 
      Height          =   6300
      ItemData        =   "Form5.frx":0000
      Left            =   510
      List            =   "Form5.frx":0002
      TabIndex        =   3
      Top             =   2625
      Width           =   8310
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3075
      TabIndex        =   2
      Top             =   585
      Width           =   1875
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   825
      TabIndex        =   5
      Top             =   2070
      Width           =   6810
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Μεταφορικά"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   780
      TabIndex        =   4
      Top             =   1350
      Width           =   2130
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PP BARCODE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   780
      TabIndex        =   1
      Top             =   630
      Width           =   2130
   End
End
Attribute VB_Name = "epist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fR As New ADODB.Recordset

Private Sub Form_Load()

  'gConnect = "DSN=ELTA;UID=sa;pwd=;"
  Set gdb = New Connection
  gdb.CursorLocation = adUseClient
  gdb.Open gConnect   ' "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=C:\salonika\taxypliromes.mdb;"




End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
     SendKeys "{TAB}"
  End If
   
  


End Sub

Private Sub Text1_LostFocus()

If Len(Text1.Text) = 0 Then
   Text1.SetFocus
   Exit Sub
End If




   fR.Open "select * from pel where c1='" + Text1.Text + "'", gdb, adOpenDynamic, adLockOptimistic
   If fR.RecordCount = 0 Then
       MsgBox "ΔΕΝ ΒΡΕΘΗΚΕ ΤΟ ΡΡ"
       Text1.SetFocus
       fR.Close
       Text1.Text = ""
       Exit Sub
   End If
   
'   If Val(Text2.Text) = 0 Then
'       MsgBox "ΔΕΝ ΔΟΘΗΚΕ ΑΞΙΑ ΜΕΤΑΦΟΡΙΚΩΝ"  NB220044244GRNB220044244GR


'       Text2.SetFocus
'       fR.Close
'       Text1.Text = ""
'       Exit Sub
'
'   End If
   
   
   If fR("metaforika") > 0 Then
      If fR("pliromi") = 1 Then
        MsgBox "πληρωθηκε " + Format(fR("d3"), "dd/mm/yyyy")
      Else
         MsgBox "Eπέστρεψε απλήρωτο  " + Format(fR("d3"), "dd/mm/yyyy")
      End If
      Text1.Text = ""
       Text1.SetFocus
       fR.Close
       Exit Sub
   End If
   
   
   
   
   If Left(fR("cpliromi"), 5) = "ΑΚΥΡΟ" Then
      MsgBox "ΕΧΕΙ ΗΔΗ ΑΚΥΡΩΘΕΙ  " + Format(fR("d3"), "dd/mm/yyyy")
      Text1.Text = ""
       Text1.SetFocus
       fR.Close
       Exit Sub
   End If
   
   
   
   
   
   
   Dim n2
   n2 = List1.ListCount
   
   
   List1.AddItem Format(n2 + 1, "000") + " " + fR("C1") + " " + fR("EPO"), 0
       
  Dim n As Long
  If Option1.Value = True Then
      plir = "1"
      cplir = "PAID"
  Else
      plir = "2"
      cplir = "ΑΚΥΡΟ"
  End If
  
  Text2.Text = Replace(Text2.Text, ",", ".")
    
  gdb.Execute "update pel set cpliromi='" + cplir + "',d3=GETDATE(),metaforika =" + Str(Val(Text2.Text)) + ",pliromi=" + plir + " where c1='" + Text1.Text + "'", n
  
       
  If n = 0 Then
      MsgBox "ΔΕΝ ΚΑΤΑΧΩΡΗΘΗΚΕ"
  End If
  
fR.Close

       
       Text1.Text = ""
       
       
       
       
   Text1.SetFocus
       
       
       
       
End Sub

Private Sub Text2_GotFocus()
   Text2.BackColor = vbYellow
   
End Sub

Private Sub Text2_LostFocus()
   Text2.BackColor = vbWhite
   
End Sub
