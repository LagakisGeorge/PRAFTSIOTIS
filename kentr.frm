VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   10515
   ClientLeft      =   240
   ClientTop       =   750
   ClientWidth     =   15240
   LinkTopic       =   "MDIForm1"
   Picture         =   "kentr.frx":0000
   Begin VB.Menu Arx 
      Caption         =   "������"
      Begin VB.Menu p1 
         Caption         =   "E������� ���������"
      End
      Begin VB.Menu p2 
         Caption         =   "���������� ������� ����"
      End
      Begin VB.Menu ex 
         Caption         =   "������"
      End
   End
   Begin VB.Menu ��� 
      Caption         =   "��������"
      Begin VB.Menu ep1 
         Caption         =   "��������/����������"
      End
   End
   Begin VB.Menu bohu 
      Caption         =   "��������� ������"
      Begin VB.Menu b1 
         Caption         =   "N��� �������"
      End
      Begin VB.Menu d2 
         Caption         =   "������� ����� ���������"
      End
      Begin VB.Menu newprakt 
         Caption         =   "�������� ���� ������� �����������"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b1_Click()
 form1.Show
End Sub

Private Sub d2_Click()
   Form4.Show
End Sub
Private Sub ep1_Click()
   epist.Show
End Sub
Private Sub ex_Click()
  End
End Sub
Private Sub MDIForm_Load()
Dim a$
Open "c:\mercpath.txt" For Input As #1
Input #1, a$
Close #1


  gConnect = a$  '"DSN=ELTA;UID=sa;pwd=;"
  Set gdb = New Connection
  gdb.CursorLocation = adUseClient
  gdb.Open gConnect   ' "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=C:\salonika\taxypliromes.mdb;"
















End Sub

Private Sub newprakt_Click()
   bohu53.Show
   
End Sub

Private Sub p1_Click()
  fORM2.Show
End Sub

Private Sub p2_Click()
   Form3.Show
End Sub

