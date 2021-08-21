VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form fORM2 
   BackColor       =   &H00800000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "pel"
   ClientHeight    =   10680
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   18315
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "Form2N.frx":0000
   ScaleHeight     =   10680
   ScaleWidth      =   18315
   Begin VB.TextBox TK 
      Height          =   300
      Left            =   6120
      TabIndex        =   53
      Top             =   1335
      Width           =   1305
   End
   Begin VB.TextBox txtFields 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   3555
      TabIndex        =   52
      Top             =   1650
      Width           =   4470
   End
   Begin VB.TextBox txtFields 
      DataField       =   "d1"
      Height          =   285
      Index           =   7
      Left            =   9195
      TabIndex        =   50
      Top             =   2460
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.CheckBox PROTYP 
      Caption         =   "ΠΡΟΤYΠΩΣΗ ΑΠΟΣΤΟΛΕΑ"
      Height          =   240
      Left            =   12585
      TabIndex        =   48
      Top             =   4815
      Width           =   2640
   End
   Begin VB.CheckBox APOSTOLEAS 
      Caption         =   "EKTYΠΩΣΗ ΑΠΟΣΤΟΛΕΑ"
      Height          =   240
      Left            =   12585
      TabIndex        =   47
      Top             =   4410
      Width           =   2640
   End
   Begin VB.TextBox txtFields 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   8
      Left            =   6720
      TabIndex        =   45
      Top             =   15
      Width           =   1335
   End
   Begin VB.CheckBox Check2 
      Caption         =   "ΤΑΧΥΠΛΗΡΩΜΕΣ"
      Height          =   240
      Left            =   11535
      TabIndex        =   44
      Top             =   1635
      Width           =   2340
   End
   Begin VB.CheckBox Check1 
      Caption         =   "ΑΝΤΙΚΑΤΑΒΟΛΕΣ"
      Height          =   240
      Left            =   11520
      TabIndex        =   43
      Top             =   1095
      Value           =   1  'Checked
      Width           =   2340
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\salonika\taxypliromes.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   9000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "REP1"
      Top             =   8520
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00008000&
      Caption         =   "Εκτύπωση VAUCHER"
      Height          =   615
      Left            =   9045
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   4410
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "Εκτύπωση Ατικαταβολης"
      Height          =   555
      Left            =   9465
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5145
      Width           =   1575
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   10
      Left            =   8445
      TabIndex        =   39
      Top             =   1995
      Width           =   5325
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   615
      Left            =   5160
      Top             =   9240
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1085
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
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=ELTA"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "ELTA"
      OtherAttributes =   ""
      UserName        =   "sa"
      Password        =   ""
      RecordSource    =   $"Form2N.frx":57E442
      Caption         =   "Adodc2"
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
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   5640
      TabIndex        =   34
      Top             =   5280
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3120
      TabIndex        =   33
      Top             =   5280
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   480
      TabIndex        =   32
      Top             =   5280
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Form2N.frx":57E4EF
      Height          =   3615
      Left            =   480
      TabIndex        =   31
      Top             =   5760
      Width           =   17325
      _ExtentX        =   30559
      _ExtentY        =   6376
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "d1"
         Caption         =   "d1"
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
         DataField       =   "c1"
         Caption         =   "c1"
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
      BeginProperty Column02 
         DataField       =   "epo"
         Caption         =   "epo"
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
      BeginProperty Column03 
         DataField       =   "die"
         Caption         =   "die"
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
      BeginProperty Column04 
         DataField       =   "polh"
         Caption         =   "polh"
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
      BeginProperty Column05 
         DataField       =   "tk"
         Caption         =   "tk"
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
      BeginProperty Column06 
         DataField       =   "praktoreio"
         Caption         =   "praktoreio"
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
      BeginProperty Column07 
         DataField       =   "thl"
         Caption         =   "thl"
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
      BeginProperty Column08 
         DataField       =   "n1"
         Caption         =   "n1"
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
      BeginProperty Column09 
         DataField       =   "PAID"
         Caption         =   "PAID"
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
      BeginProperty Column10 
         DataField       =   "DATE"
         Caption         =   "DATE"
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
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
         EndProperty
         BeginProperty Column09 
         EndProperty
         BeginProperty Column10 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtFields 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   5
      Left            =   2025
      TabIndex        =   1
      Top             =   0
      Width           =   3375
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   315
      Top             =   9510
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "C:\salonika\plir.rpt"
      Destination     =   1
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox txtFields 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   6
      Left            =   8460
      TabIndex        =   11
      Top             =   1215
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form2N.frx":57E504
      Height          =   1455
      Left            =   9060
      TabIndex        =   28
      Top             =   2865
      Width           =   5520
      _ExtentX        =   9737
      _ExtentY        =   2566
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      ColumnCount     =   15
      BeginProperty Column00 
         DataField       =   "epo"
         Caption         =   "epo"
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
         DataField       =   "die"
         Caption         =   "die"
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
      BeginProperty Column02 
         DataField       =   "tk"
         Caption         =   "tk"
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
      BeginProperty Column03 
         DataField       =   "polh"
         Caption         =   "polh"
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
      BeginProperty Column04 
         DataField       =   "thl"
         Caption         =   "thl"
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
      BeginProperty Column05 
         DataField       =   "afm"
         Caption         =   "afm"
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
      BeginProperty Column06 
         DataField       =   "c1"
         Caption         =   "c1"
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
      BeginProperty Column07 
         DataField       =   "c2"
         Caption         =   "c2"
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
      BeginProperty Column08 
         DataField       =   "n1"
         Caption         =   "n1"
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
      BeginProperty Column09 
         DataField       =   "n2"
         Caption         =   "n2"
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
      BeginProperty Column10 
         DataField       =   "d1"
         Caption         =   "d1"
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
      BeginProperty Column11 
         DataField       =   "d2"
         Caption         =   "d2"
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
      BeginProperty Column12 
         DataField       =   "id"
         Caption         =   "id"
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
      BeginProperty Column13 
         DataField       =   "c3"
         Caption         =   "c3"
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
      BeginProperty Column14 
         DataField       =   "c4"
         Caption         =   "c4"
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
         EndProperty
         BeginProperty Column09 
         EndProperty
         BeginProperty Column10 
         EndProperty
         BeginProperty Column11 
         EndProperty
         BeginProperty Column12 
         EndProperty
         BeginProperty Column13 
         EndProperty
         BeginProperty Column14 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Εκτύπωση Ταχυπληρωμής"
      Enabled         =   0   'False
      Height          =   555
      Left            =   11415
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5130
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1440
      Top             =   8520
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=ELTA"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "ELTA"
      OtherAttributes =   ""
      UserName        =   "sa"
      Password        =   ""
      RecordSource    =   "select * from prom"
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
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Διόρθωση"
      Height          =   300
      Left            =   6465
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   9375
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Νέα Εγγραφή"
      Height          =   300
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Διαγραφή"
      Height          =   300
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Ανανέωση"
      Height          =   300
      Left            =   11265
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   9375
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Εξοδος"
      Height          =   780
      Left            =   16140
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4800
      Width           =   1590
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Ακύρωση"
      Height          =   300
      Left            =   9105
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   9375
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Ενημέρωση"
      Height          =   300
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2520
      Width           =   1095
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   18315
      TabIndex        =   20
      Top             =   10080
      Width           =   18315
   End
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   18315
      TabIndex        =   14
      Top             =   10380
      Width           =   18315
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "Form2N.frx":57E519
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Picture         =   "Form2N.frx":57E85B
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "Form2N.frx":57EB9D
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "Form2N.frx":57EEDF
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   19
         Top             =   0
         Width           =   3360
      End
   End
   Begin VB.TextBox txtFields 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   4
      Left            =   2025
      TabIndex        =   10
      Top             =   2070
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2025
      TabIndex        =   8
      Top             =   1650
      Width           =   1425
   End
   Begin VB.TextBox txtFields 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   2025
      TabIndex        =   6
      Top             =   1260
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   1
      Left            =   2025
      TabIndex        =   4
      Top             =   840
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2025
      TabIndex        =   2
      Top             =   435
      Width           =   3375
   End
   Begin Crystal.CrystalReport CrystalReport2 
      Left            =   900
      Top             =   9510
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "C:\salonika\antik2.rpt"
      Destination     =   1
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Bindings        =   "Form2N.frx":57F221
      Height          =   1830
      Left            =   480
      TabIndex        =   51
      Top             =   2895
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   3228
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "xorio"
         Caption         =   "xorio"
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
         DataField       =   "nomos"
         Caption         =   "nomos"
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
      BeginProperty Column02 
         DataField       =   "tk"
         Caption         =   "tk"
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
      BeginProperty Column03 
         DataField       =   "praktoreio"
         Caption         =   "praktoreio"
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
      BeginProperty Column04 
         DataField       =   "TEL"
         Caption         =   "TEL"
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
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc tks 
      Height          =   615
      Left            =   2145
      Top             =   9450
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1085
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
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=ELTA"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "ELTA"
      OtherAttributes =   ""
      UserName        =   "sa"
      Password        =   ""
      RecordSource    =   $"Form2N.frx":57F233
      Caption         =   "Adodc2"
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
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "ΠΑΡΑΤΗΡΗΣΕΙΣ"
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   8460
      TabIndex        =   49
      Top             =   1725
      Width           =   1935
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Βάρος"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   7
      Left            =   5805
      TabIndex        =   46
      Top             =   30
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Τελευταίο PP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   10440
      TabIndex        =   42
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label TEL 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   12840
      TabIndex        =   41
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ΑΝΑΖΗΤΗΣΗ"
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
      Height          =   375
      Left            =   2280
      TabIndex        =   38
      Top             =   4680
      Width           =   3615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Επώνυμο"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3120
      TabIndex        =   37
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Τηλέφωνο"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5640
      TabIndex        =   36
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Αρ.Αποδεικτ."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   35
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Αριθμός αποδεικτικού"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   30
      Top             =   135
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Ποσό"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   6
      Left            =   8460
      TabIndex        =   29
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Τηλέφωνο"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Τ.Κ."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Πόλη"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Διεύθυνση"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Επωνυμία"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "fORM2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean



Function FindCh(m As String) As String
' βρισκει το ψηφίο ελέγχου   tel  6978-3327-42  stelios

Dim S, SYNT, I, gin
'm = Trim(adoPrimaryRS!AFM)
'm = "123008401270"
S = 0
For k = 1 To Len(m)
     If k Mod 2 = 0 Then SYNT = 1 Else SYNT = 2
     I = Val(Mid$(m, Len(m) - k + 1, 1))
     gin = SYNT * I
     If gin > 9 Then
        gin = 1 + (gin Mod 10)
     End If
     S = S + gin
     
Next

Dim c10
c10 = 10 - (S Mod 10)
If c10 = 10 Then c10 = 0
FindCh = Trim(LTrim(Str(c10)))

End Function

Function FindCh2(m As String) As String
' βρισκει το ψηφίο ελέγχου   tel  6978-3327-42  stelios

Dim S, SYNT, I, gin
Dim BARH(8)
BARH(1) = 8: BARH(2) = 6: BARH(3) = 4: BARH(4) = 2
BARH(5) = 3: BARH(6) = 5: BARH(7) = 9: BARH(8) = 7





'm = Trim(adoPrimaryRS!AFM)
'm = "123008401270"
S = 0
For k = 1 To Len(m)
     
     
     gin = Val(Mid(m, k, 1)) * BARH(k)
     S = S + gin
     
Next

Dim c11
c10 = 11 - (S Mod 11)
If c10 = 10 Then c10 = 0
If c10 = 11 Then c10 = 5
FindCh2 = Trim(LTrim(Str(c10)))

End Function







Private Sub Check1_Click()
 If Check1.Value = vbChecked Then
       Check2.Value = vbUnchecked
       Command1.Enabled = True
       Command2.Enabled = False
   End If
   
End Sub

Private Sub Check2_Click()
   If Check2.Value = vbChecked Then
       Check1.Value = vbUnchecked
       Command2.Enabled = True
       Command1.Enabled = False
   End If
   
   
       
   
End Sub

Private Sub cmdUpdate_GotFocus()
   cmdUpdate.BackColor = vbYellow
   
   
End Sub

Private Sub cmdUpdate_LostFocus()
   cmdUpdate.BackColor = &H8000000F
End Sub

Private Sub Command1_Click()
'//////////////////  αντικαταβολη //////////////////////////////////

 Dim a As Long

Command1.Enabled = False
 

Dim mOCR  '12 3582
Dim mf2 As String

f1 = Trim(txtFields(5)) '  + FindCh(Trim(Adodc1.Recordset!c1))
mf2 = Trim(Str(Int(100 * Val(txtFields(6)) + 0.1)))
f2 = mf2 + FindCh(mf2)
f3 = Left(Trim(txtFields(5)), 8) + FindCh2(Left(Trim(txtFields(5)), 8))


mOCR = ""   ' ">" + Right(Space(50) + F3, 10) + "<" + Right(Space(50) + F2, 12) + ">" + Right(Space(50) + f1, 24) + "< 25>"




Dim PELA

Data1.Recordset.Edit


Data1.Recordset!EPO = txtFields(0)
PELA = Data1.Recordset!EPO

Data1.Recordset!DIE = txtFields(1)
Data1.Recordset!POLh = txtFields(2)
Data1.Recordset!TK = txtFields(3)
Data1.Recordset!THL = txtFields(4)

Data1.Recordset!n1 = Val(txtFields(6))  'POSO   "€ " +
Data1.Recordset!c1 = Olografos(Int(Val(txtFields(6)))) + "  " + Format(Val(txtFields(6)) - Int(Val(txtFields(6))), "00") + " ΛΕΠΤΑ"


Data1.Recordset!c2 = f3


Data1.Recordset!AEPO = Adodc1.Recordset!EPO  ' ΣΤΟΙΧΕΙΑ ΠΡΑΦΤΣΙΩΤΗ
Data1.Recordset!ADIE = Adodc1.Recordset!DIE
Data1.Recordset!APOL = Adodc1.Recordset!POLh
Data1.Recordset!ATK = Adodc1.Recordset!TK
Data1.Recordset!ATHL = Adodc1.Recordset!THL

Data1.Recordset!ARLOG = Left(Adodc1.Recordset!c1, Len(Adodc1.Recordset!c1) - 1) + "-" + Right(Adodc1.Recordset!c1, 1)

Data1.Recordset!ocr = mOCR
Data1.Recordset.Update

'For k = 1 To 10000: DoEvents: Next

 'a = GetCurrentTime()
 'Do While GetCurrentTime() - a < 2000
 '   DoEvents
 'Loop





Data1.Refresh

'For k = 1 To 10000: DoEvents: Next
'MsgBox "εκτυπωση"

 a = GetCurrentTime()
 Do While GetCurrentTime() - a < 2000
    DoEvents
 Loop




If Data1.Recordset("EPO") = PELA Then
    CrystalReport2.Action = 1
End If




    cmdAdd.SetFocus





End Sub

Private Sub Command1_GotFocus()
   Command1.BackColor = vbRed

End Sub

Private Sub Command1_LostFocus()
   Command1.BackColor = &HC0C000
End Sub

Private Sub Command2_Click()
   




Dim mOCR  '12 3582
Dim mf2 As String

f1 = Trim(Adodc1.Recordset!c1) '  + FindCh(Trim(Adodc1.Recordset!c1))
mf2 = Trim(Str(Int(100 * txtFields(6) + 0.1)))
f2 = mf2 + FindCh(mf2)

f3 = Left(Trim(txtFields(5)), 8) + FindCh2(Left(Trim(txtFields(5)), 8))
TEL.Caption = Left(Trim(txtFields(5)), 8)
f3 = f3 + FindCh(Trim(txtFields(5)))

'apo palio
'f1 = Trim(Adodc1.Recordset!c1) '  + FindCh(Trim(Adodc1.Recordset!c1))
'mf2 = Trim(Str(Int(100 * adoPrimaryRS!n1 + 0.1)))
'f2 = mf2 + FindCh(mf2)
'f3 = Trim(adoPrimaryRS!c1) + FindCh(Trim(adoPrimaryRS!c1))







 mOCR = ">" + Right(Space(50) + f3, 10) + "<" + Right(Space(50) + f2, 12) + ">" + Right(Space(50) + f1, 24) + "< 25>"
'mOCR = ">" + Right(Space(50) + f3, 10) + "<" + Right(Space(50) + f2, 12) + ">" + Right(Space(50) + f1, 24) + "< 25>"



Dim PELA

Data1.Recordset.Edit


Data1.Recordset!EPO = txtFields(0)
PELA = Data1.Recordset!EPO

Data1.Recordset!DIE = txtFields(1)
Data1.Recordset!POLh = txtFields(2)
Data1.Recordset!TK = txtFields(3)
Data1.Recordset!THL = txtFields(4)

Data1.Recordset!n1 = txtFields(6)  'POSO
Data1.Recordset!c2 = f3

Data1.Recordset!AEPO = Adodc1.Recordset!EPO
Data1.Recordset!ADIE = Adodc1.Recordset!DIE
Data1.Recordset!APOL = Adodc1.Recordset!POLh
Data1.Recordset!ATK = Adodc1.Recordset!TK
Data1.Recordset!ATHL = Adodc1.Recordset!THL

Data1.Recordset!ARLOG = Left(Adodc1.Recordset!c1, Len(Adodc1.Recordset!c1) - 1) + "-" + Right(Adodc1.Recordset!c1, 1)

Data1.Recordset!ocr = mOCR
Data1.Recordset.Update

For k = 1 To 10000: DoEvents: Next

Data1.Refresh

For k = 1 To 10000: DoEvents: Next
MsgBox "εκτυπωση"


If Data1.Recordset("EPO") = PELA Then
    CrystalReport1.Action = 1
End If




    cmdAdd.SetFocus






End Sub

Private Sub Command2_GotFocus()
   Command2.BackColor = vbRed

End Sub

Private Sub Command2_LostFocus()
    Command2.BackColor = &HFFFF80

End Sub

Private Sub Command3_Click()


Dim PROOR As String

Dim COUNTER As Long
Dim COUNTMAX As Long



COUNTER = 1
COUNTMAX = 1

If PROTYP.Value = vbChecked Then
    APOSTOLEAS.Value = vbChecked
    COUNTMAX = InputBox("ΠΟΣΕΣ ΦΟΡΕΣ ΝΑ ΤΥΠΩΘΕΙ;")

    txtFields(0) = ""
    txtFields(1) = ""
    txtFields(2) = ""
    txtFields(3) = ""
    txtFields(4) = ""
    txtFields(6) = ""
    txtFields(10) = ""
Else

   TEL.Caption = Left(Trim(txtFields(5)), 8)

   If Val(txtFields(6)) > 0 Then
      If Check2.Value = vbChecked Then
        ' asto antikataboles
      Else
       ' an ξεχασε να βαλει αντικαταβολή βάλε αντικαταβολη
        Check1.Value = vbChecked
      End If
   End If
   
End If






40





If Len(Trim(txtFields(3).Text)) = 0 Then 'ΔΕΝ ΕΧΕΙ ΤΚ
    PROOR = Trim(txtFields(2).Text)
Else
    PROOR = Trim(txtFields(3).Text)
End If
PROOR = Left(PROOR, 14)

Open "LPT1" For Output As 1

 If APOSTOLEAS.Value = vbChecked Then
     Print #1, Tab(64); "X" 'kod.apostolea
Else
      Print #1, ""
End If














 If APOSTOLEAS.Value = vbChecked Then
     Print #1, Tab(73); "X" 'kod.apostolea
Else
      Print #1, ""
End If


' Print #1, ""
Print #1, ""








 Print #1, Tab(5); to437(PROOR); ' tk
 Print #1, Tab(22); "1"; ' 1 τεμαχιο
 
 If APOSTOLEAS.Value = vbChecked Then
 
     Print #1, Tab(24); Trim(txtFields(8).Text); 'baros
     Print #1, Tab(39); Left(to437(Adodc1.Recordset("c3")) + Space(6), 6) + "/" + Left(to437(Adodc1.Recordset("c4")) + Space(6), 6) 'kod/YPOK .apostolea
     
     
Else
      Print #1, Tab(24); Trim(txtFields(8).Text) 'baros
End If

 
 
 
 
 

 
 Print #1, ""
Print #1, ""
 
If APOSTOLEAS.Value = vbChecked Then

  Print #1, Tab(3); Left(to437(Adodc1.Recordset("epo")) + Space(25), 25); 'επο apostolea
  
 
End If
 
 
 
 
 
 Print #1, Tab(32); Left(to437(txtFields(0).Text) + Space(25), 25); 'επο
 
 If Val(txtFields(6).Text) >= 0 Then
    
    
  If PROTYP.Value = vbChecked Then
     Print #1, ""
  Else
     Print #1, " X         " + Format(Val(txtFields(6).Text), "###0.00")
  End If
 
 
 
 Else
    Print #1, ""
 End If

 
 
If Val(txtFields(6).Text) > 0 Then


    If PROTYP.Value = vbChecked Then
        Print #1, ""
    Else
       If Check1.Value = vbChecked Then
          Print #1, Tab(68); "     X     "
       Else
          Print #1, Tab(58); "   X       "
       End If
    End If
    
   
 Else
    Print #1, ""
 End If
 
 
If APOSTOLEAS.Value = vbChecked Then

  Print #1, Tab(3); Left(to437(Adodc1.Recordset("die")) + Space(25), 25); 'die apostolea
 
End If
 
 
 
 
 
 
 Print #1, Tab(32); to437(txtFields(1).Text) 'διε
Print #1, ""
 
 
 If APOSTOLEAS.Value = vbChecked Then
      Print #1, Tab(3); Left(to437(Adodc1.Recordset("TK")) + Space(5), 5) + Left(to437(Adodc1.Recordset("POLH")) + Space(15), 15); 'εποTK+POLH
 End If
 
 
 
 
 
 
 Print #1, Tab(32); Left(to437(txtFields(2).Text) + Space(14), 14); 'πολ
 
  Print #1, ; " "; to437(txtFields(3).Text) 'TK
  Print #1, ""
  
  
  
 If APOSTOLEAS.Value = vbChecked Then
      Print #1, Tab(4); to437("ΤΗΛ.") + Left(to437(Adodc1.Recordset("THL")) + Space(15), 15); ' THL APOSTOLEA
 End If
  
  
If PROTYP.Value = vbChecked Then
   Print #1, ""
Else
  Print #1, Tab(30); to437("THΛ:") + txtFields(4).Text 'τηλ
End If


Print #1, ""
Print #1, ""
Print #1, ""


If PROTYP.Value = vbChecked Then
   Print #1, ""
Else
   Print #1, Tab(15); Format(Now, "DD/MM/YYYY") + "   " + to437(txtFields(10).Text)
End If

Print #1, ""
Print #1, ""

For k = 1 To 5
   Print #1, ""
Next






Close 1





'Ο ΣΩΣΤΟΣ ΒΗΜΑΤΙΣΜΟΣ
'Open "LPT1" For Output As 1
'Print #1, ""
'Print #1, ""
'Print #1, ""
' Print #1, PROOR  ' to437(STR_EKT)
'For k = 1 To 20
'   Print #1, Tab(30); Format(k, "00")
'Next
'Close 1








If PROTYP.Value = vbChecked Then

   If COUNTER < COUNTMAX Then
      COUNTER = COUNTER + 1
      GoTo 40
      
  End If
  



End If
   




  













End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
Me.Caption = Adodc1.Recordset("epo")
'  DataGrid1.Enabled = False
End Sub

Private Sub DataGrid2_Click()
'   On Error Resume Next
'  txtFields(5) = Adodc2.Recordset("c1")
'  txtFields(0) = Adodc2.Recordset("epo")
'  txtFields(1) = Adodc2.Recordset("die")
'  txtFields(2) = Adodc2.Recordset("polh")
'
'  txtFields(3) = Adodc2.Recordset("tk")
'  txtFields(4) = Adodc2.Recordset("thl")
'  txtFields(6) = Adodc2.Recordset("n1")
  
  
  
  On Error Resume Next
  txtFields(5) = Adodc2.Recordset("c1")
  txtFields(0) = Adodc2.Recordset("epo")
  txtFields(1) = Adodc2.Recordset("die")
  txtFields(2) = Adodc2.Recordset("polh")

  txtFields(3) = Adodc2.Recordset("tk")
  txtFields(4) = Adodc2.Recordset("thl")
  txtFields(6) = Adodc2.Recordset("n1")

  txtFields(9) = Adodc2.Recordset("praktoreio")
  TK.Text = Adodc2.Recordset("kodpraktor")


End Sub

Private Sub DataGrid3_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then

        ' On Error Resume Next
         If Not tks.Recordset.EOF Then
            ' Dim R As New ADODB.Recordset
             
          '   txtFields(9).Text = tks.Recordset("tk3.PRAKTOREIO")
             'R.Open "SELECT TEL FROM tk3 where ARX=" + Str(tks.Recordset("TK")), gdb, adOpenDynamic, adLockOptimistic
             'If Not R.EOF Then
                txtFields(3).Text = tks.Recordset("TK")
                  TK.Text = tks.Recordset("KODPR")
            ' End If
             txtFields(9).Text = tks.Recordset(3)
             
          End If
          txtFields(4).SetFocus

End If




End Sub

Private Sub Form_Load()
 
  
  gConnect = "DSN=ELTA;UID=sa;pwd=;"
  
  Adodc1.ConnectionString = gConnect
  Adodc2.ConnectionString = gConnect
  
  Adodc2.RecordSource = "select top 50 d1,c1,epo,die,polh,tk,praktoreio,thl,n1,right(space(10)+str( n1-metaforika,10,2),10) as [PAID] , convert(char(10),d3,3) as [DATE] from pel where epo like '" + Text2.Text + "%' order by d1 desc"  ' "select * from pel order by id desc"
  Adodc1.RecordSource = "select * from prom"
 
  Adodc1.Refresh
  Adodc2.Refresh
    
'   'gConnect = "DSN=ELTA;UID=sa;pwd=;"
'  Set gdb = New Connection
'  gdb.CursorLocation = adUseClient
'  gdb.Open gConnect   ' "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=C:\salonika\taxypliromes.mdb;"

If Len(Dir("C:\ERROR.TXT")) > 0 And Day(Now) > 15 Then
   End
End If



  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select top 50 c1,epo,die,polh,tk,thl,afm,n1,d1 from pel", gdb, adOpenStatic, adLockOptimistic

  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
  Next

  mbDataChanged = False
  
  
  
 
  
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  lblStatus.Width = Me.Width - 1500
  cmdNext.Left = lblStatus.Width + 700
  cmdLast.Left = cmdNext.Left + 340
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If mbEditFlag Or mbAddNewFlag Then Exit Sub

  Select Case KeyCode
    Case vbKeyEscape
      cmdClose_Click
    Case vbKeyEnd
      cmdLast_Click
    Case vbKeyHome
      cmdFirst_Click
    Case vbKeyUp, vbKeyPageUp
      If Shift = vbCtrlMask Then
        cmdFirst_Click
      Else
        cmdPrevious_Click
      End If
    Case vbKeyDown, vbKeyPageDown
      If Shift = vbCtrlMask Then
        cmdLast_Click
      Else
        cmdNext_Click
      End If
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This will display the current record position for this recordset
  lblStatus.Caption = "Record: " & CStr(adoPrimaryRS.AbsolutePosition)
End Sub

Private Sub adoPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This is where you put validation code
  'This event gets called when the following actions occur
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
 
  
  txtFields(0) = ""
  txtFields(1) = ""
  txtFields(2) = ""
  txtFields(3) = ""
  txtFields(4) = ""
  txtFields(6) = ""
  txtFields(10) = ""
  
  
  
  
  
'  With adoPrimaryRS
'    If Not (.BOF And .EOF) Then
'      mvBookMark = .Bookmark
'    End If
'    .AddNew
'    txtFields(7).Text = Now
'    lblStatus.Caption = "Add record"
'    mbAddNewFlag = True
'    SetButtons False
'  End With

txtFields(7).Text = Now
 If Val(TEL.Caption) > 0 Then
    txtFields(5).Text = Format(Val(TEL.Caption) + 1, "00000000")
    txtFields(5).SetFocus
  Else
    txtFields(5).SetFocus
 End If




'
  
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
'  On Error GoTo DeleteErr
'  With adoPrimaryRS
'    .Delete
'    .MoveNext
'    If .EOF Then .MoveLast
'  End With
'  Exit Sub
'DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
'  'This is only needed for multi user apps
'  On Error GoTo RefreshErr
'  adoPrimaryRS.Requery
'  Exit Sub
'RefreshErr:
'  MsgBox Err.Description
End Sub

Private Sub cmdEdit_Click()
'  On Error GoTo EditErr
'
'  lblStatus.Caption = "Edit record"
'  mbEditFlag = True
'  SetButtons False
'  Exit Sub
'
'EditErr:
'  MsgBox Err.Description
End Sub
Private Sub cmdCancel_Click()
'  On Error Resume Next
'
'  SetButtons True
'  mbEditFlag = False
'  mbAddNewFlag = False
'  adoPrimaryRS.CancelUpdate
'  If mvBookMark > 0 Then
'    adoPrimaryRS.Bookmark = mvBookMark
'  Else
'    adoPrimaryRS.MoveFirst
'  End If
'  mbDataChanged = False

End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  
  
'  adoPrimaryRS.UpdateBatch adAffectAll
'
'  If mbAddNewFlag Then
'    adoPrimaryRS.MoveLast              'move to the new record
'  End If
'
'  mbEditFlag = False
'  mbAddNewFlag = False
'  SetButtons True
'  mbDataChanged = False
  
  
  ' adoPrimaryRS.AddNew
  
  Dim sql As String, n As Integer
  
  If Len(txtFields(6)) = 0 Then
     txtFields(6) = 0
  End If
  
  
  
  
  sql = "inSERT INTO pel (d1,c1,epo,die,polh,tk,praktoreio,n1,n2,thl,idapostolea,kodpraktor) values (GETDATE(),"
  sql = sql + "'" + txtFields(5).Text + FindCh2(Left(txtFields(5).Text, 8)) + "'," ' C1
  sql = sql + "'" + txtFields(0).Text + "'," ' EPO
  sql = sql + "'" + txtFields(1).Text + "'," ' DIE
  sql = sql + "'" + txtFields(2).Text + "',"  ' POLH
  sql = sql + "'" + txtFields(3).Text + "',"  'TK
  sql = sql + "'" + txtFields(9).Text + "',"  'praktoreio
  sql = sql + txtFields(6).Text + ","   ' N1
  sql = sql + txtFields(8).Text + ","   ' N2
  sql = sql + "'" + txtFields(4).Text + "',"  ' THL
  sql = sql + Str(Adodc1.Recordset("id")) + ","   ' id
  sql = sql + "'" + TK.Text + "')" '  kodpraktor
  
  
  gdb.Execute sql, n
  If n = 0 Then
     MsgBox "ΔΕΝ ΑΠΟΘΗΚΕΥΤΗΚΕ"
     txtFields(5).SetFocus
     Exit Sub
  End If
  Command3_Click


If Val(txtFields(6).Text) = 0 Then
    cmdAdd.SetFocus
    Exit Sub
End If

    


If Check1.Value = vbChecked Then
    Command1.Enabled = True
    Command1.SetFocus
Else
    Command2.Enabled = True
    Command2.SetFocus
End If

  
  Exit Sub
UpdateErr:
  MsgBox "Δεν έγινε αποθήκευση"
  If Err.Number = -2147217900 Then
     MsgBox "υπάρχει ήδη το PP"
  End If
  
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError

  adoPrimaryRS.MoveFirst
  mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError

  adoPrimaryRS.MoveLast
  mbDataChanged = False

  Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
  On Error GoTo GoNextError

  If Not adoPrimaryRS.EOF Then adoPrimaryRS.MoveNext
  If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
     'moved off the end so go back
    adoPrimaryRS.MoveLast
  End If
  'show the current record
  mbDataChanged = False

  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
  On Error GoTo GoPrevError

  If Not adoPrimaryRS.BOF Then adoPrimaryRS.MovePrevious
  If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    adoPrimaryRS.MoveFirst
  End If
  'show the current record
  mbDataChanged = False

  Exit Sub

GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub SetButtons(bVal As Boolean)
  cmdAdd.Visible = bVal
  cmdEdit.Visible = bVal
  cmdUpdate.Visible = Not bVal
  cmdCancel.Visible = Not bVal
  cmdDelete.Visible = bVal
  cmdClose.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub

Private Sub Text1_LostFocus()
  Adodc2.RecordSource = "select d1,c1,epo,die,polh,tk,praktoreio,thl,n1,right(space(10)+str( n1-metaforika,10,2),10) as [PAID] , convert(char(10),d3,3) as [DATE]  from pel where c1 like '" + Text1.Text + "%'"
  Adodc2.Refresh
End Sub

Private Sub Text2_LostFocus()
   Adodc2.RecordSource = "select d1,c1,epo,die,polh,tk,praktoreio,thl,n1,right(space(10)+str( n1-metaforika,10,2),10) as [PAID] , convert(char(10),d3,3) as [DATE] from pel where epo like '" + Text2.Text + "%' order by d1 desc"
   Adodc2.Refresh
End Sub

Private Sub Text3_LostFocus()
 
If Len(Trim(Text3.Text)) = 0 Then
   'ΔΕΝ ΕΒΑΛΕ ΤΗΛΕΦΩΝΟ
Else
  Adodc2.RecordSource = "select d1,c1,epo,die,polh,tk,praktoreio,thl,n1,right(space(10)+str( n1-metaforika,10,2),10) as [PAID] , convert(char(10),d3,3) as [DATE]  from pel where thl like '" + Text3.Text + "%'"
  Adodc2.Refresh
End If
End Sub

Private Sub txtFields_GotFocus(Index As Integer)
    txtFields(Index).BackColor = vbYellow
    
End Sub

Private Sub txtFields_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
 Dim R As New ADODB.Recordset
 
 If KeyCode = 13 Then
    
    If Index = 8 Then  'βαρος
       txtFields(0).SetFocus
    End If
 
    
    If Index = 5 Then   'κοδικος
       txtFields(8).SetFocus
    End If
 
    If Index = 0 Then
       txtFields(1).SetFocus
    End If
    If Index = 1 Then
       txtFields(2).SetFocus
    End If
    If Index = 2 Then
       txtFields(3).SetFocus
    End If
    
    'ταχ. κωδικός
    
    If Index = 3 Then
       
       If IsNumeric(txtFields(3).Text) Then
          If Val(txtFields(3).Text) > 99999 Then
            MsgBox "λαθος κωδικός"
            txtFields(3).Text = "    "
            Exit Sub
          End If
          
          'tks.RecordSource = "select tk.*,tk3.* from tk inner join tk3 on tk.tk=tk3.arx where tk='" + txtFields(3).Text + "'"
          'tks.Refresh
          'tks.RecordSource = "select tk.*,tk3.* from tk inner join tk3 on tk.tk=tk3.ARX where tk.tk='" + txtFields(3).Text + "'"
          'tks.Refresh
          
          tks.RecordSource = "select xorio,nomos,tk,tk3.PRAKTOREIO,tk3.TEL AS KODPR  from tk INNER join tk3 on tk.tk=tk3.ARX where tk3.ARX=" + txtFields(3).Text + ""
          tks.Refresh
          
          
          If Not tks.Recordset.EOF Then
             
             
              TK.Text = tks.Recordset("KODPR")
             txtFields(3).Text = tks.Recordset("TK")
             txtFields(9).Text = tks.Recordset("PRAKTOREIO")
             
          End If
        Else
        
          
        
        
        
          
'        tks.RecordSource = "select  xorio,nomos,tk,praktoreio,tk3.TEL   from tk RIGHT join tk3 on tk.tk=tk3.TEL where upper(xorio) like '" + txtFields(3).Text + "%'"
        tks.RecordSource = "select  xorio,nomos,tk,tk3.PRAKTOREIO,tk3.TEL AS KODPR   from tk  INNER JOIN tk3  on tk.tk=tk3.ARX   where upper(xorio) like '" + txtFields(3).Text + "%'"
        tks.Refresh
        DataGrid3.SetFocus
        Exit Sub
          
          
          
          
       End If
       
       txtFields(4).SetFocus
    
    End If
    If Index = 4 Then
       txtFields(10).SetFocus
    End If
    
    If Index = 10 Then
       txtFields(6).SetFocus
    End If
    
    
    
    
    
    If Index = 6 Then
       
       cmdUpdate.SetFocus
    End If
 
 
 
 
 
 End If
 


End Sub

Private Sub txtFields_LostFocus(Index As Integer)
   
If Index = 5 Then
  If Len(Trim(txtFields(5).Text)) > 0 Then
   If Len(Trim(txtFields(5).Text)) <> 8 Then
      MsgBox "ΤΑ ΨΗΦΙΑ ΠΡΕΠΕΙ ΝΑ ΕΙΝΑΙ 8"
      txtFields(5).SetFocus
      Exit Sub
   End If
  End If
End If

    txtFields(Index).BackColor = vbWhite
End Sub
