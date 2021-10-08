VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmfull 
   BackColor       =   &H00FFFF00&
   Caption         =   "Full History Of Patient"
   ClientHeight    =   7440
   ClientLeft      =   3120
   ClientTop       =   5460
   ClientWidth     =   20505
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   20505
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   7560
      Top             =   6840
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\Projects\Visual Basic\Hospital.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\Projects\Visual Basic\Hospital.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *  from fulldetails"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmfull.frx":0000
      Height          =   4935
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   19935
      _ExtentX        =   35163
      _ExtentY        =   8705
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
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
            LCID            =   3072
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
            LCID            =   3072
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      TabIndex        =   2
      Top             =   6000
      Width           =   1335
   End
   Begin VB.TextBox txtsearch 
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   360
      Width           =   3495
   End
   Begin VB.CommandButton cmdsearch 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin VB.Image Image3 
      Height          =   615
      Left            =   5400
      Picture         =   "frmfull.frx":0015
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   5880
      Picture         =   "frmfull.frx":40A3
      Stretch         =   -1  'True
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmfull"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdback_Click()
Dim ans As String
ans = MsgBox("Are You Sure You Want To Go Back?", vbYesNo)
If ans = vbYes Then
 Unload frmfull
 frmhos.Show
 End If
End Sub

Private Sub cmdsearch_Click()
Adodc1.RecordSource = "select * from fulldetails where patientid='" + txtsearch.Text + "'"
Adodc1.Refresh
If txtsearch.Text = "" Then
    MsgBox " Please enter ID to search!!! "
    Exit Sub
Else
    If Adodc1.Recordset.EOF = True Or Adodc1.Recordset.BOF = True Then
    MsgBox "Record Not Found!!!"
Else
    Adodc1.Caption = Adodc1.RecordSource
    Set DataGrid1.DataSource = Adodc1.Recordset
End If
End If
End Sub


Private Sub Form_Load()
    'Adodc1.Recordset.AddNew
    Set DataGrid1.DataSource = Adodc1.Recordset
    Adodc1.Refresh
End Sub

