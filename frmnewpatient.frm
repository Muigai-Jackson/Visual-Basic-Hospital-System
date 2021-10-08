VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmnewpatient 
   BackColor       =   &H00FFFF00&
   Caption         =   "Add New Patient Details"
   ClientHeight    =   9885
   ClientLeft      =   6195
   ClientTop       =   2445
   ClientWidth     =   12675
   LinkTopic       =   "Form1"
   ScaleHeight     =   9885
   ScaleWidth      =   12675
   Begin VB.CommandButton cmddelete 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   22
      Top             =   6120
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmnewpatient.frx":0000
      Height          =   2415
      Left            =   720
      TabIndex        =   21
      Top             =   7200
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   4260
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2040
      Top             =   6840
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      RecordSource    =   "patients"
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
   Begin VB.CommandButton cmdnew 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   20
      Top             =   6120
      Width           =   1095
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
      Height          =   495
      Left            =   7440
      TabIndex        =   19
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   18
      Top             =   6120
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      DataField       =   "bloodpressure"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2640
      TabIndex        =   17
      Top             =   4680
      Width           =   2775
   End
   Begin VB.TextBox Text8 
      DataField       =   "temperature"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2640
      TabIndex        =   16
      Top             =   3960
      Width           =   2775
   End
   Begin VB.TextBox Text7 
      DataField       =   "anydisease"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   2640
      TabIndex        =   15
      Top             =   5160
      Width           =   2775
   End
   Begin VB.TextBox Text6 
      DataField       =   "age"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2640
      TabIndex        =   14
      Top             =   3360
      Width           =   2775
   End
   Begin VB.TextBox Text5 
      DataField       =   "address"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2640
      TabIndex        =   13
      Top             =   2760
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      DataField       =   "telno"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2640
      TabIndex        =   12
      Top             =   2160
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      DataField       =   "idnumber"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2640
      TabIndex        =   11
      Top             =   1560
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      DataField       =   "patientname"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2640
      TabIndex        =   10
      Top             =   960
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      DataField       =   "patientid"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2640
      TabIndex        =   9
      Top             =   360
      Width           =   2775
   End
   Begin VB.Image Image4 
      Height          =   855
      Left            =   4320
      Picture         =   "frmnewpatient.frx":0015
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   855
   End
   Begin VB.Image Image3 
      Height          =   615
      Left            =   6720
      Picture         =   "frmnewpatient.frx":9AAA
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   2280
      Picture         =   "frmnewpatient.frx":DB38
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   120
      Picture         =   "frmnewpatient.frx":11794
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   615
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFF00&
      Caption         =   "Any Major Disease:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFF00&
      Caption         =   "Blood Pressure:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   4680
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFF00&
      Caption         =   "Temperature:"
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
      Left            =   240
      TabIndex        =   6
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFF00&
      Caption         =   "Age:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFF00&
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF00&
      Caption         =   "Telephone Number:"
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
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      Caption         =   "National ID No:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Patient Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Patient ID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frmnewpatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
frmhos.Show
End Sub

Private Sub cmdback_Click()
Dim ans As String
ans = MsgBox("Are You Sure You Want To Go Back?", vbYesNo)
If ans = vbYes Then
 Adodc1.Refresh
 Unload frmnewpatient
 frmhos.Show
End If
End Sub

Private Sub cmddelete_Click()
Adodc1.Recordset.Delete
MsgBox "Successfully Deleted!!!?"
End Sub

Private Sub cmdnew_Click()
Adodc1.Recordset.AddNew
End Sub

Private Sub cmdsave_Click()
Adodc1.Recordset.Fields("patientid") = Text1.Text
Adodc1.Recordset.Fields("patientname") = Text2.Text
Adodc1.Recordset.Fields("idnumber") = Text3.Text
Adodc1.Recordset.Fields("telno") = Text4.Text
Adodc1.Recordset.Fields("address") = Text5.Text
Adodc1.Recordset.Fields("age") = Text6.Text
Adodc1.Recordset.Fields("temperature") = Text8.Text
Adodc1.Recordset.Fields("bloodpressure") = Text9.Text
Adodc1.Recordset.Fields("anydisease") = Text7.Text
Adodc1.Recordset.Update
Dim anss As String
anss = MsgBox("Details Saved Successfully!!!?")
Adodc1.Recordset.AddNew
End Sub



Private Sub Form_Load()
Adodc1.Recordset.AddNew
End Sub
