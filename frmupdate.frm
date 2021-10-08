VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmupdate 
   BackColor       =   &H00FFFF00&
   Caption         =   "Add New Patient Details"
   ClientHeight    =   7365
   ClientLeft      =   6195
   ClientTop       =   2445
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   ScaleHeight     =   458.878
   ScaleMode       =   0  'User
   ScaleWidth      =   7740
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2280
      Top             =   6960
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
      RecordSource    =   "select * from patients"
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
      Left            =   4920
      TabIndex        =   21
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdupdate 
      Caption         =   "Update"
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
      Left            =   1200
      TabIndex        =   20
      Top             =   6480
      Width           =   1215
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
      Left            =   5760
      TabIndex        =   19
      Top             =   240
      Width           =   1095
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
      Left            =   2760
      TabIndex        =   18
      Top             =   840
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
      Left            =   2760
      TabIndex        =   17
      Top             =   1320
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
      Left            =   2760
      TabIndex        =   16
      Top             =   1920
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
      Left            =   2760
      TabIndex        =   15
      Top             =   2520
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
      Left            =   2760
      TabIndex        =   14
      Top             =   3120
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
      Left            =   2760
      TabIndex        =   13
      Top             =   3720
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
      Left            =   2760
      TabIndex        =   12
      Top             =   5520
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
      Left            =   2760
      TabIndex        =   11
      Top             =   4320
      Width           =   2775
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
      Left            =   2760
      TabIndex        =   10
      Top             =   5040
      Width           =   2775
   End
   Begin VB.TextBox txtsearch 
      Height          =   405
      Left            =   2760
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   4200
      Picture         =   "frmupdate.frx":0000
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   600
      Picture         =   "frmupdate.frx":408E
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   6960
      Picture         =   "frmupdate.frx":C4AF
      Stretch         =   -1  'True
      Top             =   240
      Width           =   495
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
      Left            =   360
      TabIndex        =   9
      Top             =   1440
      Width           =   1815
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
      Left            =   360
      TabIndex        =   8
      Top             =   2040
      Width           =   1935
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
      Left            =   360
      TabIndex        =   7
      Top             =   2640
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
      Left            =   480
      TabIndex        =   6
      Top             =   3240
      Width           =   2055
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
      Left            =   360
      TabIndex        =   5
      Top             =   3840
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
      Left            =   360
      TabIndex        =   4
      Top             =   4440
      Width           =   2295
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
      Left            =   360
      TabIndex        =   3
      Top             =   5160
      Width           =   2175
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
      Left            =   360
      TabIndex        =   2
      Top             =   5880
      Width           =   2175
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
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "frmupdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdback_Click()
Dim ans As String
ans = MsgBox("Are You Sure You Want To Go Back?", vbYesNo)
If ans = vbYes Then
 Unload frmupdate
 frmhos.Show
 End If
End Sub

Private Sub cmdsearch_Click()
Adodc1.RecordSource = "select * from patients where patientid='" + txtsearch.Text + "'"
Adodc1.Refresh
If txtsearch.Text = "" Then
    MsgBox " Please enter ID to search!!! "
    Exit Sub
Else
    If Adodc1.Recordset.EOF = True Or Adodc1.Recordset.BOF = True Then
    MsgBox "Record Not Found!!!"
Else
    Adodc1.Caption = Adodc1.RecordSource
End If
End If
End Sub

Private Sub cmdupdate_Click()
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
MsgBox "Details Updated Successfully!!!?"
Adodc1.Refresh
End Sub

Private Sub Form_Load()
Adodc1.Recordset.AddNew
End Sub
