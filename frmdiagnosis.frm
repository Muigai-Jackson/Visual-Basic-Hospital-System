VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmdiagnosis 
   BackColor       =   &H00FFFF00&
   Caption         =   "Patient Diagnosis"
   ClientHeight    =   7650
   ClientLeft      =   5685
   ClientTop       =   2460
   ClientWidth     =   12495
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   12495
   Begin VB.CheckBox chkward 
      BackColor       =   &H00FFFF00&
      Caption         =   "Yes"
      DataField       =   "wardreq"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8640
      TabIndex        =   16
      Top             =   2160
      Width           =   855
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   6240
      Top             =   4560
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   582
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\Projects\Visual Basic\Hospital.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\Projects\Visual Basic\Hospital.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from patients"
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   6240
      Top             =   3840
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   582
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\Projects\Visual Basic\Hospital.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\Projects\Visual Basic\Hospital.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from patientsreport"
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
   Begin VB.ComboBox cbotype 
      DataField       =   "wardtype"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmdiagnosis.frx":0000
      Left            =   8400
      List            =   "frmdiagnosis.frx":000D
      TabIndex        =   15
      Text            =   "General"
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      DataField       =   "medicine"
      DataSource      =   "Adodc1"
      Height          =   975
      Left            =   2880
      TabIndex        =   14
      Top             =   4560
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      DataField       =   "diagnosis"
      DataSource      =   "Adodc1"
      Height          =   855
      Left            =   2880
      TabIndex        =   13
      Top             =   3480
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      DataField       =   "syptoms"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2880
      TabIndex        =   12
      Top             =   2760
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      DataField       =   "bloodgroup"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2880
      TabIndex        =   11
      Top             =   2160
      Width           =   2535
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
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
      Left            =   7080
      TabIndex        =   10
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmssave 
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
      Left            =   1920
      TabIndex        =   9
      Top             =   5880
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmdiagnosis.frx":0027
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   1720
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
      Caption         =   "Patient Details"
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
      Left            =   4680
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txtsearch 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
   Begin VB.Image Image3 
      Height          =   615
      Left            =   6360
      Picture         =   "frmdiagnosis.frx":003C
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   1080
      Picture         =   "frmdiagnosis.frx":40CA
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label lblward 
      BackColor       =   &H00FFFF00&
      Caption         =   "Ward Type:"
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
      Left            =   6120
      TabIndex        =   8
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFF00&
      Caption         =   "Ward Required?"
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
      Left            =   6120
      TabIndex        =   7
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF00&
      Caption         =   "Medicine:"
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
      TabIndex        =   6
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      Caption         =   "Diagnosis:"
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
      TabIndex        =   5
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Syptoms:"
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
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Blood Group:"
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
      TabIndex        =   3
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   5880
      Picture         =   "frmdiagnosis.frx":C4EB
      Stretch         =   -1  'True
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmdiagnosis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkward_Click()
If chkward.Value = 1 Then
    lblward.Visible = True
    cbotype.Visible = True
    Else
    lblward.Visible = False
    cbotype.Visible = False
    End If
End Sub
Private Sub cmdclose_Click()
Dim ans As String
ans = MsgBox("Are You Sure You Want To Go Back?", vbYesNo)
If ans = vbYes Then
 Unload frmdiagnosis
 frmhos.Show
 End If
End Sub
Private Sub cmdsearch_Click()
Adodc2.RecordSource = "select * from patients where patientid='" + txtsearch.Text + "'"
Adodc2.Refresh
If txtsearch.Text = "" Then
    MsgBox " Please enter ID to search!!! "
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    chkward.Value = 0
    Exit Sub
Else
    If Adodc2.Recordset.EOF = True Or Adodc2.Recordset.BOF = True Then
    MsgBox "Record Not Found!!!"
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    chkward.Value = 0
Else
    Adodc2.Caption = Adodc2.RecordSource
    Set DataGrid1.DataSource = Adodc2.Recordset
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    chkward.Value = 0
End If
End If
End Sub

Private Sub cmssave_Click()
Adodc1.Recordset.Fields("patientid") = txtsearch.Text
Adodc1.Recordset.Fields("bloodgroup") = Text2.Text
Adodc1.Recordset.Fields("syptoms") = Text3.Text
Adodc1.Recordset.Fields("diagnosis") = Text4.Text
Adodc1.Recordset.Fields("medicine") = Text5.Text
If chkward.Value = 1 Then
   Adodc1.Recordset.Fields("wardreq") = True
   Adodc1.Recordset.Fields("wardtype") = cbotype.Text
Else
    Adodc1.Recordset.Fields("wardreq") = False
    Adodc1.Recordset.Fields("wardtype") = "None"
End If
Adodc1.Recordset.Update
MsgBox "Details Updated Successfully!!!?"
End Sub

Private Sub Form_Load()
Adodc2.Recordset.AddNew
Adodc1.Recordset.AddNew
Adodc2.RecordSource = "select * from patients"
Adodc1.RecordSource = "select * from patientsreport"
Adodc2.Refresh
chkward.Value = 0
Set DataGrid1.DataSource = Adodc2.Recordset
    lblward.Visible = False
    cbotype.Visible = False
End Sub
