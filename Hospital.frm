VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmhospital 
   BackColor       =   &H00FFFF00&
   Caption         =   "BenjaminWellness Center"
   ClientHeight    =   5370
   ClientLeft      =   8340
   ClientTop       =   4755
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   5235
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1320
      Top             =   5400
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
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
      RecordSource    =   "select * from login"
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
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
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
      Left            =   3360
      TabIndex        =   6
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton cmdlogin 
      Caption         =   "Login"
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
      TabIndex        =   5
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox txtpassword 
      DataField       =   "password"
      DataSource      =   "Adodc1"
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   3720
      Width           =   2295
   End
   Begin VB.TextBox txtname 
      DataField       =   "username"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Image pichide 
      Height          =   495
      Left            =   4440
      Picture         =   "Hospital.frx":0000
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   735
   End
   Begin VB.Image picshow 
      Height          =   495
      Left            =   4440
      Picture         =   "Hospital.frx":17D1C
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   735
   End
   Begin VB.Image Image3 
      Height          =   615
      Left            =   240
      Picture         =   "Hospital.frx":219D0
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   2640
      Picture         =   "Hospital.frx":28B59
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   2265
      Left            =   600
      Picture         =   "Hospital.frx":437DC
      Stretch         =   -1  'True
      Top             =   600
      Width           =   3960
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   0
      X2              =   5160
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      Caption         =   "Password:"
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
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Username:"
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
      TabIndex        =   1
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      Caption         =   "BENJAMIN WELNESS CENTER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmhospital"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdexit_Click()
Dim ans As String
ans = MsgBox("Are You Sure You Want To Exit?", vbYesNo)
If ans = vbYes Then
 End
End If
End Sub

Private Sub cmdlogin_Click()
Adodc1.RecordSource = "select * from login where username='" + txtname.Text + "' and password='" + txtpassword.Text + "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Or Adodc1.Recordset.BOF = True Then
        MsgBox "Login Failed, Try Again..!!!", vbCritical, "Please Enter Correct Creditential.."
        Adodc1.Recordset.AddNew
    Else
        frmhos.Show
        Unload frmhospital
End If
End Sub

Private Sub Form_Load()
Adodc1.Recordset.AddNew
pichide.Visible = False
End Sub

Private Sub pichide_Click()
picshow.Visible = True
pichide.Visible = False
txtpassword.PasswordChar = "*"
End Sub

Private Sub picshow_Click()
txtpassword.PasswordChar = ""
pichide.Visible = True
picshow.Visible = False
End Sub

