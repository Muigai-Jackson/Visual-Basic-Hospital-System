VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmload 
   BackColor       =   &H00FFFF00&
   Caption         =   "BWC"
   ClientHeight    =   5790
   ClientLeft      =   8070
   ClientTop       =   5010
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   7335
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6120
      Top             =   240
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   5040
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   873
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6720
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label lblper 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   4440
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   3495
      Left            =   600
      Picture         =   "frmload.frx":0000
      Stretch         =   -1  'True
      Top             =   840
      Width           =   6015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      Caption         =   "BENJAMIN WELLNESS CENTER"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "frmload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 5
lblper.Caption = ProgressBar1.Value & "%"
If ProgressBar1.Value = ProgressBar1.Max Then
Timer1.Enabled = False
frmhospital.Show
Unload frmload
End If
End Sub
