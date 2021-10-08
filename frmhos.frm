VERSION 5.00
Begin VB.Form frmhos 
   BackColor       =   &H00FFFF00&
   Caption         =   "BenjaminWellness Center"
   ClientHeight    =   8670
   ClientLeft      =   2400
   ClientTop       =   2610
   ClientWidth     =   12525
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   12525
   Begin VB.CommandButton cmdfull 
      Caption         =   "Patient Full History"
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
      Left            =   1680
      TabIndex        =   7
      Top             =   5160
      Width           =   2415
   End
   Begin VB.CommandButton cmddiagnosis 
      Caption         =   "Add Diagnosis"
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
      Left            =   1680
      TabIndex        =   6
      Top             =   4200
      Width           =   2535
   End
   Begin VB.CommandButton cmdlogout 
      Caption         =   "Logout"
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
      Left            =   1680
      TabIndex        =   5
      Top             =   7080
      Width           =   2415
   End
   Begin VB.CommandButton cmdinfo 
      Caption         =   "Hospital Information"
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
      Left            =   1680
      TabIndex        =   4
      Top             =   6120
      Width           =   2535
   End
   Begin VB.CommandButton cmdupdate 
      Caption         =   "Update Record"
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
      Left            =   1680
      TabIndex        =   3
      Top             =   3240
      Width           =   2535
   End
   Begin VB.CommandButton cmddetails 
      Caption         =   "Patient Details"
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
      Left            =   1680
      TabIndex        =   2
      Top             =   2280
      Width           =   2535
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "Add New Patient"
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
      Left            =   1680
      TabIndex        =   1
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Image Image7 
      Height          =   855
      Left            =   600
      Picture         =   "frmhos.frx":0000
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   855
   End
   Begin VB.Image Image6 
      Height          =   855
      Left            =   600
      Picture         =   "frmhos.frx":17348
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   855
   End
   Begin VB.Image Image5 
      Height          =   855
      Left            =   720
      Picture         =   "frmhos.frx":1F769
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   855
   End
   Begin VB.Image Image4 
      Height          =   855
      Left            =   600
      Picture         =   "frmhos.frx":2A1FE
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   975
   End
   Begin VB.Image Image3 
      Height          =   735
      Left            =   720
      Picture         =   "frmhos.frx":2DDAD
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   855
      Left            =   720
      Picture         =   "frmhos.frx":31E3B
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   720
      Picture         =   "frmhos.frx":358C0
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   765
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      Caption         =   "BENJAMIN WELLNESS CENTER "
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   9615
   End
End
Attribute VB_Name = "frmhos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmddetails_Click()
frmdetails.Show
End Sub

Private Sub cmddiagnosis_Click()
frmdiagnosis.Show
End Sub

Private Sub cmdfull_Click()
frmfull.Show
End Sub

Private Sub cmdinfo_Click()
frmhistory.Show
End Sub

Private Sub cmdlogout_Click()
Dim ans As String
ans = MsgBox("Are You Sure You Want To Exit?", vbYesNo)
If ans = vbYes Then
    frmhospital.Show
    Unload frmhos
End If
End Sub

Private Sub cmdnew_Click()
frmnewpatient.Show
End Sub

Private Sub cmdupdate_Click()
frmupdate.Show
End Sub
