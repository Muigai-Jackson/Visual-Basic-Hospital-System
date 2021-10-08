VERSION 5.00
Begin VB.Form frmhistory 
   BackColor       =   &H00FFFF00&
   Caption         =   "About "
   ClientHeight    =   8595
   ClientLeft      =   5835
   ClientTop       =   2325
   ClientWidth     =   12540
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   12540
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
      Left            =   5520
      TabIndex        =   0
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   4680
      Picture         =   "frmhistory.frx":0000
      Stretch         =   -1  'True
      Top             =   7800
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      Caption         =   $"frmhistory.frx":408E
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7575
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   12255
   End
End
Attribute VB_Name = "frmhistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdback_Click()
Dim ans As String
ans = MsgBox("Are You Sure You Want To Go Back?", vbYesNo)
If ans = vbYes Then
 Unload frmhistory
 frmhos.Show
 End If
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

