VERSION 5.00
Begin VB.UserControl UserControl1 
   ClientHeight    =   4110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   4110
   ScaleWidth      =   4800
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   315
      Left            =   3060
      TabIndex        =   1
      Top             =   480
      Width           =   315
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label2"
      Height          =   1635
      Left            =   300
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   315
      Left            =   300
      TabIndex        =   0
      Top             =   480
      Width           =   2715
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Sub Command1_Click()
    Label2.Visible = Not Label2.Visible
End Sub

