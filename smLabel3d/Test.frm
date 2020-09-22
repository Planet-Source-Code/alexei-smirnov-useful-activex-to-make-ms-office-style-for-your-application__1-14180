VERSION 5.00
Object = "*\AsmLabel3d.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   9660
   StartUpPosition =   3  'Windows Default
   Begin smLabel3d.smLabel smlbl3d1 
      Height          =   555
      Left            =   840
      Top             =   420
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   979
      Style           =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   2820
      Width           =   1995
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   435
      Left            =   720
      TabIndex        =   0
      Top             =   1860
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    smlbl3d1.Caption = "Style 0"
    smlbl3d1.Style = 0
    Me.Refresh
End Sub

Private Sub Command2_Click()
    smlbl3d1.Caption = "Style 1"
    smlbl3d1.Style = 1
    Me.Refresh
End Sub


Private Sub Form_Load()
    smlbl3d1.Caption = "Style is " & smlbl3d1.Style
End Sub
