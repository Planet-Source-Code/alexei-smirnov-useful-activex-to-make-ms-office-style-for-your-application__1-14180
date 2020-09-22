VERSION 5.00
Object = "{FE39306C-FD26-11D2-8003-00104BD28E91}#7.0#0"; "smCombo.ocx"
Object = "*\AsmProgress\smProgress.vbp"
Object = "*\AsmLabel3d\smLabel3d.vbp"
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ActiveX Test"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9060
   Icon            =   "Test.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   9060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command10 
      Caption         =   "Caption"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   11
      Top             =   4380
      Width           =   1035
   End
   Begin smComboBox.smCombo smCombo1 
      Height          =   315
      Left            =   240
      TabIndex        =   10
      Top             =   1860
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   556
   End
   Begin smLabel3d.smLabel smLabel1 
      Height          =   315
      Left            =   240
      Top             =   240
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   556
   End
   Begin smProgress.smProgresser smProgresser1 
      Height          =   315
      Left            =   240
      TabIndex        =   9
      Top             =   2940
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   556
      Caption         =   "Progress..."
      Style           =   1
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   3960
      Width           =   1035
   End
   Begin VB.CommandButton Command8 
      Caption         =   "StepIt!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   3540
      Width           =   1035
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Alignment"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   4440
      Width           =   1035
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Style"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   4020
      Width           =   1035
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1560
      Top             =   4020
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Progress"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   3600
      Width           =   1035
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Enabled"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2580
      TabIndex        =   3
      Top             =   2220
      Width           =   1035
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ItemData"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2580
      TabIndex        =   2
      Top             =   1800
      Width           =   1035
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Style 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2580
      TabIndex        =   1
      Top             =   660
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Style 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2580
      TabIndex        =   0
      Top             =   240
      Width           =   1035
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    smLabel1.Caption = "Style is 0"
    smLabel1.Style = 0
End Sub

Private Sub Command10_Click()
    smProgresser1.Caption = "Goto..."
End Sub

Private Sub Command2_Click()
    smLabel1.Caption = "Style is 1"
    smLabel1.Style = 1
End Sub

Private Sub Command3_Click()
    smCombo1.Enabled = Not smCombo1.Enabled
End Sub

Private Sub Command4_Click()
    With smCombo1
        If .ItemData = 12345 Then
            .ItemData = "Text string"
        Else
            .ItemData = 12345
        End If
        .Text = .ItemData
    End With
End Sub

Private Sub Command5_Click()
    With smProgresser1
        .Position = 0
        .Max = 103
    End With
    Timer1.Enabled = True
End Sub

Private Sub Command6_Click()
    smProgresser1.Style = IIf(smProgresser1.Style = 0, 1, 0)
End Sub

Private Sub Command7_Click()
    
Dim ix As Byte

    ix = smProgresser1.Alignment
    ix = ix + 1
    If ix > 2 Then ix = 0
    smProgresser1.Alignment = ix
    
End Sub

Private Sub Command8_Click()
    
    smProgresser1.Max = 10
    smProgresser1.StepIt
    
End Sub

Private Sub Command9_Click()
    Timer1.Enabled = False
    smProgresser1.Position = 0
End Sub

Private Sub Form_Load()
    smLabel1.Caption = "This is a Smirnoff Label3d Control. Style is " & smLabel1.Style
    smCombo1.Caption = "ComboBox Caption"
End Sub

Private Sub smCombo1_GetMore()
    MsgBox "Get Me More!"
    smCombo1.Text = "Get Me More!"
End Sub

Private Sub Timer1_Timer()
    smProgresser1.Position = smProgresser1.Position + 1
End Sub
