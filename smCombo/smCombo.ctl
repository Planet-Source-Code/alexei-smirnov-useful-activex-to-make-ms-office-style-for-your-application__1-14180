VERSION 5.00
Begin VB.UserControl smCombo 
   ClientHeight    =   2100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   2100
   ScaleWidth      =   4800
   Begin VB.Image imgBut 
      Height          =   315
      Left            =   1800
      Top             =   240
      Width           =   315
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   0
      Picture         =   "smCombo.ctx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   0
      Picture         =   "smCombo.ctx":014E
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   1755
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   45
   End
End
Attribute VB_Name = "smCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Default Property Values:
Const m_def_ItemData = 0
Const m_def_Caption = ""
Const m_def_Text = ""
Const m_def_Enabled = True

'Property Variables:
Dim m_ItemData As Variant
'Dim m_Caption As String
'Dim m_Text As String

'Event Declarations:
Event GetMore()

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,2,0
Public Property Get ItemData() As Variant
Attribute ItemData.VB_Description = "Returns/sets the Combo ItemData"
Attribute ItemData.VB_MemberFlags = "400"
    ItemData = m_ItemData
End Property

Public Property Let ItemData(ByVal New_ItemData As Variant)
    If Ambient.UserMode = False Then Err.Raise 387
    m_ItemData = New_ItemData
    PropertyChanged "ItemData"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets the Combo Enabled"
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";List"
Attribute Enabled.VB_UserMemId = -514
    Enabled = imgBut.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    ChFace New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the Combo Caption"
    Caption = Label2.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    'm_Caption = New_Caption
    Label2.Caption = New_Caption
    PropertyChanged "Caption"
    UserControl_Resize
End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the Combo Text"
    'Text = m_Text
    Text = Label1.Caption
End Property

Public Property Let Text(ByVal New_Text As String)
    'm_Text = New_Text
    Label1.Caption = New_Text
    PropertyChanged "Text"
End Property

Private Sub ImgBut_Click()
    RaiseEvent GetMore
End Sub

Private Sub ChFace(ByVal Enabled As Boolean)

    If Enabled Then
        imgBut.Picture = Image1.Picture
        With Label1
            '.BackColor = &HFFFFFF
            .ForeColor = &H80000012
        End With
    Else
        imgBut.Picture = Image2.Picture
        With Label1
            '.BackColor = &H8000000F
            .ForeColor = &H808080
        End With
    End If
    imgBut.Enabled = Enabled
    
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_ItemData = m_def_ItemData
    Label1.Caption = m_def_Caption
    Label2.Caption = m_def_Text
    ChFace m_def_Enabled
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

Dim bEn As Boolean

    m_ItemData = PropBag.ReadProperty("ItemData", m_def_ItemData)
    Label2.Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    Label1.Caption = PropBag.ReadProperty("Text", m_def_Text)
    bEn = PropBag.ReadProperty("Enabled", m_def_Enabled)
    ChFace bEn

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    Call PropBag.WriteProperty("ItemData", m_ItemData, m_def_ItemData)
    Call PropBag.WriteProperty("Caption", Label2.Caption, m_def_Caption)
    Call PropBag.WriteProperty("Text", Label1.Caption, m_def_Text)
    Call PropBag.WriteProperty("Enabled", imgBut.Enabled, m_def_Enabled)
    ChFace imgBut.Enabled
    
End Sub

Private Sub UserControl_Resize()

Dim lw&
    
    lw = UserControl.ScaleWidth
        
    If Len(Label2.Caption) Then
        UserControl.Height = 555
        Label1.Move 0, 240, lw
        imgBut.Move lw - 280, 270
    Else
        UserControl.Height = 315
        Label1.Move 0, 0, lw
        imgBut.Move lw - 280, 30
    End If
    
End Sub

'--------------------------------------------------------------
' ABOUT Event
'--------------------------------------------------------------
Public Sub ShowAboutBox()
Attribute ShowAboutBox.VB_UserMemId = -552
    dlgAbout.Show vbModal
    Unload dlgAbout
    Set dlgAbout = Nothing
End Sub
