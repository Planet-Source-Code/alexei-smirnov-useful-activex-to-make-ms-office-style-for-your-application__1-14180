VERSION 5.00
Begin VB.UserControl smLabel 
   Alignable       =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   1185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   HasDC           =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   4680
   Begin VB.Line Line41 
      BorderColor     =   &H00808080&
      X1              =   4260
      X2              =   4260
      Y1              =   120
      Y2              =   960
   End
   Begin VB.Line Line31 
      BorderColor     =   &H00FFFFFF&
      X1              =   260
      X2              =   260
      Y1              =   960
      Y2              =   120
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   4320
      Y1              =   140
      Y2              =   140
   End
   Begin VB.Label Label1 
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
      Height          =   735
      Left            =   300
      TabIndex        =   0
      Top             =   180
      Width           =   3975
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   4320
      X2              =   4320
      Y1              =   120
      Y2              =   960
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   240
      X2              =   240
      Y1              =   120
      Y2              =   960
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   4320
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   240
      X2              =   4320
      Y1              =   120
      Y2              =   120
   End
End
Attribute VB_Name = "smLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Default Property Values:
Const m_def_Style = 0
Const m_def_Caption = ""

'Property Variables:
Dim m_Style As Integer
'Dim m_Caption As String

'Event Declarations:
'--------------------------------------------------------------
' ABOUT Event
'--------------------------------------------------------------
Public Sub ShowAboutBox()
Attribute ShowAboutBox.VB_UserMemId = -552
    dlgAbout.Show vbModal
    Unload dlgAbout
    Set dlgAbout = Nothing
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'Dim vDate As Variant, iYear As Integer
'
'    vDate = Date
'    iYear = Year(vDate)
'    If (Month(vDate) > 7 And iYear = 1999) Or iYear > 1999 Then
'        ShowAboutBox
'    End If
    
    'Caption = PropBag.ReadProperty("Caption", "")
    Label1.Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    m_Style = PropBag.ReadProperty("Style", m_def_Style)
    ChangeStyle
    UserControl_Resize
    
End Sub

Private Sub UserControl_Resize()
    
    Line1.X1 = 10
    Line1.X2 = UserControl.ScaleWidth - 10
    Line1.Y1 = 10
    Line1.Y2 = 10
    
    Line2.X1 = 10
    Line2.X2 = UserControl.ScaleWidth - 10
    Line2.Y1 = UserControl.ScaleHeight - 10
    Line2.Y2 = UserControl.ScaleHeight - 10
    
    Line3.X1 = 10
    Line3.X2 = 10
    Line3.Y1 = 10
    Line3.Y2 = UserControl.ScaleHeight - 10
    
    Line4.X1 = UserControl.ScaleWidth - 10
    Line4.X2 = UserControl.ScaleWidth - 10
    Line4.Y1 = 10
    Line4.Y2 = UserControl.ScaleHeight - 10
    
    Select Case m_Style
    Case 0
    
    Case 1
        Line11.X1 = 30
        Line11.X2 = UserControl.ScaleWidth - 15
        Line11.Y1 = 30
        Line11.Y2 = 30
        
        Line31.X1 = 30
        Line31.X2 = 30
        Line31.Y1 = 30
        Line31.Y2 = UserControl.ScaleHeight - 15
        
        Line41.X1 = UserControl.ScaleWidth - 30
        Line41.X2 = UserControl.ScaleWidth - 30
        Line41.Y1 = 10
        Line41.Y2 = UserControl.ScaleHeight - 10
        
    End Select
    
    Label1.Move 50, 50, UserControl.ScaleWidth - 50, UserControl.ScaleHeight - 50
        
End Sub

Public Property Get Caption() As String
Attribute Caption.VB_Description = "smLabel3d Caption"
    Caption = Label1.Caption
End Property

Public Property Let Caption(ByVal vNewValue As String)
    Label1.Caption = vNewValue
    PropertyChanged "Caption"
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", Label1.Caption, m_def_Caption)
    Call PropBag.WriteProperty("Style", m_Style, m_def_Style)
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Caption = m_def_Caption
    m_Style = m_def_Style
    ChangeStyle
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,3,0,0
Public Property Get Style() As Integer
Attribute Style.VB_Description = "smLabel3d Style"
    'If Ambient.UserMode Then Err.Raise 393
    Style = m_Style
End Property

'--------------------------------------------------------------
' Change label3d properties
'--------------------------------------------------------------
Private Sub ChangeStyle()
    
    Select Case m_Style
    Case 0
        Line11.Visible = False
        Line31.Visible = False
        Line41.Visible = False
        Line2.BorderColor = &HFFFFFF
    Case 1
        Line11.Visible = True
        Line31.Visible = True
        Line41.Visible = True
        Line2.BorderColor = &H808080
    End Select
    
End Sub

'--------------------------------------------------------------
' STYLE Property
'--------------------------------------------------------------
Public Property Let Style(ByVal New_Style As Integer)
    
    'If Ambient.UserMode Then Err.Raise 382
    If New_Style <> 0 And New_Style <> 1 Then New_Style = 0
    m_Style = New_Style
    PropertyChanged "Style"
    
    ChangeStyle
    UserControl_Resize
    
End Property

