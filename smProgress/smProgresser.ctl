VERSION 5.00
Begin VB.UserControl smProgresser 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   4320
      X2              =   4320
      Y1              =   240
      Y2              =   1080
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   4320
      Y1              =   1020
      Y2              =   1020
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   240
      X2              =   240
      Y1              =   240
      Y2              =   1080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   240
      X2              =   4320
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "ss"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   300
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800000&
      Height          =   375
      Left            =   300
      TabIndex        =   1
      Top             =   360
      Width           =   1395
   End
End
Attribute VB_Name = "smProgresser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Default Property Values:
Const m_def_Style = 0
Const m_def_Caption = ""
Const m_def_Position = 0
Const m_def_Max = 0
Const m_def_Step = 1
Const m_def_Align = 2

'Property Variables:
Dim m_Style As Byte
Dim m_Position As Long
Dim m_Caption As String
Dim m_Max As Long
Dim m_Alignment As Byte

Dim m_Step As Integer

Public Property Get Step() As Integer
Attribute Step.VB_Description = "Set a step size"
    Step = m_Step
End Property

Public Property Let Step(ByVal vNewValue As Integer)
    If vNewValue > 0 Then
        m_Step = vNewValue
        PropertyChanged "Step"
    End If
End Property

'--------------------------------------------------------------
' StepIt Event
'--------------------------------------------------------------
Public Sub StepIt()
Attribute StepIt.VB_Description = "StepIt Event"

Dim CurPos&
        
    CurPos = Position
    Position = CurPos + m_Step
    
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

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Alignment = PropBag.ReadProperty("Alignment", m_def_Align)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    m_Style = PropBag.ReadProperty("Style", m_def_Style)
    m_Step = PropBag.ReadProperty("Step", m_def_Step)
    m_Max = PropBag.ReadProperty("Max", m_def_Max)
    ChangeStyle
    
End Sub

Private Sub UserControl_Resize()
    
    UserControl.Height = 315
    
    Line1.X1 = 10
    Line1.X2 = UserControl.ScaleWidth - 10
    Line1.Y1 = 10
    Line1.Y2 = 10
    
    Line2.X1 = 10
    Line2.X2 = UserControl.ScaleWidth - 10
    Line2.Y1 = 305
    Line2.Y2 = 305
    
    Line3.X1 = 10
    Line3.X2 = 10
    Line3.Y1 = 10
    Line3.Y2 = 305
    
    Line4.X1 = UserControl.ScaleWidth - 10
    Line4.X2 = UserControl.ScaleWidth - 10
    Line4.Y1 = 10
    Line4.Y2 = 305
    
    Label1.Move 50, 50, UserControl.ScaleWidth - 50, 265
    Label2.Move 50, 50
    Label2.Height = 235
    ChPosition
            
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal vNewValue As String)
    m_Caption = vNewValue
    PropertyChanged "Caption"
    ChangeStyle
End Property

Public Property Get Position() As Long
    Position = m_Position
End Property

Public Property Let Position(ByVal vNewValue As Long)
    m_Position = vNewValue
    PropertyChanged "Position"
    ChPosition
End Property

Private Sub ChPosition()

Dim lVal!

    If m_Position <= 0 Or m_Max <= 0 Then
        Label2.Visible = False
        If m_Style = 0 Then Label1.Caption = "0%"
    Else
        Label2.Visible = True
        If m_Position > m_Max Then
            If m_Style = 0 Then Label1.Caption = "100%"
        Else
            lVal = m_Position / m_Max
            Label2.Width = (UserControl.ScaleWidth - 80) * lVal
            If m_Style = 0 Then Label1.Caption = Round(lVal * 100) & "%"
        End If
    End If
    
End Sub

Public Property Get Max() As Long
    Max = m_Max
End Property

Public Property Let Max(ByVal vNewValue As Long)
    m_Max = vNewValue
    PropertyChanged "Max"
    ChPosition
End Property

Public Property Get Alignment() As Byte
    Alignment = m_Alignment
End Property

Public Property Let Alignment(ByVal vNewValue As Byte)
    If vNewValue >= 0 And vNewValue <= 2 Then
        m_Alignment = vNewValue
        PropertyChanged "Alignment"
        Label1.Alignment = vNewValue
    End If
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("Style", m_Style, m_def_Style)
    Call PropBag.WriteProperty("Alignment", m_Alignment, m_def_Align)
    Call PropBag.WriteProperty("Step", m_Step, m_def_Step)
    Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Alignment = m_def_Align
    m_Caption = m_def_Caption
    m_Style = m_def_Style
    m_Step = m_def_Step
    m_Max = m_def_Max
    ChangeStyle
End Sub

Public Property Get Style() As Byte
Attribute Style.VB_Description = "Style Property"
    Style = m_Style
End Property

'--------------------------------------------------------------
' Change smProgress properties
'--------------------------------------------------------------
Private Sub ChangeStyle()
    
    Select Case m_Style
    Case 0
        With Label1
            If m_Position <= 0 And m_Max <= 0 Then
                .Caption = "0%"
            Else
                .Caption = Round(m_Position * 100 / m_Max) & "%"
            End If
        End With
    Case 1
        With Label1
            .Caption = m_Caption
        End With
    End Select
    Label1.Alignment = m_Alignment
    
End Sub

'--------------------------------------------------------------
' STYLE Property
'--------------------------------------------------------------
Public Property Let Style(ByVal New_Style As Byte)
    
    If New_Style <> 0 And New_Style <> 1 Then New_Style = 0
    
    m_Style = New_Style
    PropertyChanged "Style"
    
    ChangeStyle
        
End Property

