VERSION 5.00
Begin VB.UserControl pcsScrollbar 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1320
   ScaleHeight     =   102
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   88
   Begin VB.ComboBox cb 
      Height          =   315
      ItemData        =   "pcsScrollbar.ctx":0000
      Left            =   1920
      List            =   "pcsScrollbar.ctx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   120
      Picture         =   "pcsScrollbar.ctx":0021
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Timer tmrDey 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   600
      Top             =   960
   End
   Begin VB.Timer tmrOver 
      Enabled         =   0   'False
      Interval        =   3
      Left            =   120
      Top             =   960
   End
End
Attribute VB_Name = "pcsScrollbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
    X       As Long
    Y       As Long
End Type

Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

' API Declarations
Private Declare Function SetRect Lib "user32" (lpRect As RECT, _
    ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, _
    ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, _
    ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
    ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
    
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Enum ENUMPCSSBTAGSTATE
    [pstNormal]
    [pstUpOver]
    [pstUpDown]
    [pstDownOver]
    [pstDownDown]
    [pstMoveBar]
End Enum

Dim m_rDown             As RECT
Dim m_rBar              As RECT

Dim m_eState            As ENUMPCSSBTAGSTATE

Dim m_Max               As Double
Dim m_Value             As Double
Dim m_Subtract          As Single

Dim mDC                 As Long

Event Scroll()
Event MouseLeave()

' Draw picasa scrollbar
Private Sub DrawScrollbar()
    Dim lW&, lH&, mbT&, mbH&
    
    UserControl.Cls
    
    lW = UserControl.ScaleWidth
    lH = UserControl.ScaleHeight
    
    ' Set rect to draw
    SetRect m_rDown, 0, lH - 24, 0, 24                                    ' Down
    
    mbH = (lH - 48) / m_Max
    If mbH < 30 Then mbH = 30
    mbT = 24 + (m_Value / m_Max * (lH - 48 - mbH))
    If mbT < 24 Then mbT = 24
    If mbT > lH - 24 - mbH Then mbT = lH - 24 - mbH
    SetRect m_rBar, 0, mbT, 16, mbH

    ' Draw background
    StretchBlt hDC, 0, 0, 16, lH, mDC, 48, 24, 16, 24, vbSrcCopy
    
    ' Draw button
    
    BitBlt hDC, 0, 0, 16, 24, mDC, 0, 0, vbSrcCopy                        ' Up
    BitBlt hDC, 0, m_rDown.Top, 16, 24, mDC, 0, 24, vbSrcCopy             ' Down
    
    Select Case m_eState
        Case 1: ' Up-Over
            BitBlt hDC, 0, 0, 16, 24, mDC, 16, 0, vbSrcCopy
        Case 2: ' Up-Down
            BitBlt hDC, 0, 0, 16, 24, mDC, 32, 0, vbSrcCopy
        Case 3: ' Down-Over
            BitBlt hDC, 0, m_rDown.Top, 16, 24, mDC, 16, 24, vbSrcCopy
        Case 4: ' Down-Down
            BitBlt hDC, 0, m_rDown.Top, 16, 24, mDC, 32, 24, vbSrcCopy
    End Select
    
    ' Draw bar
    StretchBlt hDC, 0, mbT, 16, mbH, mDC, 48, 2, 16, 2, vbSrcCopy         ' Middle
    BitBlt hDC, 0, mbT, 16, 2, mDC, 48, 0, vbSrcCopy                      ' Top
    BitBlt hDC, 0, mbT + mbH - 2, 16, 2, mDC, 48, 22, vbSrcCopy           ' Bottom
    
    UserControl.Refresh
End Sub

' Check mouse over state
Private Function isMouseOver() As Boolean
On Error GoTo NoMouse:
    Dim mPoint As POINTAPI
    GetCursorPos mPoint
    isMouseOver = WindowFromPoint(mPoint.X, mPoint.Y) = hWnd
NoMouse: End Function

Private Sub cb_Change()
    cb_Click
End Sub

Private Sub cb_Click()
    If cb.ListIndex = 0 Then Value = Value - 1
    If cb.ListIndex = 2 Then Value = Value + 1
    cb.ListIndex = 1
End Sub

Private Sub tmrDey_Timer()
    Static n%
    If m_eState <> pstUpDown And m_eState <> pstDownDown Then
        n = 0
        tmrDey.Enabled = False
    End If
    If n < 1000 Then n = n + 50
    If n >= 1000 Then
        If m_eState = pstUpDown Then Value = Value - 1
        If m_eState = pstDownDown Then Value = Value + 1
    End If
End Sub

' Mouse leave state
Private Sub tmrOver_Timer()
    If Not isMouseOver And m_eState <> pstMoveBar Then
        m_eState = pstNormal
        DrawScrollbar
        tmrOver.Enabled = False
    End If
End Sub

Private Sub UserControl_InitProperties()
On Error Resume Next
    CreateScrollbar
    Max = 100
    Value = 0
    cb.ListIndex = 1
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Value to subtract
    If Button = vbLeftButton Then
        If (X >= 0 And X <= 16) Then
            If (Y >= m_rBar.Top And Y <= m_rBar.Top + m_rBar.Bottom) Then
                m_eState = pstMoveBar
                m_Subtract = Y - m_rBar.Top
            End If
        End If
    End If
    Call UserControl_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tmrOver.Enabled = True
    
    If m_eState <> pstMoveBar Then
        ' Up/Down button state
        If (X >= 0 And X <= 16) Then
            If (Y >= 0 And Y <= 24) Then
                m_eState = IIf(Button = vbLeftButton, pstUpDown, pstUpOver)
            ElseIf (Y >= m_rDown.Top And Y <= m_rDown.Top + 24) Then
                m_eState = IIf(Button = vbLeftButton, pstDownDown, pstDownOver)
            Else
                m_eState = pstNormal
            End If
        End If
    
        If Not isMouseOver Then m_eState = pstNormal
        
    Else
        ' Move scrollbar
        Dim valTmp As Double
        valTmp = FormatNumber((Y - (24 + m_Subtract)) / _
                    (UserControl.ScaleHeight - 48 - m_rBar.Bottom) * m_Max, 5)
        If valTmp < 0 Then valTmp = 0
        If valTmp > m_Max Then valTmp = m_Max
        Value = valTmp
    End If
    
    If tmrDey.Enabled = False Then
        If m_eState = pstUpDown Then Value = Value - 1
        If m_eState = pstDownDown Then Value = Value + 1
        tmrDey.Enabled = True
    End If
    
    If isMouseOver Then cb.SetFocus
    
    DrawScrollbar
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Restore normal state
    m_eState = pstNormal
    Call UserControl_MouseMove(vbRightButton, Shift, X, Y)
End Sub

' Limited control size
Private Sub UserControl_Resize()
On Error Resume Next
    UserControl.Width = 240
    DrawScrollbar
End Sub

Private Function hDC&()
    hDC = UserControl.hDC
End Function

Private Function hWnd&()
    hWnd = UserControl.hWnd
End Function

' Max property
Public Property Get Max() As Double
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Double)
    If New_Max < 1 Then New_Max = 1
    If Value > New_Max Then Value = New_Max
    m_Max = New_Max
    DrawScrollbar
    PropertyChanged "Max"
End Property

' Value property
Public Property Get Value() As Double
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Double)
    If New_Value < 0 Then New_Value = 0
    If New_Value > m_Max Then New_Value = m_Max
    m_Value = New_Value
    RaiseEvent Scroll
    DrawScrollbar
    PropertyChanged "Value"
End Property

Private Sub UserControl_Show()
    CreateScrollbar
    DrawScrollbar
    cb.ListIndex = 1
End Sub

Private Sub UserControl_Terminate()
On Error Resume Next
    DeleteDC mDC
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
    PropBag.WriteProperty "Max", m_Max, 0
    PropBag.WriteProperty "Value", m_Value, 0
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    m_Max = PropBag.ReadProperty("Max", 0)
    m_Value = PropBag.ReadProperty("Value", 0)
End Sub

Private Sub CreateScrollbar()
On Error Resume Next
    Dim mBitmap
    mDC = CreateCompatibleDC(UserControl.hDC)
    mBitmap = CreateCompatibleBitmap(UserControl.hDC, 64, 48)
    SelectObject mDC, mBitmap
    BitBlt mDC, 0, 0, 64, 48, pic.hDC, 0, 0, vbSrcCopy
    DeleteObject mBitmap
End Sub

