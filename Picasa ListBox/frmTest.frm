VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTest 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Picasa ListBox Style - siguri92@yahoo.com.vn"
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11175
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   617
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   745
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog Dlg1 
      Left            =   9720
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCM 
      Caption         =   "Clear All Item"
      Height          =   615
      Index           =   1
      Left            =   9360
      TabIndex        =   46
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton cmdCM 
      Caption         =   "Remove Selected Item"
      Height          =   615
      Index           =   0
      Left            =   9360
      TabIndex        =   45
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add 10,000 Items"
      Height          =   495
      Index           =   1
      Left            =   9360
      TabIndex        =   42
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add 5,000 Items"
      Height          =   495
      Index           =   0
      Left            =   9360
      TabIndex        =   41
      Top             =   120
      Width           =   1695
   End
   Begin vbprjPCSListBox.imgList imgLarge 
      Left            =   9360
      Top             =   3960
      _ExtentX        =   2540
      _ExtentY        =   1270
      ImageSize       =   48
      ImageCount      =   1
      Image0          =   "frmTest.frx":0000
      Image1          =   "frmTest.frx":2452
   End
   Begin vbprjPCSListBox.imgList imgSmall 
      Left            =   9360
      Top             =   4800
      _ExtentX        =   4233
      _ExtentY        =   423
      ImageSize       =   16
      ImageCount      =   9
      Image0          =   "frmTest.frx":48A4
      Image1          =   "frmTest.frx":4CF6
      Image2          =   "frmTest.frx":5148
      Image3          =   "frmTest.frx":559A
      Image4          =   "frmTest.frx":59EC
      Image5          =   "frmTest.frx":5E3E
      Image6          =   "frmTest.frx":6290
      Image7          =   "frmTest.frx":66E2
      Image8          =   "frmTest.frx":6B34
      Image9          =   "frmTest.frx":6F86
   End
   Begin VB.Frame frm 
      Caption         =   "Custom All Default Items"
      Height          =   7815
      Index           =   0
      Left            =   5880
      TabIndex        =   1
      Top             =   0
      Width           =   3375
      Begin VB.TextBox txtInput 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   3
         Left            =   2520
         TabIndex        =   40
         Text            =   "4"
         Top             =   7320
         Width           =   735
      End
      Begin VB.TextBox txtInput 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   2
         Left            =   2520
         MaxLength       =   4
         TabIndex        =   39
         Text            =   "30"
         Top             =   6600
         Width           =   735
      End
      Begin VB.TextBox txtInput 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   1
         Left            =   2520
         MaxLength       =   4
         TabIndex        =   38
         Text            =   "25"
         Top             =   5760
         Width           =   735
      End
      Begin VB.TextBox txtInput 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   0
         Left            =   2520
         MaxLength       =   3
         TabIndex        =   37
         Text            =   "20"
         Top             =   2840
         Width           =   735
      End
      Begin VB.HScrollBar hsOp 
         Height          =   135
         Index           =   1
         Left            =   2520
         Max             =   255
         TabIndex        =   36
         Top             =   3280
         Value           =   255
         Width           =   735
      End
      Begin VB.HScrollBar hsOp 
         Height          =   135
         Index           =   0
         Left            =   2520
         Max             =   255
         TabIndex        =   35
         Top             =   2570
         Value           =   255
         Width           =   735
      End
      Begin VB.PictureBox picCol 
         BackColor       =   &H00D3D3D3&
         Height          =   255
         Index           =   7
         Left            =   3000
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   34
         Top             =   6200
         Width           =   255
      End
      Begin VB.PictureBox picCol 
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   3000
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   33
         Top             =   5400
         Width           =   255
      End
      Begin VB.PictureBox picCol 
         BackColor       =   &H00737373&
         Height          =   255
         Index           =   5
         Left            =   3000
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   32
         Top             =   5040
         Width           =   255
      End
      Begin VB.PictureBox picCol 
         BackColor       =   &H008B6425&
         Height          =   255
         Index           =   4
         Left            =   3000
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   31
         Top             =   4680
         Width           =   255
      End
      Begin VB.PictureBox picCol 
         BackColor       =   &H00F3F3F3&
         Height          =   255
         Index           =   3
         Left            =   3000
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   30
         Top             =   4320
         Width           =   255
      End
      Begin VB.PictureBox picCol 
         BackColor       =   &H004D4D4D&
         ForeColor       =   &H004D4D4D&
         Height          =   255
         Index           =   2
         Left            =   3000
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   29
         Top             =   2160
         Width           =   255
      End
      Begin VB.PictureBox picCol 
         BackColor       =   &H004D4D4D&
         Height          =   255
         Index           =   1
         Left            =   3000
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   28
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton cmdAct 
         Caption         =   "Select"
         Height          =   255
         Index           =   4
         Left            =   2520
         TabIndex        =   27
         Top             =   3960
         Width           =   735
      End
      Begin VB.CommandButton cmdAct 
         Caption         =   "Select"
         Height          =   255
         Index           =   3
         Left            =   2520
         TabIndex        =   26
         Top             =   3600
         Width           =   735
      End
      Begin VB.CommandButton cmdAct 
         Caption         =   "Select"
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   25
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton cmdAct 
         Caption         =   "Select"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   24
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdAct 
         Caption         =   "Browse"
         Height          =   255
         Index           =   0
         Left            =   2400
         TabIndex        =   23
         Top             =   720
         Width           =   855
      End
      Begin VB.PictureBox picCol 
         BackColor       =   &H00E8E8E8&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3000
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   22
         Top             =   360
         Width           =   255
      End
      Begin VB.CheckBox chkAS 
         Caption         =   "Auto Spacing Separator"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   7080
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Max Header Pin"
         Height          =   210
         Index           =   18
         Left            =   120
         TabIndex        =   20
         Top             =   7440
         Width           =   1125
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Separator Spacing"
         Height          =   210
         Index           =   17
         Left            =   120
         TabIndex        =   19
         Top             =   6640
         Width           =   1350
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Separator Color"
         Height          =   210
         Index           =   16
         Left            =   120
         TabIndex        =   18
         Top             =   6240
         Width           =   1140
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Item Spacing"
         Height          =   210
         Index           =   15
         Left            =   120
         TabIndex        =   17
         Top             =   5790
         Width           =   915
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Item Selected Fore Color"
         Height          =   210
         Index           =   14
         Left            =   120
         TabIndex        =   16
         Top             =   5400
         Width           =   1755
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Item Fore Color"
         Height          =   210
         Index           =   13
         Left            =   120
         TabIndex        =   15
         Top             =   5040
         Width           =   1080
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Item Selected Back Color"
         Height          =   210
         Index           =   12
         Left            =   120
         TabIndex        =   14
         Top             =   4680
         Width           =   1785
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Item Back Color"
         Height          =   210
         Index           =   11
         Left            =   120
         TabIndex        =   13
         Top             =   4320
         Width           =   1110
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Item Selected Font"
         Height          =   210
         Index           =   10
         Left            =   120
         TabIndex        =   12
         Top             =   3960
         Width           =   1320
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Item Font"
         Height          =   210
         Index           =   9
         Left            =   120
         TabIndex        =   11
         Top             =   3600
         Width           =   645
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Item Opacity"
         Height          =   210
         Index           =   8
         Left            =   120
         TabIndex        =   10
         Top             =   3240
         Width           =   885
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Item Height"
         Height          =   210
         Index           =   7
         Left            =   120
         TabIndex        =   9
         Top             =   2880
         Width           =   780
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Header Opacity"
         Height          =   210
         Index           =   6
         Left            =   120
         TabIndex        =   8
         Top             =   2520
         Width           =   1125
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Header Expand Fore Color"
         Height          =   210
         Index           =   5
         Left            =   120
         TabIndex        =   7
         Top             =   2160
         Width           =   1905
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Header Fore Color"
         Height          =   210
         Index           =   4
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   1320
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Header Expand Font"
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   1470
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Header Font"
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   885
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Background Picture"
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1410
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "BackColor"
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
   End
   Begin vbprjPCSListBox.pcsListBox pcs 
      Height          =   9015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   15901
      BackColor       =   15263976
      Background      =   "frmTest.frx":73D8
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HeaderExpandFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderForeColor =   5066061
      HeaderExpandForeColor=   5066061
      HeaderOpacity   =   255
      MaxHeaderPin    =   4
      BeginProperty ItemFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty ItemSelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ItemHeight      =   20
      ItemBackColor   =   15987699
      ItemForeColor   =   7566195
      ItemSelBackColor=   9135141
      ItemSelForeColor=   16777215
      ItemOpacity     =   255
      ItemSpacing     =   25
      SeparatorColor  =   13882323
      SeparatorSpacing=   30
      AutoSpacingSeparator=   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   $"frmTest.frx":73F4
      Height          =   1695
      Left            =   9360
      TabIndex        =   47
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label lblInf 
      Caption         =   "Action : None"
      Height          =   255
      Index           =   1
      Left            =   5880
      TabIndex        =   44
      Top             =   8880
      Width           =   3375
   End
   Begin VB.Label lblInf 
      Caption         =   "Total :"
      Height          =   975
      Index           =   0
      Left            =   5880
      TabIndex        =   43
      Top             =   7920
      Width           =   3375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00D3D3D3&
      Height          =   9045
      Left            =   105
      Top             =   105
      Width           =   5685
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Const strH1$ = "All Programs;Accessories[s];Calculator;Command Prompt;Notepad;Paint;Snipping Tool;Sticky Notes;Windows Explorer;Word Pad"
Private Const strH2$ = "Programs[s];CCleaner;Format Factory;Internet Download Manager;Microsoft Office;My SQL;NetBeans;WampServer"
Private Const colCh$ = "16443110;16775416;16775408;16449525;15794160;13826810;13499135;14481663;14745599;15794175"
Private Const strWs$ = "http://ask.com;http://imageshack.us;http://sourceforge.net;http://goole.com.vn;http://mail.yahoo.com;http://caulacbovb.com;http://vbforums.com;http://planetsourcecode.com"

Private Sub chkAS_Click()
    pcs.AutoSpacingSeparator = CBool(chkAS.Value)
    pcs.RedrawAll
End Sub

Private Sub cmdAct_Click(Index As Integer)
    Dim font As StdFont
    If Index = 0 _
        Then _
            Dlg1.Filter = "Image File (*.bmp;*.jpg;*.gif;*.ico)|*.bmp;*.jpg;*.gif;*.ico"
            Dlg1.ShowOpen
            If Dlg1.FileName = "" Then Exit Sub
            Set pcs.Background = LoadPicture(Dlg1.FileName)
            pcs.RedrawAll
            If MsgBox("You need reduce opacity to see background picture! Reduce now?", _
                        vbInformation + vbYesNo) = vbYes Then
                hsOp(0).Value = 200
                hsOp(1).Value = 200
            End If
            Exit Sub
    'Select font
    Dlg1.ShowFont
    Set font = Me.font
    With font
        .Bold = Dlg1.FontBold
        .Italic = Dlg1.FontItalic
        .Name = Dlg1.FontName
        .Size = Dlg1.FontSize
        .Strikethrough = Dlg1.FontStrikethru
        .Underline = Dlg1.FontUnderline
    End With
    SetFont Index - 1, font
End Sub

Private Sub cmdAdd_Click(Index As Integer)
    Dim i&, j&, t&
    t = GetTickCount
    For j = 0 To 50
        pcs.AddHeader "k" & j, "New Header " & Format(j, ("00"))
        For i = 1 To IIf(Index = 0, 100, 200)
            pcs.AddItem pcs.HeaderCount - 1, "Item - " & Format(i, ("0000"))
        Next
    Next j
    pcs.RedrawAll   'Remember
    
    MsgBox "Fill " & IIf(Index = 0, 5000, 10000) & " items with 50 headers in " & (GetTickCount - t) / 1000 & " sec"
End Sub

Private Sub cmdCM_Click(Index As Integer)
    Dim a
    a = Split(pcs.SelectedItem, ":")
    Select Case Index
        Case 0: pcs.RemoveItem CLng(a(0)), CLng(a(1))
        Case 1: pcs.Clear
    End Select
End Sub

Private Sub Form_Load()
    Dim i%, a, tmp$
    
    ' Init listbox
    pcs.InitializeListBox
    
    ' Add image
    
    ' Large
    For i = 0 To imgLarge.ImageCount - 1
        pcs.AddImage imgLarge.ImageData(i)
    Next
    
    ' Small
    For i = 0 To imgSmall.ImageCount - 1
        pcs.AddImage imgSmall.ImageData(i)
    Next
    
    a = Split(strH1, ";")
    
    ' Add non-header items
    Dim arr
    arr = Array("Information", "Internet Explorer")
    For i = 0 To UBound(arr)
        pcs.AddItem -1, arr(i), i
        ' Custom item size
        pcs.CustomItemHeight(-1, i) = 52
        ' Custom item spacing
        pcs.CustomItemSpacing(-1, i) = 52
        ' Custom font
        Set pcs.CustomItemFont(-1, i) = picCol(0).font
        Set pcs.CustomItemSelFont(-1, i) = picCol(0).font
    Next
    
    ' Add new header
    pcs.AddHeader "ALLP", a(0) & " (Normal Item && Separator Type)"
    
    ' Add header items (relative can be header key or header index)
    For i = 1 To UBound(a)
        tmp = Replace$(a(i), "[s]", "")
        pcs.AddItem 0, tmp, imgLarge.ImageCount - 2 + i, Right$(a(i), 3) = "[s]"
    Next
    
    a = Split(strH2, ";")
    
    ' Add header items
    For i = 0 To UBound(a)
        tmp = Replace$(a(i), "[s]", "")
        pcs.AddItem 0, tmp, 10, Right$(a(i), 3) = "[s]"
    Next
    
    ' Add colorful header
    pcs.AddHeader "CORL", "Colorful Items (Custom BackColor)"
    a = Split(colCh, ";")
    
    ' Add colorful items
    For i = 0 To UBound(a)
        pcs.AddItem 1, "Colorful item (" & i & ")"
        ' Custom item back color & selected back color
        pcs.CustomItemBackColor(1, i) = CLng(a(i))
        pcs.CustomItemSelBackColor(1, i) = DarkenColor(CLng(a(i)), 20)
        ' Custom item selected fore color
        pcs.CustomItemSelForeColor(1, i) = DarkenColor(CLng(a(i)), 100)
    Next
    
    ' Add website header
    pcs.AddHeader "WEBS", "Web sites"
    a = Split(strWs, ";")
    
    ' Add website items
    For i = 0 To UBound(a)
        pcs.AddItem 2, a(i), 11
        ' Colorize ;)
        pcs.CustomItemBackColor(2, i) = IIf(i Mod 2, 16777200, 15130800)
    Next
    
    ' Sorting website items
    pcs.SortHeaderItems (2)
    
    pcs.RedrawAll
End Sub

Private Function DarkenColor&(ByVal Color&, Value%)
    Dim R%, G%, B%
    Long2RGB Color, R, G, B
    R = R - Value: If R < 0 Then R = 0
    G = G - Value: If G < 0 Then G = 0
    B = B - Value: If B < 0 Then B = 0
    DarkenColor = RGB(R, G, B)
End Function

' Convert long color to r,g,b color
Private Sub Long2RGB(ByVal LongValue As Long, R As Integer, G As Integer, B As Integer)
On Error Resume Next
    R = (LongValue And &HFF)
    G = (((LongValue And &HFF00) - (LongValue And &HFF0000)) \ 256)
    B = ((LongValue And &HFF0000) \ 65536)
End Sub

Private Sub hsOp_Change(Index As Integer)
    hsOp_Scroll Index
End Sub

Private Sub hsOp_Scroll(Index As Integer)
    Select Case Index
        Case 0: pcs.HeaderOpacity = hsOp(Index).Value
        Case 1: pcs.ItemOpacity = hsOp(Index).Value
    End Select
    pcs.RedrawAll
End Sub

Private Sub pcs_Change()
    Dim i&, n As Double
    n = 0
    For i = 0 To pcs.HeaderCount - 1
        n = n + pcs.ItemCount(i)
    Next
    lblInf(0).Caption = "No header items : " & pcs.ItemCount(-1) & vbCrLf & _
                        "Total headers : " & pcs.HeaderCount & vbCrLf & _
                        "Total all headers items : " & n
End Sub

Private Sub pcs_HeaderClick(Index As Long)
    lblInf(1).Caption = "Action : " & "Header Click " & Index
End Sub

Private Sub pcs_ItemClick(Relative As Long, Index As Long)
    lblInf(1).Caption = "Action : " & "Item Click " & Relative & ":" & Index
End Sub

Private Sub pcs_Scroll()
    lblInf(1).Caption = "Action : " & "Scrolling"
End Sub

Sub ChangeColor(ByVal Index%, lColor&)
    Select Case Index
        ' Back color
        Case 0: pcs.BackColor = lColor
        ' Header fore color
        Case 1: pcs.HeaderForeColor = lColor
        ' Header selected fore color
        Case 2: pcs.HeaderExpandForeColor = lColor
        ' Item back color
        Case 3: pcs.ItemBackColor = lColor
        ' Item selected back color
        Case 4: pcs.ItemSelBackColor = lColor
        ' Item fore color
        Case 5: pcs.ItemForeColor = lColor
        ' Item selected fore color
        Case 6: pcs.ItemSelForeColor = lColor
        ' Separator color
        Case 7: pcs.SeparatorColor = lColor
    End Select
    pcs.RedrawAll
End Sub

Sub SetFont(ByVal Index%, font As StdFont)
    Select Case Index
        ' Header font
        Case 0: Set pcs.HeaderFont = font
        ' Header expand font
        Case 1: Set pcs.HeaderExpandFont = font
        ' Item font
        Case 2: Set pcs.ItemFont = font
        ' Item selected font
        Case 3: Set pcs.ItemSelFont = font
    End Select
    pcs.RedrawAll
End Sub

Private Sub picCol_Click(Index As Integer)
    Dlg1.Color = picCol(Index).BackColor
    Dlg1.ShowColor
    If Dlg1.Color <> picCol(Index).BackColor Then
        ChangeColor Index, Dlg1.Color
        picCol(Index).BackColor = Dlg1.Color
    End If
End Sub

Private Sub txtInput_Change(Index As Integer)
On Error Resume Next
    Select Case Index
        Case 0 ' Item height
            pcs.ItemHeight = CSng(txtInput(Index).Text)
        Case 1 ' Item spacing left
            pcs.ItemSpacing = CSng(txtInput(Index).Text)
        Case 2 ' Separator spacing
            pcs.SeparatorSpacing = CSng(txtInput(Index).Text)
            chkAS.Value = 0: Call chkAS_Click
        Case 3 ' Max pin
            pcs.MaxHeaderPin = CInt(txtInput(Index).Text)
    End Select
    pcs.RedrawAll
End Sub

Private Sub txtInput_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr(1, "1234567890" & Chr$(vbKeyBack), Chr$(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
