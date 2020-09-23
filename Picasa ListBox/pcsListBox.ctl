VERSION 5.00
Begin VB.UserControl pcsListBox 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1260
   ScaleHeight     =   108
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   84
   Begin vbprjPCSListBox.pcsScrollbar pSB 
      Height          =   1215
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   2143
      Max             =   100
   End
   Begin VB.PictureBox pHdr 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      DrawWidth       =   3550
      Height          =   330
      Left            =   0
      Picture         =   "pcsListBox.ctx":0000
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.PictureBox picDraw 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   0
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "pcsListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'PICASA LISTBOX - GOOGLE PICASA LISTBOX STYLE
'CREATE BY SIGURI92
'E-MAIL: SIGURI92@YAHOO.COM.VN

Option Explicit

Private Type RGBTRIPLE
    rgbBlue              As Byte
    rgbGreen             As Byte
    rgbRed               As Byte
End Type

Private Type RGBQUAD
    rgbBlue              As Byte
    rgbGreen             As Byte
    rgbRed               As Byte
    rgbAlpha             As Byte
End Type

Private Type BITMAP
    bmType               As Long
    bmWidth              As Long
    bmHeight             As Long
    bmWidthBytes         As Long
    bmPlanes             As Integer
    bmBitsPixel          As Integer
    bmBits               As Long
End Type

Private Type BITMAPINFOHEADER
    biSize               As Long
    biWidth              As Long
    biHeight             As Long
    biPlanes             As Integer
    biBitCount           As Integer
    biCompression        As Long
    biSizeImage          As Long
    biXPelsPerMeter      As Long
    biYPelsPerMeter      As Long
    biClrUsed            As Long
    biClrImportant       As Long
End Type

Private Type BITMAPINFO
    bmiHeader            As BITMAPINFOHEADER
    bmiColors            As RGBTRIPLE
End Type

' API TransparentBlt
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, _
    ByVal hPalette As Long, ByRef pccolorref As Long) As Long
    
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long) As Long
    
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, _
    ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, _
    ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
    
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long

Private Declare Function GetObject Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, _
    ByVal nCount As Long, ByRef lpObject As Any) As Long
    
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, _
    ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
    ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
    
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, _
    ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, _
    lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
    
Private Declare Function GetNearestColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long

Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, _
    ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, _
    ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, _
    Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long

Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, _
    ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

' API Alpha Blending
Private Const AC_SRC_OVER = &H0

Private Type BLENDFUNCTION
    BlendOp                 As Byte
    BlendFlags              As Byte
    SourceConstantAlpha     As Byte
    AlphaFormat             As Byte
End Type

Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hDC As Long, ByVal lInt As Long, _
    ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal hDC As Long, _
    ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, _
    ByVal BLENDFUNCT As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)
    
Private Type RECT
    Left            As Long
    Top             As Long
    Right           As Long
    Bottom          As Long
End Type

Private Declare Function DrawText Lib "user32" Alias "DrawTextW" (ByVal hDC As Long, _
    ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, _
    ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

' API For Custom Draws
Private Type POINTAPI
    X               As Long
    Y               As Long
End Type

Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, _
    ByVal hBrush As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, _
    lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, _
    ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, _
    ByVal nWidth As Long, ByVal crColor As Long) As Long
' Item type
Private Type TYPEPCSLISTBOXITEM
    StrText                 As Variant
    intImage                As Integer
    blnSeparator            As Boolean
    ' Customs
    sngHeight               As Single
    sngSpacing              As Single
    lngFont                 As font
    lngSelFont              As font
    lngBackColor            As OLE_COLOR
    lngForeColor            As OLE_COLOR
    lngSelBackColor         As OLE_COLOR
    lngSelForeColor         As OLE_COLOR
    cRect                   As RECT
    ' Using default
    b_defHeight             As Boolean
    b_defFont               As Boolean
    b_defSelFont            As Boolean
    b_defBackColor          As Boolean
    b_defForeColor          As Boolean
    b_defSelBackColor       As Boolean
    b_defSelForeColor       As Boolean
    b_defSpacing            As Boolean
    ' Event
    b_Selected              As Boolean
End Type

' Header type
Private Type TYPEPCSLISTBOXHEADER
    strKey                  As String
    StrText                 As String
    blnExpand               As Boolean
    blnPin                  As Boolean
    cRect                   As RECT
    dblTotalHeight          As Double
    colItems()              As TYPEPCSLISTBOXITEM
    lngItemCount            As Long
    b_isPin                 As Boolean
    ' Customs
    lngFont                 As font
    lngExpandFont           As font
    lngForeColor            As OLE_COLOR
    lngExpandForeColor      As OLE_COLOR
    ' Using default
    b_defFont               As Boolean
    b_defExpandFont         As Boolean
    b_defForeColor          As Boolean
    b_defExpandForeColor    As Boolean
End Type

' Image Type
Private Type TYPEPCSLISTBOXIMAGE
    srcImage                As Picture
    lngWidth                As Long
    lngHeight               As Long
End Type

Private Const CLR_INVALID   As Long = &HFFFF
Private Const DI_NORMAL     As Long = &H3
Dim m_AutoRedraw            As Boolean
' For header
Dim m_HeaderCount           As Long
Dim m_Header()              As TYPEPCSLISTBOXHEADER
Dim m_HeaderFont            As font
Dim m_HeaderExpandFont      As font
Dim m_HeaderForeColor       As OLE_COLOR
Dim m_HeaderExpandForeColor As OLE_COLOR
' Noheader items & header customs
Dim m_ItemCount             As Long
Dim m_Item()                As TYPEPCSLISTBOXITEM
Dim m_NoHeaderCount         As Long
Dim m_NoHeaderTHeight       As Double
Dim m_HeaderOpacity         As Byte
Dim m_MaxHeaderPin          As Integer
Dim m_HeaderPin()           As Long
Dim m_HeaderPinRect()       As RECT
Dim m_PinReachMax           As Boolean
' Default item values
Dim m_ItemHeight            As Single
Dim m_ItemFont              As font
Dim m_ItemSelFont           As font
Dim m_ItemBackColor         As OLE_COLOR
Dim m_ItemForeColor         As OLE_COLOR
Dim m_ItemSelBackColor      As OLE_COLOR
Dim m_ItemSelForeColor      As OLE_COLOR
Dim m_SeparatorSpacing      As Single
Dim m_AutoSpacingSeparator  As Boolean
' Item customs
Dim m_ItemOpacity           As Byte
Dim m_SeparatorColor        As OLE_COLOR
Dim m_ItemSpacing           As Single
' Header image
Dim m_hDC(1)                As Long
' Picture list
Dim m_Image()               As TYPEPCSLISTBOXIMAGE
Dim m_ImageCount            As Integer
' For last selected
Dim m_SelIndex(1)        As Long

' ListBox events
Public Event Click()
Public Event DblClick()
Public Event ItemClick(Relative&, Index&)
Public Event HeaderClick(Index&)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event Scroll()
Public Event Resize()
Public Event Change()


' Prepare for new draw
Public Sub InitializeListBox()
    m_HeaderCount = -1
    m_ItemCount = -1
    m_ImageCount = -1
    ReleaseSelected
End Sub

' Add new image
Public Function AddImage(ByVal lPicture As StdPicture) As Boolean
On Error GoTo NoImage:                                              ' Resume next
    m_ImageCount = m_ImageCount + 1
    ReDim Preserve m_Image(m_ImageCount)
    With m_Image(m_ImageCount)
        Set .srcImage = lPicture
        .lngWidth = ScaleX(lPicture.Width, vbHimetric, vbPixels)    ' Image width
        .lngHeight = ScaleY(lPicture.Height, vbHimetric, vbPixels)  ' Image height
    End With
    AddImage = True
    RaiseEvent Change
NoImage: End Function

' Image count
Public Property Get ImageCount%()
    ImageCount = m_ImageCount + 1
End Property

' Add new header
Public Sub AddHeader(HeaderKey$, HeaderText, Optional HeaderExpand As Boolean = True, _
                        Optional HeaderPin As Boolean = True)
    m_HeaderCount = m_HeaderCount + 1
    ReDim Preserve m_Header(m_HeaderCount)
    ' Set header values
    With m_Header(m_HeaderCount)
        .strKey = Trim$(HeaderKey)
        .StrText = HeaderText
        .blnExpand = HeaderExpand
        .blnPin = HeaderPin
        .lngItemCount = -1
        .dblTotalHeight = 0
        .b_isPin = False
        ' Make default
        .b_defExpandFont = True
        .b_defExpandForeColor = True
        .b_defFont = True
        .b_defForeColor = True
    End With
    RaiseEvent Change
End Sub

' Add new item (ItemRelative can be Key or Index of Header, default is -1)
Public Sub AddItem(ItemRelative, ItemText, Optional ItemPicture% = -1, Optional ItemSeparator As Boolean = False)
    Dim i&, bHeader As Boolean
    Dim tmpItem As TYPEPCSLISTBOXITEM
    ' Use for set new item values
    With tmpItem
        .StrText = ItemText
        .intImage = ItemPicture
        .blnSeparator = ItemSeparator
        ' Make default
        .b_defBackColor = True
        .b_defFont = True
        .b_defForeColor = True
        .b_defHeight = True
        .b_defSelBackColor = True
        .b_defSelFont = True
        .b_defSelForeColor = True
        .b_defSpacing = True
    End With
    
    ' If no header(key or index) match ItemRelative
    ' then add to no header item group
    If m_HeaderCount > -1 Then                                                      ' If header avaiable
        If IsNumeric(ItemRelative) Then                                             ' Check the key or index
            If CLng(ItemRelative) <= m_HeaderCount And CLng(ItemRelative >= 0) Then ' Using header index
                m_Header(ItemRelative).lngItemCount = m_Header(ItemRelative).lngItemCount + 1
                ReDim Preserve m_Header(ItemRelative).colItems(m_Header(ItemRelative).lngItemCount)
                m_Header(ItemRelative).colItems(m_Header(ItemRelative).lngItemCount) = tmpItem
                RaiseEvent Change
                Exit Sub
            End If
        Else
            For i = 0 To m_HeaderCount
                If LCase$(Trim$(ItemRelative)) = LCase$(m_Header(i).strKey) Then    ' Using header key
                    m_Header(i).lngItemCount = m_Header(i).lngItemCount + 1
                    ReDim Preserve m_Header(i).colItems(m_Header(i).lngItemCount)
                    m_Header(i).colItems(m_Header(i).lngItemCount) = tmpItem
                    RaiseEvent Change
                    Exit Sub
                End If
            Next
        End If
    End If
    
    ' No header item group
    m_ItemCount = m_ItemCount + 1
    ReDim Preserve m_Item(m_ItemCount)
    m_Item(m_ItemCount) = tmpItem
    RaiseEvent Change
End Sub

' Calculate avaiable height for scrollbar visible
Private Sub CalHeightRECT()
On Error Resume Next
    Dim i&, j&
    ' Use for scrollbar visible
    Dim dH As Double
    
    dH = 0
    m_NoHeaderTHeight = 0
    
    ' No header items
    If m_ItemCount > -1 Then
        For i = 0 To m_ItemCount
            m_NoHeaderTHeight = m_NoHeaderTHeight + IIf(m_Item(i).b_defHeight, _
                                                        m_ItemHeight, m_Item(i).sngHeight)
        Next
    End If
    
    ' Headers
    If m_HeaderCount > -1 Then
        For i = 0 To m_HeaderCount
            m_Header(i).dblTotalHeight = 0                      ' Return to zero
        Next
    ' Header items
        For i = 0 To m_HeaderCount
            If m_Header(i).blnExpand Then
                For j = 0 To m_Header(i).lngItemCount
                    m_Header(i).dblTotalHeight = m_Header(i).dblTotalHeight + _
                        IIf(m_Header(i).colItems(j).b_defHeight, m_ItemHeight, _
                            m_Header(i).colItems(j).sngHeight)  ' Default or custom size?
                Next j
            End If
        Next
        ' Header always visible
        For i = 0 To m_HeaderCount
            dH = dH + m_Header(i).dblTotalHeight + 22
        Next
    End If
    
    Dim lW& ' Temp width
    
    ' Size of no header item group
    dH = dH + m_NoHeaderTHeight
    
    ' Is scrollbar visible?
    If dH > UserControl.ScaleHeight Then
        pSB.Visible = True
        pSB.Max = dH - UserControl.ScaleHeight
        'Move scrollbar to right edge
        pSB.Move UserControl.ScaleWidth - pSB.Width, 0, pSB.Width, UserControl.ScaleHeight
    Else
        'Make invisible scrollbar
        pSB.Visible = False
        pSB.Value = 0
    End If
    
    ' Item width depend on lW
    lW = UserControl.ScaleWidth - IIf(pSB.Visible, pSB.Width, 0)
    
    ' Create header image for blending
    Call CreateTheHeader(lW)
    
    ' No header items
    If m_ItemCount > -1 Then
        ' Set rect for each item
        SetRect m_Item(0).cRect, 25, 0, lW, IIf(m_Item(0).b_defHeight, m_ItemHeight, m_Item(0).sngHeight)
        If m_ItemCount > 0 Then
            For i = 1 To m_ItemCount
                With m_Item(i).cRect
                    .Left = m_ItemSpacing
                    .Top = m_Item(i - 1).cRect.Bottom
                    .Right = lW
                    .Bottom = .Top + IIf(m_Item(i).b_defHeight, m_ItemHeight, m_Item(i).sngHeight)
                End With
            Next i
        End If
    End If
                
    ' Work when header avaiable
    If m_HeaderCount > -1 Then
        ' Init for first header
        With m_Header(0).cRect
            .Left = 0
            .Top = m_NoHeaderTHeight
            .Right = lW
            .Bottom = .Top + 22
        End With

        ' Set rect for each header
        For i = 1 To m_HeaderCount
            With m_Header(i).cRect
                .Left = 0
                .Top = m_Header(i - 1).cRect.Bottom + m_Header(i - 1).dblTotalHeight
                .Right = lW
                .Bottom = .Top + 22
            End With
        Next
        
        ' Set header items rect
        For i = 0 To m_HeaderCount
            If m_Header(i).blnExpand Then
                If m_Header(i).lngItemCount <> -1 Then
                    ' First item
                    With m_Header(i).colItems(0).cRect
                        .Left = m_ItemSpacing
                        .Top = m_Header(i).cRect.Bottom
                        .Right = lW
                        ' Customs for default?
                        .Bottom = .Top + IIf(m_Header(i).colItems(0).b_defHeight, m_ItemHeight, _
                                                m_Header(i).colItems(0).sngHeight)
                    End With
                    ' Another items
                    If m_Header(i).lngItemCount > 0 Then
                        For j = 1 To m_Header(i).lngItemCount
                            With m_Header(i).colItems(j).cRect
                                .Left = m_ItemSpacing
                                .Top = m_Header(i).colItems(j - 1).cRect.Bottom
                                .Right = lW
                                ' Customs for default?
                                .Bottom = .Top + IIf(m_Header(i).colItems(j).b_defHeight, m_ItemHeight, _
                                                        m_Header(i).colItems(j).sngHeight)
                            End With
                        Next j
                    End If
                End If
            End If
        Next
    End If
End Sub

' Set font and text color for  item
Private Sub SetItemFontColor(IptItem As TYPEPCSLISTBOXITEM)
    If Not IptItem.b_Selected Or IptItem.blnSeparator Then
        Set picDraw.font = IIf(IptItem.b_defFont, m_ItemFont, IptItem.lngFont)
        SetTextColor picDraw.hDC, IIf(IptItem.b_defForeColor, m_ItemForeColor, IptItem.lngForeColor)
    Else
        Set picDraw.font = IIf(IptItem.b_defSelFont, m_ItemSelFont, IptItem.lngSelFont)
        SetTextColor picDraw.hDC, IIf(IptItem.b_defSelForeColor, m_ItemSelForeColor, _
                                        IptItem.lngSelForeColor)
    End If
End Sub

' Set font and text color for header
Private Sub SetHeaderFontColor(IptHeader As TYPEPCSLISTBOXHEADER)
    If Not IptHeader.blnExpand Then
        Set picDraw.font = IIf(IptHeader.b_defFont, m_HeaderFont, IptHeader.lngFont)
        SetTextColor picDraw.hDC, IIf(IptHeader.b_defForeColor, m_HeaderForeColor, IptHeader.lngForeColor)
    Else
        Set picDraw.font = IIf(IptHeader.b_defExpandFont, m_HeaderExpandFont, IptHeader.lngExpandFont)
        SetTextColor picDraw.hDC, IIf(IptHeader.b_defExpandForeColor, m_HeaderExpandForeColor, _
                                        IptHeader.lngExpandForeColor)
    End If
End Sub

' Draw item or separator func
Private Sub AutoDrawItem(ByVal StrText$, intImage%, bSeparator As Boolean, cRect As RECT)
    Dim lSpacing&
    If bSeparator Then       ' Draw separator
        lSpacing = CLng(m_SeparatorSpacing)
        If m_AutoSpacingSeparator Then lSpacing = picDraw.TextWidth(StrText) + 8
        ' Draw separator line
        DrawSeparator picDraw.hDC, lSpacing, cRect.Top, cRect.Right, cRect.Bottom
        cRect.Left = 0
        cRect.Right = lSpacing
        ' Draw text into dc
        DrawText picDraw.hDC, StrPtr(StrText), -1, cRect, &H20 Or &H4 Or &H1
    Else                                 ' Draw normal
        ' Draw text into dc
        DrawText picDraw.hDC, StrPtr(StrText), -1, cRect, &H20 Or &H4
    End If
    
    ' Draw item image
    If intImage > -1 And intImage <= m_ImageCount And Not bSeparator Then
        If Is32BitBMP(m_Image(intImage).srcImage) Then            ' Draw 32-bit bitmap
            TransBlt32 picDraw.hDC, cRect.Left - m_Image(intImage).lngWidth - 3, _
                                    cRect.Top + ((cRect.Bottom - cRect.Top) - _
                                    (m_Image(intImage).lngHeight)) / 2, _
                                    m_Image(intImage).srcImage
        Else                                                      ' Draw normal bitmap
            TransBlt picDraw.hDC, cRect.Left - m_Image(intImage).lngWidth - 3, _
                                    cRect.Top + ((cRect.Bottom - cRect.Top) - _
                                    (m_Image(intImage).lngHeight)) / 2, _
                                    m_Image(intImage).srcImage
        End If
    End If
End Sub

' Draw visible item in area
Private Sub DrawVisibleItems()
On Error Resume Next
    Dim lstTop As Double, i&, j&, nA As Double
    Dim mRect As RECT
    
    ' Starting point
    lstTop = pSB.Value
    picDraw.Cls ' Clean up
    
    ' Draw no header items first
    If m_ItemCount > -1 Then
        ' If neccessary
        'nA = m_Item(m_ItemCount).cRect.Top - m_Item(m_ItemCount).cRect.Top
        'nA = nA + IIf(m_Item(m_ItemCount).b_defHeight, m_ItemHeight, m_Item(m_ItemCount).sngHeight)
        'If lstTop - nA < nA Then
            For i = 0 To m_ItemCount
                ' Exit for if neccessary
                If m_Item(i).cRect.Top >= lstTop + UserControl.ScaleHeight Then Exit For
                If m_Item(i).cRect.Top >= lstTop - IIf(m_Item(i).b_defHeight, _
                                    m_ItemHeight, m_Item(i).sngHeight) Then
                    ' Set font & color
                    SetItemFontColor m_Item(i)
                    ' Set rect
                    With mRect
                        .Left = IIf(m_Item(i).b_defSpacing, m_ItemSpacing, m_Item(i).sngSpacing)
                        .Top = m_Item(i).cRect.Top - lstTop
                        .Right = m_Item(i).cRect.Right
                        .Bottom = .Top + IIf(m_Item(i).b_defHeight, m_ItemHeight, m_Item(i).sngHeight)
                    End With
                    ' Create item opacity if neccessary
                    If m_ItemOpacity > 0 Then _
                        CreateItem picDraw.hDC, 0, mRect.Top, mRect.Right, _
                                    mRect.Bottom - mRect.Top, m_ItemOpacity, m_Item(i)
                    ' Draw item & image
                    AutoDrawItem m_Item(i).StrText, m_Item(i).intImage, m_Item(i).blnSeparator, mRect
                End If
            Next
        'End If
    End If

    If m_HeaderCount > -1 Then
        Dim lastTop&, lMin&
        Dim n%
        lastTop = 0
        n = -1
        m_PinReachMax = False
        ' Redim new list of header pin
        If m_MaxHeaderPin > 0 Then
            lMin = IIf(m_HeaderCount < (m_MaxHeaderPin - 1), m_HeaderCount, (m_MaxHeaderPin - 1))
            ReDim m_HeaderPin(lMin)
            ReDim m_HeaderPinRect(lMin)
            ' Set default value
            For i = 0 To UBound(m_HeaderPin)
                m_HeaderPin(i) = -1
                SetRect m_HeaderPinRect(i), 0, i * 22, UserControl.ScaleWidth - IIf(pSB.Visible, pSB.Width, 0), _
                                                i * 22 + 22
            Next
            ' Check for pin header first
            For i = 0 To m_HeaderCount
                ' Set no pinning first
                m_Header(i).b_isPin = False
                If m_Header(i).blnPin Then                  ' If header is pin mode
                    If m_Header(i).cRect.Top - lstTop < lastTop Then
                        lastTop = lastTop + 22
                        n = n + 1
                        If lastTop > UBound(m_HeaderPin) * 22 Then _
                            lastTop = UBound(m_HeaderPin) * 22
                        ' Is pinning
                        m_Header(i).b_isPin = True
                        If UBound(m_HeaderPin) > 0 Then
                            If n < UBound(m_HeaderPin) + 1 Then
                                m_HeaderPin(n) = i
                            Else                            'Add from bottom
                                For j = 0 To UBound(m_HeaderPin) - 1
                                    m_HeaderPin(j) = m_HeaderPin(j + 1)
                                Next j
                                m_HeaderPin(UBound(m_HeaderPin)) = i
                                m_PinReachMax = True
                            End If
                        Else                                ' Only one item
                            m_HeaderPin(0) = i
                        End If
                    End If
                End If
            Next
        End If

        ' Draw header & header items
        For i = 0 To m_HeaderCount
            If m_Header(i).cRect.Top > lstTop + UserControl.ScaleHeight Then Exit For
            If m_Header(i).cRect.Top >= lstTop - 22 Then
                If Not m_Header(i).b_isPin Then         ' Do not draw if header pinning
                    AlphaBlending picDraw.hDC, m_Header(i).cRect.Left, m_Header(i).cRect.Top - lstTop, _
                                    m_Header(i).cRect.Right, 22, _
                                    m_hDC(IIf(m_Header(i).blnExpand, 1, 0)), m_HeaderOpacity
                End If
                ' Set font & color
                SetHeaderFontColor m_Header(i)
                With mRect
                    .Left = 25
                    .Top = m_Header(i).cRect.Top - lstTop
                    .Right = m_Header(i).cRect.Right
                    .Bottom = .Top + 22
                End With
                DrawText picDraw.hDC, StrPtr(m_Header(i).StrText), -1, mRect, &H20 Or &H4
            End If
        Next
        ' Draw header items
        For i = 0 To m_HeaderCount
            If m_Header(i).blnExpand Then                                   ' If header expanding
                'nA = m_Header(i).colItems(m_Header(i).lngItemCount).cRect.Top - m_Header(i).colItems(0).cRect.Top
                'nA = nA + IIf(m_Header(i).colItems(m_Header(i).lngItemCount).b_defHeight, _
                    m_ItemHeight, m_Header(i).colItems(m_Header(i).lngItemCount).sngHeight)
                'If lstTop - nA < nA Then
                    For j = 0 To m_Header(i).lngItemCount
                        '  Exit for if needed
                        If m_Header(i).colItems(j).cRect.Top > lstTop + UserControl.ScaleHeight Then Exit For
                        If m_Header(i).colItems(j).cRect.Top >= lstTop - IIf(m_Header(i).colItems(j).b_defHeight, m_ItemHeight, _
                                                                m_Header(i).colItems(j).sngHeight) Then
                            ' Set font & color
                            SetItemFontColor m_Header(i).colItems(j)
                            With mRect
                                .Left = IIf(m_Header(i).colItems(j).b_defSpacing, _
                                    m_ItemSpacing, m_Header(i).colItems(j).sngSpacing)
                                .Top = m_Header(i).colItems(j).cRect.Top - lstTop
                                .Right = m_Header(i).colItems(j).cRect.Right
                                .Bottom = .Top + IIf(m_Header(i).colItems(j).b_defHeight, m_ItemHeight, _
                                                    m_Header(i).colItems(j).sngHeight)
                            End With
                            ' Opacity if needed
                            If m_ItemOpacity > 0 Then _
                                CreateItem picDraw.hDC, 0, mRect.Top, mRect.Right, _
                                            mRect.Bottom - mRect.Top, m_ItemOpacity, m_Header(i).colItems(j)
                            ' Draw text & image
                            AutoDrawItem m_Header(i).colItems(j).StrText, m_Header(i).colItems(j).intImage, _
                                            m_Header(i).colItems(j).blnSeparator, mRect
                        End If
                    Next j
                'End If
            End If
        Next
        
        ' Header pinning?
        lastTop = 0
        For i = 0 To UBound(m_HeaderPin)
            If m_HeaderPin(i) <> -1 Then
                If m_Header(m_HeaderPin(i)).cRect.Top - lstTop < lastTop Then
                    AlphaBlending picDraw.hDC, m_Header(m_HeaderPin(i)).cRect.Left, lastTop, _
                                    m_Header(m_HeaderPin(i)).cRect.Right, 22, _
                                    m_hDC(IIf(m_Header(m_HeaderPin(i)).blnExpand, 1, 0)), m_HeaderOpacity
                    SetHeaderFontColor m_Header(m_HeaderPin(i))
                    With mRect
                        .Left = 25
                        .Top = lastTop
                        .Right = m_Header(m_HeaderPin(i)).cRect.Right
                        .Bottom = .Top + 22
                    End With
                    lastTop = lastTop + 22
                    DrawText picDraw.hDC, StrPtr(m_Header(m_HeaderPin(i)).StrText), -1, mRect, &H20 Or &H4
                End If
            End If
        Next
    End If
End Sub

' Redraw with 2 state
Private Sub Redraw(Optional bCalulateRect As Boolean = False)
On Error Resume Next
    ' Resize picturebox
    picDraw.Move 0, 0, IIf(pSB.Visible, UserControl.ScaleWidth - pSB.Width, _
                            UserControl.ScaleWidth), UserControl.ScaleHeight
    If Ambient.UserMode Then
        If bCalulateRect Then Call CalHeightRECT
        Call DrawVisibleItems
        UserControl.Refresh
    End If
End Sub

' Force redraw
Public Sub RedrawAll()
    Redraw True
End Sub

' Create header image for blending
Private Sub CreateTheHeader(lWidth&)
On Error Resume Next
    Dim hBitmap&
    ' Clean up first
    DeleteDC m_hDC(0)                                                           ' Inactive
    DeleteDC m_hDC(1)                                                           ' Active
    ' Create DC to draw
    m_hDC(0) = CreateCompatibleDC(UserControl.hDC)
    hBitmap = CreateCompatibleBitmap(UserControl.hDC, lWidth, 22)
    SelectObject m_hDC(0), hBitmap
    ' Inactive
    StretchBlt m_hDC(0), 0, 0, lWidth, 22, pHdr.hDC, 0, 0, 1, 22, vbSrcCopy     ' Background
    BitBlt m_hDC(0), 0, 0, 22, 22, pHdr.hDC, 1, 0, vbSrcCopy                    ' Button
    DeleteObject hBitmap
    ' Create DC to draw
    m_hDC(1) = CreateCompatibleDC(UserControl.hDC)
    hBitmap = CreateCompatibleBitmap(UserControl.hDC, lWidth, 22)
    SelectObject m_hDC(1), hBitmap
    ' Active
    StretchBlt m_hDC(1), 0, 0, lWidth, 22, pHdr.hDC, 0, 0, 1, 22, vbSrcCopy     ' Background
    BitBlt m_hDC(1), 0, 0, 22, 22, pHdr.hDC, 23, 0, vbSrcCopy                   ' Button
    
    DrawLineAPI m_hDC(0), lWidth - 1, 0, lWidth - 1, 22, &H9B9B9B               ' Draw right line
    DrawLineAPI m_hDC(1), lWidth - 1, 0, lWidth - 1, 22, &H9B9B9B               ' Draw right line
    DeleteObject hBitmap
End Sub

Private Sub CreateItem(ByVal DstDC&, lLeft&, lTop&, lW&, lH&, _
    bOpacity As Byte, SrcItem As TYPEPCSLISTBOXITEM)
On Error Resume Next
    Dim tDC&, tBitmap&
    Dim lColor&
    ' Create device context
    tDC = CreateCompatibleDC(UserControl.hDC)
    tBitmap = CreateCompatibleBitmap(UserControl.hDC, lW, lH)
    SelectObject tDC, tBitmap
    ' Create brush & paint
    lColor = IIf((SrcItem.b_Selected And Not SrcItem.blnSeparator), _
                    IIf(SrcItem.b_defSelBackColor, m_ItemSelBackColor, SrcItem.lngSelBackColor), _
                    IIf(SrcItem.b_defBackColor, m_ItemBackColor, SrcItem.lngBackColor))
    DrawOpaqueRect tDC, 0, 0, lW, lH, lColor
    AlphaBlending DstDC, lLeft, lTop, lW, lH, tDC, bOpacity
    ' Clean up
    DeleteObject tBitmap
    DeleteDC tDC
End Sub

' Draw listbox separator
Private Sub DrawSeparator(ByVal DstDC&, L&, t&, R&, B&)
    Dim dT&
    dT = Fix((B - t) / 2)
    ' Draw 2 lines
    DrawLineAPI DstDC, L, t + dT - 1, R, t + dT - 1, m_SeparatorColor
    DrawLineAPI DstDC, L, t + dT + 1, R, t + dT + 1, m_SeparatorColor
End Sub

' Draw filled rectangle
Private Sub DrawOpaqueRect(DstDC&, L&, t&, W&, H&, lColor&)
    Dim hBrush&
    Dim cRect As RECT
    SetRect cRect, L, t, W, H
    hBrush = CreateSolidBrush(lColor)
    FillRect DstDC, cRect, hBrush
    ' Clean
    DeleteObject hBrush
End Sub

' Draw API line with color
Private Sub DrawLineAPI(ByVal DstDC&, X1&, Y1&, X2&, Y2&, lColor)
    Dim mPoint As POINTAPI
    Dim mPen&, mOldPen&
    
    mPen = CreatePen(0&, 1, lColor)
    mOldPen = SelectObject(DstDC, mPen)
    MoveToEx DstDC, X1, Y1, mPoint
    LineTo DstDC, X2, Y2
    
    SelectObject DstDC, mOldPen
    DeleteObject mPen
    DeleteObject mOldPen
End Sub

Private Sub AlphaBlending(ByVal DstDC&, lLeft&, lTop&, lWidth&, lHeight, SrcDC&, SrcAlpha As Byte)
On Error Resume Next
    Dim BF As BLENDFUNCTION, lBF&
    With BF
        .BlendOp = AC_SRC_OVER
        .BlendFlags = 0
        .SourceConstantAlpha = SrcAlpha
        .AlphaFormat = 0
    End With
    RtlMoveMemory lBF, BF, 4
    AlphaBlend DstDC, lLeft, lTop, lWidth, lHeight, SrcDC, 0, 0, lWidth, lHeight, lBF
End Sub

Private Function TranslateColor&(ByVal clrColor As OLE_COLOR, Optional ByRef hPalette& = 0)
    ' System color code to long rgb
    If OleTranslateColor(clrColor, hPalette, TranslateColor) Then TranslateColor = CLR_INVALID
End Function

Private Sub TransBlt(ByVal DstDC&, DstX&, DstY&, ByVal SrcPic As StdPicture, _
    Optional MaskColor As Boolean = False, Optional ByVal TransColor& = -1)
On Error Resume Next
    ' Routine : To make transparent and grayscale images
    ' Author  : Gonkuchi
    ' Modified by Dana Seaman

    Dim B&, H&, F&, i&, newW&
    Dim TmpDC&, TmpBmp&, TmpObj&
    Dim Sr2DC&, Sr2Bmp&, Sr2Obj&
    Dim DataDest() As RGBTRIPLE
    Dim DataSrc() As RGBTRIPLE
    Dim Info As BITMAPINFO
    Dim BrushRGB As RGBTRIPLE
    Dim gCol&, hOldOb&
    Dim SrcDC&, tObj&, tTT&
    Dim DstW&, DstH&

    DstW = ScaleX(SrcPic.Width, vbHimetric, vbPixels)
    DstH = ScaleY(SrcPic.Height, vbHimetric, vbPixels)

    SrcDC = CreateCompatibleDC(DstDC)

    If SrcPic.Type = vbPicTypeBitmap Then   'Check if it's an icon or a bitmap
        tObj = SelectObject(SrcDC, SrcPic)
    Else
        Dim hBrush&
        tObj = SelectObject(SrcDC, CreateCompatibleBitmap(DstDC, DstW, DstH))
        hBrush = CreateSolidBrush(TransColor)
        DrawIconEx SrcDC, 0, 0, SrcPic.Handle, DstW, DstH, 0, hBrush, DI_NORMAL
        DeleteObject hBrush
    End If

    TmpDC = CreateCompatibleDC(SrcDC)
    Sr2DC = CreateCompatibleDC(SrcDC)
    TmpBmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
    Sr2Bmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
    TmpObj = SelectObject(TmpDC, TmpBmp)
    Sr2Obj = SelectObject(Sr2DC, Sr2Bmp)
    
    ReDim DataDest(DstW * DstH * 3 - 1)
    ReDim DataSrc(UBound(DataDest))
    
    With Info.bmiHeader
        .biSize = Len(Info.bmiHeader)
        .biWidth = DstW
        .biHeight = DstH
        .biPlanes = 1
        .biBitCount = 24
    End With

    BitBlt TmpDC, 0, 0, DstW, DstH, DstDC, DstX, DstY, vbSrcCopy
    BitBlt Sr2DC, 0, 0, DstW, DstH, SrcDC, 0, 0, vbSrcCopy
    GetDIBits TmpDC, TmpBmp, 0, DstH, DataDest(0), Info, 0
    GetDIBits Sr2DC, Sr2Bmp, 0, DstH, DataSrc(0), Info, 0

    ' No Maskcolor to use
    If Not MaskColor Then TransColor = -1

    newW = DstW - 1

    For H = 0 To DstH - 1
        F = H * DstW
        For B = 0 To newW
            i = F + B
            If GetNearestColor(DstDC, CLng(DataSrc(i).rgbRed) + 256& * DataSrc(i).rgbGreen + 65536 * DataSrc(i).rgbBlue) <> TransColor Then
                DataDest(i) = DataSrc(i)
            End If
        Next B
    Next H

    ' Paint it!
    SetDIBitsToDevice DstDC, DstX, DstY, DstW, DstH, 0, 0, 0, DstH, DataDest(0), Info, 0

    Erase DataDest, DataSrc
    DeleteObject SelectObject(TmpDC, TmpObj)
    DeleteObject SelectObject(Sr2DC, Sr2Obj)
    If SrcPic.Type = vbPicTypeIcon Then DeleteObject SelectObject(SrcDC, tObj)
    DeleteDC TmpDC
    DeleteDC Sr2DC
    DeleteObject tObj
    DeleteDC SrcDC
End Sub

Private Sub TransBlt32(ByVal DstDC&, DstX&, DstY&, ByVal SrcPic As StdPicture)
On Error Resume Next
    ' Routine : Renders 32 bit Bitmap
    ' Author  : Dana Seaman
    Dim B&, H&, F&, i&, newW&
    Dim TmpDC&, TmpBmp&, TmpObj&
    Dim Sr2DC&, Sr2Bmp&, Sr2Obj&
    Dim DataDest()  As RGBQUAD
    Dim DataSrc()   As RGBQUAD
    Dim Info        As BITMAPINFO
    Dim BrushRGB    As RGBQUAD
    Dim gCol&, hOldOb&
    Dim PicBlend    As Boolean
    Dim SrcDC&, tObj&, tTT&
    Dim a1&, a2&, DstW&, DstH&

    DstW = ScaleX(SrcPic.Width, vbHimetric, vbPixels)
    DstH = ScaleY(SrcPic.Height, vbHimetric, vbPixels)
    
    SrcDC = CreateCompatibleDC(DstDC)

    tObj = SelectObject(SrcDC, SrcPic)

    TmpDC = CreateCompatibleDC(SrcDC)
    Sr2DC = CreateCompatibleDC(SrcDC)
    TmpBmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
    Sr2Bmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
    TmpObj = SelectObject(TmpDC, TmpBmp)
    Sr2Obj = SelectObject(Sr2DC, Sr2Bmp)

    With Info.bmiHeader
        .biSize = Len(Info.bmiHeader)
        .biWidth = DstW
        .biHeight = DstH
        .biPlanes = 1
        .biBitCount = 32
        .biSizeImage = 4 * ((DstW * .biBitCount + 31) \ 32) * DstH
    End With
    
    ReDim DataDest(Info.bmiHeader.biSizeImage - 1)
    ReDim DataSrc(UBound(DataDest))

    BitBlt TmpDC, 0, 0, DstW, DstH, DstDC, DstX, DstY, vbSrcCopy
    BitBlt Sr2DC, 0, 0, DstW, DstH, SrcDC, 0, 0, vbSrcCopy
    GetDIBits TmpDC, TmpBmp, 0, DstH, DataDest(0), Info, 0
    GetDIBits Sr2DC, Sr2Bmp, 0, DstH, DataSrc(0), Info, 0

    newW = DstW - 1

    For H = 0 To DstH - 1
        F = H * DstW
        For B = 0 To newW
            i = F + B
            With DataDest(i)
                If DataSrc(i).rgbAlpha = 255 Then
                    DataDest(i) = DataSrc(i)
                ElseIf DataSrc(i).rgbAlpha > 0 Then
                    a1 = DataSrc(i).rgbAlpha
                    a2 = 255 - a1
                    .rgbRed = (a2 * .rgbRed + a1 * DataSrc(i).rgbRed) \ 256
                    .rgbGreen = (a2 * .rgbGreen + a1 * DataSrc(i).rgbGreen) \ 256
                    .rgbBlue = (a2 * .rgbBlue + a1 * DataSrc(i).rgbBlue) \ 256
                End If
            End With
        Next B
    Next H

    ' Paint it!
    SetDIBitsToDevice DstDC, DstX, DstY, DstW, DstH, 0, 0, 0, DstH, DataDest(0), Info, 0

    Erase DataDest, DataSrc
    DeleteObject SelectObject(TmpDC, TmpObj)
    DeleteObject SelectObject(Sr2DC, Sr2Obj)
    If SrcPic.Type = vbPicTypeIcon Then DeleteObject SelectObject(SrcDC, tObj)
    DeleteDC TmpDC
    DeleteDC Sr2DC
    DeleteObject tObj
    DeleteDC SrcDC
End Sub

' Check for 32bit BMP
Private Function Is32BitBMP(pPicture As Picture) As Boolean
    Dim uBI As BITMAP
    If pPicture.Type = vbPicTypeBitmap Then
        Call GetObject(pPicture.Handle, Len(uBI), uBI)
        Is32BitBMP = uBI.bmBitsPixel = 32
    End If
End Function

Private Sub picDraw_Click()
    RaiseEvent Click
End Sub

Private Sub picDraw_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub picDraw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Dim i&, j&, unArea&
    Dim lstTop As Double
    RaiseEvent MouseDown(Button, Shift, X, Y)
    If Button <> 1 Then Exit Sub
    lstTop = pSB.Value
    unArea = 0  ' Unavailable items (for header pin)
    ' Pin item action
    For i = 0 To UBound(m_HeaderPin)
        If m_HeaderPin(i) > -1 Then
            unArea = unArea + 22
            If Y >= m_HeaderPinRect(i).Top And _
                Y <= m_HeaderPinRect(i).Bottom And m_Header(m_HeaderPin(i)).b_isPin Then
                    ' Move to top of header
                    pSB.Value = m_Header(m_HeaderPin(i)).cRect.Top - (IIf(m_PinReachMax, i + 1, i) * 22)
                    RaiseEvent HeaderClick(m_HeaderPin(i))
                Exit Sub    ' You can't click multiple items at one time
            End If
        End If
    Next
    ' Header action
    For i = 0 To m_HeaderCount
        If m_Header(i).cRect.Top - lstTop > UserControl.ScaleHeight Then Exit For
        If m_Header(i).cRect.Top - lstTop > -22 Then
            If Y >= m_Header(i).cRect.Top - lstTop And Y <= m_Header(i).cRect.Bottom - lstTop And _
                Y > unArea - 22 Then
                    RaiseEvent HeaderClick(i)
                    If m_Header(i).cRect.Top - lstTop < unArea Then ' Move to this header
                        pSB.Value = m_Header(i).cRect.Top - unArea
                        Exit Sub
                    Else
                        ' Expand item
                        m_Header(i).blnExpand = Not m_Header(i).blnExpand
                        GoTo ENDSUB:
                    End If
            End If
        End If
    Next
    ' Item action
    ' No header items
    If m_ItemCount > -1 Then
        For i = 0 To m_ItemCount
            If m_Item(i).cRect.Top - lstTop > UserControl.ScaleHeight Then Exit Sub
            If m_Item(i).cRect.Top - lstTop >= IIf(m_Item(i).b_defHeight, _
                                                    -m_ItemHeight, -m_Item(i).sngHeight) Then
                If Not m_Item(i).blnSeparator Then
                    If Y >= m_Item(i).cRect.Top - lstTop And Y <= m_Item(i).cRect.Bottom - lstTop Then
                        RaiseEvent ItemClick(-1, i)
                        ReleaseSelected
                        m_Item(i).b_Selected = True
                        If m_Item(i).cRect.Top - lstTop < unArea Then _
                            pSB.Value = m_Item(i).cRect.Top - unArea
                        m_SelIndex(0) = -1: m_SelIndex(1) = i
                        GoTo ENDSUB:
                    End If
                End If
            End If
        Next
    End If
    ' Header items
    For i = 0 To m_HeaderCount
        If m_Header(i).cRect.Top - lstTop > UserControl.ScaleHeight Then Exit Sub
        If m_Header(i).blnExpand And m_Header(i).lngItemCount > -1 Then
            For j = 0 To m_Header(i).lngItemCount
                If m_Header(i).colItems(j).cRect.Top - lstTop > UserControl.ScaleHeight Then Exit For
                If Not m_Header(i).colItems(j).blnSeparator Then
                    If Y >= m_Header(i).colItems(j).cRect.Top - lstTop And _
                        Y <= m_Header(i).colItems(j).cRect.Bottom - lstTop Then
                        RaiseEvent ItemClick(i, j)
                        ReleaseSelected
                        m_Header(i).colItems(j).b_Selected = True
                        If m_Header(i).colItems(j).cRect.Top - lstTop < unArea Then _
                            pSB.Value = m_Header(i).colItems(j).cRect.Top - unArea
                        m_SelIndex(0) = i: m_SelIndex(1) = j
                        GoTo ENDSUB:
                    End If
                End If
            Next j
        End If
    Next
ENDSUB:
    RedrawAll
End Sub

' Clear last selected item
Private Sub ReleaseSelected()
On Error GoTo NEXTRELEASE:                                  ' Resume next
    If m_SelIndex(0) = -1 And m_SelIndex(1) > -1 Then       ' No header item
        m_Item(m_SelIndex(1)).b_Selected = False
    ElseIf m_SelIndex(0) > -1 Then                          ' Header item
        m_Header(m_SelIndex(0)).colItems(m_SelIndex(1)).b_Selected = False
    End If
NEXTRELEASE:
    m_SelIndex(0) = -1
    m_SelIndex(1) = -1
End Sub

Private Sub picDraw_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub picDraw_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub pSB_Scroll()
    'Redraw listbox
    RaiseEvent Scroll
    Redraw
End Sub

' Default property

' Auto redraw
Public Property Get AutoRedraw() As Boolean
    AutoRedraw = m_AutoRedraw
End Property

Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)
    m_AutoRedraw = New_AutoRedraw
    PropertyChanged "AutoRedraw"
End Property

' Back color
Public Property Get BackColor() As OLE_COLOR
    BackColor = picDraw.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    picDraw.BackColor = New_BackColor
    RedrawAll
    PropertyChanged "BackColor"
End Property

' Background picture
Public Property Get Background() As Picture
    Set Background = picDraw.Picture
End Property

Public Property Set Background(ByVal New_Background As Picture)
    Set picDraw.Picture = New_Background
    RedrawAll
    PropertyChanged "Background"
End Property

' Header font
Public Property Get HeaderFont() As font
    Set HeaderFont = m_HeaderFont
End Property

Public Property Set HeaderFont(ByVal New_HeaderFont As font)
    Set m_HeaderFont = New_HeaderFont
    RedrawAll
    PropertyChanged "HeaderFont"
End Property

' Header expand font
Public Property Get HeaderExpandFont() As font
    Set HeaderExpandFont = m_HeaderExpandFont
End Property

Public Property Set HeaderExpandFont(ByVal New_HeaderExpandFont As font)
    Set m_HeaderExpandFont = New_HeaderExpandFont
    RedrawAll
    PropertyChanged "HeaderExpandFont"
End Property

' Header fore color
Public Property Get HeaderForeColor() As OLE_COLOR
    HeaderForeColor = m_HeaderForeColor
End Property

Public Property Let HeaderForeColor(ByVal New_HeaderForeColor As OLE_COLOR)
    m_HeaderForeColor = New_HeaderForeColor
    RedrawAll
    PropertyChanged "HeaderForeColor"
End Property

' Header expand fore color
Public Property Get HeaderExpandForeColor() As OLE_COLOR
    HeaderExpandForeColor = m_HeaderExpandForeColor
End Property

Public Property Let HeaderExpandForeColor(ByVal New_HeaderExpandForeColor As OLE_COLOR)
    m_HeaderExpandForeColor = New_HeaderExpandForeColor
    RedrawAll
    PropertyChanged "HeaderExpandForeColor"
End Property

'Max header pin
Public Property Get MaxHeaderPin%()
    MaxHeaderPin = m_MaxHeaderPin
End Property

Public Property Let MaxHeaderPin(ByVal New_MaxHeaderPin%)
    m_MaxHeaderPin = New_MaxHeaderPin
    RedrawAll
    PropertyChanged "MaxHeaderPin"
End Property

'Header opacity
Public Property Get HeaderOpacity() As Byte
    HeaderOpacity = m_HeaderOpacity
End Property

Public Property Let HeaderOpacity(ByVal New_HeaderOpacity As Byte)
    m_HeaderOpacity = New_HeaderOpacity
    RedrawAll
    PropertyChanged "HeaderOpacity"
End Property

' Item font
Public Property Get ItemFont() As font
    Set ItemFont = m_ItemFont
End Property

Public Property Set ItemFont(ByVal New_ItemFont As font)
    Set m_ItemFont = New_ItemFont
    RedrawAll
    PropertyChanged "ItemFont"
End Property

' Item sel font
Public Property Get ItemSelFont() As font
    Set ItemSelFont = m_ItemSelFont
End Property

Public Property Set ItemSelFont(ByVal New_ItemSelFont As font)
    Set m_ItemSelFont = New_ItemSelFont
    RedrawAll
    PropertyChanged "ItemSelFont"
End Property

' Item height
Public Property Get ItemHeight!()
    ItemHeight = m_ItemHeight
End Property

Public Property Let ItemHeight(ByVal New_ItemHeight!)
    m_ItemHeight = New_ItemHeight
    RedrawAll
    PropertyChanged "ItemHeight"
End Property

' Item back color
Public Property Get ItemBackColor() As OLE_COLOR
    ItemBackColor = m_ItemBackColor
End Property

Public Property Let ItemBackColor(ByVal New_ItemBackColor As OLE_COLOR)
    m_ItemBackColor = New_ItemBackColor
    RedrawAll
    PropertyChanged "ItemBackColor"
End Property

' Item fore color
Public Property Get ItemForeColor() As OLE_COLOR
    ItemForeColor = m_ItemForeColor
End Property

Public Property Let ItemForeColor(ByVal New_ItemForeColor As OLE_COLOR)
    m_ItemForeColor = New_ItemForeColor
    RedrawAll
    PropertyChanged "ItemForeColor"
End Property

' Item sel back color
Public Property Get ItemSelBackColor() As OLE_COLOR
    ItemSelBackColor = m_ItemSelBackColor
End Property

Public Property Let ItemSelBackColor(ByVal New_ItemSelBackColor As OLE_COLOR)
    m_ItemSelBackColor = New_ItemSelBackColor
    RedrawAll
    PropertyChanged "ItemSelBackColor"
End Property

' Item sel fore color
Public Property Get ItemSelForeColor() As OLE_COLOR
    ItemSelForeColor = m_ItemSelForeColor
End Property

Public Property Let ItemSelForeColor(ByVal New_ItemSelForeColor As OLE_COLOR)
    m_ItemSelForeColor = New_ItemSelForeColor
    RedrawAll
    PropertyChanged "ItemSelForeColor"
End Property

' Item opacity
Public Property Get ItemOpacity() As Byte
    ItemOpacity = m_ItemOpacity
End Property

Public Property Let ItemOpacity(ByVal New_ItemOpacity As Byte)
    m_ItemOpacity = New_ItemOpacity
    RedrawAll
    PropertyChanged "ItemOpacity"
End Property

' Item spacing
Public Property Get ItemSpacing!()
    ItemSpacing = m_ItemSpacing
End Property

Public Property Let ItemSpacing(ByVal New_ItemSpacing!)
    m_ItemSpacing = New_ItemSpacing
    RedrawAll
    PropertyChanged "ItemSpacing"
End Property

' Separator color
Public Property Get SeparatorColor() As OLE_COLOR
    SeparatorColor = m_SeparatorColor
End Property

Public Property Let SeparatorColor(ByVal New_SeparatorColor As OLE_COLOR)
    m_SeparatorColor = New_SeparatorColor
    RedrawAll
    PropertyChanged "SeparatorColor"
End Property

' Separator spacing
Public Property Get SeparatorSpacing!()
    SeparatorSpacing = m_SeparatorSpacing
End Property

Public Property Let SeparatorSpacing(ByVal New_SeparatorSpacing!)
    m_SeparatorSpacing = New_SeparatorSpacing
    RedrawAll
    PropertyChanged "SeparatorSpacing"
End Property

' Auto spacing
Public Property Get AutoSpacingSeparator() As Boolean
    AutoSpacingSeparator = m_AutoSpacingSeparator
End Property

Public Property Let AutoSpacingSeparator(ByVal New_AutoSpacingSeparator As Boolean)
    m_AutoSpacingSeparator = New_AutoSpacingSeparator
    RedrawAll
    PropertyChanged "AutoSpacingSeparator"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    With PropBag
        m_AutoRedraw = .ReadProperty("AutoRedraw", 0)
        picDraw.BackColor = .ReadProperty("BackColor", 0)
        Set picDraw.Picture = .ReadProperty("Background", 0)
        Set m_HeaderFont = .ReadProperty("HeaderFont", 0)
        Set m_HeaderExpandFont = .ReadProperty("HeaderExpandFont", 0)
        m_HeaderForeColor = .ReadProperty("HeaderForeColor", 0)
        m_HeaderExpandForeColor = .ReadProperty("HeaderExpandForeColor", 0)
        m_HeaderOpacity = .ReadProperty("HeaderOpacity", 0)
        m_MaxHeaderPin = .ReadProperty("MaxHeaderPin", 0)
        Set m_ItemFont = .ReadProperty("ItemFont", 0)
        Set m_ItemSelFont = .ReadProperty("ItemSelFont", 0)
        m_ItemHeight = .ReadProperty("ItemHeight", 0)
        m_ItemBackColor = .ReadProperty("ItemBackColor", 0)
        m_ItemForeColor = .ReadProperty("ItemForeColor", 0)
        m_ItemSelBackColor = .ReadProperty("ItemSelBackColor", 0)
        m_ItemSelForeColor = .ReadProperty("ItemSelForeColor", 0)
        m_ItemOpacity = .ReadProperty("ItemOpacity", 0)
        m_ItemSpacing = .ReadProperty("ItemSpacing", 0)
        m_SeparatorColor = .ReadProperty("SeparatorColor", 0)
        m_SeparatorSpacing = .ReadProperty("SeparatorSpacing", 0)
        m_AutoSpacingSeparator = .ReadProperty("AutoSpacingSeparator", 0)
    End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
    With PropBag
        .WriteProperty "AutoRedraw", m_AutoRedraw, 0
        .WriteProperty "BackColor", picDraw.BackColor, 0
        .WriteProperty "Background", picDraw.Picture, 0
        .WriteProperty "HeaderFont", m_HeaderFont, 0
        .WriteProperty "HeaderExpandFont", m_HeaderExpandFont, 0
        .WriteProperty "HeaderForeColor", m_HeaderForeColor, 0
        .WriteProperty "HeaderExpandForeColor", m_HeaderExpandForeColor, 0
        .WriteProperty "HeaderOpacity", m_HeaderOpacity, 0
        .WriteProperty "MaxHeaderPin", m_MaxHeaderPin, 0
        .WriteProperty "ItemFont", m_ItemFont, 0
        .WriteProperty "ItemSelFont", m_ItemSelFont, 0
        .WriteProperty "ItemHeight", m_ItemHeight, 0
        .WriteProperty "ItemBackColor", m_ItemBackColor, 0
        .WriteProperty "ItemForeColor", m_ItemForeColor, 0
        .WriteProperty "ItemSelBackColor", m_ItemSelBackColor, 0
        .WriteProperty "ItemSelForeColor", m_ItemSelForeColor, 0
        .WriteProperty "ItemOpacity", m_ItemOpacity, 0
        .WriteProperty "ItemSpacing", m_ItemSpacing, 0
        .WriteProperty "SeparatorColor", m_SeparatorColor, 0
        .WriteProperty "SeparatorSpacing", m_SeparatorSpacing, 0
        .WriteProperty "AutoSpacingSeparator", m_AutoSpacingSeparator, 0
    End With
End Sub

' Init custom property
Private Sub UserControl_InitProperties()
On Error Resume Next
    BackColor = &HE8E8E8
    Set HeaderFont = Ambient.font
        HeaderFont.Bold = True
    Set HeaderExpandFont = Ambient.font
        HeaderExpandFont.Bold = True
    HeaderForeColor = &H4D4D4D
    HeaderExpandForeColor = &H4D4D4D
    HeaderOpacity = 255
    MaxHeaderPin = 4
    Set ItemFont = Ambient.font
    Set ItemSelFont = Ambient.font
    ItemBackColor = &HF3F3F3
    ItemForeColor = &H737373
    ItemSelBackColor = &H8B6425
    ItemSelForeColor = &HFFFFFF
    ItemHeight = 20
    ItemOpacity = 255
    ItemSpacing = 25
    SeparatorSpacing = 30
    SeparatorColor = &HD3D3D3
    AutoSpacingSeparator = True
    RedrawAll
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    RaiseEvent Resize
    RedrawAll
End Sub

Private Sub UserControl_Terminate()
On Error Resume Next
    ' Clean up
    DeleteDC m_hDC(0)
    DeleteDC m_hDC(1)
End Sub

' More customs

' Get item info
Private Function GetItemInfo(ByVal Relative&, ByVal Index&) As TYPEPCSLISTBOXITEM
    If Relative = -1 Then
        If m_ItemCount > -1 And m_ItemCount >= Index Then
            GetItemInfo = m_Item(Index)
        End If
    Else
        If m_HeaderCount >= Relative Then
            GetItemInfo = m_Header(Relative).colItems(Index)
        End If
    End If
End Function

' Item text
Public Property Get ItemText$(ByVal Relative&, ByVal Index&)
    ItemText = GetItemInfo(Relative, Index).StrText
End Property

Public Property Let ItemText(ByVal Relative&, ByVal Index&, ByVal New_Text$)
    If Relative = -1 Then
        If m_ItemCount > -1 And m_ItemCount >= Index Then
            m_Item(Index).StrText = New_Text
        End If
    Else
        If m_HeaderCount >= Relative Then
            m_Header(Relative).colItems(Index).StrText = New_Text
        End If
    End If
    If m_AutoRedraw Then RedrawAll
End Property

' Item image
Public Property Get ItemImage%(ByVal Relative&, ByVal Index&)
    ItemImage = GetItemInfo(Relative, Index).intImage
End Property

Public Property Let ItemImage(ByVal Relative&, ByVal Index&, ByVal New_Image%)
    If Relative = -1 Then
        If m_ItemCount > -1 And m_ItemCount >= Index Then
            m_Item(Index).intImage = New_Image
        End If
    Else
        If m_HeaderCount >= Relative Then
            m_Header(Relative).colItems(Index).intImage = New_Image
        End If
    End If
    If m_AutoRedraw Then RedrawAll
End Property

' Item separator type
Public Property Get ItemSeparator(ByVal Relative&, ByVal Index&) As Boolean
    ItemSeparator = GetItemInfo(Relative, Index).blnSeparator
End Property

Public Property Let ItemSeparator(ByVal Relative&, ByVal Index&, ByVal New_Separator As Boolean)
    If Relative = -1 Then
        If m_ItemCount > -1 And m_ItemCount >= Index Then
            m_Item(Index).blnSeparator = New_Separator
        End If
    Else
        If m_HeaderCount >= Relative Then
            m_Header(Relative).colItems(Index).blnSeparator = New_Separator
        End If
    End If
    If m_AutoRedraw Then RedrawAll
End Property

' Item height
Public Property Get CustomItemHeight!(ByVal Relative&, ByVal Index&)
    CustomItemHeight = GetItemInfo(Relative, Index).sngHeight
End Property

Public Property Let CustomItemHeight(ByVal Relative&, ByVal Index&, ByVal New_CustomItemHeight!)
    If Relative = -1 Then
        If m_ItemCount > -1 And m_ItemCount >= Index Then
            m_Item(Index).sngHeight = New_CustomItemHeight
            m_Item(Index).b_defHeight = False
        End If
    Else
        If m_HeaderCount >= Relative Then
            m_Header(Relative).colItems(Index).sngHeight = New_CustomItemHeight
            m_Header(Relative).colItems(Index).b_defHeight = False
        End If
    End If
    If m_AutoRedraw Then RedrawAll
End Property

' Item font
Public Property Get CustomItemFont(ByVal Relative&, ByVal Index&) As font
    Set CustomItemFont = GetItemInfo(Relative, Index).lngFont
End Property

Public Property Set CustomItemFont(ByVal Relative&, ByVal Index&, ByVal New_CustomItemFont As font)
    If Relative = -1 Then
        If m_ItemCount > -1 And m_ItemCount >= Index Then
            Set m_Item(Index).lngFont = New_CustomItemFont
            m_Item(Index).b_defFont = False
        End If
    Else
        If m_HeaderCount >= Relative Then
            Set m_Header(Relative).colItems(Index).lngFont = New_CustomItemFont
            m_Header(Relative).colItems(Index).b_defFont = False
        End If
    End If
    If m_AutoRedraw Then RedrawAll
End Property

' Item sel font
Public Property Get CustomItemSelFont(ByVal Relative&, ByVal Index&) As font
    Set CustomItemSelFont = GetItemInfo(Relative, Index).lngSelFont
End Property

Public Property Set CustomItemSelFont(ByVal Relative&, ByVal Index&, ByVal New_CustomItemSelFont As font)
    If Relative = -1 Then
        If m_ItemCount > -1 And m_ItemCount >= Index Then
            Set m_Item(Index).lngSelFont = New_CustomItemSelFont
            m_Item(Index).b_defSelFont = False
        End If
    Else
        If m_HeaderCount >= Relative Then
            Set m_Header(Relative).colItems(Index).lngSelFont = New_CustomItemSelFont
            m_Header(Relative).colItems(Index).b_defSelFont = False
        End If
    End If
    If m_AutoRedraw Then RedrawAll
End Property

' Item back color
Public Property Get CustomItemBackColor(ByVal Relative&, ByVal Index&) As OLE_COLOR
    CustomItemBackColor = GetItemInfo(Relative, Index).lngBackColor
End Property

Public Property Let CustomItemBackColor(ByVal Relative&, ByVal Index&, ByVal New_CustomItemBackColor As OLE_COLOR)
    If Relative = -1 Then
        If m_ItemCount > -1 And m_ItemCount >= Index Then
            m_Item(Index).lngBackColor = New_CustomItemBackColor
            m_Item(Index).b_defBackColor = False
        End If
    Else
        If m_HeaderCount >= Relative Then
            m_Header(Relative).colItems(Index).lngBackColor = New_CustomItemBackColor
            m_Header(Relative).colItems(Index).b_defBackColor = False
        End If
    End If
    If m_AutoRedraw Then RedrawAll
End Property

' Item sel back color
Public Property Get CustomItemSelBackColor(ByVal Relative&, ByVal Index&) As OLE_COLOR
    CustomItemSelBackColor = GetItemInfo(Relative, Index).lngSelBackColor
End Property

Public Property Let CustomItemSelBackColor(ByVal Relative&, ByVal Index&, ByVal New_CustomItemSelBackColor As OLE_COLOR)
    If Relative = -1 Then
        If m_ItemCount > -1 And m_ItemCount >= Index Then
            m_Item(Index).lngSelBackColor = New_CustomItemSelBackColor
            m_Item(Index).b_defSelBackColor = False
        End If
    Else
        If m_HeaderCount >= Relative Then
            m_Header(Relative).colItems(Index).lngSelBackColor = New_CustomItemSelBackColor
            m_Header(Relative).colItems(Index).b_defSelBackColor = False
        End If
    End If
    If m_AutoRedraw Then RedrawAll
End Property

' Item fore color
Public Property Get CustomItemForeColor(ByVal Relative&, ByVal Index&) As OLE_COLOR
    CustomItemForeColor = GetItemInfo(Relative, Index).lngForeColor
End Property

Public Property Let CustomItemForeColor(ByVal Relative&, ByVal Index&, ByVal New_CustomItemForeColor As OLE_COLOR)
    If Relative = -1 Then
        If m_ItemCount > -1 And m_ItemCount >= Index Then
            m_Item(Index).lngForeColor = New_CustomItemForeColor
            m_Item(Index).b_defForeColor = False
        End If
    Else
        If m_HeaderCount >= Relative Then
            m_Header(Relative).colItems(Index).lngForeColor = New_CustomItemForeColor
            m_Header(Relative).colItems(Index).b_defForeColor = False
        End If
    End If
    If m_AutoRedraw Then RedrawAll
End Property

' Item sel fore color
Public Property Get CustomItemSelForeColor(ByVal Relative&, ByVal Index&) As OLE_COLOR
    CustomItemSelForeColor = GetItemInfo(Relative, Index).lngSelForeColor
End Property

Public Property Let CustomItemSelForeColor(ByVal Relative&, ByVal Index&, ByVal New_CustomItemSelForeColor As OLE_COLOR)
    If Relative = -1 Then
        If m_ItemCount > -1 And m_ItemCount >= Index Then
            m_Item(Index).lngSelForeColor = New_CustomItemSelForeColor
            m_Item(Index).b_defSelForeColor = False
        End If
    Else
        If m_HeaderCount >= Relative Then
            m_Header(Relative).colItems(Index).lngSelForeColor = New_CustomItemSelForeColor
            m_Header(Relative).colItems(Index).b_defSelForeColor = False
        End If
    End If
    If m_AutoRedraw Then RedrawAll
End Property

' Item custom spacing
Public Property Get CustomItemSpacing!(ByVal Relative&, ByVal Index&)
    CustomItemSpacing = GetItemInfo(Relative, Index).sngSpacing
End Property

Public Property Let CustomItemSpacing(ByVal Relative&, ByVal Index&, ByVal New_CustomItemSpacing!)
    If Relative = -1 Then
        If m_ItemCount > -1 And m_ItemCount >= Index Then
            m_Item(Index).sngSpacing = New_CustomItemSpacing
            m_Item(Index).b_defSpacing = False
        End If
    Else
        If m_HeaderCount >= Relative Then
            m_Header(Relative).colItems(Index).sngSpacing = New_CustomItemSpacing
            m_Header(Relative).colItems(Index).b_defSpacing = False
        End If
    End If
    If m_AutoRedraw Then RedrawAll
End Property

' Item selected
Public Property Get ItemSelected(ByVal Relative&, ByVal Index&) As Boolean
    ItemSelected = GetItemInfo(Relative, Index).b_Selected
End Property

Public Property Let ItemSelected(ByVal Relative&, ByVal Index&, ByVal bSelected As Boolean)
    If bSelected Then ReleaseSelected
    If Relative = -1 Then
        If m_ItemCount > -1 And m_ItemCount >= Index Then
            m_Item(Index).b_Selected = bSelected
        End If
    Else
        If m_HeaderCount >= Relative Then
            m_Header(Relative).colItems(Index).b_Selected = bSelected
        End If
    End If
    If m_AutoRedraw Then RedrawAll
End Property

' Get selected item
Public Property Get SelectedItem$()
    SelectedItem = m_SelIndex(0) & ":" & m_SelIndex(1)
End Property

Private Sub ShortArray(InputArray() As TYPEPCSLISTBOXITEM, bDesc As Boolean)
    Dim i&, j&
    Dim tmp$
    For i = 0 To UBound(InputArray) - 1
        For j = 0 To UBound(InputArray) - 1
            If bDesc Then   ' Sort descending
                If InputArray(j).StrText < InputArray(j + 1).StrText Then
                    tmp = InputArray(j).StrText
                    InputArray(j).StrText = InputArray(j + 1).StrText
                    InputArray(j + 1).StrText = tmp
                End If
            Else            ' Sort acsending
                If InputArray(j).StrText > InputArray(j + 1).StrText Then
                    tmp = InputArray(j).StrText
                    InputArray(j).StrText = InputArray(j + 1).StrText
                    InputArray(j + 1).StrText = tmp
                End If
            End If
    Next j, i
End Sub

Public Sub SortHeaderItems(ByVal HeaderIndex&, Optional bSortDescending As Boolean)
On Error Resume Next
    Call ShortArray(m_Header(HeaderIndex).colItems(), bSortDescending)
End Sub

Public Sub SortNoHeaderItems(Optional bSortDescending As Boolean)
On Error Resume Next
    Call ShortArray(m_Item(), bSortDescending)
End Sub

' Remove selected item
Public Sub RemoveItem(Relative&, Index&)
On Error Resume Next
    Dim i&
    
    ReleaseSelected
    
    If Relative = -1 Then   ' Remove no header item
        If Index > m_ItemCount Then Exit Sub
        If Index = m_ItemCount Then                 'If last item
            GoTo SKIP1:
        Else
            For i = Index To m_ItemCount - 1
                m_Item(i) = m_Item(i + 1)
                m_Item(i).b_Selected = False
            Next
        End If
SKIP1:
        m_ItemCount = m_ItemCount - 1
        ReDim Preserve m_Item(m_ItemCount)
    Else                    ' Remove header item
        If Relative > m_HeaderCount Then Exit Sub
        If Index > m_Header(Relative).lngItemCount Then Exit Sub
        If Index = m_Header(Relative).lngItemCount Then
            GoTo SKIP2:
        Else
            For i = Index To m_Header(Relative).lngItemCount - 1
                m_Header(Relative).colItems(i) = m_Header(Relative).colItems(i + 1)
                m_Header(Relative).colItems(i).b_Selected = False
            Next
        End If
SKIP2:
        m_Header(Relative).lngItemCount = m_Header(Relative).lngItemCount - 1
        ReDim Preserve m_Header(Relative).colItems(m_Header(Relative).lngItemCount)
    End If
    RaiseEvent Change
    Redraw True
End Sub

' Remove selected header
Public Sub RemoveHeader(Index&)
On Error Resume Next
    Dim i&
    If Index > m_HeaderCount Or Index < 0 Then Exit Sub
    If Index = m_HeaderCount Then
        GoTo SKIP:
    Else
        For i = Index To m_HeaderCount - 1
            m_Header(i) = m_Header(i + 1)
        Next
    End If
SKIP:
    m_HeaderCount = m_HeaderCount - 1
    ReDim Preserve m_Header(m_HeaderCount)
    RaiseEvent Change
    Redraw
End Sub

' Clear all
Public Sub Clear()
On Error Resume Next
    Dim i&, tmp%
    Erase m_Header
    Erase m_Item
    tmp = m_ImageCount
    Call InitializeListBox
    m_ImageCount = tmp
    Redraw True
    RaiseEvent Change
End Sub

' Item width
Public Property Get ItemWidth&()
    ItemWidth = UserControl.ScaleWidth - IIf(pSB.Visible, pSB.Width, 0)
End Property

' Header/Non-Header item count
Public Function ItemCount&(ByVal HeaderIndex&)
On Error Resume Next
    If HeaderIndex = -1 Then    ' Get non-header item count
        ItemCount = m_ItemCount + 1
    Else
        If HeaderIndex > m_HeaderCount Then
            ItemCount = -1
        Else
            ItemCount = m_Header(HeaderIndex).lngItemCount + 1
        End If
    End If
End Function

' Header count
Public Function HeaderCount&()
    HeaderCount = m_HeaderCount + 1
End Function
