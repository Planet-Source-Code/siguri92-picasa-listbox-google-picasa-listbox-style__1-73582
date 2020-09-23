VERSION 5.00
Begin VB.UserControl imgList 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   2265
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4320
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   151
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   288
End
Attribute VB_Name = "imgList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_ImgCount%
Dim m_Img() As StdPicture

Dim m_ImgSize!

Private Sub AddImage(ByVal New_Image As StdPicture)
    Dim i%
    m_ImgCount = m_ImgCount + 1
    ReDim Preserve m_Img(m_ImgCount)
    Set m_Img(m_ImgCount) = New_Image
    PropertyChanged "ImageCount"
    For i = 0 To m_ImgCount
        PropertyChanged "Image" & i
    Next
    Redraw
End Sub

Public Property Get ImageSize!()
    ImageSize = m_ImgSize
End Property

Public Property Let ImageSize(ByVal New_ImageSize!)
    m_ImgSize = New_ImageSize
    Redraw
    PropertyChanged "ImageSize"
End Property

Public Property Get Image() As StdPicture
    '
End Property

Public Property Set Image(ByVal New_Image As StdPicture)
    AddImage New_Image
End Property

Public Property Get ImageCount%()
    ImageCount = m_ImgCount + 1
End Property

Public Property Let ImageCount(ByVal New_ImageCount%)
    '
End Property


Private Sub Redraw()
On Error Resume Next
    Dim i%
    UserControl.Width = ScaleX((m_ImgCount + 1) * m_ImgSize, vbPixels, vbTwips)
    UserControl.Height = ScaleY(m_ImgSize, vbPixels, vbTwips)
    Cls
    For i = 0 To m_ImgCount
        UserControl.PaintPicture m_Img(i), i * m_ImgSize, 0, m_ImgSize, m_ImgSize, _
                                    0, 0, m_ImgSize, m_ImgSize
    Next
End Sub

Private Sub UserControl_InitProperties()
    m_ImgSize = 16
    m_ImgCount = -1
    Redraw
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    Dim i%
    m_ImgSize = PropBag.ReadProperty("ImageSize", 0)
    m_ImgCount = PropBag.ReadProperty("ImageCount", 0)
    If m_ImgCount > -1 Then
        ReDim m_Img(m_ImgCount)
        For i = 0 To m_ImgCount
            Set m_Img(i) = PropBag.ReadProperty("Image" & i, 0)
        Next
    End If
End Sub

Private Sub UserControl_Resize()
    Redraw
End Sub

Private Sub UserControl_Show()
    Redraw
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
    Dim i%
    PropBag.WriteProperty "ImageSize", m_ImgSize, 0
    PropBag.WriteProperty "ImageCount", m_ImgCount, 0
    If m_ImgCount > -1 Then
        For i = 0 To m_ImgCount
            PropBag.WriteProperty "Image" & i, m_Img(i), 0
        Next
    End If
End Sub

Public Property Get ImageData(Index) As StdPicture
On Error Resume Next
    Set ImageData = m_Img(Index)
End Property
