VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Picasa ListBox Style - siguri92@yahoo.com.vn"
   ClientHeight    =   8535
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5190
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
   ScaleHeight     =   569
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   346
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNext 
      Caption         =   "Give Me More..."
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   7920
      Width           =   2055
   End
   Begin vbprjPCSListBox.pcsListBox Picasa 
      Height          =   7335
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   12938
      BackColor       =   15263976
      Background      =   "frmMain.frx":0000
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
   Begin VB.Shape Shape1 
      BorderColor     =   &H00D3D3D3&
      Height          =   7365
      Left            =   105
      Top             =   465
      Width           =   4965
   End
   Begin VB.Image imgOrg 
      Height          =   240
      Index           =   2
      Left            =   1080
      Picture         =   "frmMain.frx":001C
      Top             =   7920
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgOrg 
      Height          =   240
      Index           =   1
      Left            =   720
      Picture         =   "frmMain.frx":045E
      Top             =   7920
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgOrg 
      Height          =   495
      Index           =   0
      Left            =   120
      Picture         =   "frmMain.frx":08A0
      Top             =   7800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblTip 
      AutoSize        =   -1  'True
      Caption         =   "Original Picasa Style - Sorry I don't have Albums && Project images ;("
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4875
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const strOrg$ = "Albums;People;Project;Folders;Other Stuff"
Dim i&, j&

Private Sub cmdNext_Click()
    frmTest.Show
    Unload Me
End Sub

Private Sub Form_Load()
    Dim arr
    
    ' Init first
    Picasa.InitializeListBox
    
    ' Add image
    For i = 0 To imgOrg.UBound
        Picasa.AddImage imgOrg(i).Picture
    Next
    arr = Split(strOrg, ";")
    For i = 0 To UBound(arr)
        Picasa.AddHeader "k" & i, arr(i)
    Next
    ' Albums item
    Picasa.AddItem 0, "Recently Updated (33)", 1
    ' People item
    Picasa.AddItem 1, "Unnamed (5)", 0
    ' Custom height & padding
    Picasa.CustomItemHeight(1, 0) = 50
    Picasa.CustomItemSpacing(1, 0) = 41
    ' Project item
    Picasa.AddItem 2, "Screen capture", 1
    ' Folder item
    For i = 0 To 109
        Picasa.AddItem 3, IIf(i Mod 10, "Wallpaper " & (Fix(i / 10) + 2000), (Fix(i / 10) + 2000)), 2, i Mod 10 = 0
    Next
    ' Other stuff
    Picasa.AddItem 4, "Examples", 2
    Picasa.RedrawAll
End Sub
