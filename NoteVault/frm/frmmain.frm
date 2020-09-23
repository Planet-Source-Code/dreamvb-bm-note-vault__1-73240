VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmmain 
   BackColor       =   &H00FFDBBF&
   Caption         =   "BM Note Vault"
   ClientHeight    =   6915
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10575
   Icon            =   "frmmain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   10575
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox txtNote 
      Height          =   870
      Left            =   2505
      TabIndex        =   15
      Top             =   3675
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1535
      _Version        =   393217
      BackColor       =   -2147483624
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmmain.frx":5C1A
   End
   Begin VB.PictureBox pGrip2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   165
      Left            =   6030
      Picture         =   "frmmain.frx":5C9C
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   14
      Top             =   4005
      Visible         =   0   'False
      Width           =   165
   End
   Begin NoteVault.DmDownload DmDownload1 
      Left            =   9375
      Top             =   1455
      _ExtentX        =   423
      _ExtentY        =   423
   End
   Begin VB.PictureBox pIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   105
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   13
      Top             =   5865
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pIcons 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   405
      Picture         =   "frmmain.frx":5E6A
      ScaleHeight     =   240
      ScaleWidth      =   5760
      TabIndex        =   12
      Top             =   5880
      Visible         =   0   'False
      Width           =   5760
   End
   Begin MSComctlLib.ImageList LstIcons 
      Left            =   9480
      Top             =   750
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   16711935
      _Version        =   393216
   End
   Begin NoteVault.Tray Tray1 
      Left            =   8955
      Top             =   1440
      _ExtentX        =   529
      _ExtentY        =   529
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   8130
      Top             =   2055
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   7
      ImageHeight     =   4
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":A6AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":A75E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pGrip1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   165
      Left            =   6015
      Picture         =   "frmmain.frx":A810
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   10
      Top             =   3795
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox sBar1 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF80FF&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   0
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   705
      TabIndex        =   9
      Top             =   6615
      Width           =   10575
      Begin VB.Label lblItems 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   11
         Top             =   45
         Width           =   135
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   8880
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox pTitle3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   2505
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   194
      TabIndex        =   7
      Top             =   3300
      Width           =   2910
      Begin VB.Image imgInfo 
         Height          =   240
         Left            =   60
         Picture         =   "frmmain.frx":A9DE
         Top             =   75
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblNoteTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   345
         TabIndex        =   8
         Top             =   60
         Width           =   240
      End
   End
   Begin MSComctlLib.ListView LstV 
      Height          =   2490
      Left            =   2505
      TabIndex        =   6
      Top             =   750
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   4392
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "LstIcons"
      SmallIcons      =   "LstIcons"
      ColHdrIcons     =   "ImageList2"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Title"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Creation Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Last Updated"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Priority"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.PictureBox pTitle2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   2505
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   194
      TabIndex        =   4
      Top             =   345
      Width           =   2910
      Begin VB.Label lblGroup 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   60
         TabIndex        =   5
         Top             =   60
         Width           =   240
      End
   End
   Begin MSComctlLib.ImageList tvImg 
      Left            =   8145
      Top             =   1410
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":AA60
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":ADB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":B104
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8145
      Top             =   810
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":B456
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":B7A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":BB3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":BE8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":C1DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":C570
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":C8C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":CC14
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tv1 
      Height          =   4515
      Left            =   15
      TabIndex        =   3
      Top             =   750
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   7964
      _Version        =   393217
      Indentation     =   635
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "tvImg"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox pTitle1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   15
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   164
      TabIndex        =   1
      Top             =   345
      Width           =   2460
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "My Notes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   75
         TabIndex        =   2
         Top             =   60
         Width           =   990
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NEW"
            Object.ToolTipText     =   "New Database"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "OPEN"
            Object.ToolTipText     =   "Open Database"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "BACKUP"
            Object.ToolTipText     =   "Backup Database"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ADD_GROUP"
            Object.ToolTipText     =   "Add Group"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DELETE_GROUP"
            Object.ToolTipText     =   "Delete Group"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ADD_NOTE"
            Object.ToolTipText     =   "New Note"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "EDIT_NOTE"
            Object.ToolTipText     =   "Edit Note"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DEL_NOTE"
            Object.ToolTipText     =   "Delete Note"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin VB.Image imgTexture2 
      Height          =   390
      Left            =   5565
      Picture         =   "frmmain.frx":CF66
      Top             =   4215
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgMenu 
      Height          =   1665
      Left            =   6285
      Picture         =   "frmmain.frx":D6F8
      Top             =   3495
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.Image imgTexture1 
      Height          =   390
      Left            =   5550
      Picture         =   "frmmain.frx":E26F
      Top             =   3720
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "New..."
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "Backup As"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBlank1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuGroup 
      Caption         =   "GroupMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuNewNote3 
         Caption         =   "New Note"
      End
      Begin VB.Menu mnuBlank5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewGroup 
         Caption         =   "New Group"
      End
      Begin VB.Menu mnuRename 
         Caption         =   "Rename"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu mnuNote 
      Caption         =   "NoteGroup"
      Visible         =   0   'False
      Begin VB.Menu mnuEditNote 
         Caption         =   "Edit"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuNewNote 
         Caption         =   "New Note"
      End
      Begin VB.Menu mnuDel 
         Caption         =   "Delete"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuBlank2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMove 
         Caption         =   "Move"
      End
   End
   Begin VB.Menu mnuAction 
      Caption         =   "&Action"
      Begin VB.Menu mnuNotes 
         Caption         =   "&Notes"
         Begin VB.Menu mnuEdit1 
            Caption         =   "Edit"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuNewNote1 
            Caption         =   "New Note"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuDelete2 
            Caption         =   "Delete"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuBlank3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuMove1 
            Caption         =   "Move"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuGroups 
         Caption         =   "Groups"
         Begin VB.Menu mnuNew1 
            Caption         =   "New Group"
         End
         Begin VB.Menu mnuRename1 
            Caption         =   "Rename"
         End
         Begin VB.Menu mnuDelete1 
            Caption         =   "&Delete"
         End
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuMyNotes 
         Caption         =   "My Notes"
      End
      Begin VB.Menu mnuToolbar 
         Caption         =   "Toolbar"
      End
      Begin VB.Menu mnuStatusbar 
         Caption         =   "Status Bar"
      End
      Begin VB.Menu mnuNoteInfo 
         Caption         =   "Note Information"
      End
      Begin VB.Menu mnuBlank7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLayout 
         Caption         =   "Layout"
         Begin VB.Menu mnuReport 
            Caption         =   "Report"
         End
         Begin VB.Menu mnuIcons 
            Caption         =   "Icons"
         End
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuPws 
         Caption         =   "Set Database Password"
      End
      Begin VB.Menu mnuCompact 
         Caption         =   "Compact Repair Database"
      End
      Begin VB.Menu mnuBlank8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuReadme 
         Caption         =   "Readme"
      End
      Begin VB.Menu mnuBlank6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUpdate 
         Caption         =   "Check for updates"
      End
      Begin VB.Menu mnuBlank9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVisit 
         Caption         =   "Visit BM Note Vault Home page"
      End
      Begin VB.Menu mnuContact 
         Caption         =   "Contact Us"
      End
      Begin VB.Menu mnuBlank10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About..."
      End
   End
   Begin VB.Menu mnuTray 
      Caption         =   "Tray"
      Visible         =   0   'False
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu mnuAbout1 
         Caption         =   "About"
      End
      Begin VB.Menu mnuBlank4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit1 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGroupID As Integer
Private mNoteID As Integer

'Store Treeview and Listview Selected Index
Private mTvID As Integer
Private mLstID As Integer
Private oWinState As Integer

Private HasLoaded As Boolean
Private mButton1 As MouseButtonConstants
Private mButton2 As MouseButtonConstants
Private mFileAttr As Long
Private Const Version As Integer = 7 '= ver 1.1 + build 5

Public Sub SetSkin(sType As Integer)
    'This part set the skin for the program
    SkinType = sType
    
    If (sType = 0) Then
        frmmain.BackColor = &HFFDBBF
        hBrush1 = CreatePatternBrush(imgTexture1.Picture.Handle)
        SetMenuColour frmmain.Hwnd, 0, True
        Call SetToolbarBG(Toolbar1.Hwnd, imgTexture1.Picture)
        'Set label forecolors
        lblTitle.ForeColor = vbBlack
        lblGroup.ForeColor = vbBlack
        lblNoteTitle.ForeColor = vbBlack
    Else
        hBrush1 = CreatePatternBrush(imgTexture2.Picture.Handle)
        Call pTitle1.Cls
        Call pTitle2.Cls
        Call pTitle3.Cls
        SetMenuColour frmmain.Hwnd, vbWhite, True
        Call SetToolbarBG(Toolbar1.Hwnd, imgTexture2.Picture)
        frmmain.BackColor = vbButtonFace
        'Set label forecolors
        lblTitle.ForeColor = vbWhite
        lblGroup.ForeColor = vbWhite
        lblNoteTitle.ForeColor = vbWhite
    End If
    
    Call sBar1_Resize
    Call Form_Resize
    
End Sub

Public Sub SetNodeColor(ByVal Color As Long)
Dim nNode As Node
    For Each nNode In tv1.Nodes
        nNode.ForeColor = Color
    Next nNode
End Sub

Private Sub HideNoteInformation(ByVal tHide As Boolean)
    'Show / hide controls
    pTitle3.Visible = tHide
    txtNote.Visible = tHide
    'Resize form
    Call Form_Resize
End Sub

Private Sub HideMyNotes(ByVal tHide As Boolean)
    'Hide / show Notes
    pTitle1.Visible = tHide
    tv1.Visible = tHide
    
    If (Not tHide) Then
        pTitle2.Left = 15
        pTitle3.Left = 15
        LstV.Left = 15
        txtNote.Left = 15
    Else
        LstV.Left = 2505
        pTitle2.Left = 2505
        pTitle3.Left = 2505
        txtNote.Left = 2505
    End If
    
    Call Form_Resize

End Sub

Private Sub HideStatusBar(ByVal tHide As Boolean)
    'Hide / show statusbar
    sBar1.Visible = tHide
    'Set statusbar height
    sBar1.Height = IIf(Not tHide, 0, 300)
    'Call form resize
    Call Form_Resize
End Sub

Private Sub HideToolBar(ByVal tHide As Boolean)
    'Hide / show toolbar
    Toolbar1.Visible = tHide
    
    If (Not tHide) Then
        pTitle1.Top = 0
        pTitle2.Top = 0
        tv1.Top = (pTitle1.Top + pTitle1.Height) + 30
        LstV.Top = tv1.Top
        LstV.Height = (pTitle3.Top - pTitle3.Height) - 60
    Else
        pTitle1.Top = 345
        pTitle2.Top = 345
        tv1.Top = 750
        LstV.Top = 750
        LstV.Height = 2925
    End If
    
    'Call form resize
    Call Form_Resize
    
End Sub

Public Sub SetIcons()
Dim Ret As Long
Dim Count As Integer

    For Count = 0 To 23
        Ret = BitBlt(pIcon.hdc, 0, 0, 16, 16, pIcons.hdc, (16 * Count), 0, vbSrcCopy)
        Call pIcon.Refresh
        'Add the image to the image list
        LstIcons.ListImages.Add , , pIcon.Image
    Next Count

End Sub

Private Function InTreeview(ByVal sFind As String) As Integer
Dim nNode As Node

    'Check if node text is in the list
    For Each nNode In tv1.Nodes
        If (nNode.Key <> "M_TOP") Then
            If LCase(nNode.Text) = LCase(sFind) Then
                InTreeview = 1
                Exit For
            End If
        End If
    Next nNode
    
    'Clear up
    Set nNode = Nothing
    
End Function

Private Function GetDLGName(Optional ShowOpen As Boolean = True, Optional Title As String = "Open", Optional Filter As String)
On Error GoTo CanErr:
    'Show open or save dialog.
    With CD1
        .CancelError = True
        .DialogTitle = Title
        .Filter = Filter
        
        If (ShowOpen) Then
           Call .ShowOpen
        Else
           Call .ShowSave
        End If
        
        GetDLGName = .FileName
        .FileName = vbNullString
    End With
    
    Exit Function
CanErr:
    If (Err.Number = cdlCancel) Then
        Err.Clear
    End If
End Function

Private Sub ClickButton(ByVal Index As Integer)
Dim mButton As Button
On Error Resume Next
    Set mButton = Toolbar1.Buttons(Index)
    Call Toolbar1_ButtonClick(mButton)
    'Destroy mButton
    Set mButton = Nothing
End Sub

Private Sub ClickNode(ByVal Index As Integer)
On Error Resume Next
Dim nNode As Node
    Set nNode = tv1.Nodes(Index)
    nNode.Selected = True
    Call tv1_NodeClick(nNode)
    Call tv1.SetFocus
    'Destroy node
    Set nNode = Nothing
End Sub

Private Sub ClickItem(ByVal Index As Integer)
Dim lItem As ListItem
On Error Resume Next
    Set lItem = LstV.ListItems(Index)
    lItem.Selected = True
    Call LstV_ItemClick(lItem)
    Call LstV.SetFocus
    'Destroy Listitem
    Set lItem = Nothing
End Sub

Private Sub DmDownload1_DownloadComplete(mCurBytes As Long, mMaxBytes As Long, LocalFile As String)
Dim fp As Long
Dim sBuff As String
On Error Resume Next
    
    'Check for update
    fp = FreeFile
    Open LocalFile For Binary As #fp
        sBuff = Space(LOF(fp))
        Get #fp, , sBuff
    Close #fp
    
    'Delete the update file we do not need it
    Call Kill(LocalFile)
    
    If Val(sBuff) > Version Then
        If MsgBox("There is a new update available do you want to download it now?", vbYesNo Or vbQuestion, frmmain.Caption) = vbYes Then
            Call OpenUrl("http://www.bm-it-software.co.uk/dload/setup.zip")
        End If
    Else
        Call MsgBox("No new updates are available.", vbInformation, frmmain.Caption)
    End If
    
    'Clear up
    sBuff = ""
End Sub

Private Sub Form_Activate()
    If (HasLoaded) Then
        'Set treeview font and color
        Call SetFont(mConfig.GroupFont, tv1)
        
        'Check for nodes and select the first child
        If (tv1.Nodes.Count = 1) Then
            Call ClickNode(1)
        Else
            Call ClickNode(2)
        End If
        If (LstV.ListItems.Count) Then
            Call ClickItem(1)
        End If
        HasLoaded = False
    End If
End Sub

Private Sub Form_Initialize()
Dim x As Long
    x = InitCommonControls
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    'Check for quick exit
    If (KeyCode = vbKeyEscape) And (mConfig.QuickExit = 1) Then
        Call mnuExit_Click
    End If
End Sub

Private Sub Form_Load()
Dim ch As ColumnHeader

    If IsAppOpen Then
        MsgBox "An Instance of BM Note Vault is already open.", vbInformation, App.ProductName
        Call Unload(frmmain)
        Exit Sub
    End If
    
    'Disable buttons
    Toolbar1.Buttons(5).Enabled = False
    Toolbar1.Buttons(6).Enabled = False
    Toolbar1.Buttons(8).Enabled = False
    Toolbar1.Buttons(9).Enabled = False
    Toolbar1.Buttons(10).Enabled = False
    'Disable menu items
    mnuNew1.Enabled = False
    mnuRename1.Enabled = False
    mnuDelete1.Enabled = False
    mnuPws.Enabled = False
    mnuCompact.Enabled = False
    
    'Load config
    Call LoadConfig
    'Load the list icons
    Call SetIcons
    
    'Select first Column header
    Set ch = LstV.ColumnHeaders(1)
    Call LstV_ColumnClick(ch)
    'Setup listview view
    LstV.View = mConfig.View
    'Check if loading program maxsized
    frmmain.WindowState = IIf(mConfig.Maxsized, vbMaximized, vbNormal)
    
    'Set control fonts
    Call SetFont(mConfig.RecordFont, LstV)
    'Create menu brush
    hBrush2 = CreatePatternBrush(imgMenu.Picture.Handle)
    'Default skin
    SkinType = mConfig.Skin
    'Set menu background
    Call SetSkin(SkinType)
    
    'Setup tray control
    Tray1.ToolTip = frmmain.Caption
    Set Tray1.Icon = frmmain.Icon
    
    'Set captions
    lblNoteTitle.Caption = ""
    lblGroup.Caption = ""
    lblItems.Caption = ""
    
    'Check if opening from command line
    If Len(Command$) Then
        'Replace "
        dbFile = Replace(Command$, Chr$(34), "", , , vbBinaryCompare)
    Else
        dbFile = mConfig.DatabaseFile
    End If
    
    'Set view toolbar
    mnuToolbar.Checked = mConfig.Toolbar
    Call mnuToolbar_Click
    'Set statusbar
    mnuStatusbar.Checked = mConfig.StatusBar
    Call mnuStatusbar_Click
    'Set My Notes bar
    mnuMyNotes.Checked = mConfig.NotesBar
    Call mnuMyNotes_Click
    'Set note information
    mnuNoteInfo.Checked = mConfig.NoteInfo
    Call mnuNoteInfo_Click
    
    If Not FindFile(dbFile) Then
        'Put error here
        Exit Sub
    End If
    
    'Get file attr
    mFileAttr = GetAttr(dbFile)
    
    'Check if the file is read only
    If (mFileAttr And vbReadOnly) Then
        'Remove read only state
        Call SetAttr(dbFile, vbNormal)
    End If
    
    'Open database
    Call OpenDataBaseA
    
    If (Not dbOpen) Then
        'Put error here
    Else
        'Get password
        Pass1 = GetPassword
        'Check for password
        If Len(Pass1) = 0 Then
            'Load the groups into Treeview
            Call LoadGroups(tv1)
        Else
            'Show password dialog
            'Call frmmain.Show
            Call frmmain.Show
            Call frmPws1.Show(vbModal, frmmain)
            
            dbOpen = (Pass1 = Pass2)
            'Only load the groups if the password was correct
            If (dbOpen) Then
                Call LoadGroups(tv1)
            End If
        End If
    End If
    
    'Enable / toolbar and menu items
    Toolbar1.Buttons(3).Enabled = dbOpen
    Toolbar1.Buttons(5).Enabled = dbOpen
    mnuBackup.Enabled = dbOpen
    mnuNew1.Enabled = dbOpen
    mnuPws.Enabled = dbOpen
    mnuCompact.Enabled = dbOpen
    HasLoaded = True

End Sub

Private Sub Form_Resize()
On Error Resume Next
    
    If (frmmain.WindowState <> 1) Then
        oWinState = frmmain.WindowState
    Else
        If (mConfig.MinToTray) Then
            frmmain.Visible = False
            Tray1.Visible = True
        End If
    End If

    'Resize controls
    pTitle2.Width = (frmmain.ScaleWidth - pTitle2.Left)
    pTitle3.Width = pTitle2.Width
    LstV.Width = pTitle2.Width
    txtNote.Width = pTitle2.Width
    
    tv1.Height = (frmmain.ScaleHeight - sBar1.Height - tv1.Top)
    
    If (txtNote.Visible) Then
        LstV.Height = (frmmain.ScaleHeight - sBar1.Height - LstV.Top) / 2
    Else
        LstV.Height = tv1.Height
    End If
    
    pTitle3.Top = LstV.Height + (360 * 2) + 60
    txtNote.Top = (pTitle3.Top + pTitle3.Height) + 10
    txtNote.Height = (frmmain.ScaleHeight - sBar1.Height - txtNote.Top)
    
    'Texture textboxes
    If (SkinType = 0) Then
        Call TextPBox(pTitle1)
        Call TextPBox(pTitle2)
        Call TextPBox(pTitle3)
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CloseDB
    
    'Restore file attr
    If FindFile(dbFile) Then
        Call SetAttr(dbFile, mFileAttr)
    End If
    
    Set frmmain = Nothing
End Sub

Private Sub LstV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Static sSort As Integer
Dim I As Long

    sSort = (Not sSort)
    LstV.SortKey = ColumnHeader.Index - 1
    LstV.SortOrder = Abs(sSort)
    LstV.Sorted = True
    
    For I = 0 To LstV.ColumnHeaders.Count - 1
      If I = LstV.SortKey Then
            Call ShowHeaderIcon(LstV, LstV.SortKey, LstV.SortOrder, True)
        Else
            Call ShowHeaderIcon(LstV, I, 0, False)
      End If
   Next
   
   I = 0

End Sub

Private Sub LstV_DblClick()
    If (mButton2 = vbLeftButton) And (LstV.ListItems.Count) Then
        Call ClickButton(9)
    End If
End Sub

Private Sub LstV_ItemClick(ByVal Item As MSComctlLib.ListItem)

    'Store Listview ID
    mLstID = Item.Index
    'Store Note ID
    mNoteID = Val(Mid$(Item.Key, 2))
    'Store Note Info
    mNoteInfo.NoteTitle = Item.Text
    mNoteInfo.NoteAdded = Item.SubItems(1)
    mNoteInfo.NoteLastModifed = Item.SubItems(2)
    mNoteInfo.NoteData = GetNote(mNoteID)
    mNoteInfo.NoteColor = Item.ForeColor
    mNoteInfo.NoteIcon = Val(Item.Tag)
    mNoteInfo.NotePriority = Item.SubItems(3)
    
    'Display the note in the textbox
    txtNote.TextRTF = mNoteInfo.NoteData
    'Update note display label
    lblNoteTitle.Caption = mNoteInfo.NoteTitle & " Information"
    'Show info icon
    imgInfo.Visible = True
    'Enable Toolbar buttons
    Toolbar1.Buttons(9).Enabled = True
    Toolbar1.Buttons(10).Enabled = True
    'Enable menu items
    mnuEdit1.Enabled = True
    mnuDelete2.Enabled = True
    mnuMove1.Enabled = True
    
    If (mButton2 = vbRightButton) Then
        Call PopupMenu(mnuNote)
    End If
    
End Sub

Private Sub LstV_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete note if delete key is pressed
    If (KeyCode = vbKeyDelete) And Toolbar1.Buttons(10).Enabled Then
        Call ClickButton(10)
    End If
    'Rename note if F2 key is pressed
    If (KeyCode = vbKeyF2) Or (KeyCode = vbKeyReturn) And Toolbar1.Buttons(9).Enabled Then
        Call ClickButton(9)
    End If
    
End Sub

Private Sub LstV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mButton2 = Button

    If (mButton2 = vbRightButton) And (tv1.Nodes.Count > 1) Then
        mnuEditNote.Enabled = LstV.ListItems.Count
        mnuDel.Enabled = mnuEditNote.Enabled
        mnuMove.Enabled = mnuEditNote.Enabled
        Call PopupMenu(mnuNote)
    End If
    
End Sub

Private Sub mnuAbout_Click()
    'Show about form
    Call frmSplash.Show(vbModal, frmmain)
End Sub

Private Sub mnuAbout1_Click()
    'Show about form
    Call mnuAbout_Click
End Sub

Private Sub mnuBackup_Click()
Dim lzFilename As String
    'Get save filename form dialog.
    lzFilename = GetDLGName(False, "Backup", mFilter)
    
    If Len(lzFilename) Then
        'Close the database first
        Call CloseDB
        'Make a copy of the database
        Call FileCopy(dbFile, lzFilename)
        'Reopen the database
        Call OpenDataBaseA
        'Clear up
        lzFilename = vbNullString
    End If
End Sub

Private Sub mnuCompact_Click()
Dim tmpFile As String
    'Create temp file
    tmpFile = GetPathFormFile(dbFile) & "tmp.mdb"
    
    If (dbOpen) Then
        Call CloseDB
    End If
    'Set attr to notmal
    Call SetAttr(dbFile, vbNormal)
    'Compact the database
    Call CompactDatabase(dbFile, tmpFile, dbLangGeneral, , ";pwd=33CC66")
    'Delete original file
    Call Kill(dbFile)
    'Rename the new database as the original one
    Name tmpFile As dbFile
    'Re-open the original database
    Call OpenDataBaseA
    'Disaplay sucess message
    Call MsgBox("The database was repaired.", vbInformation, "Compact Repair Database")

End Sub

Private Sub mnuContact_Click()
    Call OpenUrl("mailto:bm-it-software@gmx.com")
End Sub

Private Sub mnuDel_Click()
    'Delete note
    Call ClickButton(10)
End Sub

Private Sub mnuDelete_Click()
    'Delete Group
    Call ClickButton(6)
End Sub

Private Sub mnuDelete1_Click()
    'Delete group
    Call mnuDelete_Click
End Sub

Private Sub mnuDelete2_Click()
    'Delete note
    Call mnuDel_Click
End Sub

Private Sub mnuEdit1_Click()
    'Edit note
    Call mnuEditNote_Click
End Sub

Private Sub mnuEditNote_Click()
    'Edit note
    Call ClickButton(9)
End Sub

Private Sub mnuExit_Click()
    Call CloseDB
    Call Unload(frmmain)
End Sub

Private Sub mnuExit1_Click()
    'Exit program
    Call mnuExit_Click
End Sub

Private Sub mnuIcons_Click()
    'Icon view
    LstV.View = lvwIcon
    Call SaveSetting("NoteVault", "Config", "ListView", lvwIcon)
End Sub

Private Sub mnuMove_Click()
    'Show move note dialog
    Call frmmove.Show(vbModal, frmmain)
    
    If (ButtonPress = vbOK) Then
        If MoveNote(mNoteID, mGroupIndex) <> 1 Then
            Call MsgBox("There was an error moving the note.", vbExclamation, "Move Note")
        Else
            Call ClickNode(mTvID)
        End If
    End If
    
    ButtonPress = vbCancel
    
End Sub

Private Sub mnuMove1_Click()
    'Move note
    Call mnuMove_Click
End Sub

Private Sub mnuMyNotes_Click()
    mnuMyNotes.Checked = (Not mnuMyNotes.Checked)
    Call HideMyNotes(mnuMyNotes.Checked)
    
    'Hide/show My Notes
    If (Not mnuMyNotes.Checked) Then
        Call SaveSetting("NoteVault", "Config", "NotesBar", 1)
    Else
        Call SaveSetting("NoteVault", "Config", "NotesBar", 0)
    End If

End Sub

Private Sub mnuNew_Click()
Dim lzFilename As String
    'Get Filename from dialog
    lzFilename = GetDLGName(False, "New Database", mFilter)
    
    If Len(lzFilename) Then
        'Create the database
        If CreateNewDatabase(lzFilename) <> 1 Then
            Call MsgBox("There was an error creating the database." & vbCrLf & "Possible reason(s):" & vbCrLf & "Database already exists." _
            , vbExclamation, "Create Database")
        End If
    End If
    
End Sub

Private Sub mnuNew1_Click()
    'Add group
    Call ClickButton(5)
End Sub

Private Sub mnuNewGroup_Click()
    'Add group
    Call ClickButton(5)
End Sub

Private Sub mnuNewNote_Click()
    'New note
    Call ClickButton(8)
End Sub

Private Sub mnuNewNote1_Click()
    'New note
    Call mnuNewNote_Click
End Sub

Private Sub mnuNewNote3_Click()
    'Add note
    Call ClickButton(8)
End Sub

Private Sub mnuNoteInfo_Click()
    mnuNoteInfo.Checked = (Not mnuNoteInfo.Checked)
    Call HideNoteInformation(mnuNoteInfo.Checked)
    
    'Hide/show  note information
    If (Not mnuNoteInfo.Checked) Then
        Call SaveSetting("NoteVault", "Config", "NoteInfo", 1)
    Else
        Call SaveSetting("NoteVault", "Config", "NoteInfo", 0)
    End If
End Sub

Private Sub mnuOpen_Click()
Dim lzFilename As String
Dim OldFile As String
    
    OldFile = dbFile
    
    'Get filename from dialog
    lzFilename = GetDLGName(, "Open Database", mFilter)
    
    If Len(lzFilename) Then
        
        If (dbOpen) Then
            'Close open database
            Call CloseDB
        End If
        
        'Store old filename
        dbFile = lzFilename
        'Open the database
        Call OpenDataBaseA
        'Check if database is opened
        If (Not dbOpen) Then
            Call MsgBox("There was an error opening the database.", vbExclamation, "Open Database Error")
            'Reopen the last database
            dbFile = OldFile
            Call OpenDataBaseA
        Else
            'Get password
            Pass1 = GetPassword
            'Check for password
            Call tv1.Nodes.Clear
            Call LstV.ListItems.Clear
            'Clear captions
            lblGroup.Caption = ""
            lblItems.Caption = ""
            'Disable buttons
            Toolbar1.Buttons(5).Enabled = False
            Toolbar1.Buttons(6).Enabled = False
            Toolbar1.Buttons(8).Enabled = False
            Toolbar1.Buttons(9).Enabled = False
            Toolbar1.Buttons(10).Enabled = False
            'Disable menu items
            mnuNew1.Enabled = False
            mnuRename1.Enabled = False
            mnuDelete1.Enabled = False
            mnuNewNote1.Enabled = False
            mnuEdit1.Enabled = False
            mnuDelete2.Enabled = False
            mnuMove1.Enabled = False
            mnuPws.Enabled = False
            mnuCompact.Enabled = False
            'Clear textbox
            txtNote.Text = ""
            
            If Len(Pass1) = 0 Then
                'Load the groups into Treeview
                Call LoadGroups(tv1)
            Else
                Call frmPws1.Show(vbModal, frmmain)
                'Check for correct password
                dbOpen = (Pass1 = Pass2)
                'Check if database is opened
                If (dbOpen) Then
                    Call LoadGroups(tv1)
                End If
            End If
            
            If (tv1.Nodes.Count > 1) Then
                Call ClickNode(2)
            Else
                Call ClickNode(1)
            End If
            'Select the first item in the notes list
            If (LstV.ListItems.Count) Then
                Call ClickItem(1)
            End If
        End If
    End If
    
    'Enable / disable menu and toolbar items
    Toolbar1.Buttons(5).Enabled = tv1.Nodes.Count
    Toolbar1.Buttons(3).Enabled = dbOpen
    mnuBackup.Enabled = dbOpen
    mnuNew1.Enabled = dbOpen
    mnuPws.Enabled = dbOpen
    mnuCompact.Enabled = dbOpen
End Sub

Private Sub mnuOptions_Click()
    'Show options dialog
    Call frmOptions.Show(vbModal, frmmain)
    
    If (ButtonPress = vbOK) Then
        'Set font for controls
        Call SetFont(mConfig.GroupFont, tv1)
        Call SetFont(mConfig.RecordFont, LstV)
    End If
    
End Sub

Private Sub mnuPws_Click()
    'Show set password dialog
    Call frmPws2.Show(vbModal, frmmain)
    If (ButtonPress = vbOK) Then
        Call SetPassword(Pass2)
    End If
End Sub

Private Sub mnuReadme_Click()
Dim sFile As String
    sFile = FixPath(App.Path) & "readme.txt"
    'Check if file is found
    If Not FindFile(sFile) Then
        Call MsgBox("File not found:" & vbCrLf & sFile, vbCritical, "File Not Found")
    Else
        'Open readme file
        Call OpenUrl(sFile)
    End If
End Sub

Private Sub mnuRename_Click()
    mGroupAdd = False
    'Rename Group
    Call frmrename.Show(vbModal, frmmain)
    'Check if OK was pressed
    If (ButtonPress <> vbOK) Then
        Exit Sub
    Else
        If Len(mGroupName) Then
            'Do nothing if the name is the same
            If LCase(tv1.SelectedItem.Text) = LCase(mGroupName) Then
                Exit Sub
            End If
            'Check if group is already in the list
            If InTreeview(mGroupName) Then
                Call MsgBox("The group '" & mGroupName & "' " & "is already in the list." & vbCrLf & _
                "Please try a different name.", vbExclamation, "Add Group")
                Exit Sub
            End If
            'Do rename group
            If RenameGroup(mGroupName, mGroupID) <> 1 Then
                Call MsgBox("There was an error while renaming '" & tv1.SelectedItem.Text & "'", vbExclamation, "Rename Group")
            Else
                tv1.Nodes(mTvID).Text = mGroupName
            End If
        End If
    End If
    
    ButtonPress = vbCancel
    'Clear up
    mGroupName = vbNullString
End Sub

Private Sub mnuRename1_Click()
    'Rename group
    Call mnuRename_Click
End Sub

Private Sub mnuReport_Click()
    'Report view
    LstV.View = lvwReport
    Call SaveSetting("NoteVault", "Config", "ListView", lvwReport)
End Sub

Private Sub mnuRestore_Click()
    'Restore
    Call Tray1_MouseDown(vbLeftButton)
End Sub

Private Sub mnuStatusbar_Click()
    mnuStatusbar.Checked = (Not mnuStatusbar.Checked)
    Call HideStatusBar(mnuStatusbar.Checked)
    
    'Hide/show  statusbar
    If (Not mnuStatusbar.Checked) Then
        Call SaveSetting("NoteVault", "Config", "StatusBar", 1)
    Else
        Call SaveSetting("NoteVault", "Config", "StatusBar", 0)
    End If

End Sub

Private Sub mnuToolbar_Click()
    mnuToolbar.Checked = (Not mnuToolbar.Checked)
    Call HideToolBar(mnuToolbar.Checked)
    
    'Hide/show  toolbar
    If (Not mnuToolbar.Checked) Then
        Call SaveSetting("NoteVault", "Config", "Toolbar", 1)
    Else
        Call SaveSetting("NoteVault", "Config", "Toolbar", 0)
    End If
    
End Sub

Private Sub mnuUpdate_Click()
    'Check for updates
    Call DmDownload1.DownloadFile("http://www.bm-it-software.co.uk/dload/dmvault.txt", _
    FixPath(App.Path) & "update.txt", vbAsyncTypeByteArray)
End Sub

Private Sub mnuVisit_Click()
    Call OpenUrl("http://www.bm-it-software.co.uk/")
End Sub

Private Sub sBar1_Resize()
    'Draw satusbar
    Call TextPBox(sBar1)
    If (SkinType = 0) Then
        TransparentBlt sBar1.hdc, (sBar1.ScaleWidth - 12), (sBar1.ScaleHeight - 12), 11, 11, pGrip1.hdc, 0, 0, 11, 11, vbMagenta
    Else
        TransparentBlt sBar1.hdc, (sBar1.ScaleWidth - 12), (sBar1.ScaleHeight - 12), 11, 11, pGrip2.hdc, 0, 0, 11, 11, vbMagenta
    End If
    
    sBar1.Refresh
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Ans As VbMsgBoxResult
Dim sGroup As String

    Select Case Button.Key
        Case "NEW"
            'Create new database
            Call mnuNew_Click
        Case "OPEN"
            'Open database
            Call mnuOpen_Click
        Case "BACKUP"
            'Do Backup
            Call mnuBackup_Click
        Case "ADD_GROUP"
            'Show add group form
            mGroupAdd = True
            Call frmrename.Show(vbModal, frmmain)
            
            'Check if ok was pressed
            If (ButtonPress <> vbOK) Then
                mGroupAdd = False
                Exit Sub
            Else
                'Check for group name
                If Len(mGroupName) Then
                    'Check if group is already in the list
                    If InTreeview(mGroupName) Then
                        Call MsgBox("The group '" & mGroupName & "' " & "is already in the list." & vbCrLf & _
                        "Please try a different name.", vbExclamation, "New Group")
                        ButtonPress = vbCancel
                        Exit Sub
                    End If
                    'Add the group
                    If AddGroup(mGroupName) <> 1 Then
                        Call MsgBox("There was an error adding the new group '" & mGroupName & "'", vbExclamation, "New Group")
                    Else
                        'Reload te groups
                        Call LoadGroups(tv1)
                        If (tv1.Nodes.Count = 2) Then
                            'Click second node
                            Call ClickNode(2)
                        ElseIf (mTvID = 1) Then
                            'Click last index
                            Call ClickNode(tv1.Nodes.Count)
                        Else
                            'Click exsiting node
                            Call ClickNode(mTvID)
                        End If
                    End If
                End If
                mGroupAdd = False
            End If
        Case "DELETE_GROUP"
            'Delete group
            Ans = MsgBox("Warning you are about to delete the group '" & mGroupName & "'" & _
            vbCrLf & vbCrLf & "Deleting this group will also delete all the notes in this group." _
            & vbCrLf & vbCrLf & "Are you sure you want to delete this group.", vbYesNo Or vbQuestion, "Delete Group")
    
            If (Ans = vbYes) Then
                If DeleteGroup(mGroupID) <> 1 Then
                    Call MsgBox("There was an error deleting the group '" & mGroupName & "'", vbExclamation, "Delete Group")
                Else
                    'Reload te groups
                    Call LoadGroups(tv1)
                    If (tv1.Nodes.Count > 1) Then
                        'Click second node
                        Call ClickNode(2)
                    Else
                        'Click first node
                        Call ClickNode(1)
                    End If
                End If
            End If
            
        Case "ADD_NOTE"
            mAddNote = True
            'Show add note form
            Call frmAddNote.Show(vbModal, frmmain)
            
            If (ButtonPress = vbOK) Then
                If AddNote(mGroupID) <> 1 Then
                    Call MsgBox("There was an error adding the new note.", vbExclamation, "New Note")
                Else
                    'Load the notes for each group
                    Call LoadNotes(LstV, mGroupID)
                    'Select the last index
                    Call ClickItem(LstV.ListItems.Count)
                End If
            End If
        Case "EDIT_NOTE"
            mAddNote = False
            'Show edit note form
            Call frmAddNote.Show(vbModal, frmmain)
            
            If (ButtonPress = vbOK) Then
                If EditNote(mNoteID) <> 1 Then
                    Call MsgBox("There was an error editing the note.", vbExclamation, "Edit Note")
                Else
                    'Load the notes for each group
                    Call LoadNotes(LstV, mGroupID)
                    
                    'Check if we are in all notes selection
                    If (tv1.SelectedItem.Index = 1) Then
                        Call ClickNode(1)
                    End If
                    
                    'Select last index
                    Call ClickItem(mLstID)
                End If
            End If
        Case "DEL_NOTE"
            'Ask user if they want to delete the note
            If MsgBox("Are you sure you want to delete '" & LstV.SelectedItem.Text & "'", vbYesNo Or vbQuestion, "Delete Note") = vbYes Then
                If DeleteNote(mNoteID) <> 1 Then
                    Call MsgBox("There was an error deleting '" & LstV.SelectedItem.Text & "'", vbExclamation, "Delete Note")
                Else
                    Call ClickNode(mTvID)
                End If
            End If
    End Select
    
    ButtonPress = vbCancel
End Sub

Private Sub Tray1_MouseDown(Button As Integer)
    If (Button = vbLeftButton) Then
        Tray1.Visible = False
        frmmain.WindowState = oWinState
        frmmain.Visible = True
    End If
    'Show popup menu
    If (Button = vbRightButton) Then
        Call PopupMenu(mnuTray)
    End If
End Sub

Private Sub tv1_DblClick()
    If (mButton1 = vbLeftButton) And tv1.SelectedItem.Key <> "M_TOP" Then
        'Rename group
        If mnuRename1.Enabled Then
            'Rename group
            Call mnuRename_Click
        End If
    End If
End Sub

Private Sub tv1_KeyDown(KeyCode As Integer, Shift As Integer)

    If (tv1.SelectedItem.Key = "M_TOP") Then
        Exit Sub
    End If
    'Delete group
    If (KeyCode = vbKeyDelete) Then
        Call ClickButton(6)
    End If
    'Rename group
    If (KeyCode = vbKeyF2) And mnuRename1.Enabled Then
        'Rename group
        Call mnuRename_Click
    End If
    
End Sub

Private Sub tv1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mButton1 = Button
End Sub

Private Sub tv1_NodeClick(ByVal Node As MSComctlLib.Node)
    
    txtNote.Text = ""
    lblNoteTitle.Caption = ""
    imgInfo.Visible = False
    lblGroup.Caption = Node.Text
    
    'Disable buttons
    Toolbar1.Buttons(9).Enabled = False
    Toolbar1.Buttons(10).Enabled = False
    
    'Store Node Index
    mTvID = Node.Index
    
    If (Node.Key = "M_TOP") Then
        'Diplay all notes
        Call LoadNotes(LstV, -1)
        mnuDelete1.Enabled = False
        mnuRename1.Enabled = False
        mnuNewNote1.Enabled = False
        mnuEdit1.Enabled = False
        mnuDelete2.Enabled = False
        mnuMove1.Enabled = False
        Toolbar1.Buttons(6).Enabled = False
        Toolbar1.Buttons(8).Enabled = False
        mnuMove.Enabled = False
    Else
        'Store group name
        mGroupName = Node.Text
        'Enable buttons
        Toolbar1.Buttons(5).Enabled = True
        Toolbar1.Buttons(6).Enabled = True
        Toolbar1.Buttons(8).Enabled = True
        mnuMove.Enabled = True
        'Enable group menu items
        mnuDelete1.Enabled = True
        mnuRename1.Enabled = True
        mnuNewNote1.Enabled = True
        'Disable menu items
        mnuEdit1.Enabled = False
        mnuDelete2.Enabled = False
        mnuMove1.Enabled = False
        'Clear the text box
        txtNote.Text = ""
        'Extract the NodeID
        mGroupID = Val(Mid$(Node.Key, 2))
        'Load the notes for each group
        Call LoadNotes(LstV, mGroupID)
        'Display PopupMenu Menu
        If (mButton1 = vbRightButton) Then
            Call PopupMenu(mnuGroup)
        End If
    End If
    
    mButton1 = vbLeftButton
    lblItems.Caption = LstV.ListItems.Count & " Items(s)"
End Sub

