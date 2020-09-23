VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Fonts"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   195
      TabIndex        =   8
      Top             =   2490
      Width           =   4530
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   165
         ScaleHeight     =   360
         ScaleWidth      =   3300
         TabIndex        =   14
         Top             =   1215
         Width           =   3330
         Begin VB.Label lblRecordFont 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "#1"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   60
            Width           =   195
         End
      End
      Begin VB.CommandButton cmdFont2 
         Caption         =   "Edit"
         Height          =   390
         Left            =   3570
         TabIndex        =   13
         ToolTipText     =   "Choose Font"
         Top             =   1215
         Width           =   750
      End
      Begin VB.CommandButton cmdFont1 
         Caption         =   "Edit"
         Height          =   390
         Left            =   3570
         TabIndex        =   12
         ToolTipText     =   "Choose Font"
         Top             =   495
         Width           =   750
      End
      Begin VB.PictureBox pFontView 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   165
         ScaleHeight     =   360
         ScaleWidth      =   3300
         TabIndex        =   10
         Top             =   495
         Width           =   3330
         Begin VB.Label lblGroupFont 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "#0"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   60
            Width           =   195
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Record Font"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   210
         TabIndex        =   16
         Top             =   960
         Width           =   885
      End
      Begin VB.Label lblGroup 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group Font"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   210
         TabIndex        =   9
         Top             =   240
         Width           =   810
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   4710
      Top             =   105
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      Height          =   350
      Left            =   2670
      TabIndex        =   6
      Top             =   4485
      Width           =   945
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   350
      Left            =   3780
      TabIndex        =   5
      Top             =   4485
      Width           =   945
   End
   Begin VB.Frame Frame1 
      Caption         =   "General"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2340
      Left            =   195
      TabIndex        =   0
      Top             =   135
      Width           =   4530
      Begin VB.ComboBox cboSkin 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1275
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1770
         Width           =   1785
      End
      Begin VB.CheckBox chkTray 
         Caption         =   "Minimize to tray"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   195
         TabIndex        =   17
         Top             =   1500
         Width           =   2775
      End
      Begin VB.CheckBox chkMaxSize 
         Caption         =   "Open program maximized"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   195
         TabIndex        =   7
         Top             =   1245
         Width           =   2775
      End
      Begin VB.CheckBox chkExit 
         Caption         =   "Allow quick exit using Esc key"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   195
         TabIndex        =   4
         Top             =   960
         Width           =   3690
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   ". . ."
         Height          =   345
         Left            =   3675
         TabIndex        =   3
         ToolTipText     =   "Open"
         Top             =   525
         Width           =   510
      End
      Begin VB.TextBox txtDatabase 
         Height          =   340
         Left            =   195
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   525
         Width           =   3435
      End
      Begin VB.Label lblSkin 
         AutoSize        =   -1  'True
         Caption         =   "Choose Skin:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   270
         TabIndex        =   18
         Top             =   1830
         Width           =   945
      End
      Begin VB.Label lblDatabase 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Database Filename:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   195
         TabIndex        =   1
         Top             =   285
         Width           =   1410
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DoFont(ByVal TheLabel As Label, oFlags As Long, IsGroup As Boolean)
On Error GoTo CanErr:

    With CD1
        .CancelError = True
        'Set flags
        .flags = oFlags
        'Setup font properties
        .FontName = TheLabel.FontName
        .FontSize = TheLabel.FontSize
        .FontBold = TheLabel.FontBold
        .FontItalic = TheLabel.FontItalic
        .FontUnderline = TheLabel.FontUnderline
        .FontStrikethru = TheLabel.FontStrikethru
        'Check if group
        If IsGroup Then
            .Color = TheLabel.ForeColor
        End If
        
        'Show font dialog
        Call .ShowFont
        
        'Set group font strings
        If (IsGroup) Then
            mConfig.GroupFont = .FontName & "," & .FontSize _
            & "," & Abs(.FontBold) & "," & Abs(.FontItalic) & "," & Abs(.FontUnderline) _
            & "," & Abs(.FontStrikethru) & "," & .Color
        Else
            'Set record font strings
            mConfig.RecordFont = .FontName & "," & .FontSize _
            & "," & Abs(.FontBold) & "," & Abs(.FontItalic) & "," & Abs(.FontUnderline) _
            & "," & Abs(.FontStrikethru)
        End If
        
        'Set caption
        TheLabel.Caption = .FontName & "," & .FontSize
        'Set label font
        TheLabel.FontName = .FontName
        TheLabel.FontSize = .FontSize
        TheLabel.FontBold = .FontBold
        TheLabel.FontItalic = .FontItalic
        TheLabel.FontUnderline = .FontUnderline
        TheLabel.FontStrikethru = .FontStrikethru
        'Check if group
        If IsGroup Then
            TheLabel.ForeColor = .Color
        End If
    End With
    Exit Sub
CanErr:
    If (Err.Number = cdlCancel) Then
        Err.Clear
    End If
End Sub

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

Private Sub cboSkin_Click()
    Call frmmain.SetSkin(cboSkin.ListIndex)
End Sub

Private Sub cmdCancel_Click()
    'Exit
    ButtonPress = vbCancel
    Call Unload(frmOptions)
End Sub

Private Sub cmdFont1_Click()
    Call DoFont(lblGroupFont, cdlCFBoth Or cdlCFApply Or cdlCFEffects, True)
End Sub

Private Sub cmdFont2_Click()
    Call DoFont(lblRecordFont, cdlCFBoth Or cdlCFApply, False)
End Sub

Private Sub cmdok_Click()
    'Store config
    If Len(txtDatabase.Text) = 0 Then
        Call Beep
        txtDatabase.BackColor = &HFFDBBF
        Call txtDatabase.SetFocus
        Exit Sub
    Else
        'Store config
        mConfig.DatabaseFile = txtDatabase.Text
        mConfig.QuickExit = chkExit.Value
        mConfig.Maxsized = chkMaxSize.Value
        mConfig.MinToTray = chkTray.Value
        mConfig.Skin = cboSkin.ListIndex
        'Save config
        Call Saveconfig
        'Exit
        ButtonPress = vbOK
        Call Unload(frmOptions)
    End If
End Sub

Private Sub cmdOpen_Click()
Dim lzFilename As String

    lzFilename = GetDLGName(, "Select Database", mFilter)
    
    If Len(lzFilename) Then
        'Set text box properties
        txtDatabase.Text = lzFilename
        txtDatabase.ToolTipText = lzFilename
    End If
    'Clear up
    lzFilename = vbNullString
End Sub

Private Sub Form_Load()
    'Remove icon
    Set frmOptions.Icon = Nothing
    'Set textbox
    txtDatabase.Text = mConfig.DatabaseFile
    txtDatabase.ToolTipText = txtDatabase.Text
    'Set checkbox
    chkExit.Value = mConfig.QuickExit
    'Set mazsized checkbox
    chkMaxSize.Value = mConfig.Maxsized
    'Set tray minsize option
    chkTray.Value = mConfig.MinToTray
    'Add skin names
    Call cboSkin.AddItem("Cool Blue")
    Call cboSkin.AddItem("Windows 2000")
    'Set index
    cboSkin.ListIndex = mConfig.Skin
    'Setup fonts
    Call SetFont(mConfig.GroupFont, lblGroupFont)
    Call SetFont(mConfig.RecordFont, lblRecordFont)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmOptions = Nothing
End Sub

Private Sub txtDatabase_LostFocus()
    If (txtDatabase.BackColor = &HFFDBBF) Then
        txtDatabase.BackColor = vbWhite
    End If
End Sub
