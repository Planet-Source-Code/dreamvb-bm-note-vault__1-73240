VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAddNote 
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8670
   Icon            =   "frmAddNote.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   8670
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox pImage 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   7395
      ScaleHeight     =   615
      ScaleWidth      =   705
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   525
      Visible         =   0   'False
      Width           =   705
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Height          =   390
      Left            =   4830
      TabIndex        =   18
      Top             =   1650
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_BOLD"
            Object.ToolTipText     =   "Bold"
            ImageIndex      =   6
            Style           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_ITALIC"
            Object.ToolTipText     =   "Italic"
            ImageIndex      =   7
            Style           =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_UNDERLINE"
            Object.ToolTipText     =   "Underline"
            ImageIndex      =   8
            Style           =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_COL"
            Object.ToolTipText     =   "Text Color"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_LEFT"
            Object.ToolTipText     =   "Align Left"
            ImageIndex      =   10
            Style           =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_CENTER"
            Object.ToolTipText     =   "Align Center"
            ImageIndex      =   11
            Style           =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_RIGHT"
            Object.ToolTipText     =   "Align Right"
            ImageIndex      =   12
            Style           =   1
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_BULLET"
            Object.ToolTipText     =   "Bullets"
            ImageIndex      =   13
            Style           =   1
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_IMAGE"
            Object.ToolTipText     =   "Image"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_DATE"
            Object.ToolTipText     =   "Insert Date Time"
            ImageIndex      =   15
         EndProperty
      EndProperty
   End
   Begin NoteVault.Line3D Line3D2 
      Height          =   30
      Left            =   45
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1560
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   53
   End
   Begin VB.ComboBox cboSize 
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
      Left            =   4050
      Style           =   2  'Dropdown List
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Font Size"
      Top             =   1680
      Width           =   720
   End
   Begin VB.ComboBox cboFont 
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
      Left            =   2010
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Font"
      Top             =   1680
      Width           =   1920
   End
   Begin RichTextLib.RichTextBox txtNote 
      Height          =   2595
      Left            =   45
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2085
      Width           =   7950
      _ExtentX        =   14023
      _ExtentY        =   4577
      _Version        =   393217
      ScrollBars      =   2
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmAddNote.frx":058A
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   390
      Left            =   90
      TabIndex        =   13
      Top             =   1650
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_NEW"
            Object.ToolTipText     =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_OPEN"
            Object.ToolTipText     =   "Open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_CUT"
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_CPY"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_PASTE"
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5550
      Top             =   930
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddNote.frx":0601
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddNote.frx":0713
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddNote.frx":0A65
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddNote.frx":0DB7
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddNote.frx":1109
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddNote.frx":145B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddNote.frx":17AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddNote.frx":1AFF
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddNote.frx":1E51
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddNote.frx":21A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddNote.frx":24F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddNote.frx":2847
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddNote.frx":2B99
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddNote.frx":2EEB
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddNote.frx":2FFD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageCombo CboIcon 
      Height          =   330
      Left            =   4035
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Choose icon"
      Top             =   1095
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Locked          =   -1  'True
   End
   Begin VB.ComboBox cboPriority 
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
      Left            =   2325
      Style           =   2  'Dropdown List
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Choose priority"
      Top             =   1095
      Width           =   1605
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   6165
      Top             =   945
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   ". . ."
      Height          =   315
      Left            =   1500
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Choose Color"
      Top             =   1125
      Width           =   420
   End
   Begin VB.PictureBox PColor 
      BackColor       =   &H00000000&
      Height          =   360
      Left            =   150
      ScaleHeight     =   300
      ScaleWidth      =   1740
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Choose Color"
      Top             =   1110
      Width           =   1800
   End
   Begin VB.PictureBox pBottom 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   8670
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4905
      Width           =   8670
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   375
         Left            =   4605
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   105
         Width           =   1035
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Close"
         Height          =   375
         Left            =   5730
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   105
         Width           =   1035
      End
      Begin NoteVault.Line3D Line3D1 
         Height          =   30
         Left            =   0
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   53
      End
   End
   Begin VB.TextBox txtTitle 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   150
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   390
      Width           =   6585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note icon"
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
      Left            =   4050
      TabIndex        =   12
      Top             =   855
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Priority"
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
      Left            =   2325
      TabIndex        =   9
      Top             =   855
      Width           =   495
   End
   Begin VB.Label lblColor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note color:"
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
      Left            =   165
      TabIndex        =   6
      Top             =   855
      Width           =   780
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Note Title:"
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
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   705
   End
End
Attribute VB_Name = "frmAddNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private TxtCls As dTxtHelper

Private Sub ClickButton(ByVal Index As Integer)
Dim mButton As Button
On Error Resume Next
    Set mButton = Toolbar2.Buttons(Index)
    Call Toolbar2_ButtonClick(mButton)
    'Destroy mButton
    Set mButton = Nothing
End Sub

Private Sub EnableButtons()
    Toolbar1.Buttons(4).Enabled = Len(txtNote.SelText)
    Toolbar1.Buttons(5).Enabled = Len(txtNote.SelText)
    Toolbar1.Buttons(6).Enabled = TxtCls.CanPaste
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

Public Function FindInList(Cbo As ComboBox, StrFind As String) As Integer
Dim x As Integer
Dim Idx As Integer
    'Locate an items index in a combobox
    For x = 0 To Cbo.ListCount
        If LCase(StrFind) = LCase(Cbo.List(x)) Then
            Idx = x
            Exit For
        End If
    Next x
    
    FindInList = Idx
End Function

Private Function GetColor() As Long
On Error GoTo CanErr:
    With CD1
        .CancelError = True
        Call .ShowColor
        GetColor = .Color
    End With
    
    Exit Function
CanErr:
    GetColor = -1
End Function

Private Sub cboFont_Click()
On Error Resume Next
    txtNote.SelFontName = cboFont.Text
    txtNote.SetFocus
End Sub

Private Sub cboSize_Click()
On Error Resume Next
    txtNote.SelFontSize = Val(cboSize.Text)
    txtNote.SetFocus
End Sub

Private Sub cmdCancel_Click()
    ButtonPress = vbCancel
    'Unload program
    Call Unload(frmAddNote)
End Sub

Private Sub cmdok_Click()
Dim sTitle As String

    ButtonPress = vbOK
    'Check for note title
    sTitle = Trim(txtTitle.Text)
    If Len(sTitle) = 0 Then
        Call Beep
        txtTitle.BackColor = &HFFDBBF
        Call txtTitle.SetFocus
        Exit Sub
    Else
        'Store note Info
        With mNoteInfo
            .NoteTitle = sTitle
            .NoteColor = PColor.BackColor
            .NoteAdded = Date & " " & Format(Time, "hh:mm")
            .NotePriority = cboPriority.Text
            .NoteIcon = CboIcon.SelectedItem.Index
            .NoteLastModifed = .NoteAdded
            .NoteData = txtNote.TextRTF
        End With
    End If
    
    'Unload program
    Call Unload(frmAddNote)
End Sub

Private Sub cmdOpen_Click()
Dim iCol As Long
    iCol = GetColor
    If (iCol <> -1) Then
        PColor.BackColor = iCol
    End If
End Sub

Private Sub Form_Load()
Dim Idx As Integer
Dim Count As Integer

    'Setup the editor
    Set TxtCls = New dTxtHelper
    TxtCls.SetEditor = txtNote
    'Add fonts to combobox
    For Count = 0 To Screen.FontCount - 1
        Call cboFont.AddItem(Screen.Fonts(Count))
    Next Count
    
    'Add font sizes
    Call cboSize.AddItem("8")
    Call cboSize.AddItem("9")
    Call cboSize.AddItem("10")
    Call cboSize.AddItem("11")
    Call cboSize.AddItem("12")
    Call cboSize.AddItem("14")
    Call cboSize.AddItem("16")
    Call cboSize.AddItem("18")
    Call cboSize.AddItem("20")
    Call cboSize.AddItem("22")
    Call cboSize.AddItem("24")
    Call cboSize.AddItem("26")
    Call cboSize.AddItem("28")
    Call cboSize.AddItem("36")
    Call cboSize.AddItem("48")
    Call cboSize.AddItem("72")
    'Set index
    cboSize.ListIndex = 0
    cboFont.ListIndex = FindInList(cboFont, "arial")

    'Set the combobox ImageList
    Set CboIcon.ImageList = frmmain.LstIcons
    'Add the icons to the list
    For Count = 1 To 24
        CboIcon.ComboItems.Add , , , Count
    Next Count

    'Add note Priority
    cboPriority.AddItem "High"
    cboPriority.AddItem "Above Normal"
    cboPriority.AddItem "Normal"
    cboPriority.AddItem "Below Normal"
    cboPriority.AddItem "Low"
    
    'Check if adding a note or editing
    If (mAddNote) Then
        frmAddNote.Caption = "New Note"
        cboPriority.ListIndex = 0
        'Set icon index
        CboIcon.ComboItems(1).Selected = True
    Else
        frmAddNote.Caption = "Edit " & mNoteInfo.NoteTitle
        'Fill text boxes
        With mNoteInfo
            txtTitle.Text = .NoteTitle
            txtNote.TextRTF = .NoteData
            PColor.BackColor = .NoteColor
            'Find priority value
            Idx = FindInList(cboPriority, .NotePriority)
            If (Idx) Then
                cboPriority.ListIndex = Idx
            Else
                cboPriority.ListIndex = 0
            End If
            'Set icon index
            CboIcon.ComboItems(.NoteIcon).Selected = True
        End With
    End If
    
    Toolbar2.Buttons(6).Value = tbrPressed
    Call EnableButtons
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    'Resize controls
    txtTitle.Width = (frmAddNote.ScaleWidth - txtTitle.Left) - 120
    txtNote.Width = (frmAddNote.ScaleWidth - txtNote.Left) - 60
    txtNote.Height = (frmAddNote.ScaleHeight - pBottom.Height - txtNote.Top) - 120
    Line3D2.Width = (frmAddNote.ScaleWidth - Line3D2.Left) - 60
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAddNote = Nothing
End Sub

Private Sub pBottom_Resize()
    'Resize controls
    Line3D1.Width = pBottom.ScaleWidth
    cmdCancel.Left = (pBottom.ScaleWidth - cmdCancel.Width) - 60
    cmdOK.Left = (cmdCancel.Left - cmdOK.Width) - 120
End Sub

Private Sub PColor_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (Button = vbLeftButton) Then
        Call cmdOpen_Click
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim lFile As String

    Select Case Button.Key
        Case "M_NEW"
            'Start new note content
            If Len(txtNote.Text) Then
                If MsgBox("Do you want to clear the content.", vbYesNo Or vbQuestion, frmAddNote.Caption) = vbYes Then
                    txtNote.Text = ""
                End If
            End If
        Case "M_OPEN"
            'Open file
                lFile = GetDLGName(, , "Rich Text Files(*.rtf)|*.rtf|")
                If Len(lFile) Then
                    Call txtNote.LoadFile(lFile, 0)
                End If
        Case "M_CUT"
            'Cut text
            Call TxtCls.Cut
        Case "M_CPY"
            'Copy text
            Call TxtCls.Copy
            'Enable/disable buttons
            Call EnableButtons
        Case "M_PASTE"
            'Paste text
            Call TxtCls.Paste
    End Select
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim oCol As Long
Dim lFile As String

    Select Case Button.Key
        Case "M_BOLD"
            'Text bold
            txtNote.SelBold = Button.Value
        Case "M_ITALIC"
            'Text italic
            txtNote.SelItalic = Button.Value
        Case "M_UNDERLINE"
            'Text underline
            txtNote.SelUnderline = Button.Value
        Case "M_COL"
            'Do text forecolor
            oCol = GetColor()
            'Check for vailt color
            If (oCol <> -1) Then
                txtNote.SelColor = oCol
                txtNote.SetFocus
            End If
        Case "M_LEFT", "M_CENTER", "M_RIGHT"
            'Do sel alignment
            Toolbar2.Buttons(6).Value = tbrUnpressed
            Toolbar2.Buttons(7).Value = tbrUnpressed
            Toolbar2.Buttons(8).Value = tbrUnpressed
            Button.Value = tbrPressed
            If (Button.Index = 6) Then txtNote.SelAlignment = 0
            If (Button.Index = 7) Then txtNote.SelAlignment = 2
            If (Button.Index = 8) Then txtNote.SelAlignment = 1
        Case "M_BULLET"
            'Bullet point
            txtNote.SelBullet = Button.Value
        Case "M_IMAGE"
            'Image
            lFile = GetDLGName(, "Open Picture", "Bitmap Files(*.bmp)|*.bmp|Gif Files(*.gif)|*.gif|JPEG Files(*.jpg)|*.jpg|")
            If Len(lFile) Then
                pImage.Picture = LoadPicture(lFile)
                Call Clipboard.Clear
                'Copy picture to clipboard
                Call Clipboard.SetData(pImage.Picture)
                'Place image on editor
                Call TxtCls.Paste
                Call Clipboard.Clear
            End If
        Case "M_DATE"
            'insert date time
            txtNote.SelText = Now
    End Select
End Sub

Private Sub txtNote_Click()
    Select Case txtNote.SelAlignment
        Case 0
            Call ClickButton(6)
        Case 1
            Call ClickButton(8)
        Case 2
            Call ClickButton(7)
    End Select
End Sub

Private Sub txtNote_SelChange()
On Error Resume Next
  cboFont.ListIndex = FindInList(cboFont, txtNote.SelFontName)
  cboSize.ListIndex = FindInList(cboSize, txtNote.SelFontSize)
  Toolbar2.Buttons(1).Value = Abs(txtNote.SelBold)
  Toolbar2.Buttons(2).Value = Abs(txtNote.SelItalic)
  Toolbar2.Buttons(3).Value = Abs(txtNote.SelUnderline)
  'Enable/disable buttons
  Call EnableButtons
End Sub

Private Sub txtTitle_Click()
    If (txtTitle.BackColor = &HFFDBBF) Then
        txtTitle.BackColor = vbWhite
    End If
End Sub

Private Sub txtTitle_KeyPress(KeyAscii As Integer)

    If (KeyAscii = vbKeyTab) Then
        KeyAscii = 0
        Call txtNote.SetFocus
    End If
    
    If (KeyAscii = 13) Then
        KeyAscii = 0
        Call cmdok_Click
    End If
End Sub
