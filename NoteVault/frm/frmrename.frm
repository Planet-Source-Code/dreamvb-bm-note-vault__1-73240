VERSION 5.00
Begin VB.Form frmrename 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
      Height          =   350
      Left            =   2760
      TabIndex        =   3
      Top             =   825
      Width           =   1215
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Height          =   350
      Left            =   1395
      TabIndex        =   2
      Top             =   825
      Width           =   1215
   End
   Begin VB.TextBox txtname 
      Height          =   285
      Left            =   1140
      MaxLength       =   30
      TabIndex        =   1
      Top             =   255
      Width           =   2805
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "New name:"
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
      Left            =   225
      TabIndex        =   0
      Top             =   270
      Width           =   825
   End
End
Attribute VB_Name = "frmrename"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    'Unload this form
    ButtonPress = vbCancel
    Call Unload(frmrename)
End Sub

Private Sub cmdok_Click()
    
    If Len(Trim(txtname.Text)) = 0 Then
        Call Beep
        txtname.BackColor = &HFFDBBF
        Call txtname.SetFocus
        Exit Sub
    Else
        'Store group name
        mGroupName = Trim(txtname.Text)
        'Unload form
        ButtonPress = vbOK
        Call Unload(frmrename)
    End If
    
End Sub

Private Sub Form_Load()
    Set frmrename.Icon = Nothing

    'Check if adding or editing group
    If (mGroupAdd) Then
        frmrename.Caption = "New group"
    Else
        txtname.Text = mGroupName
        frmrename.Caption = "Rename group"
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmrename = Nothing
End Sub

Private Sub txtname_Click()
    If (txtname.BackColor = &HFFDBBF) Then
        txtname.BackColor = vbWhite
    End If
End Sub

Private Sub txtname_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Call cmdok_Click
        KeyAscii = 0
    End If
End Sub
