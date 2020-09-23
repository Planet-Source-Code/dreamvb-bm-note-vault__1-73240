VERSION 5.00
Begin VB.Form frmPws1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Password"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   350
      Left            =   3915
      TabIndex        =   3
      Top             =   1095
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   350
      Left            =   3915
      TabIndex        =   2
      Top             =   630
      Width           =   1215
   End
   Begin VB.TextBox txtPws 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   330
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   675
      Width           =   3330
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   195
      Picture         =   "frmPws1.frx":0000
      Top             =   150
      Width           =   315
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter the password to open the database."
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
      Left            =   615
      TabIndex        =   0
      Top             =   270
      Width           =   3585
   End
End
Attribute VB_Name = "frmPws1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    'Unload this form
    Call Unload(frmPws1)
End Sub

Private Sub cmdok_Click()
    
    Pass2 = Trim(txtPws.Text)
    'Check if password matchs
    If (Pass1 <> Pass2) Then
        txtPws.BackColor = &HFFDBBF
        txtPws.SelStart = 0
        txtPws.SelLength = Len(txtPws.Text)
        Call txtPws.SetFocus
        Call MsgBox("The password is incorrect. Please try again.", vbExclamation, "BM Note Vault")
    Else
        Call Unload(frmPws1)
    End If
    
End Sub

Private Sub Form_Load()
    'Remove icon
    Set frmPws1.Icon = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPws1 = Nothing
End Sub

Private Sub txtPws_Click()
    If (txtPws.BackColor = &HFFDBBF) Then
        txtPws.BackColor = vbWhite
    End If
End Sub

Private Sub txtPws_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Call cmdok_Click
        KeyAscii = 0
    End If
End Sub
