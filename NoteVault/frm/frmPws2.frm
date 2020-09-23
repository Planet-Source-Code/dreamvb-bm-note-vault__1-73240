VERSION 5.00
Begin VB.Form frmPws2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Password"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   350
      Left            =   3615
      TabIndex        =   8
      Top             =   1590
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   350
      Left            =   3615
      TabIndex        =   7
      Top             =   1125
      Width           =   1215
   End
   Begin VB.TextBox txtPass2 
      Height          =   350
      IMEMode         =   3  'DISABLE
      Left            =   300
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   2655
      Width           =   2415
   End
   Begin VB.TextBox txtPass1 
      Height          =   350
      IMEMode         =   3  'DISABLE
      Left            =   300
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1920
      Width           =   2415
   End
   Begin VB.TextBox txtOld 
      Height          =   350
      IMEMode         =   3  'DISABLE
      Left            =   300
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1140
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   150
      Picture         =   "frmPws2.frx":0000
      Top             =   240
      Width           =   315
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password:"
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
      Left            =   300
      TabIndex        =   5
      Top             =   2400
      Width           =   1395
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Password:"
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
      Left            =   300
      TabIndex        =   3
      Top             =   1665
      Width           =   1185
   End
   Begin VB.Label lblTitle2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Old Password:"
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
      Left            =   300
      TabIndex        =   1
      Top             =   885
      Width           =   1080
   End
   Begin VB.Label lblTitle1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter a password that will be used to open the database"
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
      Left            =   555
      TabIndex        =   0
      Top             =   360
      Width           =   4110
   End
End
Attribute VB_Name = "frmPws2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    'Unload this form
    ButtonPress = vbCancel
    Call Unload(frmPws2)
End Sub

Private Sub cmdok_Click()
Dim mPass1 As String
Dim mPass2 As String
Dim mPass3 As String
    
    mPass1 = txtOld.Text
    mPass2 = txtPass1.Text
    mPass3 = txtPass2.Text
    
    If (mPass1 <> Pass1) Then
        Call MsgBox("The password is incorrect. Please try again.", vbExclamation, frmPws2.Caption)
        txtOld.SetFocus
        Exit Sub
    End If
    
    'Check if the passwords match
    If (mPass2 <> mPass3) Then
        Call MsgBox("The passwords do not match Pease try again.", vbExclamation, frmPws2.Caption)
        Exit Sub
    Else
        Pass2 = mPass2
        If Len(Pass2) = 0 Then Pass2 = ""
        ButtonPress = vbOK
        Call Unload(frmPws2)
    End If
    
End Sub

Private Sub Form_Load()
    'Remove icon
    Set frmPws2.Icon = Nothing
    
    'Enable / Disable textbox
    txtOld.Enabled = Len(GetPassword)
    'Set Textbox backcolor
    If (Not txtOld.Enabled) Then
        txtOld.BackColor = vbButtonFace
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPws2 = Nothing
End Sub

Private Sub txtOld_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        KeyAscii = 0
        Call cmdok_Click
    End If
End Sub

Private Sub txtPass1_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        KeyAscii = 0
        Call cmdok_Click
    End If
End Sub

Private Sub txtPass2_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        KeyAscii = 0
        Call cmdok_Click
    End If
End Sub
