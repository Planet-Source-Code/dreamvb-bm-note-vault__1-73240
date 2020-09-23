VERSION 5.00
Begin VB.Form frmmove 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Move"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4125
   Icon            =   "frmmove.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   350
      Left            =   3150
      TabIndex        =   3
      Top             =   630
      Width           =   870
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   350
      Left            =   3150
      TabIndex        =   2
      Top             =   1080
      Width           =   870
   End
   Begin VB.ComboBox cboMove 
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
      Left            =   540
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   660
      Width           =   2505
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   135
      Picture         =   "frmmove.frx":058A
      Top             =   30
      Width           =   480
   End
   Begin VB.Label lblMove 
      AutoSize        =   -1  'True
      Caption         =   "Move to:"
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
      Left            =   540
      TabIndex        =   1
      Top             =   420
      Width           =   615
   End
End
Attribute VB_Name = "frmmove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboMove_Click()
    'Store the group index
    mGroupIndex = cboMove.ItemData(cboMove.ListIndex)
End Sub

Private Sub cmdCancel_Click()
    ButtonPress = vbCancel
    Call Unload(frmmove)
End Sub

Private Sub cmdok_Click()
    ButtonPress = vbOK
    Call Unload(frmmove)
End Sub

Private Sub Form_Load()
Dim nNode As Node
    
    For Each nNode In frmmain.tv1.Nodes
        If (nNode.Key <> "M_TOP") And LCase(nNode.Text) <> LCase(mGroupName) Then
            'Add group title
            Call cboMove.AddItem(nNode.Text)
            'Add group ID
            cboMove.ItemData(cboMove.ListCount - 1) = Val(Mid$(nNode.Key, 2))
        End If
    Next nNode
    
    'Enable / Disable button
    cmdOK.Enabled = cboMove.ListCount
    'Select first index if any groups are found
    If (cboMove.ListCount) Then
        cboMove.ListIndex = 0
    End If
    
    'Destroy node
    Set nNode = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmmove = Nothing
End Sub
