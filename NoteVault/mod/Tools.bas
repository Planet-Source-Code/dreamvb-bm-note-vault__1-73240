Attribute VB_Name = "Tools"
Option Explicit

Private Type Cfg
    DatabaseFile As String
    QuickExit As Integer
    Maxsized As Integer
    MinToTray As Integer
    Toolbar As Integer
    View As Integer
    NoteInfo As Integer
    StatusBar As Integer
    NotesBar As Integer
    GroupFont As String
    RecordFont As String
    Skin As Integer
End Type

Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type HD_ITEM
   mask As Long
   cxy As Long
   pszText As String
   hbm As Long
   cchTextMax As Long
   fmt As Long
   lParam As Long
   iImage As Long
   iOrder As Long
End Type

Private Type MENUINFO
   cbSize As Long
   fMask As Long
   dwStyle As Long
   cyMax As Long
   hbrBack As Long
   dwContextHelpID As Long
   dwMenuData As Long
End Type

Public Declare Function InitCommonControls Lib "Comctl32.dll" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As Rect, ByVal hBrush As Long) As Long
Public Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CreatePatternBrush Lib "gdi32.dll" (ByVal hBitmap As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Public Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Long
Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As Rect, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SetMenuInfo Lib "user32.dll" (ByVal hmenu As Long, ByRef LPCMENUINFO As MENUINFO) As Long
Private Declare Function DrawMenuBar Lib "user32.dll" (ByVal Hwnd As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long
Private Declare Function GetMenu Lib "user32.dll" (ByVal Hwnd As Long) As Long

Private Const LVM_FIRST As Long = &H1000
Private Const HDM_FIRST As Long = &H1200

Private Const LVM_GETHEADER As Long = (LVM_FIRST + 31)
Private Const HDI_IMAGE As Long = &H20
Private Const HDI_FORMAT As Long = &H4
Private Const HDF_STRING As Long = &H4000
Private Const HDF_IMAGE As Long = &H800
Private Const HDF_BITMAP_ON_RIGHT As Long = &H1000
Private Const HDM_SETITEMA As Long = (HDM_FIRST + 4)
'Menu consts
Private Const MIM_APPLYTOSUBMENUS As Long = &H80000000
Private Const MIM_BACKGROUND As Long = &H2
'Database filter const
Public Const mFilter = "Note Vault Files(*.nvf)|*.nvf|"

Public mAddNote As Boolean
Public mGroupAdd As Boolean
Public mConfig As Cfg
Public ButtonPress As VbMsgBoxResult
Public mGroupName As String
Public mGroupIndex As Integer
Public hBrush1 As Long
Public hBrush2 As Long
Public SkinType As Integer
'Database password variables
Public Pass1 As String
Public Pass2 As String

Public Function IsAppOpen() As Boolean
    'Check if app is ready open
    If App.PrevInstance Then
        IsAppOpen = True
    Else
        IsAppOpen = False
    End If
End Function

Public Sub OpenUrl(ByVal URL As String)
Dim Ret As Long
    Ret = ShellExecute(frmmain.Hwnd, "open", URL, vbNullString, vbNullString, 1)
End Sub

Public Function FixPath(ByVal lPath As String) As String
    If Right$(lPath, 1) = "\" Then
        FixPath = lPath
    Else
        FixPath = lPath & "\"
    End If
End Function

Public Function FindFile(ByVal lzFilename As String) As Boolean
On Error Resume Next
    'Returns true if file if found
    If Len(lzFilename) = 0 Then Exit Function
    FindFile = (GetAttr(lzFilename) And vbNormal) = vbNormal
    Err.Clear
End Function

Public Function GetPathFormFile(ByVal FileName As String) As String
Dim sPos As Integer
    sPos = InStrRev(FileName, "\", Len(FileName), vbBinaryCompare)
    
    If (sPos > 0) Then
        GetPathFormFile = Left$(FileName, sPos)
    Else
        GetPathFormFile = FileName
    End If
    
End Function

Public Function GetFilename(ByVal FileName As String) As String
Dim sPos As Integer
    sPos = InStrRev(FileName, "\", Len(FileName), vbBinaryCompare)
    
    If (sPos > 0) Then
        GetFilename = Mid$(FileName, sPos + 1)
    End If
    
End Function

Public Sub TextPBox(pBox As PictureBox)
Dim mRect As Rect
Dim Ret As Long
Dim LineColor As Long

    If SkinType = 0 Then
        LineColor = &HCB9E7B
    Else
        LineColor = &HB6BDC1
    End If
    
    With pBox
        Call .Cls
        Ret = SetRect(mRect, 0, 0, .ScaleWidth, .ScaleHeight)
        Ret = FillRect(.hdc, mRect, hBrush1)
        'Draw line
        pBox.Line (0, 0)-(.ScaleWidth - 1, .ScaleHeight - 1), LineColor, B
        Call .Refresh
    End With
    
End Sub

Public Sub ShowHeaderIcon(LstV As ListView, colNo As Long, IconIdx As Long, ShowColIcon As Boolean)
Dim hHeader As Long
Dim Retval As Long
Dim LstHd As HD_ITEM
   
   hHeader = SendMessage(LstV.Hwnd, LVM_GETHEADER, 0&, ByVal 0&)
   
   With LstHd
      .mask = HDI_IMAGE Or HDI_FORMAT
      .pszText = LstV.ColumnHeaders(colNo + 1).Text
      
       If ShowColIcon Then
         .fmt = HDF_STRING Or HDF_IMAGE Or HDF_BITMAP_ON_RIGHT
         .iImage = IconIdx
       Else
         .fmt = HDF_STRING
      End If
      
   End With
   
   Retval = SendMessage(hHeader, HDM_SETITEMA, colNo, LstHd)
   
End Sub

Public Function SetMenuColour(ByVal hwndfrm As Long, _
                               ByVal dwColour As Long, _
                               ByVal bIncludeSubmenus As Boolean) As Boolean

  'set application menu colour
   Dim mi As MENUINFO
   Dim flags As Long
   Dim clrref As Long
   Dim Ret As Long
   
  'convert a Windows colour (OLE colour)
  'to a valid RGB colour if required
   clrref = TranslateOLEtoRBG(dwColour)
   
  'we're changing the background,
  'so at a minimum set this flag
   flags = MIM_BACKGROUND
   
   If bIncludeSubmenus Then
     'MIM_BACKGROUND only changes
     'the back colour of the main
     'menu bar, unless this flag is set
      flags = flags Or MIM_APPLYTOSUBMENUS
   End If

  'fill in struct, assign to menu,
  'and force a redraw with the
  'new attributes
   With mi
        .cbSize = Len(mi)
        .fMask = flags
        If (dwColour <> 0) Then
            .hbrBack = CreateSolidBrush(dwColour)
        Else
            .hbrBack = hBrush2
        End If
      
   End With

   Ret = SetMenuInfo(GetMenu(hwndfrm), mi)
   Ret = DrawMenuBar(hwndfrm)

End Function

Public Sub SetFont(ByVal FontStr As String, TheObj As Object)
Dim vFontInfo As Variant
On Error Resume Next

    'Split up the font info
    vFontInfo = Split(FontStr, ",")
    
    'Set object caption
    If TypeName(TheObj) = "Label" Then
        TheObj.Caption = vFontInfo(0) & "," & vFontInfo(1)
        'Set label foecolor
        If UBound(vFontInfo) = 6 Then
            TheObj.ForeColor = Val(vFontInfo(6))
        End If
    End If

    With TheObj.Font
        .Name = vFontInfo(0)
        .Size = Val(vFontInfo(1))
        .Bold = Val(vFontInfo(2))
        .Italic = Val(vFontInfo(3))
        .Underline = Val(vFontInfo(4))
        .Strikethrough = Val(vFontInfo(5))
    End With
    
    If TypeName(TheObj) = "TreeView" Then
        If UBound(vFontInfo) = 6 Then
            Call frmmain.SetNodeColor(Val(vFontInfo(6)))
        Else
            Call frmmain.SetNodeColor(vbBlack)
        End If
    End If
    
    'Clear up
    Erase vFontInfo
End Sub

Private Function TranslateOLEtoRBG(ByVal dwOleColour As Long) As Long
Dim Ret As Long
   Ret = OleTranslateColor(dwOleColour, 0, TranslateOLEtoRBG)
End Function

Public Sub LoadConfig()
    'Load config
    mConfig.DatabaseFile = GetSetting("NoteVault", "Config", "Database", FixPath(App.Path) & "notes.nvf")
    mConfig.QuickExit = Val(GetSetting("NoteVault", "Config", "QuickExit", "0"))
    mConfig.Maxsized = Val(GetSetting("NoteVault", "Config", "Maxsized", "0"))
    mConfig.MinToTray = Val(GetSetting("NoteVault", "Config", "MinToTray", "1"))
    mConfig.Toolbar = Val(GetSetting("NoteVault", "Config", "Toolbar", "0"))
    mConfig.View = Val(GetSetting("NoteVault", "Config", "ListView", "3"))
    mConfig.NoteInfo = Val(GetSetting("NoteVault", "Config", "NoteInfo", "0"))
    mConfig.StatusBar = Val(GetSetting("NoteVault", "Config", "StatusBar", "0"))
    mConfig.NotesBar = Val(GetSetting("NoteVault", "Config", "NotesBar", "0"))
    mConfig.GroupFont = GetSetting("NoteVault", "Config", "GroupFont", "Arial,8,0,0,0,0,0")
    mConfig.RecordFont = GetSetting("NoteVault", "Config", "RecordFont", "Arial,8,0,0,0,0")
    mConfig.Skin = GetSetting("NoteVault", "Config", "Skin", "0")
End Sub

Public Sub Saveconfig()
    'Save config
    Call SaveSetting("NoteVault", "Config", "Database", mConfig.DatabaseFile)
    Call SaveSetting("NoteVault", "Config", "QuickExit", mConfig.QuickExit)
    Call SaveSetting("NoteVault", "Config", "MinToTray", mConfig.MinToTray)
    Call SaveSetting("NoteVault", "Config", "Maxsized", mConfig.Maxsized)
    Call SaveSetting("NoteVault", "Config", "GroupFont", mConfig.GroupFont)
    Call SaveSetting("NoteVault", "Config", "RecordFont", mConfig.RecordFont)
    Call SaveSetting("NoteVault", "Config", "Skin", mConfig.Skin)
End Sub
