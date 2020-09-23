Attribute VB_Name = "DBase"
Option Explicit

Public db As Database
Public dbFile As String
Public dbOpen As Boolean

Private Type TNoteInfo
    NoteTitle As String
    NoteAdded As String
    NoteLastModifed As String
    NoteColor As Long
    NoteData As String
    NotePriority As String
    NoteIcon As Integer
End Type

Public mNoteInfo As TNoteInfo

Public Function CreateNewDatabase(ByVal FileName As String) As Integer
Dim db1 As Database
Dim td As TableDef
On Error GoTo CreateErr:

    'Create the database
    Set db1 = CreateDatabase(FileName, dbLangGeneral, dbVersion40)

    'Open the database to add new tables
    Set db1 = OpenDatabase(FileName, False, False)
    
    'Create admin table
    Set td = db1.CreateTableDef("admin")
    
    With td
        'Add password field
        .Fields.Append .CreateField("vPassword", dbText, 50)
        .Fields(0).AllowZeroLength = True
    End With
    
    'Append admin table
    Call db1.TableDefs.Append(td)
    
    'Add the tables
    Set td = db1.CreateTableDef("Groups")
    
    'Create Groups table
    With td
        'Add ID field
        .Fields.Append .CreateField("ID", dbLong, 4)
        .Fields(0).Attributes = 49
        'Create vGroupName
        .Fields.Append .CreateField("vGroupName", dbText, 50)
    End With
    
    'Apend Group table
    Call db1.TableDefs.Append(td)
    'Create notes table
    Set td = db1.CreateTableDef("Notes")
    
    With td
        'ID field
        .Fields.Append .CreateField("ID", dbLong, 4)
        .Fields(0).Attributes = 49
        'Note title
        .Fields.Append .CreateField("nNoteTitle", dbText, 50)
        'Note added date
        .Fields.Append .CreateField("nAddDate", dbText, 50)
        'Note last update
        .Fields.Append .CreateField("nLastUpdate", dbText, 50)
        'Note Priority
        .Fields.Append .CreateField("nPriority", dbText, 50)
        'Note Icon
        .Fields.Append .CreateField("nIcon", dbInteger, 2)
        'Note Group ID
        .Fields.Append .CreateField("nGroupID", dbLong, 4)
        'Note forecolor
        .Fields.Append .CreateField("nColor", dbText, 50)
        'Note data
        .Fields.Append .CreateField("nNote", dbMemo, 255)
        .Fields(8).AllowZeroLength = True
    End With
    'Append notes table
    
    Call db1.TableDefs.Append(td)
    CreateNewDatabase = 1
    
    'Clear up
    Set td = Nothing
    Call db1.Close
    
    Exit Function
CreateErr:
    CreateNewDatabase = 0
    'Clear up
    Set td = Nothing
End Function

Public Sub CloseDB()
    'Close database
    If (dbOpen) Then
        dbOpen = False
        Call db.Close
    End If
End Sub

Public Sub OpenDataBaseA()
On Error GoTo OpenErr:
    'Open the database
    Set db = OpenDatabase(dbFile, False, False, ";pwd=33CC66")
    dbOpen = True
    Exit Sub
    'Error flag
OpenErr:
    dbOpen = False
End Sub

Public Sub LoadGroups(TTreeView As TreeView)
Dim rc As Recordset

    With TTreeView
        Call .Nodes.Clear
        Call .Nodes.Add(, tvwFirst, "M_TOP", GetFilename(dbFile), 1, 1)
        'Load store recordset
        Set rc = db.OpenRecordset("Groups")
        
        While (Not rc.EOF)
            'Add Key and Title
            Call .Nodes.Add(1, tvwChild, "k" & rc("ID"), rc("vGroupName"), 2, 3)
            'Get next record
            Call rc.MoveNext
        Wend
    End With
    
    'Clear up
    Call rc.Close
    Set rc = Nothing
End Sub

Public Sub LoadNotes(TListView As ListView, ByVal TheID As Integer)
Dim rc As Recordset
Dim PColor As Long

    With TListView
        Call .ListItems.Clear
        .Sorted = False
        
        If (TheID <> -1) Then
            'List Notes by Index
            Set rc = db.OpenRecordset("SELECT ID,nNoteTitle,nAddDate,nLastUpdate,nPriority,nIcon,nColor,nGroupID FROM Notes WHERE nGroupID=" & TheID)
        Else
            'Get all notes
            Set rc = db.OpenRecordset("Notes")
        End If
        
        While (Not rc.EOF)
            'Add Key and Note Title
            Call .ListItems.Add(, "k" & rc("ID"), rc("nNoteTitle"), Val(rc("nIcon")), Val(rc("nIcon")))
            'Add Creation Time
            .ListItems(.ListItems.Count).SubItems(1) = rc("nAddDate")
            'Add last updated time
            .ListItems(.ListItems.Count).SubItems(2) = rc("nLastUpdate")
            'Add Note Priority
            .ListItems(.ListItems.Count).SubItems(3) = rc("nPriority")
            'Add Forecolor
            .ListItems(.ListItems.Count).ForeColor = Val(rc("nColor"))
            .ListItems(.ListItems.Count).ListSubItems(1).ForeColor = Val(rc("nColor"))
            .ListItems(.ListItems.Count).ListSubItems(2).ForeColor = Val(rc("nColor"))
            .ListItems(.ListItems.Count).ListSubItems(3).ForeColor = Val(rc("nColor"))
            'Add icon as tag
            .ListItems(.ListItems.Count).Tag = rc("nIcon")
            'Add Note Priority
            .ListItems(.ListItems.Count).SubItems(3) = rc("nPriority")
             'Set Priority colors
            Select Case LCase$(.ListItems(.ListItems.Count).SubItems(3))
                Case "high"
                    PColor = vbRed
                Case "above normal"
                    PColor = &H80FF&
                Case "normal"
                    PColor = &H8000&
                Case "below normal"
                    PColor = &H808080
                Case Else
                    PColor = vbBlack
            End Select
            
            .ListItems(.ListItems.Count).ListSubItems(3).ForeColor = PColor
            'Get next record
            Call rc.MoveNext
        Wend
    End With
    
    'Clear up
    Call rc.Close
    Set rc = Nothing
End Sub

Public Function GetNote(ByVal TheID As Long) As String
Dim rc As Recordset
Dim SqlStr As String
On Error Resume Next
    
    'Build SQL string
    SqlStr = "SELECT nNote FROM Notes WHERE ID = " & TheID
    
    'Open record set
    Set rc = db.OpenRecordset(SqlStr)
    
    If (rc.RecordCount) Then
        GetNote = rc("nNote")
    End If
    
    'Clear up
    Call rc.Close
    Set rc = Nothing
    
End Function

Public Function AddNote(ByVal TheID As Integer) As Integer
Dim rc As Recordset
On Error GoTo AddErr:

    'Open the record set
    Set rc = db.OpenRecordset("Notes")
    
    With rc
        'Add new record
        Call .AddNew
        !nNoteTitle = mNoteInfo.NoteTitle
        !nGroupID = TheID
        !nPriority = mNoteInfo.NotePriority
        !nIcon = mNoteInfo.NoteIcon
        !nAddDate = mNoteInfo.NoteAdded
        !nLastUpdate = mNoteInfo.NoteLastModifed
        !nColor = mNoteInfo.NoteColor
        !nNote = mNoteInfo.NoteData
        Call .Update
    End With
    
    AddNote = 1
    
    Call rc.Close
    Set rc = Nothing
    
    Exit Function
    'Error flag
AddErr:
MsgBox "s"
    'Clear up
    Call rc.Close
    Set rc = Nothing
End Function

Public Function EditNote(ByVal TheID As Integer) As Integer
Dim rc As Recordset
On Error GoTo EditErr:

    'Open the record set
    Set rc = db.OpenRecordset("SELECT * FROM Notes WHERE ID =" & TheID)

    With rc
        'Add new record
        Call .Edit
        !nNoteTitle = mNoteInfo.NoteTitle
        !nPriority = mNoteInfo.NotePriority
        !nIcon = mNoteInfo.NoteIcon
        !nLastUpdate = mNoteInfo.NoteAdded
        !nColor = mNoteInfo.NoteColor
        !nNote = mNoteInfo.NoteData
        Call .Update
    End With
    
    EditNote = 1
    
    Call rc.Close
    Set rc = Nothing
    
    Exit Function
    'Error flag
EditErr:
    'Clear up
    EditNote = 0
    Call rc.Close
    Set rc = Nothing
End Function

Public Function MoveNote(ByVal TheID As Integer, MoveToID As Integer) As Integer
Dim rc As Recordset
On Error GoTo MoveErr:

    'Open the record set
    Set rc = db.OpenRecordset("SELECT * FROM Notes WHERE ID =" & TheID)

    With rc
        'Add new record
        Call .Edit
        !nGroupID = MoveToID
        Call .Update
    End With
    
    MoveNote = 1
    
    Call rc.Close
    Set rc = Nothing
    
    Exit Function
    'Error flag
MoveErr:
    'Clear up
    MoveNote = 0
    Call rc.Close
    Set rc = Nothing
End Function

Public Function DeleteNote(ByVal TheID As Integer) As Integer
Dim rc As Recordset
On Error GoTo DelErr:

    'Open record set
    Set rc = db.OpenRecordset("SELECT * FROM Notes WHERE ID = " & TheID)
    
    If (rc.RecordCount) Then
        'Delete the record
        rc.Delete
    End If
    
    DeleteNote = 1
    'Clear up
    Call rc.Close
    Set rc = Nothing
    
    Exit Function
    'Error flag
DelErr:
    DeleteNote = 0
    'Clear up
    Call rc.Close
    Set rc = Nothing
End Function

Public Function RenameGroup(ByVal NewName As String, ByVal TheID As Integer) As Integer
On Error GoTo EditErr:
Dim rc As Recordset
    
    'Open record set
    Set rc = db.OpenRecordset("SELECT ID,vGroupName FROM Groups WHERE ID = " & TheID)
    'Check if record was found
    If (rc.RecordCount) Then
        With rc
            'Edit the record
            Call .Edit
            !vGroupName = NewName
            Call .Update
        End With
    End If
    
    RenameGroup = 1
    
    'Clear up
    Call rc.Close
    Set rc = Nothing
    
    Exit Function
EditErr:
    RenameGroup = 0
    'Clear up
    Call rc.Close
    Set rc = Nothing
End Function

Public Function AddGroup(ByVal NewName As String) As Integer
On Error GoTo AddErr:
Dim rc As Recordset
    
    'Open record set
    Set rc = db.OpenRecordset("Groups")
    With rc
        'Edit the record
        Call .AddNew
        !vGroupName = NewName
        Call .Update
    End With
    
    AddGroup = 1
    
    'Clear up
    Call rc.Close
    Set rc = Nothing
    
    Exit Function
AddErr:
    AddGroup = 0
    'Clear up
    Call rc.Close
    Set rc = Nothing
End Function

Public Function DeleteGroup(ByVal TheID As Integer) As Integer
On Error GoTo DelErr:
Dim rc As Recordset
    
    'Open record set
    Set rc = db.OpenRecordset("SELECT * FROM Notes WHERE nGroupID =" & TheID)
    
    'Delete the notes first
    While (Not rc.EOF)
        'Delete Record
        Call rc.Delete
        'Get next reocrd
        Call rc.MoveNext
    Wend
    
    'Delete the group
    Set rc = db.OpenRecordset("SELECT ID,vGroupName FROM Groups WHERE ID=" & TheID)
    
    If (rc.RecordCount) Then
        'Delete the record
        Call rc.Delete
    End If
    
    DeleteGroup = 1
    'Clear up
    Call rc.Close
    Set rc = Nothing
    
    Exit Function
DelErr:
    DeleteGroup = 0
    'Clear up
    Call rc.Close
    Set rc = Nothing
End Function

Public Function GetPassword() As String
Dim rc As Recordset
    
    'Open admin record set
    Set rc = db.OpenRecordset("admin")
    
    If (rc.RecordCount) Then
        GetPassword = rc("vPassword")
    End If
    
    'Clear up
    Call rc.Close
    Set rc = Nothing
End Function

Public Sub SetPassword(ByVal ThePassword As String)
Dim rc As Recordset
    
    'Open recored set
    Set rc = db.OpenRecordset("admin")

    With rc
        If (rc.RecordCount) Then
            Call rc.Delete
            Call .AddNew
            !vPassword = ThePassword
            Call .Update
        Else
            Call .AddNew
            !vPassword = ThePassword
            Call .Update
        End If
    End With
    
    'Clear up
    Call rc.Close
    Set rc = Nothing
End Sub
