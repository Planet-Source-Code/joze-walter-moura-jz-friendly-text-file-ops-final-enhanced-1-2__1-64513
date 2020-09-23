VERSION 5.00
Begin VB.UserControl JzTxFile 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   1005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1545
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   1005
   ScaleWidth      =   1545
   ToolboxBitmap   =   "JzTxFile.ctx":0000
   Begin VB.Image JzImage 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   465
      Left            =   0
      Picture         =   "JzTxFile.ctx":0312
      Top             =   0
      Width           =   405
   End
End
Attribute VB_Name = "JzTxFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Friendly Text Files I-O Operations v1.0.00 ® Joze 2006/03/01."
Option Explicit
'.-------------------------------------------------------------
' JzTxFile - A Friendly type Text Files I-O Operations OCX
'`=============================================================
' Version:- 1.2.00 2006/03/04 By JOZE (from Rio de Janeiro, BR)
'
' Text Files are as Random Files: Lines can be accessed directly
' by a 'Line Number', any ordering, toward, forward, etc.
'
' No APIs, No Classes - it's a simple Activex Control to be
' dragged on the Form, with any number of instances.
'
' Not so fast code, was made to abrange inclusive Large text files.
'
' The process consists:
'
' a) Retrieve all line records (from a Original File - or from a
'    previous backup made) into a dynamic array, the WorkBench;
'    Its essential provide a significative '.FileName' content.
'    Do it using .OpenFile or .ReadFromBackup (needs previous
'    .MakeABackup) respective commands.
'
' b) So, lines will be from 1 to '.TotalLines', a all times
'    updated property. The WorkBench will be scenario for GetLine
'    String Function and PutLine, AddLine, InsertLine, RemoveLine
'    commands.
'
' c) The actual WorkBench can be any time:
'
'    - .UpDate - rewrites it over Original File.
'    - .MakeABackup - creates new same filename in other folder as
'                     .BakFileName contents (if <empty> will be
'                     auto filled from .FileName in a "/Bak" sub
'                     folder)
'    - .SaveAsNewFile - creates a new file where .NewFileName (If
'                     <empty> Will Be auto filled from .FileName
'                     in a "/New" subfolder.
'
' All Operations, when completed, will cause a 'OPComplete(Msg)'
' event. If a empty message, Operation was successfull. If a string,
' means a failled operation, wich message became with the Op Name
' and problem related. You can intercept mode administrate program
' actions about it (silent, message box, etc) or merely checks True/
' False states returning all commands.
'
' See 'JzTxFile.Txt' for a programatic Guide or see code to aply.
'
'`--------------------------------------------------------------
' Comments: I wrote this code to be as simple as possible to made
'           it comprehensive and easy to anything customize. So,
'           some redundances will be found, ok?
'
' The '.TotalLines' and 'OpenFile' names are for remember the
' excellent 1997 FSText.OCX control, freeware not openned source,
' that does similar functions.
'
' I'D LIKE RECEIVE COMMENTS OR CODE MODIFICATIONS YOU MADE, OK?
'
'
' Joze.
' jozew@globo.com
'
'
'.-------------------------------------------------------------
' JzTxFile v1.2.00 (2006/03/03) by JOZE jozew@globo.com
'==============================================================
' FIXED: <no fixes>
'
' ENHANCED:
'   - All Sub commands were rewrote as Boolean Functions for tests.
'   - 2 new commands: .OpenHole and .CloseHole to speed Insertion of
'     a group of sucessive lines after certain line in WorkBench.
'   - 3 new functions: .Exists, .BakExists and .NewExists, returns
'     the Original, the Backup and the Save As New file existences.
'   - Some code optimizations and more prevent error messages.
'
' KNOW ERRORS: <no know errors>
'
'`-------------------------------------------------------------
'
'.-------------------------------------------------------------
' JzTxFile v1.0.00 (2006/03/01) by JOZE jozew@globo.com
'==============================================================
' FIXED: <no fixes 1st version>
'
' ENHANCED: <no enhances 1st version>
'
' KNOW ERRORS: <no know errors at 1st version>
'
'`-------------------------------------------------------------
'
'Default Property Values:
Const m_def_FileSpec = vbNullString
Const m_def_TotalLines = 0
Const m_def_BakSpec = vbNullString
Const m_def_NewSpec = vbNullString
'Property Variables:
Dim m_FileSpec As String
Dim m_TotalLines As Long
Dim m_BakSpec As String
Dim m_NewSpec As String

'Path Split Work
Private BtxSpec As String    ' Complete Path
Private BtxPath As String    ' Isolated Path
Private BtxFNam As String    ' Basic filename

'THIS IS THE WorkBench
Private Btx() As String

'This is the HoleBench
Private Htx() As String

'Event Declarations:
Event OpCompletes(Message As String)
Attribute OpCompletes.VB_Description = "Occurs when any Text Operation is finished. 'Message' argument will return <Empty> =Sucessfull or <anysize> =message error."


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,VbNullString
Public Property Get FileSpec() As String
Attribute FileSpec.VB_Description = "Essential Original Input File complete pathfilename."
    FileSpec = m_FileSpec
End Property

Public Property Let FileSpec(ByVal New_FileSpec As String)
    m_FileSpec = New_FileSpec
    PropertyChanged "FileSpec"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get TotalLines() As Long
Attribute TotalLines.VB_Description = "All time maximum number of lines in workbench indicator."
    TotalLines = m_TotalLines
End Property

Public Property Let TotalLines(ByVal New_TotalLines As Long)
    m_TotalLines = New_TotalLines
    PropertyChanged "TotalLines"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,VbNullString
Public Property Get BakFileSpec() As String
Attribute BakFileSpec.VB_Description = "Path-filename to be used for backup purposes (see MakeABackup and ReadFromBackup methods). Default will compose by using the FileSpec string, same filename and 'path/Bak'."
    BakFileSpec = m_BakSpec
End Property

Public Property Let BakFileSpec(ByVal New_BakSpec As String)
    m_BakSpec = New_BakSpec
    PropertyChanged "BakFileSpec"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,VbNullString
Public Property Get NewFileSpec() As String
Attribute NewFileSpec.VB_Description = "Path-filename to be used for Output New File."
    NewFileSpec = m_NewSpec
End Property

Public Property Let NewFileSpec(ByVal New_NewSpec As String)
    m_NewSpec = New_NewSpec
    PropertyChanged "NewFileSpec"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function OpenFile() As Boolean
Attribute OpenFile.VB_Description = "Open 'FileSpec' as Input and read all lines to workbench. "
    If Exists Then
       OpenFile = LoadAFile("Open File", m_FileSpec)
    Else
       FinalizeOp "Open File: not found!"
       OpenFile = False
    End If
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13
Public Function GetLine(LineNumber As Long) As String
Attribute GetLine.VB_Description = "Returns single line string indexed by 'LineNumber' order in workbench."
    If LineNumber = 0 Or _
       LineNumber > UBound(Btx) Then
        GetLine = vbNullString
        FinalizeOp "GetLine: Out of Range"
    Else
        GetLine = Btx(LineNumber)
        FinalizeOp
    End If
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function PutLine(LineNumber As Long, txtLine As String) As Boolean
Attribute PutLine.VB_Description = "Puts the 'txtLine' over the 'LineNumber' indexed in workbench."
    If LineNumber = 0 Or _
        LineNumber > UBound(Btx) Then
       FinalizeOp "PutLine: Out of Range"
       PutLine = False
    Else
       Btx(LineNumber) = txtLine
       FinalizeOp
       PutLine = True
    End If
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function AddLine(txtLine As String) As Boolean
Attribute AddLine.VB_Description = "Appends the 'txtLine' to workbench (after end)."
    ReDim Preserve Btx(UBound(Btx) + 1) As String
    Btx(UBound(Btx)) = txtLine
    ShowTotal
    FinalizeOp
    AddLine = True
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14

Public Function RemoveLine(LineNumber As Long, Optional ToLine As Long = 0) As Boolean
Attribute RemoveLine.VB_Description = "Delete the 'LineNumber' from workbench. Optional 'ToLine' if present, will remove the range from ... to."
    Dim Arr() As String
    Dim L As Long
    Dim M As Long
    Dim unt As Long
    Dim rst As Long
    If (LineNumber = 0 Or _
        LineNumber > UBound(Btx)) Or _
        (ToLine <> 0 And _
         (ToLine < LineNumber Or ToLine > UBound(Btx))) Then
        FinalizeOp "RemoveLine: Invalid or Out of Range"
        RemoveLine = False
        Exit Function
    End If
    unt = ToLine
    If unt = 0 Then
       unt = LineNumber
    End If
    rst = UBound(Btx) - unt    ' how many lines after to exclude region
    If rst = 0 Then    ' to exclude region abranges last line
       L = LineNumber - 1
       ReDim Preserve Btx(L) As String
       FinalizeOp
       ShowTotal
       RemoveLine = True
       Exit Function
    End If
    'Irgh! Let's do hard work!
    ReDim Arr(rst) As String    'collect remaining lines
    For L = 1 To rst
        unt = unt + 1
        Arr(L) = Btx(unt)
    Next L
    L = LineNumber - 1
    ReDim Preserve Btx(L + rst) As String    ' re-cut WorkBench
    L = LineNumber
    unt = 0
    Do While L <= UBound(Btx)
        unt = unt + 1
        Btx(L) = Arr(unt)
        L = L + 1
    Loop
    FinalizeOp
    ShowTotal
    RemoveLine = True
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function InsertLine(LineNumber As Long, txtLine As String) As Boolean
Attribute InsertLine.VB_Description = "Inserts 'txtLine' AFTER 'LineNumber' of workbench (zero value will insert at begin). "
    Dim s As Long
    Dim L As Long
    If (LineNumber = 0 Or _
        LineNumber > UBound(Btx)) Then
        FinalizeOp "InsertLine: Out of Range"
        InsertLine = False
        Exit Function
    End If
    'All we need is one more line
    ReDim Preserve Btx(UBound(Btx) + 1) As String
    ' shift end pos inserting lines
    L = UBound(Btx) - 1
    Do While L > LineNumber
        s = Btx(L) 'two steps for speed purposes
        Btx(L + 1) = s
        L = L - 1
    Loop
    'now L is equal to LineNumber
    Btx(L + 1) = txtLine    'after LineNumber ...
    'ok
    FinalizeOp
    ShowTotal
    InsertLine = True
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function MakeABackup() As Boolean
Attribute MakeABackup.VB_Description = "Will save a backup of  workbench to file as indicated in BakFileSpec path (see also ReadFromBackup)."
    If NewSpecOk Then
       SaveAFile "Make a Backup", m_BakSpec
       MakeABackup = True
    Else
       FinalizeOp "Backup File Specs: Cannot Identify it!"
       MakeABackup = False
    End If
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function ReadFromBackup() As Boolean
Attribute ReadFromBackup.VB_Description = "Will reload the workbench by reading a previous backup (see MakeABackup and BakFileSpec)."
    If BakExists Then
       ReadFromBackup = LoadAFile("Read from Backup", m_BakSpec)
    Else
       FinalizeOp "Backup File Specs: Cannot Identify it!"
       ReadFromBackup = False
    End If
End Function

Private Function BakSpecOk() As Boolean
  Dim b As Boolean
   b = False
   If Len(Trim(m_BakSpec)) = 0 Then
      If SplitFileName Then
         BtxSpec = BtxPath & "\Bak\" & BtxFNam
         m_BakSpec = BtxSpec
         PropertyChanged "BakFileSpec"
         b = True
      End If
   Else
      b = True
   End If
   BakSpecOk = b
End Function

Private Function NewSpecOk() As Boolean
  Dim b As Boolean
   b = False
   If Len(Trim(m_NewSpec)) = 0 Then
      If SplitFileName Then
         BtxSpec = BtxPath & "\New\" & BtxFNam
         m_NewSpec = BtxSpec
         PropertyChanged "NewFileSpec"
         b = True
      End If
   Else
      b = True
   End If
   NewSpecOk = b
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function FileUpdate() As Boolean
Attribute FileUpdate.VB_Description = "Will save all workbench to input file indicated in FileSpec path. Previous existence will be killed."
    FileUpdate = SaveAFile("File Update", m_FileSpec)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function SaveAsNewFile() As Boolean
Attribute SaveAsNewFile.VB_Description = "Will write all workbench to output file indicated in NewFileSpec path. Previous existence will be killed."
    Dim FN As Integer
    Dim L As Long
    If NewSpecOk Then
       SaveAsNewFile = SaveAFile("Save As New File", m_NewSpec)
    Else
       FinalizeOp "Save As File Specs: Cannot Identify it!"
       SaveAsNewFile = False
    End If
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function Clear() As Boolean
Attribute Clear.VB_Description = "Will empty the workbench (all lines removed)."
    ReDim Btx(0) As String    ' empty WorkBench
    m_TotalLines = 0
    PropertyChanged "TotalLines"
    m_BakSpec = vbNullString
    PropertyChanged "BakFileSpec"
    m_NewSpec = vbNullString
    PropertyChanged "NewFileSpec"
    If Len(Trim(m_FileSpec)) = 0 Then
       Clear = False
    Else
       Clear = True
    End If
End Function

'1st redimension of arrays
Private Sub UserControl_Initialize()
   ReDim Btx(0) As String
   ReDim Htx(0) As String
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_FileSpec = m_def_FileSpec
    m_TotalLines = m_def_TotalLines
    m_BakSpec = m_def_BakSpec
    m_NewSpec = m_def_NewSpec
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_FileSpec = PropBag.ReadProperty("FileSpec", m_def_FileSpec)
    m_TotalLines = PropBag.ReadProperty("TotalLines", m_def_TotalLines)
    m_BakSpec = PropBag.ReadProperty("BakFileSpec", m_def_BakSpec)
    m_NewSpec = PropBag.ReadProperty("NewFileSpec", m_def_NewSpec)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("FileSpec", m_FileSpec, m_def_FileSpec)
    Call PropBag.WriteProperty("TotalLines", m_TotalLines, m_def_TotalLines)
    Call PropBag.WriteProperty("BakFileSpec", m_BakSpec, m_def_BakSpec)
    Call PropBag.WriteProperty("NewFileSpec", m_NewSpec, m_def_NewSpec)
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = JzImage.Width
    UserControl.Height = JzImage.Height
End Sub

'I N T E R N A L S

'This is the TotalLines communication after last Redim at WorkBench
Private Sub ShowTotal()
    m_TotalLines = UBound(Btx)
    PropertyChanged "TotalLines"
End Sub

'All Operations must be "FinalizeOp"ed - if error, with ErrMessg
Private Sub FinalizeOp(Optional ErrMessg As String = vbNullString)
    On Error GoTo 0    ' Restablishes any previous error treatment
    RaiseEvent OpCompletes(ErrMessg)
End Sub

'Simple file existing test
Private Function BtxExists(ByVal FileName As String) As Boolean
    Dim FileTam As Long
    On Error GoTo NonExistence
    FileTam = FileLen(FileName)
    BtxExists = True
    Exit Function
NonExistence:
    BtxExists = False
End Function

'used by ShureDirs function as 'reentrance'
Private Function SplitPath(sPath As String) As String
    SplitPath = Mid$(sPath, 1, InStrRev(sPath, "\", Len(sPath)) - 1)
End Function

'Be Shure path Dir exists
Private Function ShureDirs(sCompletePath As String) As String
    Dim s As String
    On Error GoTo ErroShureDirs
    s = SplitPath(sCompletePath)
    If LenB(Dir(s, vbDirectory)) = 0 Then
        s = ShureDirs(s)
        MkDir s
    End If
    ShureDirs = sCompletePath
    Exit Function
ErroShureDirs:
End Function

'a criterious function to split path-filename
Private Function SplitFileName() As Boolean
    Dim j As Integer
    Dim s As String
    Dim x As String
    Dim pv As String
    BtxPath = vbNullString
    BtxFNam = vbNullString
    BtxSpec = Trim(m_FileSpec)
    If Len(BtxSpec) > 0 Then
        'eliminates double char "\" if exists
        pv = vbNullString
        s = vbNullString
        For j = 1 To Len(BtxSpec)
            x = Mid$(BtxSpec, j, 1)
            If x = "\" Then
                If Not pv = "\" Then
                    s = s & x
                End If
            Else    'NOT X...
                s = s & x
            End If
            pv = x
        Next j
        BtxSpec = s
        'well, let's go ahead
        j = InStrRev(BtxSpec, "\")
        If j = 0 Then
            j = InStrRev(BtxSpec, ":")
        End If
        BtxFNam = vbNullString
        If j = 0 Then
            BtxPath = vbNullString
            BtxFNam = BtxSpec
        Else    'NOT J...
            BtxPath = Mid$(BtxSpec, 1, j - 1)
            BtxFNam = Mid$(BtxSpec, j + 1, Len(BtxSpec) - j)
        End If
    End If
    SplitFileName = Len(BtxPath) + Len(BtxFNam) > 0
End Function

Private Function LoadAFile(OpName As String, FileName As String) As Boolean
    Dim lin As String
    Dim L As Long
    Dim FN As Integer
    L = 0
    'Two pass on file: one to determine the number of lines
    'mode once array redimension - its fast and dynamic
    On Error GoTo R_Error
    FN = FreeFile
    Open FileName For Input As #FN
    Do While Not EOF(FN)
        L = L + 1
        Line Input #FN, lin
    Loop
    Close FN
    'now, adjusts array and get lines (preserve for speed purposes)
    ReDim Preserve Btx(L) As String
    L = 0
    On Error GoTo R_Error
    FN = FreeFile
    Open FileName For Input As #FN
    Do While Not EOF(FN)
        L = L + 1
        Line Input #FN, Btx(L)
    Loop
    Close FN
    ShowTotal
    FinalizeOp
    LoadAFile = True
    Exit Function
R_Error:
    FinalizeOp OpName & " Error"
    LoadAFile = False
End Function

Private Function SaveAFile(OpName As String, FileName As String) As Boolean
    Dim FN As Integer
    Dim L As Long
    If Len(FileName) = 0 Then
        SaveError OpName
        SaveAFile = False
        Exit Function
    End If
    ShureDirs FileName
    If BtxExists(FileName) Then
        Kill FileName
        DoEvents
    End If
    On Error GoTo sverr
    FN = FreeFile
    Open FileName For Output As #FN
    For L = 1 To UBound(Btx)
        Print #FN, Btx(L)
    Next L
    Close FN
    FinalizeOp
    SaveAFile = True
    Exit Function
sverr:
    SaveError OpName
    SaveAFile = False
End Function

Private Sub SaveError(OpName As String)
    FinalizeOp OpName & ": Cannot do it now or I-O Error"
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function Exists() As Boolean
Attribute Exists.VB_Description = "Returns True/False if original file exists in its path as 'FileName' specified."
   Exists = BtxExists(m_FileSpec)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function NewExists() As Boolean
Attribute NewExists.VB_Description = "Returns True/False if 'Save As New'l file exists in its path as 'NewFileName' specified (auto path generated)"
   If NewSpecOk Then
      NewExists = BtxExists(m_NewSpec)
   Else
      NewExists = False
   End If
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function BakExists() As Boolean
Attribute BakExists.VB_Description = "Returns True/False if backupl file exists in its path as 'BakFileName' specified (auto path generated)."
   If BakSpecOk Then
      BakExists = BtxExists(m_BakSpec)
   Else
      BakExists = False
   End If
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function OpenHole(LineNumber As Long) As Boolean
Attribute OpenHole.VB_Description = "Splits the Workbench into 2 parts, so that all '.AddLine' after this will apend lines to part-1. This is to speed insert two or more sucessive lines. When finished, must be 'CloseHole'd."
    Dim s As String
    Dim L As Long
    Dim M As Long
    If (LineNumber = 0 Or _
        LineNumber > UBound(Btx)) Then
        FinalizeOp "OpenHole: Out of Range"
        OpenHole = False
        Exit Function
    End If
    ReDim Htx(UBound(Btx)) As String 'prepare HoleBench
    M = 0
    L = LineNumber + 1
    Do While L <= m_TotalLines
       s = Btx(L)
       M = M + 1
       Htx(M) = s
       L = L + 1
    Loop
    ReDim Preserve Btx(LineNumber) As String
    ReDim Preserve Htx(M) As String
    FinalizeOp
    ShowTotal
    OpenHole = True
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function CloseHole() As Boolean
Attribute CloseHole.VB_Description = "Will join pat-2 of WorkBench than was 'OpenHole'd add 'AddLine'd for sucessive Inserts."
    Dim b As Boolean
    Dim s As String
    Dim L As Long
    Dim M As Long
    M = 1
    b = False
    L = UBound(Btx)
    ReDim Preserve Btx(L + UBound(Htx)) As String
    Do While M <= UBound(Htx)
       s = Htx(M)
       L = L + 1
       Btx(L) = s
       b = True
       M = M + 1
    Loop
    ReDim Htx(0) As String
    FinalizeOp
    ShowTotal
    CloseHole = b
End Function

'-oOo-oOo-oOo-


