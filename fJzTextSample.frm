VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form fJzTextSample 
   BackColor       =   &H00FFFFFF&
   Caption         =   "  Friendly File Text Operations"
   ClientHeight    =   6075
   ClientLeft      =   2325
   ClientTop       =   2865
   ClientWidth     =   8880
   Icon            =   "fJzTextSample.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cmd 
      Caption         =   "Open Hole at Line No."
      Height          =   300
      Index           =   11
      Left            =   120
      TabIndex        =   32
      ToolTipText     =   "After This, try sucessive AddLine's them CloseHole"
      Top             =   4275
      Width           =   2100
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "Close Hole (Rejoin)"
      Height          =   300
      Index           =   12
      Left            =   120
      TabIndex        =   31
      ToolTipText     =   "Rejoin Part-2 were 'Holed'"
      Top             =   4605
      Width           =   2100
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "Insert txtLine After Line No."
      Height          =   300
      Index           =   4
      Left            =   120
      TabIndex        =   30
      ToolTipText     =   "Enter some txtLine modifications"
      Top             =   1800
      Width           =   2100
   End
   Begin VB.TextBox txtLine 
      Height          =   285
      Left            =   2400
      TabIndex        =   26
      ToolTipText     =   "A Work Line to get, ste, update, etc"
      Top             =   5685
      Width           =   6330
   End
   Begin VB.TextBox Up_ToLine 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   915
      TabIndex        =   23
      Text            =   " 0"
      ToolTipText     =   "Optional ToLine used in Remove Op"
      Top             =   5670
      Width           =   750
   End
   Begin VB.TextBox Up_Line 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   90
      TabIndex        =   22
      Text            =   "0"
      ToolTipText     =   "LineNumber parameter"
      Top             =   5670
      Width           =   750
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "Save As New File"
      Height          =   300
      Index           =   10
      Left            =   120
      TabIndex        =   20
      ToolTipText     =   "Write workbench to a new file"
      Top             =   3870
      Width           =   2100
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "Update Original File"
      Height          =   300
      Index           =   9
      Left            =   120
      TabIndex        =   19
      ToolTipText     =   "Rewrite modified workbench over original file"
      Top             =   3570
      Width           =   2100
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "Append txtLine"
      Height          =   300
      Index           =   6
      Left            =   120
      TabIndex        =   18
      ToolTipText     =   "Put line after benchwork end"
      Top             =   2415
      Width           =   2100
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "Show Line Numbers"
      Height          =   300
      Index           =   1
      Left            =   120
      TabIndex        =   17
      ToolTipText     =   "Affects only 'Workbench from' text area"
      Top             =   690
      Width           =   2100
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "Restore from Last Backup"
      Height          =   300
      Index           =   8
      Left            =   120
      TabIndex        =   16
      ToolTipText     =   "See auto Path-filename generated"
      Top             =   3165
      Width           =   2100
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "Remove Line(s) Line No."
      Height          =   300
      Index           =   5
      Left            =   120
      TabIndex        =   15
      ToolTipText     =   "For more lines, use ToLine box"
      Top             =   2115
      Width           =   2100
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "Put txtLine OVER Line No."
      Height          =   300
      Index           =   3
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "Enter some txtLine modifications"
      Top             =   1485
      Width           =   2100
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "Make a Backup of all"
      Height          =   300
      Index           =   7
      Left            =   120
      TabIndex        =   13
      ToolTipText     =   "See auto Path-filename generated"
      Top             =   2850
      Width           =   2100
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "Get txtLine  by  Line No."
      Height          =   300
      Index           =   2
      Left            =   120
      TabIndex        =   12
      ToolTipText     =   "Enter Line No see txtLine"
      Top             =   1155
      Width           =   2100
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "Open Original File"
      Height          =   300
      Index           =   0
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "Prep: Choose Input FileName"
      Top             =   360
      Width           =   2100
   End
   Begin JzTxSample.JzTxFile Tx1 
      Left            =   2430
      Top             =   180
      _ExtentX        =   714
      _ExtentY        =   820
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   1785
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Bt_CD2 
      Caption         =   "..."
      Height          =   285
      Left            =   8295
      TabIndex        =   5
      ToolTipText     =   "Choose a type Text File"
      Top             =   780
      Width           =   405
   End
   Begin VB.CommandButton Bt_CD1 
      Caption         =   "..."
      Height          =   285
      Left            =   8295
      TabIndex        =   3
      ToolTipText     =   "Choose a type Text File"
      Top             =   120
      Width           =   405
   End
   Begin VB.TextBox Up_Bench 
      Height          =   4170
      Left            =   2400
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   1380
      Width           =   6315
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(If Truncated, no problems with files)"
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   6075
      TabIndex        =   29
      Top             =   1140
      Width           =   2550
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   " Backup Auto Filename "
      Height          =   195
      Left            =   2955
      TabIndex        =   28
      Top             =   510
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   " A txtLine"
      Height          =   255
      Left            =   1695
      TabIndex        =   27
      Top             =   5700
      Width           =   690
   End
   Begin VB.Label Label7 
      Caption         =   " To Line"
      Height          =   180
      Left            =   930
      TabIndex        =   25
      Top             =   5475
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   " Line No."
      Height          =   195
      Left            =   90
      TabIndex        =   24
      Top             =   5475
      Width           =   645
   End
   Begin VB.Label Label5 
      Caption         =   " Examples:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   150
      TabIndex        =   21
      Top             =   75
      Width           =   1080
   End
   Begin VB.Label Lb_Bak 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   4665
      TabIndex        =   11
      Top             =   450
      Width           =   3630
   End
   Begin VB.Label Lb_TotalLines 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   1170
      TabIndex        =   9
      Top             =   5040
      Width           =   945
   End
   Begin VB.Label Label4 
      Caption         =   "Total Lines:"
      Height          =   225
      Left            =   270
      TabIndex        =   8
      Top             =   5085
      Width           =   855
   End
   Begin VB.Label Lb_Out 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   4665
      TabIndex        =   7
      Top             =   780
      Width           =   3630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   " 'Save As New' Filename "
      Height          =   195
      Left            =   2850
      TabIndex        =   6
      Top             =   810
      Width           =   1800
   End
   Begin VB.Label Lb_In 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   4665
      TabIndex        =   4
      Top             =   120
      Width           =   3630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   " Watch Window from OCX workbench "
      Height          =   195
      Left            =   2445
      TabIndex        =   2
      Top             =   1140
      Width           =   2760
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   " Original Input Filename "
      Height          =   195
      Left            =   2955
      TabIndex        =   1
      Top             =   180
      Width           =   1695
   End
End
Attribute VB_Name = "fJzTextSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Bt_CD1_Click()
    CD1.DialogTitle = " Choose a type text File to be 'Original'"
    CD1.Flags = cdlOFNFileMustExist
    CD1.CancelError = True
    On Error GoTo Sai
    CD1.ShowOpen
    Lb_In.Caption = CD1.FileName
    Tx1.Clear
    Tx1.FileSpec = CD1.FileName
Sai:
    On Error GoTo 0
End Sub

Private Sub Bt_CD2_Click()
    CD1.DialogTitle = " Choose a type text File name to be 'Save As NewFile'"
    CD1.Flags = cdlOFNOverwritePrompt
    CD1.CancelError = True
    On Error GoTo Sai
    CD1.ShowSave
    Lb_Out.Caption = CD1.FileName
    Tx1.NewFileSpec = CD1.FileName
Sai:
    On Error GoTo 0
End Sub

Private Sub Cmd_Click(Index As Integer)
    Dim L As Long
    Dim M As Long
    Select Case Index
    Case 0    'Open File
        If Tx1.OpenFile Then
           Load_Up_Bench
        End If
    Case 1    'Show Line Numbers
        TextBoxRenum
    Case 2    'Get Any Line by  Number
        L = Val(Up_Line.Text)
        If L > 0 Then
            txtLine.Text = Tx1.GetLine(L)
        End If
    Case 3    'Put txtLine OVER Line No.
        L = Val(Up_Line.Text)
        If Tx1.PutLine(L, txtLine.Text) Then
           Load_Up_Bench
        End If
    Case 4    'Insert txtLine After Line No.
        L = Val(Up_Line.Text)
        If Tx1.InsertLine(L, txtLine.Text) Then
           Load_Up_Bench
        End If
    Case 5    'Remove Line(s)
        L = Val(Up_Line.Text)
        M = Val(Up_ToLine.Text)
        If Tx1.RemoveLine(L, M) Then
           Load_Up_Bench
        End If
    Case 6    'Append a Line
        Tx1.AddLine txtLine.Text
        Load_Up_Bench
    Case 7    'Make a Backup of all
        Tx1.MakeABackup
        Lb_Bak.Caption = Tx1.BakFileSpec
    Case 8    'Restore Last Backup
        Tx1.ReadFromBackup
        Lb_Bak.Caption = Tx1.BakFileSpec
        Load_Up_Bench
    Case 9    'Save Original File
        Tx1.FileUpdate
    Case 10    'Save As new Output
        Tx1.SaveAsNewFile
    Case 11
        L = Val(Up_Line.Text)
        If Tx1.OpenHole(L) Then
           Load_Up_Bench
        End If
    Case 12
        Tx1.CloseHole
        Load_Up_Bench
    End Select
End Sub

Private Sub Load_Up_Bench()
    Dim L As Long
    Up_Bench.Visible = False    'to avoid redraw
    Up_Bench.Text = vbNullString
    For L = 1 To Tx1.TotalLines
        Up_Bench.Text = Up_Bench.Text & Tx1.GetLine(L) & vbCrLf
    Next L
    Up_Bench.Refresh
    Up_Bench.Visible = True
    Lb_TotalLines = CStr(Tx1.TotalLines)
End Sub

Private Sub Form_Load()
  Me.Caption = Me.Caption & Space(1) & "v" & _
               App.Major & "." & App.Minor & _
               "." & Format(App.Revision, "00")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   End
End Sub

'Would be a Error Messages Central
'Message box can be customized and
'... this subroutine can be used display
'    other text file referent messages
Private Sub Tx1_OpCompletes(Message As String)
    If Len(Message) <> 0 Then
        Call MsgBox("Input=-" & Tx1.FileSpec _
                    & vbCrLf & "Lines=-" & Tx1.TotalLines _
                    & vbCrLf & "" _
                    & vbCrLf & Message _
                    , vbExclamation, " TEXT OPERATION ERROR:")
    End If
End Sub

Private Sub TextBoxRenum()
    Dim lin() As String
    Dim L As Long
    Up_Bench.Visible = False    'to avoid redraw
    lin = Split(Up_Bench.Text, vbCrLf)
    Up_Bench.Text = vbNullString
    For L = 0 To UBound(lin)
        lin(L) = Format(L + 1, "0000") & ":-" & lin(L)
        Up_Bench.Text = Up_Bench.Text & lin(L) & vbCrLf
    Next L
    Up_Bench.Refresh
    Up_Bench.Visible = True
End Sub
