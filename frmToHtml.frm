VERSION 5.00
Begin VB.Form frmExport 
   BorderStyle     =   0  'None
   Caption         =   "Export Database to Selected Format"
   ClientHeight    =   5448
   ClientLeft      =   996
   ClientTop       =   2004
   ClientWidth     =   8136
   Icon            =   "frmToHtml.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5448
   ScaleWidth      =   8136
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton butDoIt 
      Caption         =   "Finish"
      Height          =   330
      Left            =   2412
      TabIndex        =   8
      Top             =   3840
      Width           =   1308
   End
   Begin VB.CommandButton butSetSource 
      Caption         =   "Select Table"
      Height          =   372
      Left            =   4800
      TabIndex        =   7
      Top             =   1080
      Width           =   1236
   End
   Begin VB.TextBox txtDestPath 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2412
      TabIndex        =   5
      Top             =   3360
      Width           =   3636
   End
   Begin VB.ListBox lstTables 
      Height          =   1008
      Left            =   2412
      TabIndex        =   4
      Top             =   1080
      Width           =   2220
   End
   Begin VB.ComboBox cboExportTo 
      Height          =   288
      Left            =   2412
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2928
      Width           =   1800
   End
   Begin VB.TextBox txtSQL 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   6942
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   552
   End
   Begin VB.TextBox txtSourcePath 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   0
      Text            =   "c:\Address Book\AddressBook.mdb"
      Top             =   600
      Width           =   3636
   End
   Begin VB.Label Label1 
      Caption         =   "Step Three:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Index           =   6
      Left            =   642
      TabIndex        =   13
      Top             =   3840
      Width           =   1248
   End
   Begin VB.Label Label1 
      Caption         =   "Source File:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Index           =   0
      Left            =   642
      TabIndex        =   12
      Top             =   576
      Width           =   1176
   End
   Begin VB.Label Label1 
      Caption         =   "Step Two:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Index           =   2
      Left            =   642
      TabIndex        =   11
      Top             =   2928
      Width           =   1008
   End
   Begin VB.Label Label1 
      Caption         =   "Step One:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Index           =   3
      Left            =   642
      TabIndex        =   10
      Top             =   1080
      Width           =   1176
   End
   Begin VB.Label Label1 
      Caption         =   "Destination Path:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Index           =   4
      Left            =   648
      TabIndex        =   9
      Top             =   3360
      Width           =   1476
   End
   Begin VB.Label Label1 
      Caption         =   "Select Type of Export:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Index           =   5
      Left            =   4452
      TabIndex        =   6
      Top             =   2928
      Width           =   1932
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   7680
      X2              =   420
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label Label1 
      Caption         =   "SQL Export Code:"
      Height          =   228
      Index           =   1
      Left            =   6480
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1512
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Much of this i found on some other web site, and credit is due to the original author.
'I can not remember his name; sorry

Option Compare Text

Public sExt As String
Private dbSource As Database

Private Sub ListSourceTables()

Dim tDef As TableDef

For Each tDef In dbSource.TableDefs
    If Left$(tDef.Name, 4) <> "MSys" Then
        lstTables.AddItem tDef.Name
    End If
Next

Set tDef = Nothing

End Sub

Private Sub ResetForm()

SetSourcePath
SetDestPath
lstTables.Clear

With cboExportTo
     .AddItem "Excel 4.0;"
     .AddItem "Text;"
     .AddItem "HTML Export;"
End With

txtSQL = ""
txtSourcePath = App.Path & "\BlackBook.mdb"
txtDestPath = App.Path & "\HTML\"
    
butDoIt.Enabled = False

End Sub

Private Sub SetDestPath()

txtDestPath = App.Path & "\HTML\"

End Sub

Private Sub SetSourcePath()

txtSourcePath = App.Path & "\BlackBook.mdb"

End Sub

Private Sub butDoIt_Click()

On Error Resume Next

Kill txtDestPath & "*.*"

If Err.Number <> 0 Then
    Err.Clear
End If

dbSource.Execute txtSQL, dbFailOnError
MsgBox "BlackBook.mdb" & " " & "Has been Exported as:" & " " & cboExportTo & " " & "to" & " " & txtDestPath, vbInformation, "Export Complete:"

End Sub

Private Sub butSetSource_Click()

Dim sCheckPath As String

sCheckPath = Trim$(txtSourcePath)

If FileExist(sCheckPath) = False Then
    MsgBox "Source Jet Database not found"
    Exit Sub
End If

'pass thru, got a database file
txtSourcePath = sCheckPath
Set dbSource = Workspaces(0).OpenDatabase(sCheckPath, False, False)

ListSourceTables

End Sub

Private Function FileExist(ByVal sIn As String, Optional flags As Long = 0) As Boolean

On Error Resume Next
If flags = 0 Then
    FileExist = Dir$(sIn) <> ""
Else
    FileExist = Dir$(sIn, flags) <> ""
End If
    If Err.Number <> 0 Then
            FileExist = False
    End If
    
End Function

Private Sub cboExportTo_Click()

If cboExportTo = "Text;" Then
    sExt = ".txt"
ElseIf cboExportTo = "HTML Export;" Then
    sExt = ".htm"
End If
txtSQL = "SELECT * INTO [" & cboExportTo & "DATABASE=" & txtDestPath & "].[BlackBook" & sExt & "] " & _
    "FROM [" & lstTables & "]"
    
End Sub


Private Sub Form_Load()

ResetForm

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Set dbSource = Nothing

Unload frmExport
frmMain.Show

End Sub


Private Sub txtSQL_Change()

butDoIt.Enabled = (txtSQL <> "") And lstTables <> ""

End Sub


