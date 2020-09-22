VERSION 5.00
Begin VB.Form frmUrl 
   BorderStyle     =   0  'None
   Caption         =   "frmURL"
   ClientHeight    =   5436
   ClientLeft      =   1992
   ClientTop       =   1380
   ClientWidth     =   8136
   LinkTopic       =   "Form1"
   ScaleHeight     =   5436
   ScaleWidth      =   8136
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtUrl 
      Height          =   288
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   7692
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete URL"
      Height          =   372
      Left            =   1680
      TabIndex        =   2
      Top             =   1200
      Width           =   1092
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add URL"
      Height          =   372
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1092
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2928
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Dbl Click to Goto Web Site."
      Top             =   1800
      Width           =   7692
   End
   Begin VB.Label Label1 
      Caption         =   "Type Your Favorite URL'S Here."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   3132
   End
End
Attribute VB_Name = "frmUrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()

Dim url As String

url = txtUrl.Text
    info.rsURLS.Open

If url <> "" Then

    info.rsURLS.AddNew
    
    List1.AddItem txtUrl.Text
    info.rsURLS.Fields("URLS") = url
    
    info.rsURLS.Update
    info.rsURLS.AddNew
End If

Set info = Nothing

txtUrl.Text = "http://www."

End Sub

Private Sub cmdDelete_Click()

Dim i As Integer
Dim result As Integer
Dim url As String
Dim s As String
Dim Search_String As String

Search_String = " "


txtUrl.Text = List1.Text
url = txtUrl.Text
s = InStr(1, url, Search_String)
url = Mid(url, 1, s - 1)

result = MsgBox("Are you sure you want to Delete" & " " & txtUrl.Text & " ?", vbYesNo, "Delete URL")

If result = vbYes Then
    info.rsURLS.Open
    info.rsURLS.Find "ID = " & Trim(url)
    info.rsURLS.Delete adAffectCurrent
    info.rsURLS.Close
Else
    Exit Sub
    
End If

If List1.ListIndex = -1 Then
    Exit Sub
Else
    List1.RemoveItem List1.ListIndex
End If

txtUrl.Text = "http://www."


End Sub

Private Sub Form_Load()

info.rsURLS.Open

txtUrl.Text = "http://www."

Do While info.rsURLS.EOF = False
    List1.AddItem info.rsURLS.Fields("ID") & "  " & info.rsURLS.Fields("URLS")
    info.rsURLS.MoveNext
Loop

info.rsURLS.Close

End Sub

Private Sub Form_Unload(Cancel As Integer)

Set frmUrl = Nothing

End Sub

Private Sub List1_DblClick()

Dim GoUrl As String
Dim tmpStr As String
Dim Search_String As String
Dim space As String

Search_String = " "
tmpStr = List1.Text

space = InStr(1, tmpStr, Search_String)

GoUrl = Mid(tmpStr, space + 1)


Shell "start.exe" & GoUrl, vbHide

End Sub
