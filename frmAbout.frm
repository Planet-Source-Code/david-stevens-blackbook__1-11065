VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   Caption         =   "frmAbout"
   ClientHeight    =   5436
   ClientLeft      =   1308
   ClientTop       =   1392
   ClientWidth     =   8124
   LinkTopic       =   "Form1"
   Picture         =   "frmAbout.frx":0000
   ScaleHeight     =   5436
   ScaleWidth      =   8124
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtLegal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3492
      Left            =   1116
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1560
      Width           =   5892
   End
   Begin VB.Label lblWebSite 
      BackStyle       =   0  'Transparent
      Caption         =   "Web Site: www.planetsourcecode.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   372
      Index           =   2
      Left            =   2544
      MousePointer    =   2  'Cross
      TabIndex        =   3
      Top             =   600
      Width           =   3012
   End
   Begin VB.Label lblEmail 
      BackStyle       =   0  'Transparent
      Caption         =   "Email: stevens_dl@hotmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   372
      Index           =   1
      Left            =   2544
      MousePointer    =   2  'Cross
      TabIndex        =   2
      Top             =   360
      Width           =   2892
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Developed By: David L. Stevens"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   372
      Index           =   0
      Left            =   2544
      TabIndex        =   1
      Top             =   120
      Width           =   2892
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

txtLegal.Text = _
"Disclaimer:" & vbCrLf & _
"Code provided by David Lee Stevens 'as is', without warranties as to" & vbCrLf & _
"performance, fitness, merchantability, and any other warranty" & vbCrLf & _
"(whether expressed or implied)." & vbCrLf & _
"This source code is copyrighted by David Lee Stevens who has exclusive" & vbCrLf & _
"rights to distribute it." & vbCrLf & _
" Freeware:" & vbCrLf & _
"Code is freely redistributable for personal use in source code form," & vbCrLf & _
"or for personal or business use in a non-source code binary executable." & vbCrLf & _
"All other redistributions are prohibited without express written consent" & vbCrLf & _
"from David Lee Stevens."

End Sub

Private Sub Form_Unload(Cancel As Integer)

Set frmAbout = Nothing

End Sub

Private Sub lblEmail_Click(Index As Integer)

Send_Email_To ("stevens_dl@hotmail.com")

End Sub

Private Sub lblWebSite_Click(Index As Integer)

Shell "start.exe http://www.planetsourcecode.com", vbHide

End Sub
