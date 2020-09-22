VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Welcome to the Address Organizer."
   ClientHeight    =   5412
   ClientLeft      =   3240
   ClientTop       =   1428
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5412
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00000000&
      Height          =   5412
      Left            =   0
      ScaleHeight     =   5364
      ScaleWidth      =   4884
      TabIndex        =   0
      Top             =   0
      Width           =   4932
      Begin VB.PictureBox Picture1 
         Height          =   3156
         Left            =   516
         Picture         =   "frmSplash.frx":0000
         ScaleHeight     =   3108
         ScaleWidth      =   3840
         TabIndex        =   2
         Top             =   300
         Width           =   3888
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00000000&
         Height          =   1572
         Left            =   516
         ScaleHeight     =   1524
         ScaleWidth      =   3840
         TabIndex        =   1
         Top             =   3540
         Width           =   3888
         Begin VB.Label lblVer 
            BackColor       =   &H00000000&
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   252
            Left            =   120
            TabIndex        =   5
            Top             =   1080
            Width           =   1452
         End
         Begin VB.Label lblCRight 
            BackColor       =   &H00000000&
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   252
            Left            =   120
            TabIndex        =   4
            Top             =   654
            Width           =   3492
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   3492
            WordWrap        =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

'Used to center the form in the screen
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2

Label1.Caption = ""
Label1.Caption = "AddressBook Organizer"

lblCRight.Caption = ""
lblCRight.Caption = "www.somesight.com" & " " & "Copyright:" & "  " & "2000 "

lblVer.Caption = ""
lblVer.Caption = "Version" & " " & App.Major & "." & App.Minor & "." & App.Revision

End Sub

