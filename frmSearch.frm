VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearch 
   BorderStyle     =   0  'None
   Caption         =   "frmSearch"
   ClientHeight    =   5448
   ClientLeft      =   576
   ClientTop       =   1356
   ClientWidth     =   8136
   LinkTopic       =   "Form1"
   ScaleHeight     =   5448
   ScaleWidth      =   8136
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   288
      Left            =   3240
      TabIndex        =   1
      Top             =   1800
      Width           =   1452
   End
   Begin VB.TextBox txtSearch 
      Height          =   288
      Left            =   462
      TabIndex        =   0
      Top             =   1800
      Width           =   2652
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search By:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1452
      Left            =   462
      TabIndex        =   3
      Top             =   240
      Width           =   2652
      Begin VB.OptionButton optLastName 
         Caption         =   "Last Name"
         Height          =   252
         Left            =   120
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   840
         Width           =   1692
      End
      Begin VB.OptionButton optFirstName 
         Caption         =   "First Name"
         Height          =   252
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   360
         Width           =   2172
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3132
      Left            =   462
      TabIndex        =   2
      Top             =   2160
      Width           =   7212
      _ExtentX        =   12721
      _ExtentY        =   5525
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "ID"
         Text            =   "ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "FirstName"
         Text            =   "First Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "MiddleName"
         Text            =   "Middle Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "LastName"
         Text            =   "Last Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "Sufix"
         Text            =   "Suffix"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdSearch_Click()

MsgBox "Not Done"
'Call Search

End Sub

Private Sub Form_Unload(Cancel As Integer)

Set frmSearch = Nothing

End Sub

