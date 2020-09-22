VERSION 5.00
Begin VB.Form frmViewAddress 
   BorderStyle     =   0  'None
   Caption         =   "frmViewAddress"
   ClientHeight    =   5436
   ClientLeft      =   1428
   ClientTop       =   780
   ClientWidth     =   8148
   LinkTopic       =   "Form1"
   ScaleHeight     =   5436
   ScaleWidth      =   8148
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDone 
      Caption         =   "Finished"
      Height          =   288
      Left            =   7080
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.CommandButton cmdAddEntry 
      Caption         =   "Add Entry"
      Height          =   288
      Left            =   5880
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.Frame Frame2 
      Caption         =   "Home Address:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Left            =   120
      TabIndex        =   37
      Top             =   2004
      Width           =   7812
      Begin VB.TextBox txtStreet 
         Height          =   288
         Index           =   0
         Left            =   960
         TabIndex        =   11
         Top             =   480
         Width           =   2412
      End
      Begin VB.TextBox txtCity 
         Height          =   288
         Index           =   1
         Left            =   3540
         TabIndex        =   12
         Top             =   480
         Width           =   1092
      End
      Begin VB.TextBox txtZip 
         Height          =   288
         Index           =   1
         Left            =   5760
         TabIndex        =   14
         Top             =   480
         Width           =   1092
      End
      Begin VB.ComboBox cboState 
         Height          =   288
         Index           =   1
         Left            =   4848
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   480
         Width           =   732
      End
      Begin VB.Frame Frame6 
         Caption         =   "Home Address:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   972
         Left            =   120
         TabIndex        =   38
         Top             =   2124
         Width           =   7812
         Begin VB.TextBox txtPager 
            Height          =   288
            Index           =   2
            Left            =   5880
            TabIndex        =   41
            Top             =   480
            Width           =   1092
         End
         Begin VB.TextBox txtWPhone 
            Height          =   288
            Index           =   2
            Left            =   3420
            TabIndex        =   40
            Top             =   480
            Width           =   1092
         End
         Begin VB.TextBox txtHPhone 
            Height          =   288
            Index           =   2
            Left            =   240
            TabIndex        =   39
            Top             =   480
            Width           =   2412
         End
         Begin VB.Label lblPager 
            Caption         =   "Pager:"
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
            Index           =   2
            Left            =   5880
            TabIndex        =   44
            Top             =   240
            Width           =   1092
         End
         Begin VB.Label lblWPhone 
            Caption         =   "Work:"
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
            Index           =   2
            Left            =   3420
            TabIndex        =   43
            Top             =   240
            Width           =   996
         End
         Begin VB.Label lblHPone 
            Caption         =   "Home:"
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
            Index           =   2
            Left            =   240
            TabIndex        =   42
            Top             =   240
            Width           =   1692
         End
      End
      Begin VB.Label lblStreet 
         Caption         =   "Street:"
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
         Index           =   0
         Left            =   960
         TabIndex        =   48
         Top             =   240
         Width           =   1692
      End
      Begin VB.Label lblCity 
         Caption         =   "City:"
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
         Index           =   1
         Left            =   3540
         TabIndex        =   47
         Top             =   240
         Width           =   996
      End
      Begin VB.Label lblZip 
         Caption         =   "Zip Code:"
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
         Index           =   1
         Left            =   5760
         TabIndex        =   46
         Top             =   240
         Width           =   1092
      End
      Begin VB.Label lblState 
         Caption         =   "State:"
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
         Index           =   1
         Left            =   4848
         TabIndex        =   45
         Top             =   240
         Width           =   732
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Work Address:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Left            =   120
      TabIndex        =   32
      Top             =   3120
      Width           =   7812
      Begin VB.TextBox txtStreet 
         Height          =   288
         Index           =   1
         Left            =   960
         TabIndex        =   15
         Top             =   480
         Width           =   2412
      End
      Begin VB.TextBox txtCity 
         Height          =   288
         Index           =   0
         Left            =   3540
         TabIndex        =   16
         Top             =   480
         Width           =   1092
      End
      Begin VB.TextBox txtZip 
         Height          =   288
         Index           =   0
         Left            =   5760
         TabIndex        =   18
         Top             =   480
         Width           =   1092
      End
      Begin VB.ComboBox cboState 
         Height          =   288
         Index           =   0
         Left            =   4848
         Sorted          =   -1  'True
         TabIndex        =   17
         Top             =   480
         Width           =   732
      End
      Begin VB.Label lblStreet 
         Caption         =   "Street:"
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
         Index           =   1
         Left            =   960
         TabIndex        =   36
         Top             =   240
         Width           =   1692
      End
      Begin VB.Label lblCity 
         Caption         =   "City:"
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
         Index           =   0
         Left            =   3540
         TabIndex        =   35
         Top             =   240
         Width           =   996
      End
      Begin VB.Label lblZip 
         Caption         =   "Zip Code:"
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
         Index           =   0
         Left            =   5760
         TabIndex        =   34
         Top             =   240
         Width           =   1092
      End
      Begin VB.Label lblState 
         Caption         =   "State:"
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
         Index           =   0
         Left            =   4848
         TabIndex        =   33
         Top             =   240
         Width           =   732
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Phone Numbers:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Left            =   120
      TabIndex        =   25
      Top             =   840
      Width           =   7812
      Begin VB.TextBox txtPager 
         Height          =   288
         Index           =   0
         Left            =   2760
         TabIndex        =   7
         Top             =   480
         Width           =   1092
      End
      Begin VB.TextBox txtWPhone 
         Height          =   288
         Index           =   0
         Left            =   1500
         TabIndex        =   6
         Top             =   480
         Width           =   1092
      End
      Begin VB.TextBox txtHPhone 
         Height          =   288
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1092
      End
      Begin VB.TextBox txtOPhone 
         Height          =   288
         Index           =   1
         Left            =   6528
         TabIndex        =   10
         Top             =   480
         Width           =   1092
      End
      Begin VB.TextBox txtAMobile 
         Height          =   288
         Index           =   2
         Left            =   5268
         TabIndex        =   26
         Top             =   480
         Width           =   1092
      End
      Begin VB.TextBox txtMobile 
         Height          =   288
         Index           =   2
         Left            =   4008
         TabIndex        =   8
         Top             =   480
         Width           =   1092
      End
      Begin VB.Label lblPager 
         Caption         =   "Pager:"
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
         Index           =   0
         Left            =   2760
         TabIndex        =   31
         Top             =   240
         Width           =   1092
      End
      Begin VB.Label lblWPhone 
         Caption         =   "Work:"
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
         Index           =   0
         Left            =   1500
         TabIndex        =   30
         Top             =   240
         Width           =   996
      End
      Begin VB.Label lblHPone 
         Caption         =   "Home:"
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
         Index           =   1
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   1092
      End
      Begin VB.Label lblOPhone 
         Caption         =   "Other Phone:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   6525
         TabIndex        =   28
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblAMobile 
         Caption         =   "Sec Mobile:"
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
         Index           =   2
         Left            =   5268
         TabIndex        =   9
         Top             =   240
         Width           =   1092
      End
      Begin VB.Label lblMobile 
         Caption         =   "Mobile:"
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
         Index           =   2
         Left            =   4008
         TabIndex        =   27
         Top             =   240
         Width           =   1092
      End
   End
   Begin VB.ComboBox cboSuffix 
      Height          =   288
      Index           =   0
      Left            =   4728
      TabIndex        =   4
      Top             =   360
      Width           =   732
   End
   Begin VB.TextBox txtLName 
      Height          =   288
      Index           =   2
      Left            =   3288
      TabIndex        =   3
      Top             =   360
      Width           =   1332
   End
   Begin VB.TextBox txtMName 
      Height          =   288
      Index           =   1
      Left            =   1848
      TabIndex        =   2
      Top             =   360
      Width           =   1332
   End
   Begin VB.TextBox txtFName 
      Height          =   288
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   1332
   End
   Begin VB.Frame Frame4 
      Caption         =   "Email Addresses:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Left            =   120
      TabIndex        =   0
      Top             =   4320
      Width           =   7812
      Begin VB.TextBox txtHEmail 
         Height          =   288
         Index           =   2
         Left            =   120
         MousePointer    =   2  'Cross
         TabIndex        =   19
         ToolTipText     =   "Click to Send Email"
         Top             =   480
         Width           =   2412
      End
      Begin VB.TextBox txtWEmail 
         Height          =   288
         Index           =   3
         Left            =   2760
         MousePointer    =   2  'Cross
         TabIndex        =   20
         ToolTipText     =   "Click to Send Email"
         Top             =   480
         Width           =   2412
      End
      Begin VB.TextBox txtSecEmail 
         Height          =   288
         Index           =   4
         Left            =   5280
         MousePointer    =   2  'Cross
         TabIndex        =   21
         ToolTipText     =   "Click to Send Email"
         Top             =   480
         Width           =   2412
      End
      Begin VB.Label lblHEmail 
         Caption         =   "Home:"
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
         Index           =   2
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1692
      End
      Begin VB.Label lblWEmail 
         Caption         =   "Work:"
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
         Index           =   3
         Left            =   2760
         TabIndex        =   23
         Top             =   240
         Width           =   1692
      End
      Begin VB.Label lblSecEmail 
         Caption         =   "Secondary Email:"
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
         Index           =   4
         Left            =   5280
         TabIndex        =   22
         Top             =   240
         Width           =   1692
      End
   End
   Begin VB.Label lblSuffix 
      Caption         =   "Suffix:"
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
      Index           =   0
      Left            =   4728
      TabIndex        =   52
      Top             =   120
      Width           =   732
   End
   Begin VB.Label lblLName 
      Caption         =   "Last Name:"
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
      Index           =   2
      Left            =   3288
      TabIndex        =   51
      Top             =   120
      Width           =   1332
   End
   Begin VB.Label lblMName 
      Caption         =   "Middle Name:"
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
      Index           =   1
      Left            =   1848
      TabIndex        =   50
      Top             =   120
      Width           =   1332
   End
   Begin VB.Label lblFName 
      Caption         =   "First Name:"
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
      Index           =   0
      Left            =   360
      TabIndex        =   49
      Top             =   120
      Width           =   1308
   End
End
Attribute VB_Name = "frmViewAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const S1 As String = "Sr."
Private Const S2 As String = "Jr."
Private Const S3 As String = "I"
Private Const S4 As String = "II"
Private Const S5 As String = "III"
Private Const S6 As String = "IV"

Private Sub cmdAddEntry_Click()

Call AddEntry

End Sub

Private Sub cmdDone_Click()

cmdAddEntry.Visible = False
cmdDone.Visible = False

End Sub

Private Sub Form_Unload(Cancel As Integer)

Set frmViewAddress = Nothing

End Sub

Private Sub Form_Load()
Dim i As Integer
    
    With cboSuffix(0)
        .AddItem S1
        .AddItem S2
        .AddItem S3
        .AddItem S4
        .AddItem S5
        .AddItem S6
    End With
    
    For i = 0 To 1
          With cboState(i)
            .AddItem "AL"
            .AddItem "AK"
            .AddItem "AZ"
            .AddItem "AR"
            .AddItem "CA"
            .AddItem "CO"
            .AddItem "CT"
            .AddItem "DE"
            .AddItem "DC"
            .AddItem "FL"
            .AddItem "GA"
            .AddItem "HI"
            .AddItem "ID"
            .AddItem "IL"
            .AddItem "IN"
            .AddItem "IA"
            .AddItem "KS"
            .AddItem "KY"
            .AddItem "LA"
            .AddItem "ME"
            .AddItem "MD"
            .AddItem "MA"
            .AddItem "MN"
            .AddItem "MI"
            .AddItem "MS"
            .AddItem "MO"
            .AddItem "MT"
            .AddItem "NE"
            .AddItem "NV"
            .AddItem "NH"
            .AddItem "NJ"
            .AddItem "NM"
            .AddItem "NY"
            .AddItem "NC"
            .AddItem "ND"
            .AddItem "OH"
            .AddItem "OK"
            .AddItem "OR"
            .AddItem "PA"
            .AddItem "RI"
            .AddItem "SC"
            .AddItem "SD"
            .AddItem "TN"
            .AddItem "TX"
            .AddItem "UT"
            .AddItem "VT"
            .AddItem "VA"
            .AddItem "WA"
            .AddItem "WV"
            .AddItem "WI"
            .AddItem "WY"
        End With
    Next i
        
    For i = 0 To 1
        With cboState(i)
            .ListIndex = 3
        End With
    Next i
    
End Sub

Public Sub Clear()

txtFName(0).Text = ""
txtLName(2).Text = ""
txtMName(1).Text = ""
cboSuffix(0).Text = ""
txtHPhone(1).Text = ""
txtWPhone(0).Text = ""
txtPager(0).Text = ""
txtMobile(2).Text = ""
txtAMobile(2).Text = ""
txtOPhone(1).Text = ""
txtStreet(0).Text = ""
txtCity(1).Text = ""
cboState(1).Text = ""
txtZip(1).Text = ""
txtStreet(1).Text = ""
txtCity(0).Text = ""
cboState(0).Text = ""
txtZip(0).Text = ""
txtHEmail(2).Text = ""
txtWEmail(3).Text = ""
txtSecEmail(4).Text = ""

End Sub

Private Sub txtHEmail_Click(Index As Integer)

If txtHEmail(2).Text = "" Then
    Exit Sub
Else
    Send_Email_To (txtHEmail(2).Text)
End If

End Sub

Private Sub txtHEmail_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If txtHEmail(2).Text = "" Then
   txtHEmail(2).MousePointer = 0
Else
   txtHEmail(2).MousePointer = 2
End If

End Sub

Private Sub txtSecEmail_Click(Index As Integer)

If txtSecEmail(4).Text = "" Then
    Exit Sub
Else
    Send_Email_To (txtSecEmail(4).Text)
End If

End Sub

Private Sub txtSecEmail_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If txtSecEmail(4).Text = "" Then
   txtSecEmail(4).MousePointer = 0
Else
   txtSecEmail(4).MousePointer = 2
End If

End Sub

Private Sub txtWEmail_Click(Index As Integer)

If txtWEmail(3).Text = "" Then
    Exit Sub
Else
    Send_Email_To (txtWEmail(3).Text)
End If

End Sub

Private Sub txtWEmail_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If txtWEmail(3).Text = "" Then
   txtWEmail(3).MousePointer = 0
Else
   txtWEmail(3).MousePointer = 2
End If

End Sub
