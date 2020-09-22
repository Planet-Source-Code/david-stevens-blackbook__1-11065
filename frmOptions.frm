VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   0  'None
   Caption         =   "frmOptions"
   ClientHeight    =   5448
   ClientLeft      =   1284
   ClientTop       =   1440
   ClientWidth     =   8136
   LinkTopic       =   "Form1"
   ScaleHeight     =   5448
   ScaleWidth      =   8136
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Other Options:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2052
      Left            =   240
      TabIndex        =   7
      Top             =   3120
      Width           =   7572
   End
   Begin VB.Frame Frame2 
      Caption         =   "Default Photo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2172
      Left            =   5400
      TabIndex        =   4
      Top             =   600
      Width           =   2412
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         Height          =   252
         Left            =   1080
         TabIndex        =   6
         Top             =   1800
         Width           =   1092
      End
      Begin VB.CheckBox chkDefault 
         Caption         =   "Reset Orginal Photo."
         Height          =   252
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   2052
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Main Display Option:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2172
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   4932
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   2160
         Top             =   120
         _ExtentX        =   677
         _ExtentY        =   677
         _Version        =   393216
      End
      Begin VB.CommandButton cmdNewPhoto 
         Caption         =   "Choose Photo"
         Height          =   252
         Left            =   2520
         TabIndex        =   3
         Top             =   1800
         Width           =   2052
      End
      Begin VB.CheckBox chkMDis 
         Caption         =   "Check this to change the Main Display Photo to Your Own."
         Height          =   252
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   4572
      End
      Begin VB.Label lblInfo 
         Caption         =   "Label1"
         Height          =   732
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   4212
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkDefault_Click()

If chkDefault.Value = vbChecked Then
   cmdReset.Visible = True
Else
   cmdReset.Visible = False
End If


End Sub

Private Sub chkMDis_Click()

If chkMDis.Value = vbChecked Then
   cmdNewPhoto.Visible = True
Else
   cmdNewPhoto.Visible = False
  
End If

End Sub

Private Sub cmdNewPhoto_Click()

On Error GoTo ErrHandler

'--------- Open file of choice ----------'
                
CommonDialog1.CancelError = True
CommonDialog1.DialogTitle = "Select a Photo:"
CommonDialog1.Filter = "jpeg files|*.jpg|Bitmap files|*.bmp|All Files|*.*|"
CommonDialog1.ShowOpen


frmMain.pctMain.Picture = LoadPicture(CommonDialog1.FileName)

ErrHandler:
    Exit Sub


End Sub

Private Sub cmdReset_Click()

frmMain.pctMain.Picture = LoadPicture(App.Path & "/Test.jpg")

End Sub

Private Sub Form_Load()

chkDefault.Value = vbUnchecked
chkMDis.Value = vbUnchecked
lblInfo.Caption = ""
cmdNewPhoto.Visible = False
cmdReset.Visible = False


lblInfo.Caption = "You may only use .bmp, or .jpg files.  " & vbCrLf & _
                  "Max Dimensions are Width 676 X Hieght 556." & vbCrLf & _
                  "Scale is Pixels ."


End Sub

Private Sub Form_Unload(Cancel As Integer)

chkMDis.Value = vbUnchecked

Set frmOptions = Nothing

End Sub

