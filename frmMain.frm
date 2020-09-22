VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7080
   ClientLeft      =   348
   ClientTop       =   1824
   ClientWidth     =   11244
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   11244
   Begin MSComctlLib.ImageList imgTool 
      Left            =   9240
      Top             =   240
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":307E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":31DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3336
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3492
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5C46
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":60A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":663E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":679A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":68F6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   9720
      Top             =   240
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":74CA
            Key             =   ""
            Object.Tag             =   "&Delete"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7A66
            Key             =   ""
            Object.Tag             =   "Pre&frences"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8342
            Key             =   ""
            Object.Tag             =   "&Print"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":849E
            Key             =   ""
            Object.Tag             =   "&Update"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":85FA
            Key             =   ""
            Object.Tag             =   "&Exit"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8A56
            Key             =   ""
            Object.Tag             =   "A&dd"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":962A
            Key             =   ""
            Object.Tag             =   "&Favorite URLS"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BDDE
            Key             =   ""
            Object.Tag             =   "&Main"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BF3A
            Key             =   ""
            Object.Tag             =   "&About"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E6EE
            Key             =   ""
            Object.Tag             =   "&Search"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E84A
            Key             =   ""
            Object.Tag             =   "&Save As html"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E9A6
            Key             =   ""
            Object.Tag             =   "Co&py"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EB02
            Key             =   ""
            Object.Tag             =   "&Cut"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EC5E
            Key             =   ""
            Object.Tag             =   "Pas&te"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   312
      Left            =   0
      TabIndex        =   3
      Top             =   6768
      Width           =   11244
      _ExtentX        =   19833
      _ExtentY        =   550
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "8/28/00"
            Object.ToolTipText     =   "Date"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "2:43 PM"
            Object.ToolTipText     =   "Time"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Tag             =   "Option"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   732
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11244
      _ExtentX        =   19833
      _ExtentY        =   1291
      ButtonWidth     =   1058
      ButtonHeight    =   1164
      Appearance      =   1
      ImageList       =   "imgTool"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Main"
            Key             =   "Main"
            Object.ToolTipText     =   "Main"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sep"
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add"
            Key             =   "Add"
            Object.ToolTipText     =   "Add Contact"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
            Key             =   "Edit"
            Object.ToolTipText     =   "Edit Contact"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete Contact"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            Key             =   "Search"
            Object.ToolTipText     =   "Search For Contact"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "URL'S"
            Key             =   "URL"
            Object.ToolTipText     =   "List of URL'S"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "sep1"
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            Key             =   "About"
            Object.ToolTipText     =   "Help"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwMain 
      Height          =   5532
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   960
      Width           =   2772
      _ExtentX        =   4890
      _ExtentY        =   9758
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      Style           =   1
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bookman Old Style"
         Size            =   9.6
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox pctMain 
      Height          =   5532
      Left            =   3000
      Picture         =   "frmMain.frx":EDBA
      ScaleHeight     =   3.808
      ScaleMode       =   5  'Inch
      ScaleWidth      =   5.642
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   960
      Width           =   8172
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuHTML 
         Caption         =   "&Save As html"
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Begin VB.Menu mnuPrintFull 
            Caption         =   "Contact list"
         End
         Begin VB.Menu mnuPrintOne 
            Caption         =   "Current Contact"
         End
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "&Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Co&py"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Pas&te"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuMain 
         Caption         =   "&Main"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUrls 
         Caption         =   "&Favorite URLS"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuPrefrences 
         Caption         =   "Pre&frences"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuAdd 
         Caption         =   "A&dd"
      End
      Begin VB.Menu mnuUpdate 
         Caption         =   "&Update"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "&Search"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************'
' Program Written by:     David L. Stevens
' Date Written:           12-Jun-00 11:12:28
' Program Developed for:  vbsolutionsnow.com
'**************************************************************************************'
' Purpose:
'           To help maintain contact information, and the ability to use the info
'            easier.
'
'
'
' Assumptions:
'               Windows 95/98, This has not been tested very well yet.
'
'
'
' Modifications:
' Date:       InItials:   Purpose:'
'
'
'***************************************************************************************'
' CopyRight:
'               David L. Stevens vbsolutionsnow.com  also contained in the Help form
'
'***************************************************************************************'

'The icon in the menu code was taken and modified from planet source code.  The person who wrote
'it was a contest winner here.  I can not find his name.  This is nice code.


Option Explicit

Private mfrmChild As Form

Public Sub ShowFormAsChild(frmChild As Form)

'Morph the form to the main form
 MorphForm Me, frmChild, pctMain

'Show the form
frmChild.Show
Set mfrmChild = frmChild

End Sub

Private Sub Form_Load()

'Used to center the form in the screen
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2

'to do Make the staus bar more intuative
StatusBar1.Panels(3).Text = "Status: OK"

'declared in Public module
'set the menu icon visibility
'*Comment out the next two lines in desing time, or it will crash*'
Set CoolMenuObj = New CoolMenu
Call CoolMenuObj.Install(Me.hwnd, ImageList, True, True)

InitTree

frmMain.Caption = "AddressBook Organizer"

End Sub

Private Sub Form_Resize()

pctMain_Resize

End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim frm As Form

'*Comment out the next two lines in desing time, or it will crash*'
Call CoolMenuObj.Install(0&)
  
Set CoolMenuObj = Nothing

Set info = Nothing
    
For Each frm In Forms
    Unload frm
Next frm
   
Set frmMain = Nothing

End Sub

Private Sub mnuAbout_Click()

frmMain.ShowFormAsChild frmAbout

End Sub

Private Sub mnuExit_Click()

Set info = Nothing

Dim frm As Form

For Each frm In Forms
    Unload frm
Next frm
   
Set frmMain = Nothing

End

End Sub

Private Sub mnuHTML_Click()

frmMain.ShowFormAsChild frmExport

End Sub

Private Sub mnuMain_Click()

    If Not mfrmChild Is Nothing Then
       Unload mfrmChild
       Set mfrmChild = Nothing
    End If

    frmMain.Show

End Sub

Private Sub mnuPrefrences_Click()

frmMain.ShowFormAsChild frmOptions

End Sub

Private Sub mnuSearch_Click()

    frmMain.ShowFormAsChild frmSearch

End Sub

Private Sub mnuUpdate_Click()

    frmMain.ShowFormAsChild frmViewAddress

End Sub

Private Sub mnuUrls_Click()

    frmMain.ShowFormAsChild frmUrl

End Sub

Private Sub pctMain_Resize()

If Not mfrmChild Is Nothing Then
   mfrmChild.Top = 0
   mfrmChild.Height = pctMain.Height - Screen.TwipsPerPixelY * 4
   mfrmChild.Left = 0
   mfrmChild.Width = pctMain.Width - Screen.TwipsPerPixelX * 4
End If
   
End Sub

Public Sub InitTree()

Dim i As Integer

info.rsContactData.Open
info.rsContactData.MoveFirst

'load the main keys
tvwMain.Nodes.Add , , sTVW_MAIN & "1", "Main Display"
tvwMain.Nodes.Add , , sTVW_SEARCH & "1", "Search"
tvwMain.Nodes.Add , , sTVW_URLS & "1", "Favorite URL's"
tvwMain.Nodes.Add , , sTVW_ADDRESSES & "1", "Current Addresses"
   
Do While info.rsContactData.EOF = False
   For i = 1 To info.rsContactData.RecordCount
       tvwMain.Nodes.Add sTVW_ADDRESSES & "1", tvwChild, sTVW_NEWADDRESSES & i, info.rsContactData.Fields("FirstName") & " " & info.rsContactData.Fields("LastName") & " " & info.rsContactData.Fields("ID")
       info.rsContactData.MoveNext
   Next i
Loop
   
   
info.rsContactData.Close
tvwMain.Nodes.Item(sTVW_ADDRESSES & "1").Expanded = True

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key

Case "Exit"
        mnuExit_Click
  
Case "Print"
    
    
Case "Main"
    If Not mfrmChild Is Nothing Then
       Unload mfrmChild
       Set mfrmChild = Nothing
    End If
            
    frmMain.Show
    
Case "Add"

   frmViewAddress.cmdAddEntry.Visible = True
   frmViewAddress.cmdDone.Visible = True
   frmViewAddress.Clear
   frmMain.ShowFormAsChild frmViewAddress
    
Case "Edit"

    Call Edit

Case "Delete"

    Call Delete
    
Case "Search"
        
    frmMain.ShowFormAsChild frmSearch
    
Case "URL"

    frmMain.ShowFormAsChild frmUrl
    
Case "About"

     frmMain.ShowFormAsChild frmAbout
    
End Select

End Sub

Public Sub tvwMain_NodeClick(ByVal Node As MSComctlLib.Node)

Set mNode = Node
  
HandleNodeEvents Node
    
End Sub

Public Sub HandleNodeEvents(CurrentNode As Node)

Dim space As String
Dim Space2 As String
Dim S_String As String
Dim l As Integer

S_String = " "

    Select Case Left$(CurrentNode.Key, 1)
        Case sTVW_MAIN
            If Not mfrmChild Is Nothing Then
                Unload mfrmChild
                Set mfrmChild = Nothing
            End If
            
            frmMain.Show
       
            
'        Case sTVW_ADDADDRESS
'            If Not mfrmChild Is Nothing Then
'                Unload mfrmChild
'                Set mfrmChild = Nothing
'            End If
'            frmMain.ShowFormAsChild frmAddAddress
'            frmAddAddress.SetFocus

        Case sTVW_SEARCH
            If Not mfrmChild Is Nothing Then
                Unload mfrmChild
                Set mfrmChild = Nothing
            End If
            frmMain.ShowFormAsChild frmSearch
            frmSearch.SetFocus
            
        Case sTVW_URLS
            If Not mfrmChild Is Nothing Then
                Unload mfrmChild
                Set mfrmChild = Nothing
            End If
            frmMain.ShowFormAsChild frmUrl
            frmUrl.SetFocus
            
         Case sTVW_NEWADDRESSES
            If Not mfrmChild Is Nothing Then
                Unload mfrmChild
                Set mfrmChild = Nothing
            End If
               space = InStr(1, CurrentNode.Text, S_String)
               l = Len(CurrentNode.Text)
               FName = Mid(CurrentNode.Text, 1, space - 1)
               
               Space2 = InStr(space + 1, CurrentNode.Text, S_String)
               ID = Mid(CurrentNode.Text, Space2 + 1, 2)
               LName = Mid(CurrentNode.Text, space + 1, l - Len(FName) - Len(ID) - 2)

               Display
            
   End Select
   
End Sub

