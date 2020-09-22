Attribute VB_Name = "MainSupport"
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


Option Explicit

Public mNode As MSComctlLib.Node
Public info As New devGetData

Public Const sTVW_ADDRESSES = "A"
Public Const sTVW_NEWADDRESSES = "N"
Public Const sTVW_ADDADDRESS = "C"
Public Const sTVW_MAIN = "M"
Public Const sTVW_SEARCH = "S"
Public Const sTVW_URLS = "U"

Public Sql As String
Public FName As String
Public LName As String
Public ID As String

Sub Main()

    Set info = New devGetData
    info.CnConnect.Open

    Dim Waittime As Variant

    frmSplash.Show
    frmSplash.Refresh

    Waittime = Now
    While Now < Waittime + TimeValue("00:00:2")       ' Wait 2 seconds
    Wend
    
    Load frmMain
    Unload frmSplash
    frmMain.Show

    

End Sub
Public Sub Send_Email_To(EmailAddress As String)
 
Dim jmpEmail

On Error Resume Next
'check if the email address is valid
If Valid_Email_Address(EmailAddress) Then
   jmpEmail = Shell("start.exe mailto:" & EmailAddress, vbHide)
Else
    MsgBox "Sorry, but [" & EmailAddress & "] is not a valid email address" _
            , vbExclamation + vbOKOnly, "Invalid Email Address"
End If
  
End Sub
Public Function Valid_Email_Address(EmailAddress As String) As Boolean

Valid_Email_Address = EmailAddress Like "*@[A-Z,a-z,0-9]*.*"
    
End Function

Public Sub To_html()



End Sub

Public Sub Display()

info.rsContactData.Open
info.rsContactData.MoveFirst
info.rsContactData.Find "ID =" & ID

frmMain.ShowFormAsChild frmViewAddress
frmViewAddress.SetFocus

On Error GoTo CleanUp

frmViewAddress.txtFName(0).Text = info.rsContactData.Fields("FirstName")
frmViewAddress.txtLName(2).Text = info.rsContactData.Fields("LastName")
frmViewAddress.txtMName(1).Text = info.rsContactData.Fields("MiddleName")
frmViewAddress.cboSuffix(0).Text = info.rsContactData.Fields("Suffix")
frmViewAddress.txtHPhone(1).Text = info.rsContactData.Fields("HomePhone")
frmViewAddress.txtWPhone(0).Text = info.rsContactData.Fields("WorkPhone")
frmViewAddress.txtPager(0).Text = info.rsContactData.Fields("Pager")
frmViewAddress.txtMobile(2).Text = info.rsContactData.Fields("Mobile")
frmViewAddress.txtAMobile(2).Text = info.rsContactData.Fields("SecMobile")
frmViewAddress.txtOPhone(1).Text = info.rsContactData.Fields("OtherPhone")
frmViewAddress.txtStreet(0).Text = info.rsContactData.Fields("HomeStreet")
frmViewAddress.txtCity(1).Text = info.rsContactData.Fields("HomeCity")
frmViewAddress.cboState(1).Text = info.rsContactData.Fields("HomeState")
frmViewAddress.txtZip(1).Text = info.rsContactData.Fields("HomeZip")
frmViewAddress.txtStreet(1).Text = info.rsContactData.Fields("WorkStreet")
frmViewAddress.txtCity(0).Text = info.rsContactData.Fields("WorkCity")
frmViewAddress.cboState(0).Text = info.rsContactData.Fields("WorkState")
frmViewAddress.txtZip(0).Text = info.rsContactData.Fields("WorkZip")
frmViewAddress.txtHEmail(2).Text = info.rsContactData.Fields("HomeEmail")
frmViewAddress.txtWEmail(3).Text = info.rsContactData.Fields("WorkEmail")
frmViewAddress.txtSecEmail(4).Text = info.rsContactData.Fields("SecEmail")
  
info.rsContactData.Close

CleanUp:
    Resume Next

End Sub

Public Sub AddEntry()

Dim i As Integer
info.rsContactData.Open
info.rsContactData.AddNew

If frmViewAddress.txtFName(0).Text <> "" And frmViewAddress.txtLName(2).Text <> "" Then
    
    info.rsContactData.Fields("FirstName") = frmViewAddress.txtFName(0).Text
    info.rsContactData.Fields("LastName") = frmViewAddress.txtLName(2).Text
    info.rsContactData.Fields("MiddleName") = frmViewAddress.txtMName(1).Text
    info.rsContactData.Fields("Suffix") = frmViewAddress.cboSuffix(0).Text
    info.rsContactData.Fields("HomePhone") = frmViewAddress.txtHPhone(1).Text
    info.rsContactData.Fields("WorkPhone") = frmViewAddress.txtWPhone(0).Text
    info.rsContactData.Fields("Pager") = frmViewAddress.txtPager(0).Text
    info.rsContactData.Fields("Mobile") = frmViewAddress.txtMobile(2).Text
    info.rsContactData.Fields("SecMobile") = frmViewAddress.txtAMobile(2).Text
    info.rsContactData.Fields("OtherPhone") = frmViewAddress.txtOPhone(1).Text
    info.rsContactData.Fields("HomeStreet") = frmViewAddress.txtStreet(0).Text
    info.rsContactData.Fields("HomeCity") = frmViewAddress.txtCity(1).Text
    info.rsContactData.Fields("HomeState") = frmViewAddress.cboState(1).Text
    info.rsContactData.Fields("HomeZip") = frmViewAddress.txtZip(1).Text
    info.rsContactData.Fields("WorkStreet") = frmViewAddress.txtStreet(1).Text
    info.rsContactData.Fields("WorkCity") = frmViewAddress.txtCity(0).Text
    info.rsContactData.Fields("WorkState") = frmViewAddress.cboState(0).Text
    info.rsContactData.Fields("WorkZip") = frmViewAddress.txtZip(0).Text
    info.rsContactData.Fields("HomeEmail") = frmViewAddress.txtHEmail(2).Text
    info.rsContactData.Fields("WorkEmail") = frmViewAddress.txtWEmail(3).Text
    info.rsContactData.Fields("SecEmail") = frmViewAddress.txtSecEmail(4).Text
    
    
    info.rsContactData.Update
    info.rsContactData.AddNew
    frmMain.tvwMain.Nodes.Clear
        
    Call frmViewAddress.Clear
    
Else:

    MsgBox "You must have a First Name And Last Name to ADD a contact", vbInformation, "Missing Information"
    
End If

    Set info = Nothing
    Call frmMain.InitTree
End Sub

Public Sub Delete()

Dim i As Integer
Dim result As Integer

info.rsContactData.Open

result = MsgBox("Are you sure you want to Delete" & " " & FName & " " & LName & " " & "?", vbYesNo, "Delete Record!")

If result = vbYes Then

    info.rsContactData.Find "ID =" & ID
    info.rsContactData.Delete adAffectCurrent
    frmMain.tvwMain.Nodes.Clear
        
    Call frmViewAddress.Clear

    Set info = Nothing
    Call frmMain.InitTree
    
Else

    Set info = Nothing
    Exit Sub
    
End If

End Sub

Public Sub Edit()

info.rsContactData.Open
info.rsContactData.Find "ID =" & ID

If frmViewAddress.txtFName(0).Text <> "" And frmViewAddress.txtLName(2).Text <> "" Then

    info.rsContactData.Fields("FirstName") = frmViewAddress.txtFName(0).Text
    info.rsContactData.Fields("LastName") = frmViewAddress.txtLName(2).Text
    info.rsContactData.Fields("MiddleName") = frmViewAddress.txtMName(1).Text
    info.rsContactData.Fields("Suffix") = frmViewAddress.cboSuffix(0).Text
    info.rsContactData.Fields("HomePhone") = frmViewAddress.txtHPhone(1).Text
    info.rsContactData.Fields("WorkPhone") = frmViewAddress.txtWPhone(0).Text
    info.rsContactData.Fields("Pager") = frmViewAddress.txtPager(0).Text
    info.rsContactData.Fields("Mobile") = frmViewAddress.txtMobile(2).Text
    info.rsContactData.Fields("SecMobile") = frmViewAddress.txtAMobile(2).Text
    info.rsContactData.Fields("OtherPhone") = frmViewAddress.txtOPhone(1).Text
    info.rsContactData.Fields("HomeStreet") = frmViewAddress.txtStreet(0).Text
    info.rsContactData.Fields("HomeCity") = frmViewAddress.txtCity(1).Text
    info.rsContactData.Fields("HomeState") = frmViewAddress.cboState(1).Text
    info.rsContactData.Fields("HomeZip") = frmViewAddress.txtZip(1).Text
    info.rsContactData.Fields("WorkStreet") = frmViewAddress.txtStreet(1).Text
    info.rsContactData.Fields("WorkCity") = frmViewAddress.txtCity(0).Text
    info.rsContactData.Fields("WorkState") = frmViewAddress.cboState(0).Text
    info.rsContactData.Fields("WorkZip") = frmViewAddress.txtZip(0).Text
    info.rsContactData.Fields("HomeEmail") = frmViewAddress.txtHEmail(2).Text
    info.rsContactData.Fields("WorkEmail") = frmViewAddress.txtWEmail(3).Text
    info.rsContactData.Fields("SecEmail") = frmViewAddress.txtSecEmail(4).Text

    info.rsContactData.Update
    info.rsContactData.Close
End If

End Sub

Public Sub Search()

Dim FirstName As String
Dim LastName As String

info.rsContactData.Open

If frmSearch.optFirstName.Value = True Then

    Do While info.rsContactData.EOF = False
        frmSearch.ListView1.ListItems.Add 1, FirstName, info.rsContactData.Fields("FirstName")
        frmSearch.ListView1.ListItems.Add , LastName, info.rsContactData.Fields("LastName")
        info.rsContactData.MoveNext
    Loop

End If

info.rsContactData.Close

End Sub

