Attribute VB_Name = "MorphSupport"
Option Explicit


'
'  constant and sub/function declarations
'
Declare Function SetParent Lib "user32" _
    (ByVal hwndChild As Long, ByVal hWndNewParent As Long) As Long

'  Get/Set WindowLong stuff
Public Const GWL_WNDPROC = (-4)
Public Const GWL_HINSTANCE = (-6)
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_ID = (-12)
Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)
Public Const GWL_USERDATA = (-21)

Public Const WS_POPUP = &H80000000
Public Const WS_CHILD = &H40000000

'  window stuff
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'  Get/Set Prop stuff
Declare Function GetProp Lib "user32" Alias "GetPropA" _
    (ByVal hwnd As Long, ByVal lpString As String) As Long
Declare Function SetProp Lib "user32" Alias "SetPropA" _
    (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long


Private Const SUBFORMPROP = "Subform"
Private Const SUBFORMMARK = &H614149

Public Sub MorphForm(frmParent As Form, frmChild As Form, ctlParent As Control)

    Debug.Assert IsObject(frmParent)
    Debug.Assert TypeOf frmParent Is Form
    
    Debug.Assert IsObject(frmChild)
    Debug.Assert TypeOf frmChild Is Form
    
    Debug.Assert Not IsSubform(frmChild)
    
    '  get child form's style
    Dim styleChild As Long
    styleChild = GetWindowLong(frmChild.hwnd, GWL_STYLE)
    
    '  insure it's not a popup
    styleChild = styleChild And Not WS_POPUP
    
    '  morph it into a child
    styleChild = styleChild Or WS_CHILD
    
    '  well, now *actually* morph
    SetWindowLong frmChild.hwnd, GWL_STYLE, styleChild
    
    '  set parent to us
    SetParent frmChild.hwnd, ctlParent.hwnd
    
    'Position the form in the main form
    frmChild.Top = 0
    frmChild.Height = ctlParent.Height
    frmChild.Left = 0
    frmChild.Width = ctlParent.Width


End Sub

'
'  private helpers
'
Private Function IsSubform(frmTest As Form) As Boolean

    Debug.Assert IsObject(frmTest)
    Debug.Assert TypeOf frmTest Is Form
    
    IsSubform = IsSubformWnd(frmTest.hwnd)
    
End Function

Private Function IsSubformWnd(ByVal hwnd As Long) As Boolean

    Debug.Assert 0 <> hwnd
    
    If SUBFORMMARK = GetProp(hwnd, SUBFORMPROP) Then
        IsSubformWnd = True
    Else
        IsSubformWnd = False
    End If
    
End Function

