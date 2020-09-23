VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEpoch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As String) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const BS_LEFT As Long = &H100
Private Const BS_RIGHT As Long = &H200
Private Const BS_CENTER As Long = &H300
Private Const BS_TOP As Long = &H400
Private Const BS_BOTTOM As Long = &H800
Private Const BS_VCENTER As Long = &HC00

Private Const BS_ALLSTYLES = BS_LEFT Or BS_RIGHT Or BS_CENTER Or BS_TOP Or BS_BOTTOM Or BS_VCENTER

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Const GWL_STYLE = (-16)

Private Const ES_LOWERCASE As Long = &H10&
Private Const ES_NUMBER As Long = &H2000&
Private Const ES_UPPERCASE As Long = &H8&

'-------------------------------------------------------------------------------------------------
' ComboBox constants
'-------------------------------------------------------------------------------------------------
Private Const CB_FINDSTRING = &H14C
Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_FINDSTRINGEXACT = &H158

Public Enum bsHorizontalAlignments
    bsLeft = BS_LEFT
    bsRight = BS_RIGHT
    bsCenter = BS_CENTER
End Enum

Public Enum bsVerticalAlignments
    bsTop = BS_TOP
    bsBottom = BS_BOTTOM
    bsVCenter = BS_VCENTER
End Enum

Public Enum esTextBoxStyle
    esNumber = ES_NUMBER
    esSmallCaps = ES_LOWERCASE
    esUpCaps = ES_UPPERCASE
End Enum

Public Sub DragForm(f As Object)
Dim lRet As Long
    Call ReleaseCapture
    lRet = SendMessage(f.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Public Function WinCBFindString(ByVal pHwnd As Long, ByVal pSearch As String, ByVal flExact As Boolean) As Long
Dim llngMess As Long
    If flExact Then
        llngMess = CB_FINDSTRINGEXACT
    Else
        llngMess = CB_FINDSTRING
    End If
    WinCBFindString = SendMessageStr(pHwnd, llngMess, -1, ByVal pSearch)
End Function

Public Function WinCBShowDropDown(ByVal pHwnd As Long) As Boolean
Dim llngRet As Long
    llngRet = SendMessageStr(pHwnd, CB_SHOWDROPDOWN, 1, vbNull)
    If llngRet = -1 Then
        WinCBShowDropDown = False
    Else
        WinCBShowDropDown = True
    End If
End Function

Public Sub FormatTextBox(objText As Object, objTextFormat As esTextBoxStyle)
    Dim curstyle As Long, newstyle As Long
    'retrieve the window style
    curstyle = GetWindowLong(objText.hwnd, GWL_STYLE)
    If objTextFormat = esNumber Then
        curstyle = curstyle Or ES_NUMBER
    ElseIf objTextFormat = esSmallCaps Then
        curstyle = curstyle Or ES_LOWERCASE
    ElseIf objTextFormat = esUpCaps Then
        curstyle = curstyle Or ES_UPPERCASE
    End If
    'Set the new style
    newstyle = SetWindowLong(objText.hwnd, GWL_STYLE, curstyle)
    'refresh
    objText.Refresh
End Sub

Public Sub AlignButtonText(cmd As Object, Optional ByVal HStyle As bsHorizontalAlignments = bsCenter, Optional ByVal VStyle As bsVerticalAlignments = bsVCenter)
Dim oldStyle As Long
    ' retrieve the current style of the control
    oldStyle = GetWindowLong(cmd.hwnd, GWL_STYLE)
    ' change the style
    oldStyle = oldStyle And (Not BS_ALLSTYLES)
    ' set the style of the control to the new style
    Call SetWindowLong(cmd.hwnd, GWL_STYLE, _
    oldStyle Or HStyle Or VStyle)
    cmd.Refresh
End Sub

Public Function Encrypt_Text(s As String, OffsetValue As Integer) As String
Dim i As Integer
Dim temp As String
Dim sVal As Integer
    temp = ""
    sVal = 0
    For i = 1 To Len(s)
        sVal = (Asc(Mid(s, i, 1)) + OffsetValue)
        sVal = sVal + 25
        temp = temp & Chr(sVal)
    Next
    Encrypt_Text = temp
End Function

Public Function Decrypt_Text(s As String, OffsetValue As Integer) As String
Dim i As Integer
Dim temp As String
Dim sVal As Integer
    temp = ""
    sVal = 0
    For i = 1 To Len(s)
        sVal = (Asc(Mid(s, i, 1)) - OffsetValue)
        sVal = sVal - 25
        temp = temp & Chr(sVal)
    Next
    Decrypt_Text = temp
End Function



