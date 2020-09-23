Attribute VB_Name = "ModBrowse"
'This module is the standard Browse for Folder with a few changes
'to the BrowseCallbackProc function in order to locate the position
'of the buttons so our button/checkbox will line up
Option Explicit
Private Const BIF_STATUSTEXT = &H4&
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260
Private Const WM_USER = &H400
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELCHANGED = 2
Private Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Private Const BFFM_SETSELECTION = (WM_USER + 102)
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Const GW_NEXT = 2
Const GW_CHILD = 5
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Type BrowseInfo
  hwndOwner      As Long
  pIDLRoot       As Long
  pszDisplayName As Long
  lpszTitle      As Long
  ulFlags        As Long
  lpfnCallback   As Long
  lParam         As Long
  iImage         As Long
End Type
Private m_CurrentDirectory As String
'Public variables to broadcast info to other mods
Public LetsRecurse As Boolean 'User wants recursion
Public BFhwnd As Long 'handle of the dialog
Public butTop As Integer 'Position of the top of the dialogs' buttons
Public Arestart As Boolean 'We made a new folder and are relaunching the dialog
Dim R As RECT, Bt As RECT
Public Function BrowseForFolder(StartDir As String, owner As Long, Title As String) As String
'Standard call for the dialog
  Dim lpIDList As Long
  Dim szTitle As String
  Dim sBuffer As String
  Dim tBrowseInfo As BrowseInfo
  m_CurrentDirectory = StartDir & vbNullChar
  szTitle = Title
  With tBrowseInfo
    .hwndOwner = owner
    .lpszTitle = lstrcat(szTitle, "")
    .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN + BIF_STATUSTEXT
    .lpfnCallback = GetAddressofFunction(AddressOf BrowseCallbackProc)
  End With
  lpIDList = SHBrowseForFolder(tBrowseInfo)
  If (lpIDList) Then
    sBuffer = Space(MAX_PATH)
    SHGetPathFromIDList lpIDList, sBuffer
    sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    BrowseForFolder = sBuffer
  Else
    BrowseForFolder = ""
  End If
End Function
Private Function BrowseCallbackProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
Dim lpIDList As Long
Dim Ret As Long
Dim hwnda As Long, ClWind As String * 7
Dim sBuffer As String
On Error Resume Next
BFhwnd = hwnd
Select Case uMsg
  Case BFFM_INITIALIZED
    'If we are relaunching the dialog we want to use the existing
    'RECT to position it otherwise get new values
    If Not Arestart Then getBFSizePos hwnd
    Call MoveWindow(hwnd, R.Left, R.Top, R.Right - R.Left, R.Bottom - R.Top, True)
    Call SendMessage(hwnd, BFFM_SETSELECTION, 1, m_CurrentDirectory)
        'Lets go through all the dialogs' children till we find a Button
        hwnda = GetWindow(hwnd, GW_CHILD)
        Do While hwnda <> 0
            GetClassName hwnda, ClWind, 7
            If Left(ClWind, 6) = "Button" Then
                Call GetWindowRect(hwnda, Bt)
                butTop = Bt.Top - R.Top
                Exit Do
            End If
            hwnda = GetWindow(hwnda, GW_NEXT)
        Loop
  
  Case BFFM_SELCHANGED
    'Make the status text show the selected folder
    sBuffer = Space(MAX_PATH)
    Ret = SHGetPathFromIDList(lp, sBuffer)
    If Ret = 1 Then
      Call SendMessage(hwnd, BFFM_SETSTATUSTEXT, 0, sBuffer)
      m_CurrentDirectory = sBuffer
    End If
End Select
BrowseCallbackProc = 0
End Function
Private Function GetAddressofFunction(add As Long) As Long
  GetAddressofFunction = add
End Function
Public Sub getBFSizePos(hwnd As Long)
'Where's the window ?
 Call GetWindowRect(hwnd, R)
End Sub
