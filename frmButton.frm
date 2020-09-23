VERSION 5.00
Begin VB.Form frmButton 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1125
   LinkTopic       =   "Form2"
   ScaleHeight     =   360
   ScaleWidth      =   1125
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check1 
      Caption         =   "    INCLUDE SUBFOLDERS"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1440
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create"
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1125
   End
End
Attribute VB_Name = "frmButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form is a workaround to insert a button/checkbox
'on the Browse for Folder window.

'Win 2K compliant FileExists
Private Const INVALID_HANDLE_VALUE = -1
Private Const MAX_PATH = 260
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Declare Sub keybd_event Lib "User32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwnd As Long) As Long
Const KEYEVENTF_KEYUP = &H2
Dim R As RECT
Dim owner As Long
Dim Initpath As String
Dim Newboy As Boolean
Dim Message As String
Private Function FileExists(sSource As String) As Boolean
If Right(sSource, 2) = ":\" Then
    Dim allDrives As String
    allDrives = Space$(64)
    Call GetLogicalDriveStrings(Len(allDrives), allDrives)
    FileExists = InStr(1, allDrives, Left(sSource, 1), 1) > 0
    Exit Function
Else
    If Not sSource = "" Then
        Dim WFD As WIN32_FIND_DATA
        Dim hFile As Long
        hFile = FindFirstFile(sSource, WFD)
        FileExists = hFile <> INVALID_HANDLE_VALUE
        Call FindClose(hFile)
    Else
        FileExists = False
    End If
End If
End Function

Private Sub Check1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
LockWindowUpdate BFhwnd
Timer1.Enabled = False
'stop the timer so this form can temporarily have focus
End Sub

Private Sub Check1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Call SetWindowPos(BFhwnd, Me.hwnd, R.Left, R.Top, R.Right - R.Left, R.Bottom - R.Top, wFlags)
Timer1.Enabled = True
End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Timer1.Enabled = False
'stop the timer so this form can temporarily have focus
'the timer gets restarted by the GoNew sub
End Sub
Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
GoNew
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'return the check value to the UserControl
LetsRecurse = Check1.Value
End Sub

Private Sub Timer1_Timer()
'Keeps this form positioned over the Browse For Folder dialog
Dim z As Long
LockWindowUpdate 0
z = GetWindowRect(BFhwnd, R)
If z <> 0 Then
    Me.Visible = True
    Me.Left = (R.Left + 16) * Screen.TwipsPerPixelX
    'the butTop value was established by  the 'BrowseCallbackProc' function
    Me.Top = (R.Top + butTop) * Screen.TwipsPerPixelY
    'position this form above the dialog but not above anything else
    'in case a window is opened on top of the dialog
    Call SetWindowPos(BFhwnd, Me.hwnd, R.Left, R.Top, R.Right - R.Left, R.Bottom - R.Top, wFlags)
End If
End Sub
Public Function Browse(ownerform As Long, Create As Boolean, Recurse As Boolean, Optional iniitdir As String, Optional YourMessage As String) As String
Dim temppath As String
Message = YourMessage
If Message = "" Then Message = "Select a Folder"
Initpath = iniitdir
If Initpath = "" Then Initpath = "c:\"
owner = ownerform
If Create Or Recurse Then
againplease:
    Timer1.Enabled = True
    Command1.Visible = Create
    Check1.Visible = Recurse
    temppath = BrowseForFolder(Initpath, owner, Message)
    If Newboy = True Then
    'A new folder needs to be created
        If FileExists(temppath + "\" + Initpath) Then
            Me.Visible = False
            MsgBox "A Folder of that name already exists." + vbCrLf + "Please enter a different name", vbExclamation, "Bobo Enterprises Folder Browser"
            Initpath = temppath
            Newboy = False
            'If it exists then relaunch dialog at the
            'last selected folder
            GoTo againplease
        End If
        Initpath = temppath + "\" + Initpath
        MkDir Initpath
        Newboy = False
        GoTo againplease
    End If
    Browse = StripTerminator(temppath)
    Unload Me
Else
    'Just the standard Browse dialog required so launch it and bail out now
    Browse = StripTerminator(BrowseForFolder(Initpath, owner, Message))
    Unload Me
End If
End Function
Private Sub GoNew()
LockWindowUpdate BFhwnd
Initpath = InputBox("Enter a name for your new folder", "Create Folder")
If Initpath = "" Then
    Call SetWindowPos(BFhwnd, Me.hwnd, R.Left, R.Top, R.Right - R.Left, R.Bottom - R.Top, wFlags)
    Timer1.Enabled = True
    Exit Sub
End If
Call SetWindowPos(BFhwnd, Me.hwnd, R.Left, R.Top, R.Right - R.Left, R.Bottom - R.Top, wFlags)
Timer1.Enabled = True
'establish the current location of the dialog so we can
'launch the new dialog in the same location
getBFSizePos BFhwnd
Arestart = True
'Simulate the 'OK' button being pressed in order to
'close the current dialog retrieving selected path,
'so we can make a new folder and then relaunch the dialog
'with the new folder as the selected initial directory
keybd_event vbKeyReturn, 0, 0, 0
keybd_event vbKeyReturn, 0, KEYEVENTF_KEYUP, 0
Newboy = True
End Sub
Private Function StripTerminator(ByVal strString As String) As String
'gets rid of any null characters at the end of the returned path
    Dim intZeroPos As Integer

    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

