Attribute VB_Name = "CmnDialog"
Option Explicit
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
'API to manipulate windows and draw pics
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
'Standard API for commondialog
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Const ScrCopy = &HCC0020
Const OFN_ALLOWMULTISELECT As Long = &H200
Const OFN_CREATEPROMPT As Long = &H2000
Const OFN_ENABLEHOOK As Long = &H20
Const OFN_ENABLETEMPLATE As Long = &H40
Const OFN_ENABLETEMPLATEHANDLE As Long = &H80
Const OFN_EXPLORER As Long = &H80000
Const OFN_EXTENSIONDIFFERENT As Long = &H400
Const OFN_FILEMUSTEXIST As Long = &H1000
Const OFN_HIDEREADONLY As Long = &H4
Const OFN_LONGNAMES As Long = &H200000
Const OFN_NOCHANGEDIR As Long = &H8
Const OFN_NODEREFERENCELINKS As Long = &H100000
Const OFN_NOLONGNAMES As Long = &H40000
Const OFN_NONETWORKBUTTON As Long = &H20000
Const OFN_NOREADONLYRETURN As Long = &H8000&
Const OFN_NOTESTFILECREATE As Long = &H10000
Const OFN_NOVALIDATE As Long = &H100
Const OFN_OVERWRITEPROMPT As Long = &H2
Const OFN_PATHMUSTEXIST As Long = &H800
Const OFN_READONLY As Long = &H1
Const OFN_SHAREAWARE As Long = &H4000
Const OFN_SHAREFALLTHROUGH As Long = 2
Const OFN_SHAREWARN As Long = 0
Const OFN_SHARENOWARN As Long = 1
Const OFN_SHOWHELP As Long = &H10
Const OFS_MAXPATHNAME As Long = 260
Const OFN_SELECTED As Long = &H78
Const WM_INITDIALOG = &H110
Const SW_SHOWNORMAL = 1
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Const GW_NEXT = 2
Const GW_CHILD = 5
Const WM_GETTEXT = &HD
Const WM_GETTEXTLENGTH = &HE
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Dim OFN As OPENFILENAME

'Variables
Dim cdlhwnd As Long 'commondialog handle
Public ThePic As PictureBox 'Hidden picturebox on UserControl
Public TheTimer As Timer 'Timer on UserControl
Public hwndParent As Long
Public mDC As Long
Public mCdlW As Long
Public mCdlH As Long
Public tmpFilename As String
Public tmpFullname As String
Dim ShowPics As Boolean
Public Function ShowOpen(hParent As Long, Optional mFilter As String, Optional mflags As Long, Optional mInitDir As String, Optional mTitle As String, Optional Pictures As Boolean) As String
'standard open file code
    If mInitDir = "" Then mInitDir = "c:\"
    If mFilter = "" Then mFilter = "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    If mTitle = "" Then mTitle = App.Title
    ShowPics = Pictures
    OFN.lStructSize = Len(OFN)
    OFN.hwndOwner = hParent
    OFN.hInstance = App.hInstance
    OFN.lpstrFilter = mFilter
    OFN.lpstrFile = Space$(254)
    OFN.nMaxFile = 255
    OFN.lpstrFileTitle = Space$(254)
    OFN.nMaxFileTitle = 255
    OFN.lpstrInitialDir = mInitDir
    OFN.lpstrTitle = mTitle
    OFN.Flags = mflags Or OFN_ENABLEHOOK Or OFN_EXPLORER
    OFN.lpfnHook = DummyProc(AddressOf CdlgHook)
    If GetOpenFileName(OFN) Then
        ShowOpen = Trim$(OFN.lpstrFile)
    Else
        ShowOpen = ""
    End If
End Function
Public Function ShowSave(hParent As Long, Optional mFilter As String, Optional mflags As Long, Optional mInitDir As String, Optional mTitle As String) As String
'standard save file code
    If mInitDir = "" Then mInitDir = "c:\"
    If mFilter = "" Then mFilter = "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    If mTitle = "" Then mTitle = App.Title
    ShowPics = False
    OFN.lStructSize = Len(OFN)
    OFN.hwndOwner = hParent
    OFN.hInstance = App.hInstance
    OFN.lpstrFilter = mFilter
    OFN.lpstrFile = Space$(254)
    OFN.nMaxFile = 255
    OFN.lpstrFileTitle = Space$(254)
    OFN.nMaxFileTitle = 255
    OFN.lpstrInitialDir = mInitDir
    OFN.lpstrTitle = mTitle
    OFN.Flags = mflags Or OFN_ENABLEHOOK Or OFN_EXPLORER
    OFN.lpfnHook = DummyProc(AddressOf CdlgHook)
    If GetSaveFileName(OFN) Then
        ShowSave = Trim$(OFN.lpstrFile)
    Else
        ShowSave = ""
    End If
End Function
Private Function CdlgHook(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim hwnda As Long, ClWind As String * 5
Dim Buffer As String, Ret As Long
Dim R As RECT
Dim NewCdlL As Long
Dim NewCdlT As Long
Dim scrWidth As Long
Dim scrHeight As Long
If ShowPics Then 'hook only used for Showing Pictures
    Select Case uMsg
        Case WM_INITDIALOG
            hwndParent = GetParent(hwnd)
            mDC = GetDC(hwndParent)
            cdlhwnd = hwndParent
            If hwndParent <> 0 Then
                Call GetWindowRect(hwndParent, R)
                mCdlW = R.Right - R.Left
                mCdlH = R.Bottom - R.Top
                'centre the dialog
                scrWidth = Screen.Width \ Screen.TwipsPerPixelX
                scrHeight = Screen.Height \ Screen.TwipsPerPixelY
                NewCdlL = (scrWidth - mCdlW) \ 2
                NewCdlT = (scrHeight - mCdlH) \ 2
                'Resize dialog to accomodate pictures
                Call MoveWindow(hwndParent, NewCdlL, NewCdlT, mCdlW, mCdlH + 107, True)
                CdlgHook = 1
            End If
        Case 78 'selection has changed
            hwndParent = GetParent(hwnd)
            hwnda = GetWindow(hwndParent, GW_CHILD)
            Do While hwnda <> 0
                GetClassName hwnda, ClWind, 5
                'Whats the filename ?
                If Left(ClWind, 4) = "Edit" Then
                    tmpFilename = gettext(hwnda)
                    Exit Do
                End If
                hwnda = GetWindow(hwnda, GW_NEXT)
            Loop
            If tmpFilename <> "" Then
                Buffer = Space(255)
                'Get a path
                Ret = GetFullPathName(tmpFilename, 255, Buffer, "")
                Buffer = Left(Buffer, Ret)
                If FileExists(Buffer) Then
                    tmpFullname = Buffer
                    'Empty the picturebox
                    ThePic.Picture = LoadPicture()
                    'Paint the dialog to clear old data
                    StretchBlt mDC, 20, 250, mCdlW - 40, 80, ThePic.hdc, 0, 0, ThePic.ScaleWidth, ThePic.ScaleHeight, ScrCopy
                    'load the new picture - The UserControls' timer will take care of painting
                    ThePic.Picture = LoadPicture(Buffer)
                End If
            End If
        'cancel or close was pressed so bailout at this point
        Case 2
            TheTimer.Enabled = False
        Case 130
            TheTimer.Enabled = False
        Case Else
    End Select
End If
End Function
Private Function gettext(lngwindow As Long) As String
'Used to read the filename from the dialog
    Dim strbuffer As String, lngtextlen As Long
    Let lngtextlen& = SendMessage(lngwindow&, WM_GETTEXTLENGTH, 0&, 0&)
    Let strbuffer$ = String(lngtextlen&, 0&)
    Call SendMessageByString(lngwindow&, WM_GETTEXT, lngtextlen& + 1&, strbuffer$)
    Let gettext$ = strbuffer$
End Function
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
Private Function DummyProc(ByVal dProc As Long) As Long
'Used to implement the hook
  DummyProc = dProc
End Function

