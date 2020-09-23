VERSION 5.00
Begin VB.UserControl OpenImage 
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   InvisibleAtRuntime=   -1  'True
   Picture         =   "OpenImage.ctx":0000
   ScaleHeight     =   480
   ScaleWidth      =   480
   ToolboxBitmap   =   "OpenImage.ctx":0104
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1320
      Top             =   2760
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1095
      Left            =   2400
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   85
      TabIndex        =   0
      Top             =   2040
      Width           =   1335
   End
End
Attribute VB_Name = "OpenImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This API is used by timer1 to draw the selected image
'and print the file details on the commondialog
Option Explicit
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Const BDR_SUNKENOUTER = &H2
Const BDR_RAISEDINNER = &H4
Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Const BF_BOTTOM = &H8
Const BF_LEFT = &H1
Const BF_RIGHT = &H4
Const BF_TOP = &H2
Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Boolean, ByVal fdwUnderline As Boolean, ByVal fdwStrikeOut As Boolean, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Const TRANSPARENT = 1
Dim mFilter As String
Dim mflags As Long
Dim mInitDir As String
Dim mTitle As String
Dim parent As Long

Private Sub Timer1_Timer()
Dim h As Long, w As Long, hTxt As String, wTxt As String, szTxt As String
Dim RT As RECT
'gather some numbers so we can draw a thumbnail
'and maintain aspect ratio
If Picture1.Width >= Picture1.Height Then
    w = 80
    h = 80 * Picture1.Height / Picture1.Width
Else
    h = 80
    w = 80 * Picture1.Width / Picture1.Height
End If
'create a rectangular region for the panel below the standard dialog
SetRect RT, 10, 242, mCdlW - 18, 335
'draw a frame to hold the image and file details
DrawEdge mDC, RT, EDGE_ETCHED, BF_RECT
'paint the thumbnail in the panel
StretchBlt mDC, 320, 250, w, h, Picture1.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, ScrCopy
'get the file details for tmpFilename (set in the CdlgHook)
'If FileExists(tmpFullname) Then
If Trim(tmpFilename) <> "" Then
    szTxt = "Filesize : " + Format(Str(FileLen(tmpFullname) / 1024), "0.00") + " Kb"
    hTxt = "Image Height :" + Str(Picture1.ScaleHeight) + " Pixels"
    wTxt = "Image Width :" + Str(Picture1.ScaleWidth) + " Pixels"
Else
    szTxt = "Filesize : N/A"
    hTxt = "Image Height : N/A"
    wTxt = "Image Width : N/A"
End If
'End If
'print on a transparent background
SetBkMode mDC, TRANSPARENT
'delete old font settings and generate new font
DeleteObject SelectObject(mDC, CreateMyFont(8, 0))
'print the file details in the appropriate place
TextOut mDC, 20, 255, szTxt, Len(szTxt)
TextOut mDC, 20, 275, hTxt, Len(hTxt)
TextOut mDC, 20, 295, wTxt, Len(wTxt)
End Sub

Private Function CreateMyFont(nSize As Integer, nDegrees As Long) As Long
    CreateMyFont = CreateFont(12, 0, nDegrees * 10, 0, 400, False, False, False, 1, 0, 0, 2, 0, "MS Sans Serif")
End Function
Public Property Let ParentForm(ByVal vNewValue As Long)
parent = vNewValue
End Property
Public Function OpenImage() As String
Attribute OpenImage.VB_Description = "Brings up the Open File Dialog in Explorer View with a preview of any image files selected by the user"
Timer1.Enabled = True
OpenImage = ShowOpen(parent, "Image files " + Chr(0) + "*.bmp;*.gif;*.jpg;*.dib", 5, mInitDir, mTitle, True)
End Function
Public Function OpenFile() As String
Attribute OpenFile.VB_Description = "Standard open file dialog"
OpenFile = ShowOpen(parent, mFilter, mflags, mInitDir, mTitle, False)
End Function
Public Function SaveFile() As String
Attribute SaveFile.VB_Description = "Standard save file dialog"
SaveFile = ShowSave(parent, mFilter, mflags, mInitDir, mTitle)
End Function
Private Sub UserControl_Initialize()
Set ThePic = Picture1
Set TheTimer = Timer1
End Sub
Private Sub UserControl_Resize()
On Error Resume Next
UserControl.Width = 480
UserControl.Height = 480
End Sub
Public Property Let Filter(ByVal vNewValue As String)
Attribute Filter.VB_Description = "Example : ""Bitmaps files (*.bmp)"" + Chr(0) + ""*.bmp""+ Chr(0) + ""All Files (*.*)""+chr(0)+""*.*""\r\n"
mFilter = vNewValue
End Property
Public Property Let Flags(ByVal vNewValue As Long)
Attribute Flags.VB_Description = "Optional - set to 5 to remove the ""Open as Read only"" checkbox"
mflags = vNewValue
End Property
Public Property Let InitDir(ByVal vNewValue As String)
Attribute InitDir.VB_Description = "The initial directory presented to the user in the common dialog"
mInitDir = vNewValue
End Property
Public Property Let Title(ByVal vNewValue As String)
Attribute Title.VB_Description = "The text that appears as the caption of the common dialog window"
mTitle = vNewValue
End Property
