VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Example"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Common Dialog"
      Height          =   2535
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   4935
      Begin PrjExample.OpenImage OpenImage1 
         Left            =   4320
         Top             =   1200
         _extentx        =   847
         _extenty        =   847
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Open Images"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   11
         Top             =   1920
         Width           =   4215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Open Files"
         Height          =   375
         Left            =   2520
         TabIndex        =   10
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Save File"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Returned path :"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1560
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Browse for Folders"
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Show Browse for Folder dialog"
         Height          =   375
         Left            =   720
         TabIndex        =   5
         Top             =   1200
         Width           =   3255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Show Recurse Checkbox"
         Height          =   255
         Left            =   2280
         TabIndex        =   4
         Top             =   360
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Show Create Folder button"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2280
         TabIndex        =   3
         Top             =   720
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Open  Mode"
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Save Mode"
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   720
         Width           =   1695
      End
      Begin PrjExample.BFFC BFFC1 
         Left            =   4440
         Top             =   2160
         _extentx        =   556
         _extenty        =   503
      End
      Begin VB.Label Label1 
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   4335
      End
      Begin VB.Label Label2 
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2040
         Width           =   4215
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Even though it meant duplicating some declares/subs
'I have kept code separate so you should be able to
'compile 2 controls
'1.OpenImage.ocx (OpenImage.ctl + CmnDialog.bas)
'2.BBFC.ocx (BBFC.ctl + frmButton + ModBrowse.bas)

'Both these controls could be improved and have more
'functions added to them, but work quite well
'in their current state.

'For those interested in modifying MS controls there
'are a couple of sites that provide freeware DLLs or
'controls that achieve similar functionality.
'(I'm stubborn enough to insist on writing my own)
'http://vbaccelerator.com/
'http://www.mvps.org/ccrp/

'***************************************************
'***********Browse for Folder Example***************
'***************************************************
Private Sub cmdBrowse_Click()
Dim temp As String
With BFFC1
    .ParentForm = Me.hwnd 'Browse for folder needs the calling forms handle
    .InitDir = "C:\windows\desktop"
    .ShowCreate = Check2.Value
    .ShowRecurse = Check1.Value
    .Title = "Test"
    'If you are allowing the user to select a path
    'to open files, you dont need a 'Create new folder' option
    'Similarly, if you are allowing the user to select a path
    'to save files to, you dont need a 'Recurse' checkbox.
    If Option1.Value = True Then
        temp = .OpenBrowse
    Else
        temp = .SaveBrowse
    End If
    
End With
If temp = "" Then
    Label1 = "User Cancelled"
    Label2 = ""
Else
    Label1 = temp
    'The DoRecurse property returns the users desire to
    'recurse folders for the .OpenBrowse method - it's
    'up to you to provide code for recursion from the
    'returned path
    If BFFC1.DoRecurse Then
        Label2 = "User requests recursion"
    Else
        Label2 = ""
    End If
End If

End Sub

Private Sub Option1_Click()
Check1.Enabled = Option1.Value
Check2.Enabled = Not Check1.Enabled
End Sub

Private Sub Option2_Click()
Check2.Enabled = Option2.Value
Check1.Enabled = Not Check2.Enabled
End Sub
'***************************************************
'***************Common Dialog Example***************
'***************************************************

Private Sub Command1_Click()
'Calls the standard Commondialog for open files
'with selected pictures shown in a panel below
'the main dialog
'.ParentForm is optional - causes the dialog to be modal-usually desirable
'Other properties, with the exception of .Filter and .Flags, are also valid
'for the .OpenImage function
OpenImage1.ParentForm = Me.hwnd
Text1.Text = OpenImage1.OpenImage
End Sub

Private Sub Command2_Click()
'Calls the standard Commondialog for open files
With OpenImage1
    .ParentForm = Me.hwnd
    .Filter = "All Files " + Chr(0) + "*.*"
    .Flags = 5
    .InitDir = "C:\windows\desktop"
    .Title = "Fred all files"
    Text1.Text = .OpenFile
End With
End Sub

Private Sub Command3_Click()
'Calls the standard Commondialog for save files
With OpenImage1
    .ParentForm = Me.hwnd
    .Filter = "All Files " + Chr(0) + "*.*"
    .Flags = 5
    .InitDir = "C:\windows\desktop"
    .Title = "Fred save files"
    Text1.Text = .SaveFile
End With

End Sub

