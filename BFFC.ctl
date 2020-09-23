VERSION 5.00
Begin VB.UserControl BFFC 
   CanGetFocus     =   0   'False
   ClientHeight    =   285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   315
   HasDC           =   0   'False
   HitBehavior     =   0  'None
   InvisibleAtRuntime=   -1  'True
   Picture         =   "BFFC.ctx":0000
   ScaleHeight     =   285
   ScaleWidth      =   315
   ToolboxBitmap   =   "BFFC.ctx":04B6
   Windowless      =   -1  'True
End
Attribute VB_Name = "BFFC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Just properties here, the action happens on frmButton and ModBrowse
Option Explicit
Dim Recurse As Boolean
Dim Create As Boolean
Dim mTitle As String
Dim mInitDir As String
Dim parent As Long
Private Sub UserControl_Resize()
On Error Resume Next
UserControl.Width = 315
UserControl.Height = 285
End Sub
Public Function SaveBrowse() As String
SaveBrowse = frmButton.Browse(parent, Create, False, mInitDir, mTitle)
End Function
Public Function OpenBrowse() As String
OpenBrowse = frmButton.Browse(parent, False, Recurse, mInitDir, mTitle)
End Function
Public Property Let ShowRecurse(ByVal vNewValue As Boolean)
Attribute ShowRecurse.VB_Description = "Shows/Hides the Recurse Checkbox on the Browse for Folder dialog"
Recurse = vNewValue
End Property
Public Property Let ShowCreate(ByVal vNewValue As Boolean)
Attribute ShowCreate.VB_Description = "Shows/Hides the create new folder button on the Browse for Folder dialog"
Create = vNewValue
End Property
Public Property Let Title(ByVal vNewValue As String)
Attribute Title.VB_Description = "Descriptive text on the Browse for Folder dialog"
mTitle = vNewValue
End Property
Public Property Let InitDir(ByVal vNewValue As String)
Attribute InitDir.VB_Description = "Input to Browse for Folder dialog to specify an initial directory"
mInitDir = vNewValue
End Property
Public Property Let ParentForm(ByVal vNewValue As Long)
Attribute ParentForm.VB_Description = "The owner forms window as long"
parent = vNewValue
End Property
Public Property Get DoRecurse() As Boolean
Attribute DoRecurse.VB_Description = "Returns user selection to recurse subfolders"
DoRecurse = LetsRecurse
End Property

