VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMainForm 
   Caption         =   "VB Recent Projects"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   9105
   Icon            =   "frmMainForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   9105
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMoveUp 
      Caption         =   "Move Up"
      Height          =   375
      Left            =   2820
      TabIndex        =   3
      Top             =   5460
      Width           =   975
   End
   Begin VB.CommandButton cmdMoveDown 
      Caption         =   "Move Down"
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   5460
      Width           =   1095
   End
   Begin VB.OptionButton optVBVersion 
      Caption         =   "VB 6.0"
      Height          =   255
      Index           =   1
      Left            =   8160
      TabIndex        =   7
      Top             =   5520
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.OptionButton optVBVersion 
      Caption         =   "VB 5.0"
      Height          =   255
      Index           =   0
      Left            =   7080
      TabIndex        =   6
      Top             =   5520
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "R&emove from list"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   5460
      Width           =   1455
   End
   Begin MSComctlLib.ImageList imgIcons16 
      Left            =   120
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainForm.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainForm.frx":0466
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainForm.frx":05C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainForm.frx":071C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSaveChanges 
      Caption         =   "&Save changes"
      Height          =   375
      Left            =   4980
      TabIndex        =   5
      Top             =   5460
      Width           =   1515
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   375
      Left            =   60
      TabIndex        =   1
      Top             =   5460
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvListView 
      DragIcon        =   "frmMainForm.frx":0876
      Height          =   5355
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   9446
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgIcons16"
      SmallIcons      =   "imgIcons16"
      ColHdrIcons     =   "imgIcons16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmMainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------
' VB Recent Project Manager
' Author: Matjaz Bravc
' mbravc@hotmail.com
'
' Thanks to Brad Martinez and Steve McMahon!
'--------------------------------------------
Option Explicit

Private AnyChangesMade As Boolean

Dim InDrag As Boolean ' Flag that signals a Drag Drop operation.
Dim DraggedLVItem As ListItem ' Item that is being dragged.

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Private Declare Function ShellExecute Lib _
              "shell32.dll" Alias "ShellExecuteA" _
              (ByVal hwnd As Long, _
               ByVal lpOperation As String, _
               ByVal lpFile As String, _
               ByVal lpParameters As String, _
               ByVal lpDirectory As String, _
               ByVal nShowCmd As Long) As Long
               
Private Const SW_SHOW = 1
Private m_cHdrIcons As New cLVHeaderSortIcons

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Sub ExecuteCommand(ByVal NavTo As String)
  Dim CDir As String
  Dim hBrowse As Long
  
  CDir = CurDir
  ChDir GetPath(NavTo)
  hBrowse = ShellExecute(0&, "open", GetFileName(NavTo), "", "", SW_SHOW)
  ChDir CDir
End Sub

Private Sub CreateLVColumns()
  Dim ColHeader As ColumnHeader
  
  '--- Clear the ListView control
  lvListView.ListItems.Clear
  lvListView.ColumnHeaders.Clear
  Set ColHeader = lvListView.ColumnHeaders.Add(, "Project", "Recent projects (0)")
  ColHeader.Icon = 4
  Call lvListView.ColumnHeaders.Add(, "ProjectExists", "Project exists")
  Call lvListView.ColumnHeaders.Add(, "ProjectFile", "Project Filename")
End Sub

Private Sub cmdDelete_Click()
  Dim i As Long
  Dim Index As Long
  
  LockWindowUpdate lvListView.hwnd
  Index = lvListView.SelectedItem.Index
  For i = lvListView.ListItems.Count To 1 Step -1
    If lvListView.ListItems(i).Selected Then
      Call lvListView.ListItems.Remove(i)
    End If
  Next
  lvListView.Refresh
  If lvListView.ListItems.Count > 0 Then
    If lvListView.ListItems.Count > Index Then
      lvListView.ListItems(Index).Selected = True
    Else
      lvListView.ListItems(lvListView.ListItems.Count).Selected = True
    End If
  End If
  lvListView.ColumnHeaders(1).Text = "Recent Projects (" & lvListView.ListItems.Count & ")"
  LockWindowUpdate 0
  AnyChangesMade = True
End Sub

Private Sub cmdMoveDown_Click()
  Dim SelLVItem As ListItem
  Dim LVItem As ListItem
  
  If SelectedItems(lvListView) = 1 Then
    lvListView.Sorted = False
    LockWindowUpdate lvListView.hwnd
    Set SelLVItem = lvListView.SelectedItem
    If lvListView.SelectedItem.Index < lvListView.ListItems.Count Then
      If lvListView.SelectedItem.Index < lvListView.ListItems.Count Then
        Set LVItem = lvListView.ListItems.Add((lvListView.SelectedItem.Index + 2), , SelLVItem.Text, 0, 0)
        LVItem.Icon = SelLVItem.Icon
        LVItem.SmallIcon = SelLVItem.SmallIcon
        LVItem.ListSubItems.Add , "ProjectExists", SelLVItem.ListSubItems("ProjectExists").Text
        LVItem.ListSubItems("ProjectExists").ForeColor = SelLVItem.ListSubItems("ProjectExists").ForeColor
        LVItem.ListSubItems.Add , "ProjectFile", SelLVItem.ListSubItems("ProjectFile").Text
        LVItem.ListSubItems("ProjectFile").ForeColor = SelLVItem.ListSubItems("ProjectFile").ForeColor
        lvListView.ListItems.Remove lvListView.SelectedItem.Index
        lvListView.Refresh
        LVItem.EnsureVisible
        LVItem.Selected = True
        LVItem.Tag = SelLVItem.Text
      End If
    End If
    LockWindowUpdate 0
  End If
End Sub

Private Sub cmdMoveUp_Click()
  Dim SelLVItem As ListItem
  Dim LVItem As ListItem
  
  If SelectedItems(lvListView) = 1 Then
    lvListView.Sorted = False
    LockWindowUpdate lvListView.hwnd
    Set SelLVItem = lvListView.SelectedItem
    If lvListView.SelectedItem.Index > 1 Then
      If lvListView.SelectedItem.Index - 1 > 0 Then
        Set LVItem = lvListView.ListItems.Add((lvListView.SelectedItem.Index - 1), , SelLVItem.Text, 0, 0)
        LVItem.Icon = SelLVItem.Icon
        LVItem.SmallIcon = SelLVItem.SmallIcon
        LVItem.ListSubItems.Add , "ProjectExists", SelLVItem.ListSubItems("ProjectExists").Text
        LVItem.ListSubItems("ProjectExists").ForeColor = SelLVItem.ListSubItems("ProjectExists").ForeColor
        LVItem.ListSubItems.Add , "ProjectFile", SelLVItem.ListSubItems("ProjectFile").Text
        LVItem.ListSubItems("ProjectFile").ForeColor = SelLVItem.ListSubItems("ProjectFile").ForeColor
        lvListView.ListItems.Remove lvListView.SelectedItem.Index
        lvListView.Refresh
        LVItem.EnsureVisible
        LVItem.Selected = True
        LVItem.Tag = SelLVItem.Text
      End If
    End If
    LockWindowUpdate 0
  End If
End Sub

Private Sub cmdSaveChanges_Click()
  Dim i As Long
  Dim ItemCaption As String
  Dim LVItem As ListItem
  Dim Wait As New clsWaitCursor
  Dim CReg As clsRegistry

  Wait.SetCursor
  LockWindowUpdate lvListView.hwnd
  Set CReg = New clsRegistry
  CReg.ClassKey = HKEY_CURRENT_USER
  If optVBVersion(0).Value = True Then
    CReg.SectionKey = "Software\Microsoft\Visual Basic\5.0\RecentFiles\"
  Else
    CReg.SectionKey = "Software\Microsoft\Visual Basic\6.0\RecentFiles\"
  End If
  CReg.DeleteKey
  CReg.CreateKey
  For i = 1 To lvListView.ListItems.Count
    CReg.ValueKey = i
    CReg.ValueType = REG_SZ
    CReg.Value = lvListView.ListItems(i).Text
  Next
  LockWindowUpdate 0
  AnyChangesMade = False
End Sub

Private Sub Form_Activate()
  lvListView.SetFocus
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If AnyChangesMade Then
    If MsgBox("Save changes?      ", vbQuestion + vbYesNo, "Question") = vbYes Then
      cmdSaveChanges_Click
    End If
  End If
End Sub

Private Sub Form_Resize()
  lvListView.Move 0, 0, ScaleWidth, ScaleHeight - 500
  cmdRefresh.Move cmdRefresh.Left, ScaleHeight - cmdRefresh.Height - 60
  cmdDelete.Move cmdDelete.Left, ScaleHeight - cmdDelete.Height - 60
  cmdMoveUp.Move cmdMoveUp.Left, ScaleHeight - cmdMoveUp.Height - 60
  cmdMoveDown.Move cmdMoveDown.Left, ScaleHeight - cmdMoveDown.Height - 60
  cmdSaveChanges.Move cmdSaveChanges.Left, ScaleHeight - cmdSaveChanges.Height - 60
  optVBVersion(1).Move ScaleWidth - optVBVersion(1).Width - 60, ScaleHeight - optVBVersion(1).Height - 60
  optVBVersion(0).Move ScaleWidth - optVBVersion(0).Width - optVBVersion(1).Width - 120, ScaleHeight - optVBVersion(0).Height - 60
End Sub

Private Sub Form_Unload(Cancel As Integer)
  WriteListViewSettings lvListView
  WriteFormPos frmMainForm
End Sub

Private Sub RefreshList()
  Dim sKeys() As String
  Dim iKeys As Long
  Dim iKeysCount As Long
  Dim Wait As New clsWaitCursor
  Dim TmpStr As String
  Dim ItemCaption As String
  Dim LVItem As ListItem
  Dim CReg As clsRegistry

  If AnyChangesMade Then
    If MsgBox("Save changes?      ", vbQuestion + vbYesNo, "Question") = vbYes Then
      cmdSaveChanges_Click
    End If
  End If
  
  Wait.SetCursor
  
  lvListView.ListItems.Clear
  LockWindowUpdate lvListView.hwnd
  
  iKeysCount = 0
  Set CReg = New clsRegistry
  CReg.ClassKey = HKEY_CURRENT_USER
  If optVBVersion(0).Value = True Then
    CReg.SectionKey = "Software\Microsoft\Visual Basic\5.0\RecentFiles\"
  Else
    CReg.SectionKey = "Software\Microsoft\Visual Basic\6.0\RecentFiles\"
  End If
  CReg.EnumerateValues sKeys(), iKeysCount
  
  If (iKeysCount > 0) Then
   For iKeys = 1 To iKeysCount
     CReg.ValueKey = sKeys(iKeys)
     ItemCaption = CReg.Value
     Set LVItem = lvListView.ListItems.Add(, , ItemCaption, 0, 0)
     If FileExists(ItemCaption) Then
       LVItem.Icon = 1
       LVItem.SmallIcon = 1
       LVItem.ListSubItems.Add , "ProjectExists", "Yes"
       LVItem.ListSubItems("ProjectExists").ForeColor = vbWindowText
       LVItem.ListSubItems.Add , "ProjectFile", GetFileName(ItemCaption)
       LVItem.ListSubItems("ProjectFile").ForeColor = vbWindowText
     Else
       LVItem.Icon = 2
       LVItem.SmallIcon = 2
       LVItem.ListSubItems.Add , "ProjectExists", "No"
       LVItem.ListSubItems("ProjectExists").ForeColor = vbRed
       LVItem.ListSubItems.Add , "ProjectFile", GetFileName(ItemCaption)
       LVItem.ListSubItems("ProjectFile").ForeColor = vbRed
     End If
     LVItem.Selected = False
     LVItem.Tag = sKeys(iKeys)
   Next iKeys
  End If
  lvListView.Refresh
  lvListView.ColumnHeaders(1).Text = "Recent Projects (" & lvListView.ListItems.Count & ")"
  LockWindowUpdate 0
  If lvListView.ListItems.Count > 0 Then
    lvListView.ListItems(1).Selected = True
  End If
  AnyChangesMade = False
End Sub

Private Sub cmdRefresh_Click()
  RefreshList
End Sub

Private Sub Form_Load()
  ' Set to True if the current OS is WinNT. Tested in *every* shell function's proc.
  g_fIsWinNT = IsWinNT
  AnyChangesMade = False
  CreateLVColumns
  ReadListViewSettings lvListView, , False
  ReadFormPos frmMainForm
  RefreshList
  ' Initialize the header sort icons object.
  Set m_cHdrIcons.ListView = lvListView
  ' Set the header icons and sort the ListView
  'Call m_cHdrIcons.SetHeaderIcons(lvListView.SortKey, lvListView.SortOrder)
End Sub

Private Sub lvListView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  
  With lvListView
    ' Toggle the clicked column's sort order only if the active colum is clicked
    ' (iow, don't reverse the sort order when different columns are clicked).
    If (.SortKey = ColumnHeader.Index - 1) Then
      ColumnHeader.Tag = Not Val(ColumnHeader.Tag)
    End If
    
    ' Set sort order to that of the respective SortOrderConstants value
    .SortOrder = Abs(Val(ColumnHeader.Tag))
    
    ' Get the zero-based index of the clicked column.
    ' (ColumnHeader.Index is one-based).
    .SortKey = ColumnHeader.Index - 1
  
    ' set the header icons and sort the ListView
    Call m_cHdrIcons.SetHeaderIcons(.SortKey, .SortOrder)
    .Sorted = True
  End With
End Sub

Private Sub lvListView_DblClick()
  If lvListView.SelectedItem.ListSubItems("ProjectExists").Text = "Yes" Then
    ExecuteCommand lvListView.SelectedItem.Text
  Else
    Beep
  End If
End Sub

Sub MoveRow(ByVal pi_MoveFrom As Integer, ByVal pi_MoveTo As Integer)
  Dim li_Counter As Integer
    
  If pi_MoveFrom > pi_MoveTo Then
    'Moving up the list - so shift them all down
    For li_Counter = pi_MoveFrom To pi_MoveTo + 1 Step -1
      'lvListView.ListItems(li_Counter) = lvListView.ListItems(li_Counter - 1)
    Next li_Counter
  Else
    'Moving down the list - so shift them all up
    For li_Counter = pi_MoveFrom To pi_MoveTo - 1
      'lvListView.ListItems(li_Counter) = lvListView.ListItems(li_Counter + 1)
    Next li_Counter
  End If
End Sub

Private Sub lvListView_DragDrop(Source As Control, x As Single, y As Single)
  Dim LVItem As ListItem

  If lvListView.DropHighlight Is Nothing Then
    Set lvListView.DropHighlight = Nothing
    InDrag = False
    Exit Sub
  Else
    If DraggedLVItem = lvListView.DropHighlight Then Exit Sub
    Set LVItem = lvListView.ListItems.Add((lvListView.DropHighlight.Index), , DraggedLVItem.Text, 0, 0)
    LVItem.Icon = DraggedLVItem.Icon
    LVItem.SmallIcon = DraggedLVItem.SmallIcon
    LVItem.ListSubItems.Add , "ProjectExists", DraggedLVItem.ListSubItems("ProjectExists").Text
    LVItem.ListSubItems("ProjectExists").ForeColor = DraggedLVItem.ListSubItems("ProjectExists").ForeColor
    LVItem.ListSubItems.Add , "ProjectFile", DraggedLVItem.ListSubItems("ProjectFile").Text
    LVItem.ListSubItems("ProjectFile").ForeColor = DraggedLVItem.ListSubItems("ProjectFile").ForeColor
    lvListView.ListItems.Remove DraggedLVItem.Index
    lvListView.Refresh
    LVItem.Selected = True
    LVItem.Tag = DraggedLVItem.Text
    Set lvListView.DropHighlight = Nothing
    InDrag = False
  End If
End Sub

Private Sub lvListView_DragOver(Source As Control, x As Single, y As Single, state As Integer)

  If InDrag Then
    ' Set DropHighlight to the mouse's coordinates.
    Set lvListView.DropHighlight = lvListView.HitTest(x, y)
  End If
End Sub

Private Sub lvListView_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDelete Then
    cmdDelete_Click
  End If
End Sub

Private Sub lvListView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If lvListView.HitTest(x, y) Is Nothing Then
    lvListView.MultiSelect = False
  Else
    lvListView.MultiSelect = IIf(Shift = 0, False, True)
    Set DraggedLVItem = lvListView.HitTest(x, y) ' Set the item being dragged.
  End If
End Sub

Private Sub lvListView_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = vbLeftButton Then
    InDrag = True ' Set the flag to true.
    lvListView.Sorted = False
    Set lvListView.SelectedItem = lvListView.HitTest(x, y)
    lvListView.Drag vbBeginDrag ' Drag operation
  End If
End Sub

Private Sub lvListView_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Set lvListView.DropHighlight = Nothing
End Sub

Private Sub optVBVersion_Click(Index As Integer)
  RefreshList
End Sub
