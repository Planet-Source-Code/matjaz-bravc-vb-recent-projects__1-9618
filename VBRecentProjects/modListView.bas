Attribute VB_Name = "modListView"

Option Explicit

Public Type LVBKIMAGE
  ulFlags As Long ' Tiled or normal etc
  hBitmap As Long ' Handle to the bitmap
  pszImage As String ' File Name of bitmap
  cchImageMax As Long ' Size of bitmap
  xOffsetPercent As Long ' X Offset to display (if not tiled)
  yOffsetPercent As Long ' Y Offset to display (if not tiled)
End Type

Public Const LVM_FIRST = &H1000
Public Const LVM_SETCOLUMNWIDTH As Long = LVM_FIRST + 30
Public Const LVM_GETHEADER = (LVM_FIRST + 31)
Public Const LVM_GETCOUNTPERPAGE = (LVM_FIRST + 40)
Public Const LVM_SETITEMSTATE = (LVM_FIRST + 43)
Public Const LVM_GETITEMSTATE As Long = (LVM_FIRST + 44)
Public Const LVM_GETSELECTEDCOUNT = (LVM_FIRST + 50)
Public Const LVIS_SELECTED = &H2
Public Const LVIF_STATE = &H8
Public Const LVIS_STATEIMAGEMASK As Long = &HF000
Public Const GWL_STYLE = (-16)
Public Const HDS_HOTTRACK = &H4
Public Const LVM_SETTEXTBKCOLOR = (LVM_FIRST + 38)
Public Const LVM_SETBKIMAGE = (LVM_FIRST + 68)
Public Const CLR_NONE = &HFFFFFFFF

'--- ListView Header
Public Const HDI_BITMAP = &H10
Public Const HDI_IMAGE = &H20
Public Const HDI_ORDER = &H80
Public Const HDI_FORMAT = &H4
Public Const HDI_TEXT = &H2
Public Const HDI_WIDTH = &H1
Public Const HDI_HEIGHT = HDI_WIDTH
Public Const HDF_LEFT = 0
Public Const HDF_RIGHT = 1
Public Const HDF_IMAGE = &H800
Public Const HDF_BITMAP_ON_RIGHT = &H1000
Public Const HDF_BITMAP = &H2000
Public Const HDF_STRING = &H4000
Public Const HDM_FIRST = &H1200
Public Const HDM_SETITEM = (HDM_FIRST + 4)
Public Const HDS_BUTTONS = &H2
Public Const SWP_DRAWFRAME = &H20
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_FLAGS = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME

'--- ListView Item
Public Type LVItem
  mask As Long
  iItem As Long
  iSubItem As Long
  state As Long
  stateMask As Long
  pszText As String
  cchTextMax As Long
  iImage As Long
  lParam As Long
  iIndent As Long
End Type

'--- ListView Header Item
Private Type HD_ITEM
  mask        As Long
  cxy         As Long
  pszText     As String
  hbm         As Long
  cchTextMax  As Long
  fmt         As Long
  lParam      As Long
  iImage      As Long
  iOrder      As Long
End Type

'--- ListView Set Column Width Messages ---'
Public Enum LVSCW_Styles
  LVSCW_AUTOSIZE = -1
  LVSCW_AUTOSIZE_USEHEADER = -2
End Enum

'--- Rectangle ---
Public Type RECT
  Left   As Long
  Top    As Long
  Right  As Long
  Bottom As Long
End Type

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Function SelectLVItems(LV As ListView, Optional SelectLV As Boolean = True) As Long
  Dim LVItem As LVItem

On Error GoTo ErrorHandler
  With LVItem
    .mask = LVIF_STATE
    .state = SelectLV
    .stateMask = LVIS_SELECTED
  End With
     
 'By setting wParam to -1, the call affects all
 'ListItems. To just change a particular item, pass its index as wParam.
  Call SendMessageAny(LV.hWnd, LVM_SETITEMSTATE, -1, LVItem)
 
 'Update the result
  SelectLVItems = SendMessageLong(LV.hWnd, LVM_GETSELECTEDCOUNT, 0&, 0&)
Exit Function

ErrorHandler:
  Err.Clear
End Function

Public Sub ShowLVHeaderIcon(LV As ListView, _
                            ColNo As Long, _
                            imgIconNo As Long, _
                            Justify As Long, _
                            ShowImage As Long)
  Dim r As Long
  Dim hHeader As Long
  Dim HD As HD_ITEM
  
  'get a handle to the listview header component
  hHeader = SendMessageLong(LV.hWnd, LVM_GETHEADER, 0, 0)
  
  'set up the required structure members
  With HD
    .mask = HDI_IMAGE Or HDI_FORMAT
    .fmt = HDF_LEFT Or HDF_STRING Or Justify Or ShowImage
    .pszText = LV.ColumnHeaders(ColNo + 1).Text
    If ShowImage Then .iImage = imgIconNo
  End With
  
  'Modify the header
  r = SendMessageAny(hHeader, HDM_SETITEM, ColNo, HD)
End Sub

Public Sub SetLVHeaderHotTrack(LV As ListView)
  Dim r As Long
  Dim hHeader As Long
  Dim rstyle As Long
   
  'get a handle to the listview header component
   hHeader = SendMessageLong(LV.hWnd, LVM_GETHEADER, 0, 0)
   
  'retrieve the current style of the header
   rstyle = GetWindowLong(hHeader, GWL_STYLE)
   
  'set/toggle the hottrack style attribute
   rstyle = rstyle Xor HDS_HOTTRACK

  'set the header style
   Call SetWindowLong(hHeader, GWL_STYLE, rstyle)
End Sub

Public Sub SetLVFlatHeaders(MyParent_hWnd As Long, LV As ListView)

  Dim Style As Long
  Dim hHeader As Long
   
 'get the handle to the listview header
  hHeader = SendMessageLong(LV.hWnd, LVM_GETHEADER, 0, ByVal 0&)
  
 'get the current style attributes for the header
  Style = GetWindowLong(hHeader, GWL_STYLE)
  
 'modify the style by toggling the HDS_BUTTONS style
  Style = Style Xor HDS_BUTTONS
  
 'set the new style and redraw the listview
  If Style Then
     Call SetWindowLong(hHeader, GWL_STYLE, Style)
     Call SetWindowPos(LV.hWnd, MyParent_hWnd, 0, 0, 0, 0, SWP_FLAGS)
  End If
End Sub

Public Sub LVSetColWidth(LV As ListView, ByVal ColumnIndex As Long, ByVal Style As LVSCW_Styles)
  '------------------------------------------------------------------------------
  '--- If you include the header in the sizing then the last column will
  '--- automatically size to fill the remaining listview width.
  '------------------------------------------------------------------------------
  With LV
     ' verify that the listview is in report view and that the column exists
     If .View = lvwReport Then
        If ColumnIndex >= 1 And ColumnIndex <= .ColumnHeaders.Count Then
          Call SendMessage(.hWnd, LVM_SETCOLUMNWIDTH, ColumnIndex - 1, ByVal Style)
        End If
     End If
  End With
End Sub

Public Sub LVSetAllColWidths(LV As ListView, ByVal Style As LVSCW_Styles)
  Dim ColumnIndex As Long
  '--- loop through all of the columns in the listview and size each
  With LV
    For ColumnIndex = 1 To .ColumnHeaders.Count
      LVSetColWidth LV, ColumnIndex, Style
    Next ColumnIndex
  End With
End Sub

Public Function GetListviewVisibleCount(LV As ListView) As Long
  GetListviewVisibleCount = SendMessageLong(LV.hWnd, LVM_GETCOUNTPERPAGE, 0&, ByVal 0&)
End Function

Public Sub WriteListViewSettings(LV As ListView, Optional ListViewNm As String = "")
  Dim i As Integer
  Dim LVName As String
  
  If ListViewNm = "" Then
    LVName = LV.Name
  Else
    LVName = ListViewNm
  End If
  SaveSetting App.ProductName, LVName, "SortKey", LV.SortKey
  SaveSetting App.ProductName, LVName, "SortOrder", LV.SortOrder
  SaveSetting App.ProductName, LVName, "View", LV.View
  For i = 1 To LV.ColumnHeaders.Count
    SaveSetting App.ProductName, LVName & ".Column" & LV.ColumnHeaders(i).Index, "Position", LV.ColumnHeaders(i).Position
    SaveSetting App.ProductName, LVName & ".Column" & LV.ColumnHeaders(i).Index, "Width", LV.ColumnHeaders(i).Width
  Next
End Sub

Public Sub ReadListViewSettings(LV As ListView, Optional ListViewNm As String = "", Optional ReadSortOrder As Boolean = True)
  Dim i As Integer
  Dim LVName As String
 
  If ListViewNm = "" Then
    LVName = LV.Name
  Else
    LVName = ListViewNm
  End If
  If ReadSortOrder Then
    LV.SortKey = GetSetting(App.ProductName, LVName, "SortKey", LV.SortKey)
    LV.SortOrder = GetSetting(App.ProductName, LVName, "SortOrder", LV.SortOrder)
  End If
  LV.View = GetSetting(App.ProductName, LVName, "View", LV.View)
  For i = 1 To LV.ColumnHeaders.Count
    LV.ColumnHeaders(i).Position = GetSetting(App.ProductName, LVName & ".Column" & LV.ColumnHeaders(i).Index, "Position", LV.ColumnHeaders(i).Position)
    LV.ColumnHeaders(i).Width = GetSetting(App.ProductName, LVName & ".Column" & LV.ColumnHeaders(i).Index, "Width", LV.ColumnHeaders(i).Width)
  Next
End Sub

Public Function SelectedItems(LV As ListView) As Long
  SelectedItems = SendMessageAny(LV.hWnd, LVM_GETSELECTEDCOUNT, 0&, 0&)
End Function

