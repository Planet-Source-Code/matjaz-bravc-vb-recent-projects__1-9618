VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWaitCursor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'WaitCursor - Wait cursor class demo
'
'This Visual Basic 5.0 example program demonstrates how you can use
'a simple class to help you set an hourglass cursor. This class not
'only provides the SetCursor method, which provides a convenient
'method of setting the hourglass cursor (or a custom cursor if you
'like), but it automatically restores the previous cursor when the
'class object is destroyed.
'
'Declare your CWaitCursor object within a subroutine:
'
' Private Sub MySub()
'     Dim wait As New CWaitCursor
'     wait.SetCursor
'
'     'Perform lengthy tasks here
'
' End Sub
'
'Although you can call the Restore method to restore the previous
'cursor, it is not necessary. CWaitCursor guarantees that the cursor
'will be restored when the subroutine terminates, even if the
'subroutine terminates due to an unhandled run-time error!

Private m_nPointer As MousePointerConstants

Public Sub SetCursor(Optional nPointer As MousePointerConstants = vbHourglass)
 Screen.MousePointer = nPointer
End Sub

Public Sub Restore()
 Screen.MousePointer = m_nPointer
End Sub

Private Sub Class_Initialize()
 m_nPointer = Screen.MousePointer
End Sub

Private Sub Class_Terminate()
 Restore
End Sub
