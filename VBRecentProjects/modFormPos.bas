Attribute VB_Name = "modFormPos"

Option Explicit

Public Sub WriteFormPos(Frm As Form)
Dim Buf As String
 
 Buf = Frm.Left & "," & Frm.Top & "," & Frm.Width & "," & Frm.Height
 SaveSetting App.ProductName, "FormPosition", Frm.Name, Buf
End Sub

Public Sub CenterForm(Frm As Form)
 Frm.Move (Screen.Width - Frm.Width) \ 2, (Screen.Height - Frm.Height) \ 2
End Sub

Public Sub ReadFormPos(Frm As Form, Optional RestoreFrmSize As Boolean = True)
  Dim Buf As String
  Dim l As Integer, t As Integer
  Dim H As Integer, W As Integer
  Dim Pos As Integer
 
 Buf = GetSetting(App.ProductName, "FormPosition", Frm.Name, "")
 If Buf = "" Then 'If empty center form!
   CenterForm Frm
 Else
  Pos = InStr(Buf, ",")
  l = CInt(Left(Buf, Pos - 1))
  Buf = Mid(Buf, Pos + 1)
  Pos = InStr(Buf, ",")
  t = CInt(Left(Buf, Pos - 1))
  Buf = Mid(Buf, Pos + 1)
  Pos = InStr(Buf, ",")
  W = CInt(Left(Buf, Pos - 1))
  H = CInt(Mid(Buf, Pos + 1))
  If RestoreFrmSize Then
    Frm.Move l, t, W, H
  Else
    Frm.Move l, t
  End If
 End If
End Sub
