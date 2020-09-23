Attribute VB_Name = "modSysUtils"

Option Explicit

Private Declare Function GetSystemDirectoryA Lib "kernel32" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Function ExtractFileName(s As String) As String
'  Returns the file portion of a file + pathname
Dim i As Integer
Dim j As Integer
  i = 0
  j = 0
  i = InStr(s, "\")
  If i <> 0 Then
    Do While i <> 0
     j = i
     i = InStr(j + 1, s, "\")
    Loop
    If j = 0 Then
     ExtractFileName = ""
    Else
     ExtractFileName = Right$(s, Len(s) - j)
    End If
  Else
    ExtractFileName = s
  End If
End Function

Public Function ExtractFilePath(s As String) As String
' Returns the path portion of a file + pathname
Dim i As Integer
Dim j As Integer
  i = 0
  j = 0
  i = InStr(s, "\")
  Do While i <> 0
   j = i
   i = InStr(j + 1, s, "\")
  Loop
  If j = 0 Then
   ExtractFilePath = ""
  Else
   ExtractFilePath = Left$(s, j)
  End If
End Function

Public Function URLExtractFileName(s As String) As String
'  Returns the file portion of a file + pathname
Dim i As Integer
Dim j As Integer
   
 i = 0
 j = 0
 
 i = InStr(s, "/")
 Do While i <> 0
  j = i
  i = InStr(j + 1, s, "/")
 Loop
 If j = 0 Then
  URLExtractFileName = ""
 Else
  URLExtractFileName = Right$(s, Len(s) - j)
 End If
End Function

Public Function ExtractFileExt(s As String) As String
'  Returns the extension portion of a file
Dim i As Integer
Dim j As Integer
  i = 0
  j = 0
  i = InStr(s, ".")
  Do While i <> 0
   j = i
   i = InStr(j + 1, s, ".")
  Loop
  If j = 0 Then
   ExtractFileExt = ""
  Else
   ExtractFileExt = Right$(s, Len(s) - j)
  End If
End Function

