Attribute VB_Name = "modStringUtils"

Option Explicit

' Brought to you by:
'   Brad Martinez
'   btmtz@msn.com
'   btmtz@aol.com
'   http://members.aol.com/btmtz/vb
' ======================================================
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
                            (lpVersionInformation As OSVERSIONINFO) As Long
                                  
Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

Const VER_PLATFORM_WIN32s = 0
Const VER_PLATFORM_WIN32_WINDOWS = 1
Const VER_PLATFORM_WIN32_NT = 2

' ======================================================
' Handles overlapped source and destination blocks
Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" _
                      (pDest As Any, _
                      pSource As Any, _
                      ByVal ByteLen As Long)

' ======================================================
Declare Function IsTextUnicode Lib "advapi32" _
                            (lpBuffer As Any, _
                            ByVal cb As Long, _
                            lpi As Long) As Long
                            
Public Const IS_TEXT_UNICODE_ASCII16 = &H1
Public Const IS_TEXT_UNICODE_REVERSE_ASCII16 = &H10

Public Const IS_TEXT_UNICODE_STATISTICS = &H2
Public Const IS_TEXT_UNICODE_REVERSE_STATISTICS = &H20

Public Const IS_TEXT_UNICODE_CONTROLS = &H4
Public Const IS_TEXT_UNICODE_REVERSE_CONTROLS = &H40

Public Const IS_TEXT_UNICODE_SIGNATURE = &H8
Public Const IS_TEXT_UNICODE_REVERSE_SIGNATURE = &H80

Public Const IS_TEXT_UNICODE_ILLEGAL_CHARS = &H100
Public Const IS_TEXT_UNICODE_ODD_LENGTH = &H200
Public Const IS_TEXT_UNICODE_DBCS_LEADBYTE = &H400
Public Const IS_TEXT_UNICODE_NULL_BYTES = &H1000

Public Const IS_TEXT_UNICODE_UNICODE_MASK = &HF
Public Const IS_TEXT_UNICODE_REVERSE_MASK = &HF0
Public Const IS_TEXT_UNICODE_NOT_UNICODE_MASK = &HF00
Public Const IS_TEXT_UNICODE_NOT_ASCII_MASK = &HF000
'
' ======================================================
'

' Returns True if the current operating system is WinNT

Public Function IsWinNT() As Boolean
  Dim osvi As OSVERSIONINFO
  osvi.dwOSVersionInfoSize = Len(osvi)
  GetVersionEx osvi
  IsWinNT = (osvi.dwPlatformId = VER_PLATFORM_WIN32_NT)
End Function

' Returns string before first null char encountered (if any) from a string pointer.
' lpszStr = memory address of first byte in string
' nBytes = number of bytes to copy.
' StrConv used for both ANSII and Unicode strings
' BE CAREFULL!

Public Function GetStrFromPtr(lpszStr As Long, nBytes As Integer) As String
  ReDim ab(nBytes) As Byte   ' zero-based (nBytes + 1 elements)
  MoveMemory ab(0), ByVal lpszStr, nBytes
  GetStrFromPtr = GetStrFromBuffer(StrConv(ab(), vbUnicode))
End Function

' Returns string before first null char encountered (if any)
' from either an ANSII or Unicode string buffer.

Public Function GetStrFromBuffer(szStr As String) As String
  If IsUnicodeStr(szStr) Then szStr = StrConv(szStr, vbFromUnicode)
  If InStr(szStr, vbNullChar) Then
    GetStrFromBuffer = Left$(szStr, InStr(szStr, vbNullChar) - 1)
  Else
    ' If szStr had no null char, the Left$ function
    ' above would rtn a zero length string ("").
    GetStrFromBuffer = szStr
  End If
End Function

' Returns True if sBuffer evaluates to a Unicode string

Public Function IsUnicodeStr(sBuffer As String) As Boolean
  Dim dwRtnFlags As Long
  dwRtnFlags = IS_TEXT_UNICODE_UNICODE_MASK
  IsUnicodeStr = IsTextUnicode(ByVal sBuffer, Len(sBuffer), dwRtnFlags)
'Debug.Print "out: " & dwRtnFlags
End Function

