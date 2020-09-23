Attribute VB_Name = "modUndocShell32Defs"

Option Explicit

' ====================================================
' A demo of a dozen undocumented Shell32.dll functions
' ====================================================

' Brought to you by:
'   Brad Martinez
'   btmtz@msn.com
'   btmtz@aol.com
'   http://members.aol.com/btmtz/vb

' This demo would not have happened if it weren't for the function prototypes
' found at Chris Becke's site:

' http://www.dbn.lia.net/chris/  <chris@dbn.lia.net>

' I thank him for making this information available.

' ====================================================

' All of the Shell32.dll functions demonstrated are exported only by ordinal
' (NONAME) and are not know to be documented by Microsoft. As a result,
' they are most likely not supported by Microsoft and may very well not be
' included in future versions of Shell32.dll. Use them at your own risk.

' Each function's syntax and description was derived and tested solely by
' the author. The functions were also renamed from what may have been
' their original exported name in the debug version of the library, to slightly
' more intuitive names (since only the ordinals are shown in an export dump
' of Shell32.dll). It is suggested that developers who decide to implement
' these functions, maintain the names that are used here to avoid confusion.
' Here is the list:

' Ord   Hidden name           param bytes    Renamed to
' ===  ==========          =========    =========
' 59    _RestartDialog           12                 SHRestartSystemMB
' 60    ?                                4                 SHShutDownDialog
' 61    ?                              24                 SHRunDialog
' 62    _PickIconDlg             16                 SHChangeIconDialog

' 31    _PathFindExtension    4                  SHGetExtension
' 32    _PathAddBackslash    4                  SHAddBackslash
' 34    _PathFindFileName     4                  SHGetFileName
' 40    _PathIsRelative           4                  SHPathIsRelative
' 43    _PathIsExe                 4                  SHPathIsExe
' 45    _PathFileExists           4                  SHFileExists
' 52    _PathGetArgs             4                  SHGetPathArgs
' 92    _PathGetintPath          4                  SHGetShortPathName

' IMPORTANT NOTE: Unlike most documented Win32 API functions, the
' functions that accept string parameters (all but SHShutDownDialog),
' expect strings in either the ANSII or Unicode character set, depending
' on the Windows platform the function is called from (i.e. no separate
' ANSII "A" or Wide "W" function versions).

' In order for a function to return an accurate value (and reduce the potential
' for a fatal exception), the function must be passed ANSII strings when
' called in Win95, and must be passed Unicode strings when called in WinNT.
' Note the explicit use of the global "g_fIsWinNT" flag throughout the demo
' and the corresponding call to VB's StrConv function (equivalent to using the
' MultiByteToWideChar API) that converts ANSII strings to their Unicode
' equivalent when g_fIsWinNT evaluates to True.

' If it is found that any of the information in this demo proves to be inaccurate
' or incomplete, the author would appreciate notification at either of the email
' addresses above so that it can be corrected.
'                                             Thanks and enjoy, Brad Martinez

' Developed and tested with VB4.0a-32 on both Win95 v4.00.950a and WinNT
' v4.0 Server SP2.

' Set to True if the current OS is WinNT.
' Tested in *every* shell function's proc.
Public g_fIsWinNT As Boolean

' ======================================================
' Dialog functions (sorted by ordinal):
' ======================================================

' The "System Settings Change" message box.
' ("You must restart your computer before the new settings will take effect.")
Declare Function SHRestartSystemMB Lib "shell32" Alias "#59" _
                            (ByVal hOwner As Long, _
                            ByVal sPrompt As String, _
                            ByVal uFlags As Long) As Long

' hOwner = Message box owner, specify 0 for desktop (will be top-level)
' sPrompt = Specified prompt string placed above the default prompt.
' uFlags = Can be the following values:

' WinNT
' Appears to use ExitWindowsEx uFlags values and behave accordingly:
Public Const EWX_LOGOFF = 0
Public Const EWX_SHUTDOWN = 1   ' NT: needs SE_SHUTDOWN_NAME privilege (no def prompt)
Public Const EWX_REBOOT = 2        ' NT: needs SE_SHUTDOWN_NAME privilege
Public Const EWX_FORCE = 4
Public Const EWX_POWEROFF = 8   ' NT: needs SE_SHUTDOWN_NAME privilege

' Win95
' Any Yes selection produces the eqivalent to ExitWindowsEx(EWX_FORCE, 0) (?)
' (i.e. no WM_QUERYENDSESSION or WM_ENDSESSION is sent!).
' Other than is noted below, it was found that any other value shuts the system down
' (no reboot) and includes the default prompt.

' Shuts the system down (no reboot) and does not include the default prompt:
Public Const shrsExitNoDefPrompt = 1
' Reboots the system and includes the default prompt.
Public Const shrsRebootSystem = 2   ' = EWX_REBOOT

' Rtn vals: Yes = 6 (vbYes), No = 7 (vbNo)

'----------------------------
' The Shut Down dialog via the Start menu
Declare Function SHShutDownDialog Lib "shell32" Alias "#60" _
                            (ByVal YourGuess As Long) As Long

'----------------------------
' The Run dialog via the Start menu
Declare Function SHRunDialog Lib "shell32" Alias "#61" _
                            (ByVal hOwner As Long, _
                             ByVal Unknown1 As Long, _
                             ByVal Unknown2 As Long, _
                             ByVal szTitle As String, _
                             ByVal szPrompt As String, _
                             ByVal uFlags As Long) As Long

' hOwner = Dialog owner, specify 0 for desktop (will be top-level)
' Unknown1 = ?
' Unknown2 = ?, non-zero causes gpf! strings are ok...(?)
' szTitle = Dialog title, specify vbNullString for default ("Run")
' szPrompt = Dialog prompt, specify vbNullString for default ("Type the name...")

' If uFlags is the following constant, the string from last program run
' will not appear in the dialog's combo box (that's all I found...)

Public Const shrdNoMRUString = &H2   ' 2nd bit is set

' If there is some way to set & rtn the command line, I didn't find it...
' Always returns 0 (?)

'----------------------------
' The "Change Icon" dialog.
Declare Function SHChangeIconDialog Lib "shell32" Alias "#62" _
                            (ByVal hOwner As Long, _
                            ByVal szFilename As String, _
                            ByVal Reserved As Long, _
                            lpIconIndex As Long) As Long

' hOwner = Dialog owner, specify 0 for desktop (will be top-level)
' szFilename = The initially displayed filename, filled on selection.
'                      Should be allocated to MAX_PATH (260) in order to
'                      receive the selected filename's path.
' Reserved = ?
' lpIconIndex = Pointer to the initially displayed filename's icon index,
'                     and is filled on icon selection.

' Rtns non-zero on select, zero if cancelled.

' ======================================================
' Path functions (sorted by ordinal):
' ======================================================

' Rtns pointer to the last dot in szPath and the string following it.
' (includes the dot with the extension)
' Rtns 0 if szPath contains no dot.
' For the function to succeed, szPath should be null terminated
' and be allocated to MAX_PATH bytes (260).
' Does not check szPath for validity.
' (could be called "GetStrAtLastDot")
Declare Function SHGetExtension Lib "shell32" Alias "#31" _
                            (ByVal szPath As String) As Long

'----------------------------
' Inserts a backslash before the first null char in szPath.
' szPath is unchanged if it already contains a backslash
' before the first null char or contains no null char at all.
' Rtn pointer to?
' Does not check szPath for validity.
' (the name almost fits...)
Declare Function SHAddBackslash Lib "shell32" Alias "#32" _
                            (ByVal szPath As String) As Long

'----------------------------
' Rtn a pointer to the string in szPath after the last backslash.
' Rtns 0 if szPath contains no backslash or no char follows the last backslash.
' For the function to succeed, szPath should be null terminated
' and be allocated to MAX_PATH bytes (260).
' Does not check szPath for validity.
' (could be called "GetStrAfterLastBackslash")
Declare Function SHGetFileName Lib "shell32" Alias "#34" _
                            (ByVal szPath As String) As Long

'----------------------------
' Rtns non-zero if szPath does not evaluate to a UNC path.
' (if either the first char is not a backslash "\" or the 2nd char is not a colon ":")
' Does not check szPath for validity.
' (the name almost fits...)
Declare Function SHPathIsRelative Lib "shell32" Alias "#40" _
                            (ByVal szPath As String) As Long

'----------------------------
' Rtns non-zero if szPath has an executable extension.
' (if last 4 char are either ".exe", ".com", ".bat" or ".pif")
' Does not check szPath for validity.
' (could be called "HasExeExtension")
Declare Function SHPathIsExe Lib "shell32" Alias "#43" _
                            (ByVal szPath As String) As Long

'----------------------------
' Rtns non-zero if szPath is valid absolute UNC path.
' Accepts file, folder or network paths.
' Rtns True for a relative path only if it exists in the curdir.
' (the name actually fits...)
Declare Function SHFileExists Lib "shell32" Alias "#45" _
                            (ByVal szPath As String) As Long

'----------------------------
' Rtns a pointer to the string after first space in szPath.
' Rtns null pointer if szPath contains no space or no char
' following the first space.
' For the function to succeed, szPath should be null terminated
' and be allocated to MAX_PATH bytes (260).
' Does not check szPath for validity.
' (could be called "GetStrAfterFirstSpace")
Declare Function SHGetPathArgs Lib "shell32" Alias "#52" _
                            (ByVal szPath As String) As Long

'----------------------------
' Fills szPath w/ it's DOS (8.3) file system string.
' If successful, rtns non-zero (sometimes is a pointer to szPath, sometimes not!)
' Rtns zero if path is invalid.
' szPath must be a valid absolute path.
' Rtns non-zero for a relative path only if it exists in the curdir.
' For the function to work correctly, szPath should be null terminated
' and be allocated to MAX_PATH bytes (260).
' (the name definately fits...)
Declare Function SHGetShortPathName Lib "shell32" Alias "#92" _
                            (ByVal szPath As String) As Long

' ======================================================
' A few slightly more familiar APIs...

' Maximun long filename path length
Public Const MAX_PATH = 260

Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" _
                            (ByVal lpBuffer As String, _
                            ByVal nSize As Long) As Long

Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" _
                            (ByVal lpszFile As String, _
                            ByVal nIconIndex As Long, _
                            phiconLarge As Long, _
                            phiconSmall As Long, _
                            ByVal nIcons As Long) As Long

Declare Function DrawIconEx Lib "user32" _
                            (ByVal hDC As Long, _
                             ByVal xLeft As Long, _
                             ByVal yTop As Long, _
                             ByVal hIcon As Long, _
                             ByVal cxWidth As Long, _
                             ByVal cyWidth As Long, _
                             ByVal istepIfAniCur As Long, _
                             ByVal hbrFlickerFreeDraw As Long, _
                             ByVal diFlags As Long) As Boolean

' DrawIconEx() diFlags values:
Public Const DI_MASK = &H1
Public Const DI_IMAGE = &H2
Public Const DI_NORMAL = &H3
Public Const DI_COMPAT = &H4
Public Const DI_DEFAULTSIZE = &H8

Declare Function DestroyIcon Lib "user32" _
                            (ByVal hIcon As Long) As Long
'

' Terminates sPath w/ null chars making
' the return string MAX_PATH chars long.

Public Function MakeMaxPath(ByVal sPath As String) As String
  MakeMaxPath = sPath & String$(MAX_PATH - Len(sPath), 0)
End Function

' ======================================================
' Wrappers for Path functions (see respective API description above):

Public Function GetExtension(sPathIn) As String
  Dim sPathOut As String
  sPathOut = MakeMaxPath(sPathIn)
  If g_fIsWinNT Then sPathOut = StrConv(sPathOut, vbUnicode)
  ' Does not fill sPathOut w/ ext., just rtns ptr to ext
  GetExtension = GetStrFromPtr(SHGetExtension(sPathOut), Len(sPathOut))
End Function

Public Function NormalizePath(sPathIn As String) As String
  Dim sPathOut As String
  sPathOut = sPathIn & vbNullChar
  If g_fIsWinNT Then sPathOut = StrConv(sPathOut, vbUnicode)
  SHAddBackslash sPathOut
  NormalizePath = GetStrFromBuffer(sPathOut)
End Function

Public Function GetFileName(sPathIn As String) As String
  Dim sPathOut As String
  sPathOut = MakeMaxPath(sPathIn)
  If g_fIsWinNT Then sPathOut = StrConv(sPathOut, vbUnicode)
  GetFileName = GetStrFromPtr(SHGetFileName(sPathOut), MAX_PATH)
End Function

Public Function IsPathRelative(sPath As String) As Boolean
  If g_fIsWinNT Then
    IsPathRelative = SHPathIsRelative(StrConv(sPath, vbUnicode))
  Else
    IsPathRelative = SHPathIsRelative(sPath)
  End If
End Function

Public Function IsPathExe(sPath As String) As Boolean
  If g_fIsWinNT Then
    IsPathExe = SHPathIsExe(StrConv(sPath, vbUnicode))
  Else
    IsPathExe = SHPathIsExe(sPath)
  End If
End Function

Public Function FileExists(sPath As String) As Boolean
  If g_fIsWinNT Then
    FileExists = SHFileExists(StrConv(sPath, vbUnicode))
  Else
    FileExists = SHFileExists(sPath)
  End If
End Function

Public Function GetArgs(sPathIn As String) As String
  Dim sPathOut As String
  sPathOut = MakeMaxPath(sPathIn)   ' sPathIn
  If g_fIsWinNT Then sPathOut = StrConv(sPathOut, vbUnicode)
  GetArgs = GetStrFromPtr(SHGetPathArgs(sPathOut), Len(sPathOut))
End Function

Public Function GetPath(s As String) As String
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
      GetPath = ""
   Else
      GetPath = Left$(s, j)
   End If
End Function

Public Function GetShortPath(sPathIn As String) As String
  Dim sPathOut As String
  sPathOut = MakeMaxPath(sPathIn)   ' path could be longer...!
  If g_fIsWinNT Then sPathOut = StrConv(sPathOut, vbUnicode)
  SHGetShortPathName sPathOut
  GetShortPath = GetStrFromBuffer(sPathOut)
End Function
