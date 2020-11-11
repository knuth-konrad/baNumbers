Attribute VB_Name = "PUtil"
'$PROBHIDE ALL
'------------------------------------------------------------------------------
'Purpose  : Allgemeine Routinensammlung
'
'Prereq.  : -
'Note     : -
'
'   Author: Knuth Konrad 11.12.2012
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
Option Explicit
DefLng A-Z
'------------------------------------------------------------------------------
'*** Constants ***
'------------------------------------------------------------------------------
Private Const MAX_PATH As Long = 260

'API Konstanten für OS-Version
Private Const VER_PLATFORM_WIN32_WINDOWS As Long = 1
Private Const VER_PLATFORM_WIN32_NT As Long = 2

'API Konstanten für SHBrowseForFolder
Private Const BIF_BROWSEFORCOMPUTER As Long = &H1000
Private Const BIF_BROWSEFORPRINTER  As Long = &H2000
Private Const BIF_BROWSEINCLUDEFILES  As Long = &H4000
Private Const BIF_DONTGOBELOWDOMAIN  As Long = 2
Private Const BIF_RETURNFSANCESTORS  As Long = 8
Private Const BIF_RETURNONLYFSDIRS  As Long = 1
Private Const BIF_STATUSTEXT  As Long = 4

'API-Konstanten für FindFirstFile
Private Const FILE_ATTRIBUTE_ARCHIVE  As Long = &H20
Private Const FILE_ATTRIBUTE_COMPRESSED  As Long = &H800
Private Const FILE_ATTRIBUTE_DIRECTORY  As Long = &H10
Private Const FILE_ATTRIBUTE_HIDDEN  As Long = &H2
Private Const FILE_ATTRIBUTE_NORMAL  As Long = &H80
Private Const FILE_ATTRIBUTE_READONLY  As Long = &H1
Private Const FILE_ATTRIBUTE_SYSTEM  As Long = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY  As Long = &H100
Private Const INVALID_HANDLE_VALUE  As Long = -1

'API Konstanten für ShellAndWait
Private Const STILL_ACTIVE As Long = &H103
Private Const PROCESS_QUERY_INFORMATION As Long = &H400

'API Konstanten für ShellAndWaitApi
Private Const NORMAL_PRIORITY_CLASS As Long = &H20
Private Const INFINITE As Long = -1
Private Const STARTF_USESHOWWINDOW  As Long = &H1

Private Const MAX_COMPUTERNAME_LENGTH As Long = 15

Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK  As Long = &HFF
'------------------------------------------------------------------------------
'*** Enumeration/TYPEs ***
'------------------------------------------------------------------------------
'Type für OS-Version
Private Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
End Type

'Type für SHBrowseForFolder (BrowseForFolder1)
Private Type BROWSEINFO1
   hwndOwner As Long
   iImage As Long
   lParam As Long
   lpfn As Long
   lpszTitle As Long
   pIDLRoot As Long
   pszDisplayName As Long
   ulFlags As Long
End Type

'Type für SHBrowseForFolder (BrowseForFolder)
Private Type BROWSEINFO
   hwndOwner As Long
   pIDLRoot As Long
   pszDisplayName As Long
   lpszTitle As Long
   ulFlags As Long
   lpfnCallback As Long
   lParam As Long
   iImage As Long
End Type

'Type für FindFirstFile
Private Type FileTime
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FileTime
   ftLastAccessTime As FileTime
   ftLastWriteTime As FileTime
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

'für Funktion IsDateAny
Private Type SYSTEMTIME
   wYear As Integer
   wMonth As Integer
   wDayOfWeek As Integer
   wDay As Integer
   wHour As Integer
   wMinute As Integer
   wSecond As Integer
   wMilliseconds As Integer
End Type

' ShallAndWaitApi
Private Type STARTUPINFO
   cb As Long
   lpReserved As String
   lpDesktop As String
   lpTitle As String
   dwX As Long
   dwY As Long
   dwXSize As Long
   dwYSize As Long
   dwXCountChars As Long
   dwYCountChars As Long
   dwFillAttribute As Long
   dwFlags As Long
   wShowWindow As Integer
   cbReserved2 As Integer
   lpReserved2 As Long
   hStdInput As Long
   hStdOutput As Long
   hStdError As Long
End Type

Private Type PROCESS_INFORMATION
   hProcess As Long
   hThread As Long
   dwProcessId As Long
   dwThreadId As Long
End Type
'------------------------------------------------------------------------------
'*** Declares ***
'------------------------------------------------------------------------------
'API Deklaration für IsDateAny, DateTimeCompare
Private Declare Function SystemTimeToVariantTime Lib "OleAut32.dll" _
   (lpSystemTime As SYSTEMTIME, vbtime As Date) As Long
Private Declare Function VariantTimeToSystemTime Lib "OleAut32.dll" (ByVal vbtime As Date, lpSystemTime As SYSTEMTIME) As Long

'API Deklaration für BrowseForFolder-Dialog
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32.dll" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "SHELL32.DLL" (ByRef lpbi As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "SHELL32.DLL" Alias "SHGetPathFromIDListA" (ByVal lPidl As Long, ByVal sPath As String) As Long

'API-Deklaration für FileExist und DirExist
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

'API Deklaration für IsOSNT()
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

'API für ArrayDimension, PtrObj
Private Declare Sub RtlMoveMemory Lib "kernel32" (dest As Any, Source As Any, ByVal bytes As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
   (Destination As Any, Source As Any, ByVal Length As Long)

'API für GetTempDir
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'API für ShellAndWait
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal DesiredAccess As Long, ByVal InheritHandle As Long, ByVal ProcessId As Long) As Long

'API für GetWindowsUsername & GetWindowsComputerName
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function ShellExecute Lib "SHELL32.DLL" Alias "ShellExecuteA" _
   (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' API für ShellAndWaitApi
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal _
   lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
   ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
   ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, _
   lpStartupInfo As STARTUPINFO, lpProcessInformation As _
   PROCESS_INFORMATION) As Long

'DWORD FormatMessage(
'  DWORD dwFlags,      // source and processing options
'  LPCVOID lpSource,   // message source
'  DWORD dwMessageId,  // message identifier
'  DWORD dwLanguageId, // language identifier
'  LPTSTR lpBuffer,    // message buffer
'  DWORD nSize,        // maximum size of message buffer
'  va_list *Arguments  // array of message inserts
');

Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" _
   (ByVal dwFlags As Long, _
   lpSource As Any, _
   ByVal dwMessageId As Long, _
   ByVal dwLanguageId As Long, _
   ByVal lpBuffer As String, _
   ByVal nSize As Long, _
   Arguments As Long) As Long
'------------------------------------------------------------------------------
'*** Variables ***
'------------------------------------------------------------------------------
'==============================================================================

Public Function NormalizePath(ByVal sPath As String, Optional ByVal sDelimiter As String = "\", _
   Optional ByVal bolCheckTail As Boolean = True) As String
'------------------------------------------------------------------------------
'Purpose  : Stellt sicher, daß übergebener Pfad mit Back-/Slash
'           beginnt/endet
'
'Prereq.  : -
'Parameter: sPath          - zu überrpüfende Pfadangabe
'           sDelimiter     - Welcherart ist der Delimiter bei dem Pfad
'           bolCheckTail   - Ende des Pfads (True) oder Beginn (False)
'                            überprüfen.
'Returns  : sPath mit abschließenden Backslash
'Note     : -
'
'   Author: Bruce McKinney - Hardcore Visual Basic 5
'   Source: -
'  Changed: 09.08.1999, Knuth Konrad
'           - Parameterübergabe ByVal und als String anstatt als Variant
'------------------------------------------------------------------------------
   
   If bolCheckTail = True Then
      If Right$(sPath, 1) <> sDelimiter Then
         NormalizePath = sPath & sDelimiter
      Else
         NormalizePath = sPath
      End If
   Else
      If Left$(sPath, 1) <> sDelimiter Then
         NormalizePath = sDelimiter & sPath
      Else
         NormalizePath = sPath
      End If
   End If
   
End Function
'==============================================================================

Public Function DenormalizePath(ByVal sPath As String, Optional ByVal sDelimiter As String = "\", _
   Optional ByVal bolCheckTail As Boolean = True) As String
'------------------------------------------------------------------------------
'Purpose  : Stellt sicher, daß übergebener Pfad ohne Back-/Slash
'           beginnt/endet
'
'Prereq.  : -
'Parameter: sPath          - zu überrpüfende Pfadangabe
'           sDelimiter     - Welcherart ist der Delimiter bei dem Pfad
'           bolCheckTail   - Ende des Pfads (True) oder Beginn (False)
'                            überprüfen.
'Returns  : sPath ohne abschließenden Backslash
'Note     : -
'
'   Author: Bruce McKinney - Hardcore Visual Basic 5
'   Source: -
'  Changed: 09.08.1999, Knuth Konrad
'           - Parameterübergabe ByVal und als String anstatt als Variant
'------------------------------------------------------------------------------
   
   If bolCheckTail = True Then
      If Right$(sPath, 1) = sDelimiter Then
         DenormalizePath = Left$(sPath, Len(sPath) - 1)
      Else
         DenormalizePath = sPath
      End If
   Else
      If Left$(sPath, 1) = sDelimiter Then
         DenormalizePath = Mid$(sPath, 2)
      Else
         DenormalizePath = sPath
      End If
   End If
   
End Function
'==============================================================================

Public Function BrowseForFolder(Optional sCaption As String = vbNullString) As String
'------------------------------------------------------------------------------
'Purpose  : Verzeichnis-Dialog
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : ACHTUNG: Die Funktion kann auch Netzwerkpfade zurückgeben (z.B. "\\PP200\public\Tom")
'
'   Author:
'   Source: Visual Basic 6 Kochbuch, Doberenz & Kowalski, Hanser Verlag
'  Changed: -
'------------------------------------------------------------------------------
' ACHTUNG: Die Funktion kann auch Netzwerkpfade zurückgeben (z.B. "\\PP200\public\Tom")
   Dim pidl As Long
   Dim sPath As String
   Dim bi As BROWSEINFO
   
   bi.hwndOwner = Screen.ActiveForm.hWnd
   
   If Len(sCaption) > 0 Then
      bi.lpszTitle = lstrcat(sCaption, vbNullString)
   End If
   
   bi.ulFlags = BIF_RETURNONLYFSDIRS
   pidl = SHBrowseForFolder(bi)
   
   If pidl Then
      sPath = String$(MAX_PATH, 0)
      SHGetPathFromIDList pidl, sPath
      CoTaskMemFree pidl
      sPath = Left$(sPath, InStr(sPath, vbNullChar) - 1)
   End If
   
   BrowseForFolder = sPath
   
End Function
'==============================================================================

Public Function FileExist(ByVal sFile As String) As Boolean
'------------------------------------------------------------------------------
'Purpose  : Prüft ob eine Datei existiert
'
'Prereq.  : -
'Parameter: sFile -  Datei deren Existenz überprüft werden soll
'Returns  : True = Datei existiert
'           False = Datei existiert nicht
'Note     : -
'
'   Author: Knuth Konrad 17.08.1999
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Dim udtWin32FFD As WIN32_FIND_DATA
   Dim lRetval As Long
   
   lRetval = FindFirstFile(sFile, udtWin32FFD)
   If lRetval = INVALID_HANDLE_VALUE Then
      FileExist = False
      Exit Function
   Else
      lRetval = FindClose(lRetval)
      FileExist = True
   End If
   
End Function
'==============================================================================

Public Sub UnloadAllForms(ByVal frmMainForm As Form)
'------------------------------------------------------------------------------
'Purpose  : Entlädt und terminiert alle Forms eines laufenden Programms
'           außer frmMainForm. Auruf sollte normalerweise im Unload-Event
'           der Hauptform des Programms erfolgen
'
'Prereq.  : -
'Parameter: frmMainForm -  Form von der aus die Funktion aufgerufen wird.
'                          diese Form wird NICHT entladen. Das ist normaler-
'                          weise die Haupt-Programmform.
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 17.08.1999
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Dim frm As Form
   
   For Each frm In Forms
      If frm.Name <> frmMainForm.Name Then
         Unload frm
      End If
   Next frm
   
End Sub
'==============================================================================

Public Function Trim0(ByVal sText As String) As String
'------------------------------------------------------------------------------
'Purpose  : Regular Trim$() and additionaly removes also NUL
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 15.08.2017
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Trim0 = Trim$(Replace$(sText, Chr$(0), vbNullString))
End Function
'==============================================================================

Public Sub CenterFrmOnFrm(childForm As Form, ParentForm As Form)
'------------------------------------------------------------------------------
'Purpose  : Centering Forms/Dialogs in relation to the Parent
'
'Prereq.  : -
'Parameter: -
'Note     : Code to centre any form to it's parent form. This is primarily
'           for centering dialogs within a parent window where the parent
'           may be positioned anywhere on-screen.
'
'           Assumes:Place the CentreFormInParent routine into a bas
'           module sub, then call it from any form, passing the form
'           name to position (the child form name) and the parent
'           form to position within (parent form name.)
'
'           Side Effects:
'           It is often recommended that to centre a form, you should
'           set the form's left and top properties as in:
'           Me.Left = Screen.Width / 2
'           Me.Top = Screen.Height / 2
'
'           While this method is certainly not incorrect, it's execution
'           does involve 2 commands, and, on a slower system or one
'           without accelerated video, the User might see the form
'           shift position first horizontally (as the '.Left =' code
'           is executed), then vertically (as the '.Top =' code is
'           executed.) The Move command used above performs both
'           horizontal and vertical repositioning together in 1 move.
'
'   Author: VB Net (Randy Birch)
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Dim childLeft As Long
   Dim childTop As Long
   Dim parentLeft As Long
   Dim parentTop As Long
   
   parentLeft = ParentForm.Left
   parentTop = ParentForm.Top
   
   childLeft = parentLeft + ((ParentForm.Width - _
      childForm.Width) / 2)
   If childLeft < 0 Then
      childLeft = 0
   End If
   
   childTop = parentTop + ((ParentForm.Height - _
      childForm.Height) / 2.5)
   
   If childForm.WindowState <> FormWindowStateConstants.vbMinimized Then
      If childTop < 0 Then
         childTop = 0
      End If
   End If
   
   If childForm.WindowState = FormWindowStateConstants.vbNormal Then
      childForm.Move childLeft, childTop
   End If

End Sub
'==============================================================================

Public Sub CenterFrmOnScreen(oForm As Form)
'------------------------------------------------------------------------------
'Purpose  : Centers a form on the screen
'
'Prereq.  : -
'Parameter: -
'Note     : -
'
'   Author: Knuth Konrad 07.11.2013
'   Source: Derived from CenterFrmOnFrm
'  Changed: -
'------------------------------------------------------------------------------
   Dim childLeft As Long
   Dim childTop As Long
   Dim parentLeft As Long
   Dim parentTop As Long
   
   parentLeft = 0
   parentTop = 0
   
   childLeft = parentLeft + ((Screen.Width - _
      oForm.Width) / 2)
   If childLeft < 0 Then childLeft = 0
   
   childTop = parentTop + ((Screen.Height - _
      oForm.Height) / 2.5)
   
   If oForm.WindowState = FormWindowStateConstants.vbNormal Then
      oForm.Move childLeft, childTop
   End If
   
   If oForm.WindowState <> FormWindowStateConstants.vbMinimized Then
      If childTop < 0 Then childTop = 0
   End If
   
End Sub
'==============================================================================

Public Function Percent(ByVal dblPart As Double, ByVal dblTotal As Double) As Double
'------------------------------------------------------------------------------
'Purpose  : Returns the % of dblTotal given by dblPart,
'           i.e. dblTotal = 200, dblPart = 50 = 25(%)
'
'Prereq.  : -
'Parameter: dblPart  - Fraction to calculate percent from
'           dblTotal - Total
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 15.08.2017
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   
   If dblTotal <> 0 Then
      Percent = (dblPart / dblTotal) * 100
   Else
      Percent = 0
   End If
   
End Function
'==============================================================================

Public Function CreateTempFileName(Optional sPath As String = ".\", _
   Optional sPrefix As String = vbNullString, _
   Optional sFileExtension As String = "tmp") As String
'------------------------------------------------------------------------------
'Purpose  : Erzeugt einen temporären, nicht vorhanden Dateinamen
'
'Prereq.  : -
'Parameter: sPath          -  Pfad in dem die Datei erzeugt werden soll
'           sPrefix        -  Max. 3 Zeichen für den Anfang des Dateinamen
'           sFileExtension -  Endung der Datei
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 26.08.1999
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Dim sFile As String
   Dim sTemp As String
   Dim lRetval As Long
   
   On Error GoTo CreateTempFileNameError
   
   sTemp = Space$(MAX_PATH)
   'Temporäre Datei erzeugen
   lRetval = GetTempFileName(sPath, sPrefix, 0&, sTemp)
   If lRetval <> 0 Then
      sFile = Left$(sTemp, InStr(sTemp, Chr$(0)) - 1)
      'Da Windows immer eine 0 Byte große Datei anlegt -> löschen
      Kill sFile
      Do
      'Warten bis auch ein langsames Netzwerk die Datei gelöscht hat
         DoEvents
      Loop Until FileExist(sFile) = False
         
   Else
      'Internal Error
      Err.Raise 51
   End If
   
   'GetTempFileName erzeugt immer mit Endung TMP, daher auf andere Endung
   'prüfen
   If sFileExtension <> "tmp" Then
      sFile = Left$(sFile, GetExtPos(sFile) - 1) & Replace(UCase$(sFile), "TMP", sFileExtension, GetExtPos(sFile))
      'Jetzt müssen wir selbst überprüfen ob die Datei schon vorhanden ist
      Do While FileExist(sFile) = True
         sFile = CreateTempFileName(sPath, sPrefix, sFileExtension)
         DoEvents
      Loop
   End If
   
   CreateTempFileName = sFile
   
CreateTempFileNameExit:
   On Error GoTo 0
   Exit Function
   
CreateTempFileNameError:
   CreateTempFileName = vbNullString
   Err.Clear
   Resume CreateTempFileNameExit
   
End Function
'==============================================================================

Public Function GetExtPos(ByVal sSpec As String) As Long
'------------------------------------------------------------------------------
'Purpose  : Ermittelt die Position der Datei-Endung innerhalb eines
'           Dateinamens
'
'Prereq.  : -
'Parameter: sSpec - Dateiname von dem die Position der Endung ermittelt werden soll
'Returns  : Beginn der Dateiendung inkl. "."
'Note     : -
'
'   Author: Bruce McKinney, Hardcore Visual Basic 5
'   Source: -
'  Changed: 13.09.1999, Knuth Konrad
'           Integer in Long geändert
'------------------------------------------------------------------------------
   Dim lLast As Long, lExt As Long
   
   lLast = Len(sSpec)
   
   ' Parse backward to find extension or base
   For lExt = lLast + 1 To 1 Step -1
      
      Select Case Mid$(sSpec, lExt, 1)
      Case "."
         ' First . from right is extension start
         Exit For
      Case "\"
         ' First \ from right is base start
         lExt = lLast + 1
         Exit For
      End Select
   
   Next
   
   ' Negative return indicates no extension, but this
   ' is base so callers don't have to reparse.
   GetExtPos = lExt
   
End Function
'==============================================================================

Public Function ExtractFileName(ByVal sFullname As String, Optional ByVal sPathDelimiter As String = "\", _
   Optional ByVal bolIncludeExtension As Boolean = True) As String
'------------------------------------------------------------------------------
'Purpose  : Extrahiert den kompletten Dateinamen incl. Extension aus einer
'           Pfadangabe
'
'Prereq.  : -
'Parameter: sFullname            - Dateiname incl. Pfadangabe
'           sPathDelimiter       - Standard Pfadbegrenzung
'           bolIncludeExtension  - Auch Dateiendung zurückgeben?
'Returns  : Extrahierter Dateiname
'Note     : -
'
'   Author: Knuth Konrad 21.06.2000
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Dim lBackSlash As Long, sResult As String, lPos As Long
   
   lBackSlash = InStrRev(sFullname, sPathDelimiter)
   
   If lBackSlash > 0 Then
      sResult = Mid$(sFullname, lBackSlash + 1)
   Else
      sResult = sFullname
   End If
   
   If bolIncludeExtension = False Then
   ' Dateiendung nicht zurückgeben
      lPos = GetExtPos(sResult)
      If lPos > 0 Then
         sResult = Left$(sResult, lPos - 1)
      End If
   End If
   
   ExtractFileName = sResult
   
End Function
'==============================================================================

Public Function ExtractPathName(ByVal sFullname As String, Optional ByVal sPathDelimiter As String = "\") As String
'------------------------------------------------------------------------------
'Purpose  : Extrahiert den kompletten Pfad aus einer
'           Pfadangabe
'
'Prereq.  : -
'Parameter: sFullname      -  Dateiname incl. Pfadangabe
'           sPathDelimiter - Standard Pfadbegrenzung
'Returns  : Extrahierter Pfadname incl. abschließendem BackSlash
'Note     : -
'
'   Author: Knuth Konrad 21.06.2000
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Dim lBackSlash As Long
   
   lBackSlash = InStrRev(sFullname, sPathDelimiter)
   
   If lBackSlash > 0 Then
      ExtractPathName = Left$(sFullname, lBackSlash)
   Else
      ExtractPathName = sFullname
   End If
   
End Function
'==============================================================================

Public Function ExtractExtensionName(ByVal sFullname As String) As String
'------------------------------------------------------------------------------
'Purpose  : Extrahiert die Dateiendung (inkl. ".", z.B. ".txt") aus einer
'           Pfadangabe
'
'Prereq.  : -
'Parameter: sFullname   -  Dateiname incl. Pfadangabe
'Returns  : Extrahierte Dateiendung
'Note     : -
'
'   Author: Knuth Konrad 21.06.2000
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Dim lDot As Long
   
   lDot = InStrRev(sFullname, ".")
   
   If lDot > 0 Then
      ExtractExtensionName = Right$(sFullname, Len(sFullname) - (lDot - 1))
   Else
      ExtractExtensionName = "."
   End If
   
End Function
'==============================================================================

Public Function IsOSNT() As Boolean
'------------------------------------------------------------------------------
'Purpose  : Ermittelt ob das OS Windows NT oder Win9x ist
'
'Prereq.  : -
'Parameter: -
'Returns  : True  -  OS ist NT
'           False -  OS ist Win95 oder Win98
'Note     : -
'
'   Author: Knuth Konrad 02.11.1999
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
'Zur Nutzung des OSVERSIONINFO-Struktur ist zunächst ihrem Parameter dwOSVersionInfoSize
'die Größe der Struktur zu übergeben, was wir wie immer mit der VB-Anweisung Len
'bewerkstelligen. Eine Variable dieses Typs kann dann an GetVersionEx übergeben werden, um
'die interessierenden Informationen zu ermitteln:
   
   Dim OS As OSVERSIONINFO
   
   OS.dwOSVersionInfoSize = Len(OS)
   GetVersionEx OS
   
   'Zur Ermittlung des Betriebssystem reicht es aus, die Parameter dwPlatformId, dwMajorVersion
   'und dwMinorVersion auszuwerten: Der Parameter dwMajorVersion trägt für Windows NT 4,
   'Windows 95 und Windows 98 immer den Wert 4.
   '
   'Unter Windows NT ist dwPlatformId immer gleich der Konstanten VER_PLATFORM_WIN32_NT,
   'während diese unter Windows 95 und Windows 98 den Wert von
   'VER_PLATFORM_WIN32_WINDOWS annimmt. In letzterem Fall kann zwischen den beiden durch
   'Auswertung des Parameters dwMinorVersion unterschieden werden: Unter Windows 95 ist dieser
   'gleich 0, unter Windows 98 hingegen beträgt sein Wert 10. Somit ergibt sich:
   
   With OS
      If .dwMajorVersion = 4 Then
      'Windows NT, Windows 95 oder Windows 98
         If .dwPlatformId = VER_PLATFORM_WIN32_NT Then
         'Windows NT
            IsOSNT = True
            Exit Function
         End If
   
         If .dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
            IsOSNT = False
            Exit Function
         'Windows 98 oder Windows 95
   '          If .dwMinorVersion = 0 Then
   '             ' das niederwertige word von dwBuildnumber prüfen
   '             If (.dwBuildNumber And &HFFFF&) > 1000 Then
   '                MsgBox "Windows 95 >= OS R2"
   '             Else
   '                MsgBox "Windows 95 < OS R2"
   '             End If
   '          ElseIf .dwMinorVersion = 10 Then
   '             MsgBox "Windows 98"
   '          End If
         End If
      End If
   End With
   
End Function
'==============================================================================

Public Function ShortenPathText(ByVal sPath As String, _
   ByVal lMaxLen As Long) As String
'------------------------------------------------------------------------------
'Purpose  : Kürzt eine Pfadangabe auf lMaxLen Zeichen
'
'Prereq.  : -
'Parameter: sPath    -  zu kürzende Pfadangabe
'           lMaxLen  -  maximal Länge des Pfades
'Returns  : -
'Note     : -
'
'   Author: Doberenz & Kowalski 26.11.1999
'   Source: Quelle: Visual Basic 6 Kochbuch, Hanser Verlag
'  Changed: ungarische Notation und Stringfunktion statt Variant (Mid, Left...)
'------------------------------------------------------------------------------
   Dim i As Long
   Dim lLen As Long
   
   lLen = Len(sPath)
   
   ShortenPathText = sPath
   
   If Len(sPath) <= lMaxLen Then Exit Function
   
   For i = lLen - lMaxLen + 6 To lLen
      If Mid$(sPath, i, 1) = "\" Then Exit For
   Next
   
   ShortenPathText = Left$(sPath, 3) & "..." & Right$(sPath, lLen - (i - 1))
   
End Function
'==============================================================================

Public Function IsDateAny(ByVal sDate As String) As Boolean
'------------------------------------------------------------------------------
'Purpose  : Überprüft ein Datum auf dessen Gültigkeit
'
'Prereq.  : -
'Parameter: sDate -  Datum im Format dd.mm.yyyy
'Returns  : True = gültiges Datum, False = ungültiges Datum
'Note     : -
'
'   Author: Knuth Konrad 17.01.2000
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Dim st As SYSTEMTIME
   Dim dtmDate As Date
   Dim lRetval As Long
   
   With st
      .wDay = Left$(sDate, 2)
      .wMonth = Mid$(sDate, 4, 2)
      .wYear = Right$(sDate, 4)
   End With
   
   lRetval = SystemTimeToVariantTime(st, dtmDate)
   
   IsDateAny = IsDate(dtmDate) And CBool(lRetval)
   
End Function
'==============================================================================

Public Function Dimension(ByRef vnt As Variant) As Long
'------------------------------------------------------------------------------
'Purpose  : Ermittelt die Anzahl der Dimensionen eines Arrays
'
'Prereq.  : -
'Parameter: vnt   -  Arrayvariabel (z.B. a())
'Returns  : Anzahl der Dimensionen
'Note     : -
'
'   Author: Jost Schwider 21.05.2001
'   Source: http://www.vb-tec.de/arrdim.htm
'  Changed: -
'------------------------------------------------------------------------------
   Dim Ptr As Long
   
   If IsArray(vnt) Then
     Ptr = VarPtr(vnt) + 8      'VB-Array
     RtlMoveMemory Ptr, ByVal Ptr, 4 'SafeArrayDescriptor
     RtlMoveMemory Ptr, ByVal Ptr, 4 'SafeArray-Struktur
     If Ptr Then RtlMoveMemory Dimension, ByVal Ptr, 2
   Else
     Err.Raise 13 'Type mismatch
   End If
   
End Function
'==============================================================================

Public Function GetTempDir() As String
'------------------------------------------------------------------------------
'Purpose  : Ermittelt das TEMP-Verzeichnis
'
'Prereq.  : -
'Parameter: -
'Returns  : Tempverzeichnis mit abschließendem Backslash
'Note     : -
'
'   Author: Knuth Konrad 17.03.2004
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Dim sTemp As String * MAX_PATH
   Dim lRet As Long
   
   sTemp = Space$(MAX_PATH)
   lRet = GetTempPath(Len(sTemp), sTemp)
   If lRet Then
      GetTempDir = NormalizePath(Left$(sTemp, lRet))
   Else
      GetTempDir = NormalizePath(CurDir$)
   End If
   
End Function
'==============================================================================

Public Function ShellAndWait(ByVal Exec As String, Optional ByVal WindowStyle As VbAppWinStyle = vbMinimizedFocus) _
   As Long
'------------------------------------------------------------------------------
'Purpose  : Startet externes Programm und wartet auf Beendigung
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: AboutVB, Harald M. Genauck
'   Source: http://www.aboutvb.de/khw/artikel/khwshell.htm
'  Changed: -
'------------------------------------------------------------------------------
   Dim nTaskId As Long
   Dim nHProcess As Long
   Dim nExitCode As Long
   
   nTaskId = Shell(Exec, WindowStyle)
   nHProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, nTaskId)
   Do
      DoEvents
      GetExitCodeProcess nHProcess, nExitCode
   Loop While nExitCode = STILL_ACTIVE
   CloseHandle nHProcess
   ShellAndWait = nExitCode
   
End Function
'==============================================================================

Public Function ShellAndWaitApi(ByVal sExec As String, _
   Optional ByVal WindowStyle As VbAppWinStyle = vbMinimizedFocus) As Long
'------------------------------------------------------------------------------
'Purpose  : Startet externes Programm und wartet auf Beendigung
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: MS
'   Source: http://support.microsoft.com/kb/129797
'  Changed: -
'------------------------------------------------------------------------------
   Dim udtProc As PROCESS_INFORMATION
   Dim udtStart As STARTUPINFO
   Dim lRet As Long
   
   ' Initialize the STARTUPINFO structure:
   With udtStart
      .dwFlags = STARTF_USESHOWWINDOW
      .wShowWindow = WindowStyle
      .cb = Len(udtStart)
   End With
   
   lRet = CreateProcessA(vbNullString, sExec, 0&, 0&, 1&, _
      NORMAL_PRIORITY_CLASS, 0&, vbNullString, udtStart, udtProc)
   
   ' Wait for the shelled application to finish:
   lRet = WaitForSingleObject(udtProc.hProcess, INFINITE)
   GetExitCodeProcess udtProc.hProcess, lRet
   CloseHandle udtProc.hThread
   CloseHandle udtProc.hProcess
   ShellAndWaitApi = lRet
   
End Function
'==============================================================================

Public Function DateYMD(ByVal dtmDate As Date, Optional ByVal bolAppendTime As Boolean = False, _
   Optional ByVal sDateSep As String = vbNullString, Optional ByVal sDateTimeSep As String = "T") As String
'------------------------------------------------------------------------------
'Purpose  : Erstellt eine String der Form YYYYMMDD
'
'Prereq.  : -
'Parameter: Zu formatierendes Datum
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 24.03.2006
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   
   DateYMD = Format$(Year(dtmDate), "0000") & sDateSep & _
      Format$(Month(dtmDate), "00") & sDateSep & _
      Format$(Day(dtmDate), "00")
   
   If bolAppendTime = True Then
      DateYMD = DateYMD & sDateTimeSep & Format$(Hour(dtmDate), "00") & _
         Format$(Minute(dtmDate), "00") & Format$(Second(dtmDate), "00")
   End If
   
End Function
'==============================================================================

Public Function URLEncode(ByVal sStringToEncode As String, _
   Optional ByVal bolUsePlusRatherThanHexForSpace As Boolean = False) As String
'------------------------------------------------------------------------------
'Purpose  : URLEncode
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Igor
'   Source: http://www.freevbcode.com/ShowCode.asp?ID=1512
'  Changed: -
'------------------------------------------------------------------------------
   Dim sTempAns As String
   Dim lCurChr As Long
   
   lCurChr = 1
   
   Do Until lCurChr - 1 = Len(sStringToEncode)
      Select Case Asc(Mid$(sStringToEncode, lCurChr, 1))
      Case 48 To 57, 65 To 90, 97 To 122
         sTempAns = sTempAns & Mid$(sStringToEncode, lCurChr, 1)
      Case 32
         If bolUsePlusRatherThanHexForSpace = True Then
            sTempAns = sTempAns & "+"
         Else
            sTempAns = sTempAns & "%" & Hex(32)
         End If
      Case Else
         sTempAns = sTempAns & "%" & _
            Format$(Hex(Asc(Mid$(sStringToEncode, _
            lCurChr, 1))), "00")
      End Select
   
      lCurChr = lCurChr + 1
   Loop
   
   URLEncode = sTempAns
   
End Function
'==============================================================================

Public Function URLDecode(ByVal sStringToDecode As String) As String
'------------------------------------------------------------------------------
'Purpose  : URLDecode
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Igor
'   Source: http://www.freevbcode.com/ShowCode.asp?ID=1512
'  Changed: -
'------------------------------------------------------------------------------
   Dim sTempAns As String
   Dim iCurChr As Integer
   
   iCurChr = 1
   
   Do Until iCurChr - 1 = Len(sStringToDecode)
      Select Case Mid$(sStringToDecode, iCurChr, 1)
      Case "+"
         sTempAns = sTempAns & " "
      Case "%"
         sTempAns = sTempAns & Chr$(Val("&h" & Mid$(sStringToDecode, iCurChr + 1, 2)))
         iCurChr = iCurChr + 2
      Case Else
         sTempAns = sTempAns & Mid$(sStringToDecode, iCurChr, 1)
      End Select
   
      iCurChr = iCurChr + 1
   Loop
   
   URLDecode = sTempAns
   
   ' URLDecode function in Perl for reference
   ' both VB and Perl versions must return same
   '
   ' sub urldecode{
   '  local($val)=@_;
   '  $val=~s/\+/ /g;
   '  $val=~s/%([0-9A-H]{2})/pack('C',hex($1))/ge;
   '  return $val;
   ' }
   
End Function
'==============================================================================

Public Function DateBack(ByVal lAmount As Long, ByVal sUnit As String, ByVal dtmDateFrom As Date) As Date
'------------------------------------------------------------------------------
'Purpose  : Calculates a date in the past
'
'Prereq.  : -
'Parameter: lAmount     - Amount to be "substracted" from dtmDateFrom
'           sUnit       - Substract this (VB time) unit, i.e. "d", "h"
'           dtmDateFrom - Source date of calculation
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 29.07.2008
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------

   DateBack = DateAdd(sUnit, -lAmount, DateValue(dtmDateFrom))

End Function
'==============================================================================

Public Function DateTimeSetTime(ByVal dtmDate As Date, Optional ByVal bytHour As Byte = 0, Optional ByVal bytMinute As Byte = 0, _
   Optional ByVal bytSecond As Byte) As Date
'------------------------------------------------------------------------------
'Purpose  : Set the time part of a VB date variable
'
'Prereq.  : -
'Parameter: Time units
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 06.04.2010
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Dim udtSysTime As SYSTEMTIME, dtmTemp As Date
   
   If CBool(VariantTimeToSystemTime(dtmDate, udtSysTime)) = True Then
      With udtSysTime
         .wHour = bytHour
         .wMinute = bytMinute
         .wSecond = bytSecond
      End With
      
      If CBool(SystemTimeToVariantTime(udtSysTime, dtmTemp)) = True Then
         DateTimeSetTime = dtmTemp
      Else
         DateTimeSetTime = dtmDate
      End If
   End If
   
End Function
'==============================================================================

Public Function GetWindowsUserName() As String
'------------------------------------------------------------------------------
'Purpose  : Retrieve the logged on (Windows) user's name
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : Returns the user name of the process owning user
'
'   Author: Knuth Konrad 19.04.2012
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Dim lpszBuff As String, lLen As Long, lRet As Long
   
   'Get the Login User Name
   lpszBuff = Space$(1024)
   lLen = Len(lpszBuff)
   lRet = GetUserName(lpszBuff, lLen)
   
   If lRet > 0 Then
      GetWindowsUserName = Left$(lpszBuff, lLen - 1)
   End If
   
End Function
'==============================================================================

Public Function GetWindowsComputerName() As String
'------------------------------------------------------------------------------
'Purpose  : Retrieve the (local) machine's name
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 19.04.2012
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Dim lpszBuff As String, lLen As Long, lRet As Long
       
   'Get the Computer Name
   lpszBuff = Space$(1024)
   lLen = Len(lpszBuff)
   lRet = GetComputerName(lpszBuff, lLen)
   If lRet > 0 Then
      GetWindowsComputerName = Left$(lpszBuff, lLen)
   End If
   
End Function
'==============================================================================

Public Function Seconds2FormattedTime(Optional ByVal lSeconds As Long = 0, Optional ByVal dtmStartTime As Date, _
   Optional ByVal dtmEndTime As Date) As String
'------------------------------------------------------------------------------
'Purpose  : Return a formatted string of the format hh:nn:ss for a number
'           of seconds
'
'Prereq.  : -
'Parameter: lSeconds       - Number of seconds
'           dtmStartTime, dtmEndTime  - Start & end time from which to calculate
'           the time difference in seconds first
'Returns  : -
'Note     : Works either on lSeconds > 0 *or* start and end date var.
'
'   Author: Knuth Konrad 25.06.2018
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Dim lMinsSecsRemainder As Long
   
   
   If lSeconds = 0 Then
      ' DATES PASSED IN; CALCULATE SECONDS:
      lSeconds = DateDiff("s", dtmStartTime, dtmEndTime)
   End If
   
   
   If lSeconds > 59 Then
   
      If lSeconds > 3599 Then
       
            ' START THINGS OFF WITH THE HIGH-LEVEL HOURS:
            Seconds2FormattedTime = Format(CStr(Fix(lSeconds / 3600)), "00") & ":"
           
            ' FORMAT MINUTES/SECONDS FROM THE REMAINDER AFTER HOURS EXTRACTED:
            lMinsSecsRemainder = lSeconds Mod 3600
            If lMinsSecsRemainder > 59 Then
               Seconds2FormattedTime = Seconds2FormattedTime & _
                                    Format(CStr(Fix(lMinsSecsRemainder / 60)), "00") & ":" & _
                                    Format(CStr(lSeconds Mod 60), "00")
            End If
       
       Else
           ' JUST DO MINUTES/SECONDS:
           Seconds2FormattedTime = Format(CStr(Fix(lSeconds / 60)), "00") & ":" & _
                                Format(CStr(lSeconds Mod 60), "00")
       End If
       
   Else
       ' JUST DO SECONDS:
       Seconds2FormattedTime = "00:" & Format(lSeconds, "00")
   End If
   
End Function
'==============================================================================

Public Function Seconds2FormattedTime1(ByVal lSeconds As Long) As String
'------------------------------------------------------------------------------
'Purpose  : Calculates and formats an amount of seconds into days, hours, minutes, seconds
'
'Prereq.  : -
'Parameter: lSeconds - Number of seconds
'Returns  : Formatted time string
'Note     : -
'
'   Author: Knuth Konrad 06.10.2017
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Dim lDays As Long, lHours As Long, lMinutes As Long
   Dim sResult As String
   
   ' Sekunde -> Minute
   lMinutes = lSeconds / 60
   lSeconds = lSeconds Mod 60
   sResult = Format$(lSeconds, "#0 second(s)")
   
   ' Minute -> Stunde
   If lMinutes > 0 Then
   
      lHours = lMinutes / 60
      lMinutes = lMinutes Mod 60
      sResult = Format$(lMinutes, "#0 minute(s)") & ", " & sResult
      
   End If
   
   ' Stunde -> Tag
   If lHours > 0 Then
   
      lDays = lHours / 24
      lHours = lHours Mod 24
      sResult = Format$(lHours, "#0 hour(s)") & ", " & sResult
      
   End If
   
   If lDays > 0 Then
   
      Seconds2FormattedTime1 = Format$(lDays, "#0 day(s)") & ", " & sResult
   
   Else
   
      Seconds2FormattedTime1 = sResult
      
   End If
   
End Function
'==============================================================================

Public Function GetLastAPIDLLStr(ByRef lAPIErrorCode As Long) As String
'------------------------------------------------------------------------------
'Purpose  : Encapsulation of Win32 GetLastError & FormatMessage to retrieve the
'           error code & textual error message.
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 10.10.2013
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Dim lRetval As Long
   Dim sBuffer As String
   
   lAPIErrorCode = Err.LastDllError
   sBuffer = Space$(FORMAT_MESSAGE_MAX_WIDTH_MASK)
       
   lRetval = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_MAX_WIDTH_MASK, _
      ByVal vbNull, lAPIErrorCode, vbNull, sBuffer, FORMAT_MESSAGE_MAX_WIDTH_MASK, ByVal vbNull)
                           
   If lRetval Then
   ' nachfolgende vbNullChar in Fehlertext abschneiden
      GetLastAPIDLLStr = Left$(sBuffer, lRetval)
   Else
      GetLastAPIDLLStr = vbNullString
   End If
   
End Function
'==============================================================================

Public Function GetAPIErrorStr(ByVal lAPIErrorCode As Long) As String
'------------------------------------------------------------------------------
'Purpose  : Encapsulation of Win32 GetLastError & FormatMessage to retrieve the
'           error code & textual error message.
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 10.10.2013
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Dim lRetval As Long
   Dim sBuffer As String
   
   sBuffer = Space$(FORMAT_MESSAGE_MAX_WIDTH_MASK)
       
   lRetval = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_MAX_WIDTH_MASK, _
      ByVal vbNull, lAPIErrorCode, vbNull, sBuffer, FORMAT_MESSAGE_MAX_WIDTH_MASK, ByVal vbNull)
                           
   If lRetval Then
   ' nachfolgende vbNullChar in Fehlertext abschneiden
      GetAPIErrorStr = Left$(sBuffer, lRetval)
   Else
      GetAPIErrorStr = vbNullString
   End If
   
End Function
'==============================================================================

Public Function CreateObject(ByVal sClass As String, Optional ByVal sServerName As String = vbNullString) As Object
'------------------------------------------------------------------------------
'Purpose  : Override the CreateObject function in order to register what object
'           is being created in any error message that's generated.
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Darin Higgins
'   Source: http://www.darinhiggins.com/the-vb6-createobject-function/
'  Changed: 10.04.2014
'           - Reformatted to accomodate my coding style
'------------------------------------------------------------------------------
   Dim sSource As String, sDescr As String, lErrNum As Long
   
   On Error Resume Next
   
   If Len(sServerName) Then
      Set CreateObject = VBA.CreateObject(sClass, sServerName)
   Else
      Set CreateObject = VBA.CreateObject(sClass)
   End If
   
   If VBA.Err Then
      sSource = VBA.Err.Source
      sDescr = VBA.Err.Description
      lErrNum = VBA.Err
      sDescr = sDescr & " (ProgID: " & sClass
      If Len(sServerName) Then
         sDescr = sDescr & ". Instantiated on Server '" & sServerName & "'"
      End If
      sDescr = sDescr & ")"
      
      On Error GoTo 0
      
      VBA.Err.Raise lErrNum, sSource, sDescr
   End If
   
   On Error GoTo 0
   
End Function
'==============================================================================

Public Function CreateTimeStamp(Optional ByVal dtmDate As Date, Optional ByVal sDelim As String = vbNullString, _
   Optional ByVal bolDateOnly As Boolean = False, Optional sFormat As String = "yyyymmddhhnnss") As String
'------------------------------------------------------------------------------
'Purpose  : Kreiert einen TimeStampstring der Form YYYYMMDDHHNNSS
'
'Prereq.  : -
'Parameter: dtmDate     - TimeStamp kreieren von
'           sDelim      - Trennzeichen für Datumseinheiten, z.B. "."
'           bolDateOnly - Nur den Datumsanteil formatieren, nicht die Zeit
'           sFormat     - Formatstring
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 25.07.2007
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Dim dtmTemp As Date
   Dim sResult As String
   
   Const PROCEDURE_NAME As String = "PUtil:CreateTimeStamp->"
   
   If IsMissing(dtmDate) Then
      dtmTemp = Now
   Else
      dtmTemp = dtmDate
   End If
   
   If Len(sDelim) > 0 Then
      
      sResult = Format$(dtmTemp, "yyyy") & sDelim & _
         Format$(dtmTemp, "mm") & sDelim & _
         Format$(dtmTemp, "dd")
      
      If bolDateOnly = True Then
         CreateTimeStamp = sResult
      Else
         CreateTimeStamp = sResult & sDelim & _
         Format$(dtmTemp, "hhnnss")
      End If
   
   Else
   
      CreateTimeStamp = Format$(dtmTemp, sFormat)
   
   End If
   
End Function
'==============================================================================

Public Function DateTimeCompare(ByVal dtmDateBase As Date, ByVal dtmDateToCompare As Date, _
   Optional ByVal bolDateOnly As Boolean = True) As Long
'------------------------------------------------------------------------------
'Purpose  : Compares two dates with each other
'
'Prereq.  : -
'Parameter: -
'Returns  : -1 - dtmDateToCompare < dtmDateBase
'            0 - dtmDateToCompare = dtmDateBase
'            1 - dtmDateToCompare > dtmDateBase
'Note     : -
'
'   Author: Knuth Konrad 25.10.2014
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Dim udtDateBase As SYSTEMTIME, udtDateToCompare As SYSTEMTIME
   
   If bolDateOnly = True Then
   ' Compare the date part ONLY
      
      With udtDateBase
         .wDay = Day(dtmDateBase)
         .wMonth = Month(dtmDateBase)
         .wYear = Year(dtmDateBase)
      End With
      SystemTimeToVariantTime udtDateBase, dtmDateBase
      
      With udtDateToCompare
         .wDay = Day(dtmDateToCompare)
         .wMonth = Month(dtmDateToCompare)
         .wYear = Year(dtmDateToCompare)
      End With
      SystemTimeToVariantTime udtDateToCompare, dtmDateToCompare
   
   End If
   
   DateTimeCompare = Sgn(dtmDateToCompare - dtmDateBase)
   
End Function
'==============================================================================

Public Function DateTimeNewDate(ByVal lYear As Long, ByVal lMonth As Long, ByVal lDay As Long, _
   Optional ByVal lHour As Long = 0, Optional ByVal lMinute As Long = 0, _
   Optional ByVal lSecond As Long = 0, Optional ByVal lMilliSecond As Long = 0) _
   As Date
'------------------------------------------------------------------------------
'Purpose  : Creates a VB date from given values
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 25.10.2014
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Dim udt As SYSTEMTIME
   Dim dtmResult As Date
   
   With udt
      .wYear = lYear
      .wMonth = lMonth
      .wDay = lDay
      .wHour = lHour
      .wMinute = lMinute
      .wSecond = lSecond
      .wMilliseconds = lMilliSecond
   End With
   
   SystemTimeToVariantTime udt, dtmResult
   
   DateTimeNewDate = dtmResult
   
End Function
'==============================================================================

Public Function PtrObj(ByVal lpPtr As Long) As Object
'------------------------------------------------------------------------------
'Purpose  : Reverse of VB's (undocumented) ObjPtr - (re)create an object
'           from a pointer
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Unknown
'   Source: http://brianmstafford.blogspot.de/2010/03/objptr-and-ptrobj.html
'  Changed: -
'------------------------------------------------------------------------------
   Dim o As Object
   
   If lpPtr <> 0 Then
      CopyMemory o, lpPtr, 4
      Set PtrObj = o
      CopyMemory o, 0&, 4
   End If
   
End Function
'==============================================================================

Public Function IsIDE() As Boolean
'------------------------------------------------------------------------------
'Purpose  : Test if code runs within the IDE or as compiled exe
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Filyus
'   Source: http://stackoverflow.com/questions/9052024/debug-mode-in-vb-6
'  Changed: -
'------------------------------------------------------------------------------
  
   On Error Resume Next
   
   Debug.Print 0 / 0
   IsIDE = Err.Number <> 0
   
   Err.Clear

End Function
'==============================================================================
