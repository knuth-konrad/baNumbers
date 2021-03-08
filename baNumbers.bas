#If 0
----------------------------------------------------------------------------
Title          baNumbers
Author         Knuth Konrad
Language       PB/WIN 10.04
Date           10.09.2014
Purpose        Numbers/Math helpers
Update

Copyright (c)  None
               Knuth Konrad
----------------------------------------------------------------------------
#EndIf

'----------------------------------------------------------------------------
'*** PROGRAMM/COMPILE OPTIONEN ***

#Compiler PBWin 10
#Compile DLL ".\baNumbers.dll"

#Dim All
#Tools On
#Debug Error On

' Version resource file
#Include ".\baNumbersRes.inc"

DefLng A-Z

'----------------------------------------------------------------------------
'*** CONSTANTS ***
%S_Ok                                   = &H00000000
%S_False                                = &H00000001

%False = 0
%True = Not %False

%DATA_TYPE_BYTE = 0
%DATA_TYPE_CURRENCY = 1
%DATA_TYPE_DOUBLE = 2
%DATA_TYPE_INTEGER = 3
%DATA_TYPE_LONG = 4
%DATA_TYPE_SINGLE = 5
%DATA_TYPE_STRING = 6
'----------------------------------------------------------------------------
'*** #INCLUDEs ***
#Include Once "WinNT.inc"
#Include Once "VBAPI32.inc"
#Include Once "sautil.inc"
'----------------------------------------------------------------------------
'*** DECLAREs ***
'----------------------------------------------------------------------------
'*** TYPEs ***
Enum LocaleStringConstants
  locDigits = &H11
  locCurrency = &H14
  locCurSymbol = &H15
  locDate = &H1D
  locDecimal = &HE
  locList = &HC
  locMoneyDecimal = &H16
  locMoneyThousands = &H17
  locNegative = &H51
  locPositive = &H50
  locThousands = &HF
  locTime = &H1E
End Enum

Union uInteger2Word
   iValue As Integer
   wValue As Word
End Union

Union uInteger2DWord
   iValue As Integer
   dwValue As Dword
End Union

Union uLong2DWord
   lValue As Long
   dwValue As Dword
End Union

Union uLong2Quad
   lValue As Long
   qudValue As Quad
End Union
'----------------------------------------------------------------------------
'*** VARIABELN ***
' Global arrays for baSort2Arrays()

Global gaByte1(), gaByte2() As Byte
Global gaCurrency1(), gaCurrency2() As Double
Global gaDouble1(), gaDouble2() As Double
Global gaInteger1(),gaInteger2() As Integer
Global gaLong1(), gaLong2() As Long
Global gaSingle1(), gaSingle2() As Single
Global gaString1(), gaString2() As String
'----------------------------------------------------------------------------

Function LibMain(ByVal hInstance   As Long, _
                 ByVal fwdReason   As Long, _
                 ByVal lpvReserved As Long) Export As Long

Select Case fwdReason

Case %DLL_PROCESS_ATTACH
'Indicates that the DLL is being loaded by another process (a DLL
'or EXE is loading the DLL).  DLLs can use this opportunity to
'initialize any instance or global data, such as arrays.

   Trace New ".\baNumbers.tra"
   Trace On

   LibMain = 1   'success!

   'LibMain = 0   'failure!
   Exit Function

Case %DLL_PROCESS_DETACH
'Indicates that the DLL is being unloaded or detached from the
'calling application.  DLLs can take this opportunity to clean
'up all resources for all threads attached and known to the DLL.

   Trace Off
   Trace Close

   LibMain = 1   'success!

   'LibMain = 0   'failure!
   Exit Function

Case %DLL_THREAD_ATTACH
'Indicates that the DLL is being loaded by a new thread in the
'calling application.  DLLs can use this opportunity to
'initialize any thread local storage (TLS).

   LibMain = 1   'success!

   'LibMain = 0   'failure!
   Exit Function

Case %DLL_THREAD_DETACH
'Indicates that the thread is exiting cleanly.  If the DLL has
'allocated any thread local storage, it should be released.

   LibMain = 1   'success!

   'LibMain = 0   'failure!
   Exit Function

End Select

' Any message which is not handled in the above SELECT CASE reaches
' this point and is unknown.

End Function
'----------------------------------------------------------------------------
Rem <MKVBDEC>:baNumbers.dll
'----------------------------------------------------------------------------
' SortByte - Sort a single-dimension VB byte array
'
Sub baSortByte Alias "baSortByte" (psa As Dword, Optional ByVal lDescending As Long) Export

    Local l  As Long
    Local u  As Long
    Local vb As Dword

    Trace On

    l  = vbArrayLBound(psa, 1)
    u  = vbArrayUBound(psa, 1)
    vb = vbArrayFirstElem(psa)

    Dim vba(l To u) As Byte At vb

    If IsTrue(lDescending) Then
       Array Sort vba(), Descend
    Else
       Array Sort vba()
    End If

End Sub
'==============================================================================

'----------------------------------------------------------------------------
' SortInteger - Sort a single-dimension VB integer array
'
Sub baSortInteger Alias "baSortInteger" (psa As Dword, Optional ByVal lDescending As Long) Export

    Local l  As Long
    Local u  As Long
    Local vb As Dword

    Trace On

    l  = vbArrayLBound(psa, 1)
    u  = vbArrayUBound(psa, 1)
    vb = vbArrayFirstElem(psa)

    Dim vba(l To u) As Integer At vb

    If IsTrue(lDescending) Then
       Array Sort vba(), Descend
    Else
       Array Sort vba()
    End If

End Sub
'==============================================================================

'------------------------------------------------------------------------------
' SortLong - Sort a single-dimension VB long integer array
'
Sub baSortLong Alias "baSortLong" (psa As Dword, Optional ByVal lDescending As Long) Export

    Local l  As Long
    Local u  As Long
    Local vb As Dword

    Trace On

    l  = vbArrayLBound(psa, 1)
    u  = vbArrayUBound(psa, 1)
    vb = vbArrayFirstElem(psa)

    Dim vba(l To u) As Long At vb

    If IsTrue(lDescending) Then
       Array Sort vba(), Descend
    Else
       Array Sort vba()
    End If

End Sub
'==============================================================================

'------------------------------------------------------------------------------
' SortSingle - Sort a single-dimension VB single-precision array
'
Sub baSortSingle Alias "baSortSingle" (psa As Dword, Optional ByVal lDescending As Long) Export

    Local l  As Long
    Local u  As Long
    Local vb As Dword

    Trace On

    l  = vbArrayLBound(psa, 1)
    u  = vbArrayUBound(psa, 1)
    vb = vbArrayFirstElem(psa)

    Dim vba(l To u) As Single At vb

    If IsTrue(lDescending) Then
       Array Sort vba(), Descend
    Else
       Array Sort vba()
    End If

End Sub
'==============================================================================

'------------------------------------------------------------------------------
' SortDouble - Sort a single-dimension VB double-precision array
'
Sub baSortDouble Alias "baSortDouble" (psa As Dword, Optional ByVal lDescending As Long) Export

    Local l  As Long
    Local u  As Long
    Local vb As Dword

    Trace On

    l  = vbArrayLBound(psa, 1)
    u  = vbArrayUBound(psa, 1)
    vb = vbArrayFirstElem(psa)

    Dim vba(l To u) As Double At vb

    If IsTrue(lDescending) Then
       Array Sort vba(), Descend
    Else
       Array Sort vba()
    End If

End Sub
'==============================================================================

'------------------------------------------------------------------------------
' SortCurrency - Sort a single-dimension VB currency array
'
Sub baSortCurrency Alias "baSortCurrency" (psa As Dword, Optional ByVal lDescending As Long) Export

    Local l  As Long
    Local u  As Long
    Local vb As Dword

    Trace On

    l  = vbArrayLBound(psa, 1)
    u  = vbArrayUBound(psa, 1)
    vb = vbArrayFirstElem(psa)

    Dim vba(l To u) As Currency At vb

    If IsTrue(lDescending) Then
       Array Sort vba(), Descend
    Else
       Array Sort vba()
    End If

End Sub
'==============================================================================

Function baMedianByte Alias "baMedianByte" (psa As Dword) Export As Byte
'------------------------------------------------------------------------------
'Purpose  : Determines the median
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 10.09.2014
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local lRecCount, i, lCenter, lOffSet As Long
   Local iX, iY As Byte

   Local l  As Long
   Local u  As Long
   Local vb As Dword


   Trace On

   l  = vbArrayLBound(psa, 1)
   u  = vbArrayUBound(psa, 1)
   vb = vbArrayFirstElem(psa)

   Try
      Dim vba(l To u) As Byte At vb
   Catch
      Function = 0
      Exit Function
   End Try

   Array Sort vba()

   If (LBound(vba) = 0) And (UBound(vba) = 0) Then
      Function = vba(0)
      Exit Function
   End If

   lRecCount = (UBound(vba) + 1) - (LBound(vba) + 1)
   lCenter = lRecCount \ 2

   If lCenter <> 0 Then
      lOffSet = ((lRecCount + 1) / 2)
      Function = vba(lOffSet)
   Else
      lOffSet = (lRecCount / 2)
      iX = vba(lOffSet)
      iY = vba(lOffSet + 1)
      Function = CByt((iX + iY) / 2)
   End If

End Function
'==============================================================================

Function baMedianInteger Alias "baMedianInteger" (psa As Dword) Export As Integer
'------------------------------------------------------------------------------
'Purpose  : Determines the median
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 10.09.2014
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local lRecCount, i, lCenter, lOffSet As Long
   Local iX, iY As Integer

   Local l  As Long
   Local u  As Long
   Local vb As Dword


   Trace On

   l  = vbArrayLBound(psa, 1)
   u  = vbArrayUBound(psa, 1)
   vb = vbArrayFirstElem(psa)

   Try
      Dim vba(l To u) As Integer At vb
   Catch
      Function = 0
      Exit Function
   End Try

   Array Sort vba()

   If (LBound(vba) = 0) And (UBound(vba) = 0) Then
      Function = vba(0)
      Exit Function
   End If

   lRecCount = (UBound(vba) + 1) - (LBound(vba) + 1)
   lCenter = lRecCount \ 2

   If lCenter <> 0 Then
      lOffSet = ((lRecCount + 1) / 2)
      Function = vba(lOffSet)
   Else
      lOffSet = (lRecCount / 2)
      iX = vba(lOffSet)
      iY = vba(lOffSet + 1)
      Function = CInt((iX + iY) / 2)
   End If

End Function
'==============================================================================

Function baMedianLong Alias "baMedianLong" (psa As Dword) Export As Long
'------------------------------------------------------------------------------
'Purpose  : Determines the median
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 10.09.2014
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local lRecCount, i, lCenter, lOffSet As Long
   Local lX, lY As Long

   Local l  As Long
   Local u  As Long
   Local vb As Dword

   Trace On

   l  = vbArrayLBound(psa, 1)
   u  = vbArrayUBound(psa, 1)
   vb = vbArrayFirstElem(psa)

   Try
      Dim vba(l To u) As Long At vb
   Catch
      Function = 0
      Exit Function
   End Try

   Array Sort vba()

   If (LBound(vba) = 0) And (UBound(vba) = 0) Then
      Function = vba(0)
      Exit Function
   End If

   lRecCount = (UBound(vba) + 1) - (LBound(vba) + 1)
   lCenter = lRecCount \ 2

   If lCenter <> 0 Then
      lOffSet = ((lRecCount + 1) / 2)
      Function = vba(lOffSet)
   Else
      lOffSet = (lRecCount / 2)
      lX = vba(lOffSet)
      lY = vba(lOffSet + 1)
      Function = CLng((lX + lY) / 2)
   End If

End Function
'==============================================================================

Function baMedianSingle Alias "baMedianSingle" (psa As Dword) Export As Single
'------------------------------------------------------------------------------
'Purpose  : Determines the median
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 10.09.2014
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local lRecCount, i, lCenter, lOffSet As Long
   Local fX, fY As Single

   Local l  As Long
   Local u  As Long
   Local vb As Dword

   Trace On

   l  = vbArrayLBound(psa, 1)
   u  = vbArrayUBound(psa, 1)
   vb = vbArrayFirstElem(psa)

   Try
      Dim vba(l To u) As Single At vb
   Catch
      Function = 0
      Exit Function
   End Try

   Array Sort vba()

   If (LBound(vba) = 0) And (UBound(vba) = 0) Then
      Function = vba(0)
      Exit Function
   End If

   lRecCount = (UBound(vba) + 1) - (LBound(vba) + 1)
   lCenter = lRecCount \ 2

   If lCenter <> 0 Then
      lOffSet = ((lRecCount + 1) / 2)
      Function = vba(lOffSet)
   Else
      lOffSet = (lRecCount / 2)
      fX = vba(lOffSet)
      fY = vba(lOffSet + 1)
      Function = CSng((fX + fY) / 2)
   End If

End Function
'==============================================================================

Function baMedianDouble Alias "baMedianDouble" (psa As Dword) Export As Double
'------------------------------------------------------------------------------
'Purpose  : Determines the median
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 10.09.2014
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local lRecCount, i, lCenter, lOffSet As Long
   Local dblX, dblY As Double

   Local l  As Long
   Local u  As Long
   Local vb As Dword

   Trace On

   l  = vbArrayLBound(psa, 1)
   u  = vbArrayUBound(psa, 1)
   vb = vbArrayFirstElem(psa)

   Try
      Dim vba(l To u) As Double At vb
   Catch
      Function = 0
      Exit Function
   End Try

   Array Sort vba()

   If (LBound(vba) = 0) And (UBound(vba) = 0) Then
      Function = vba(0)
      Exit Function
   End If

   lRecCount = (UBound(vba) + 1) - (LBound(vba) + 1)
   lCenter = lRecCount \ 2

   If lCenter <> 0 Then
      lOffSet = ((lRecCount + 1) / 2)
      Function = vba(lOffSet)
   Else
      lOffSet = (lRecCount / 2)
      dblX = vba(lOffSet)
      dblY = vba(lOffSet + 1)
      Function = CDbl((dblX + dblY) / 2)
   End If

End Function
'==============================================================================

Function baMedianCurrency Alias "baMedianCurrency" (psa As Dword) Export As Currency
'------------------------------------------------------------------------------
'Purpose  : Determines the median
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 10.09.2014
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local lRecCount, i, lCenter, lOffSet As Long
   Local curX, curY As Currency

   Local l  As Long
   Local u  As Long
   Local vb As Dword

   Trace On

   l  = vbArrayLBound(psa, 1)
   u  = vbArrayUBound(psa, 1)
   vb = vbArrayFirstElem(psa)

   Try
      Dim vba(l To u) As Currency At vb
   Catch
      Function = 0
      Exit Function
   End Try

   Array Sort vba()

   If (LBound(vba) = 0) And (UBound(vba) = 0) Then
      Function = vba(0)
      Exit Function
   End If

   lRecCount = (UBound(vba) + 1) - (LBound(vba) + 1)
   lCenter = lRecCount \ 2

   If lCenter <> 0 Then
      lOffSet = ((lRecCount + 1) / 2)
      Function = vba(lOffSet)
   Else
      lOffSet = (lRecCount / 2)
      curX = vba(lOffSet)
      curY = vba(lOffSet + 1)
      Function = CCur((curX + curY) / 2)
   End If

End Function
'==============================================================================

Function baFormatNumber Alias "baFormatNumber" (ByVal curNumber As Currency, _
   ByVal wLangLocale As Word, ByVal wSubLangLocale As Word) Export As String

   Local lpzInputValue  As AsciiZ * 12  '18 digits, leading zero, optional leading minus, and decimal point.
   Local lpzOutputValue As AsciiZ * 40  'additional room provided for commas, etc.
   Local dwLangID As Dword

   Trace On

   dwLangID = MAKELANGID(wLangLocale, wSubLangLocale)
   lpzInputValue = LTrim$(Str$(curNumber, 10))

   GetNumberFormat dwLangID, ByVal 0, lpzInputValue, ByVal 0, lpzOutputValue, ByVal 40

   Function = lpzOutputValue

End Function
'===========================================================================

Function baFormatNumberEx Alias "baFormatNumberEx" (ByVal curNumber As Currency, _
   ByVal dwLangID As Dword) Export As String

   Local lpzInputValue  As AsciiZ * 12  '18 digits, leading zero, optional leading minus, and decimal point.
   Local lpzOutputValue As AsciiZ * 40  'additional room provided for commas, etc.

   Trace On

   lpzInputValue = LTrim$(Str$(curNumber, 10))

   GetNumberFormat dwLangID, ByVal 0, lpzInputValue, ByVal 0, lpzOutputValue, ByVal 40

   Function = lpzOutputValue

End Function
'===========================================================================

Function baFracSingle Alias "baFracSingle" (ByVal fValue As Single) Export As Single

   Function = Frac(fValue)

End Function
'===========================================================================

Function baFracDouble Alias "baFracDouble" (ByVal dblValue As Double) Export As Double

   Function = Frac(dblValue)

End Function
'===========================================================================

Function baFracCur Alias "baFracCur" (ByVal curValue As Currency) Export As Currency

   Function = Frac(curValue)

End Function
'===========================================================================

Function LocaleString Alias "LocaleString" (ByVal dwLCID As Dword, ByVal eInfo As Long) Export As String

  Dim szLocale As AsciiZ * 11
  Dim nLen As Long

  ' GetUserDefaultLCID()
  nLen = GetLocaleInfo(dwLCID, eInfo, szLocale, SizeOf(szLocale))
  LocaleString = Left$(szLocale, nLen - 1)

End Function
'===========================================================================

Function LCIDFromLangID Alias "LCIDFromLangID" (ByVal dwLangID As Dword) Export As Dword

   LCIDFromLangID = MAKELCID(dwLangID, %SORT_DEFAULT)

End Function
'===========================================================================

Function baSwapByte Alias "baSwapByte" (ByRef v1 As Byte, ByRef v2 As Byte) Export As Long
'------------------------------------------------------------------------------
'Purpose  : Swaps the contents of two variables
'
'Prereq.  : -
'Parameter: The two variables whose contens should be swapped
'Returns  : %True (Success) or %False (Error/Failure)
'Note     : -
'
'   Author: Knuth Konrad 19.04.2016
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------

   Local i, x, y As Long
   Local s As String

   Trace On
   Trace Print FuncName$

   For x = CallStkCount To 1 Step -1
      s = s & CallStk$(x)
   Next x
   Trace Print s

   Try
      Swap v1, v2
      baSwapByte = %True
   Catch
      baSwapByte = %False
      Trace Print Error$(ErrClear)
   End Try

   Trace Off

End Function
'===========================================================================

Function baSwapCurrency Alias "baSwapCurrency" (ByRef v1 As Currency, ByRef v2 As Currency) Export As Long
'------------------------------------------------------------------------------
'Purpose  : Swaps the contents of two variables
'
'Prereq.  : -
'Parameter: The two variables whose contens should be swapped
'Returns  : %True (Success) or %False (Error/Failure)
'Note     : -
'
'   Author: Knuth Konrad 19.04.2016
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------

   Local i, x, y As Long
   Local s As String

   Trace On
   Trace Print FuncName$

   For x = CallStkCount To 1 Step -1
      s = s & CallStk$(x)
   Next x
   Trace Print s

   Try
      Swap v1, v2
      baSwapCurrency = %True
   Catch
      baSwapCurrency = %False
      Trace Print Error$(ErrClear)
   End Try

   Trace Off

End Function
'===========================================================================

Function baSwapDouble Alias "baSwapDouble" (ByRef v1 As Double, ByRef v2 As Double) Export As Long
'------------------------------------------------------------------------------
'Purpose  : Swaps the contents of two variables
'
'Prereq.  : -
'Parameter: The two variables whose contens should be swapped
'Returns  : %True (Success) or %False (Error/Failure)
'Note     : -
'
'   Author: Knuth Konrad 19.04.2016
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------

   Local i, x, y As Long
   Local s As String

   Trace On
   Trace Print FuncName$

   For x = CallStkCount To 1 Step -1
      s = s & CallStk$(x)
   Next x
   Trace Print s

   Try
      Swap v1, v2
      baSwapDouble = %True
   Catch
      baSwapDouble = %False
      Trace Print Error$(ErrClear)
   End Try

   Trace Off

End Function
'===========================================================================

Function baSwapInteger Alias "baSwapInteger" (ByRef v1 As Integer, ByRef v2 As Integer) Export As Long
'------------------------------------------------------------------------------
'Purpose  : Swaps the contents of two variables
'
'Prereq.  : -
'Parameter: The two variables whose contens should be swapped
'Returns  : %True (Success) or %False (Error/Failure)
'Note     : -
'
'   Author: Knuth Konrad 19.04.2016
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------

   Local i, x, y As Long
   Local s As String

   Trace On
   Trace Print FuncName$

   For x = CallStkCount To 1 Step -1
      s = s & CallStk$(x)
   Next x
   Trace Print s

   Try
      Swap v1, v2
      baSwapInteger = %True
   Catch
      baSwapInteger = %False
      Trace Print Error$(ErrClear)
   End Try

   Trace Off

End Function
'===========================================================================

Function baSwapLong Alias "baSwapLong" (ByRef v1 As Long, ByRef v2 As Long) Export As Long
'------------------------------------------------------------------------------
'Purpose  : Swaps the contents of two variables
'
'Prereq.  : -
'Parameter: The two variables whose contens should be swapped
'Returns  : %True (Success) or %False (Error/Failure)
'Note     : -
'
'   Author: Knuth Konrad 19.04.2016
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------

   Local i, x, y As Long
   Local s As String

   Trace On
   Trace Print FuncName$

   For x = CallStkCount To 1 Step -1
      s = s & CallStk$(x)
   Next x
   Trace Print s

   Try
      Swap v1, v2
      baSwapLong = %True
   Catch
      baSwapLong = %False
      Trace Print Error$(ErrClear)
   End Try

   Trace Off

End Function
'===========================================================================

Function baSwapSingle Alias "baSwapSingle" (ByRef v1 As Single, ByRef v2 As Single) Export As Long
'------------------------------------------------------------------------------
'Purpose  : Swaps the contents of two variables
'
'Prereq.  : -
'Parameter: The two variables whose contens should be swapped
'Returns  : %True (Success) or %False (Error/Failure)
'Note     : -
'
'   Author: Knuth Konrad 19.04.2016
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------

   Local i, x, y As Long
   Local s As String

   Trace On
   Trace Print FuncName$

   For x = CallStkCount To 1 Step -1
      s = s & CallStk$(x)
   Next x
   Trace Print s

   Try
      Swap v1, v2
      baSwapSingle = %True
   Catch
      baSwapSingle = %False
      Trace Print Error$(ErrClear)
   End Try

   Trace Off

End Function
'===========================================================================

Function Int2Wrd Alias "Int2Wrd" (ByVal iValue As Integer) Export As Long
'------------------------------------------------------------------------------
'Purpose  : Converts a (negative) Int value to its (positive) Word value
'
'Prereq.  : -
'Parameter: iValue - Integer value
'Returns  : the (positive) Word value as a Long
'Note     : -
'
'   Author: Knuth Konrad 21.09.2016
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local u As uInteger2Word

   u.iValue = iValue
   Int2Wrd = u.wValue

End Function

Function Int2DWrd Alias "Int2DWrd" (ByVal iValue As Integer) Export As Currency
'------------------------------------------------------------------------------
'Purpose  : Converts a (negative) Int value to its (positive) DWord value
'
'Prereq.  : -
'Parameter: iValue - Integer value
'Returns  : the (positive) DWord value as a Currency
'Note     : -
'
'   Author: Knuth Konrad 21.09.2016
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local u As uInteger2DWord

   u.iValue = iValue
   Int2DWrd = u.dwValue

End Function

Function Lng2DWrd Alias "Lng2DWrd" (ByVal lValue As Long) Export As Currency
'------------------------------------------------------------------------------
'Purpose  : Converts a (negative) Long value to its (positive) Word value
'
'Prereq.  : -
'Parameter: lValue - Long value
'Returns  : the (positive) DWord value as a Currency
'Note     : -
'
'   Author: Knuth Konrad 21.09.2016
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local u As uLong2DWord

   u.lValue = lValue
   Lng2DWrd = u.dwValue

End Function

Function Lng2Quad Alias "Lng2Quad" (ByVal lValue As Long) Export As Currency
'------------------------------------------------------------------------------
'Purpose  : Converts a (negative) Long value to its (positive) Quad value
'
'Prereq.  : -
'Parameter: lValue - Long value
'Returns  : the (positive) Quad value as a Currency
'Note     : -
'
'   Author: Knuth Konrad 21.09.2016
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local u As uLong2Quad

   u.lValue = lValue
   Lng2Quad = u.qudValue

End Function

Function TypeOfVariant Alias "TypeOfVariant" (ByVal vntVar As Variant) Export As Long
'------------------------------------------------------------------------------
'Purpose  : Determins the subtype of a VARIANT variable, i.e. Integer, Long Integer
'
'Prereq.  : -
'Parameter: Variable to test
'Returns  : %VT_xxx variant types (buildin PB constants) or -1 for error
'Note     : -
'
'   Author: Knuth Konrad 19.04.2016
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------

   Try
      TypeOfVariant = VariantVT(vntVar)
   Catch
      TypeOfVariant = -1
   End Try

End Function
'===========================================================================

Function baRnd Alias "baRnd" () Export As Currency
'------------------------------------------------------------------------------
'Purpose  : Generate a pseudo random number within the range of 0 <= x < 1
'
'Prereq.  : -
'Parameter: -
'Returns  : Random number
'Note     : -
'
'   Author: Knuth Konrad 28.02.2017
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local curResult As Currency

   Randomize Timer

   Try
      curResult = Rnd()
   Catch
      Trace Print Error$(ErrClear)
   End Try

   Function = curResult

End Function
'===========================================================================

Sub baRndArray Alias "baRndArray" (psa As Dword) Export
'------------------------------------------------------------------------------
'Purpose  : Fill an array with pseudo random numbers within the range of
'           0 <= x < 1
'
'Prereq.  : -
'Parameter: -
'
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 28.02.2017
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local curSeed As Currency

   Local i, l, u  As Long
   Local vb As Dword

   Trace On

   l  = vbArrayLBound(psa, 1)
   u  = vbArrayUBound(psa, 1)
   vb = vbArrayFirstElem(psa)

   Dim vba(l To u) As Currency At vb

   Randomize Timer
   For i = l To u
      vba(i) = Rnd()
   Next i

End Sub
'===========================================================================

Function baRndRange Alias "baRndRange" (ByVal lLower As Long, ByVal lUpper As Long) Export As Long
'------------------------------------------------------------------------------
'Purpose  : Generate a pseudo random number within the range of
'           lLower <= x <= lUpper
'
'Prereq.  : -
'Parameter: lLower   - lowest number
'           lUpper   - highest number
'Returns  : Random number
'Note     : -
'
'   Author: Knuth Konrad 28.02.2017
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------

   Trace On

   Randomize Timer
   Function = Rnd(lLower, lUpper)

End Function
'===========================================================================

Sub baRndRangeArrayLong Alias "baRndRangeArrayLong" (psa As Dword, ByVal lLower As Long, ByVal lUpper As Long) Export
'------------------------------------------------------------------------------
'Purpose  : Fill an array with pseudo random numbers within the range of
'           lLower <= x <= lUpper
'
'Prereq.  : -
'Parameter: lLower   - lowest number
'           lUpper   - highest number
'Returns  : Array with random numbers
'Note     : -
'
'   Author: Knuth Konrad 28.02.2017
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Local i, l, u  As Long
   Local vb As Dword

   Trace On

   l  = vbArrayLBound(psa, 1)
   u  = vbArrayUBound(psa, 1)
   vb = vbArrayFirstElem(psa)

   Dim vba(l To u) As Long At vb

   Randomize Timer
   For i = l To u
      vba(i) = Rnd(lLower, lUpper)
   Next i

End Sub
'===========================================================================

'function GetNUMBERFMTForLCID(byval wLangID as word, byref udt as NUMBERFMTA) as long
'
'local wLCID as word
'local szValue as asciiz * 11
'
''Type NUMBERFMTA
''   NumDigits As Dword           ' number of decimal digits
''   LeadingZero As Dword         ' if leading zero in decimal fields
''   Grouping As Dword            ' group size left of decimal
''   lpDecimalSep As AsciiZ Ptr   ' ptr to decimal separator string
''   lpThousandSep As AsciiZ Ptr  ' ptr to thousand separator string
''   NegativeOrder As Dword       ' negative number ordering
''End Type
'
'wLCID = MAKELCID(wLangID, %SORT_DEFAULT)
'
'udt.
'
'end function
'===========================================================================

Function baSetByte Alias "baSetByte" (psa As Dword, ByVal value As Byte) Export As Long
'------------------------------------------------------------------------------
'Purpose  : Set all elements of an array to value
'
'Prereq.  : -
'Parameter: Value - Value to which all elements are set
'Returns  : True / False
'Note     : -
'
'   Author: Knuth Konrad 2020-11-11
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------

   Local l  As Long
   Local u  As Long
   Local vb As Dword

   Trace On

   l  = vbArrayLBound(psa, 1)
   u  = vbArrayUBound(psa, 1)
   vb = vbArrayFirstElem(psa)

   Dim vba(l To u) As Byte At vb

   Try
      Mat vba() = Con(value)
   Catch
      Function = %S_False
      Exit Function
   End Try

   Function = %S_Ok

End Function
'==============================================================================

Function baSetDouble Alias "baSetDouble" (psa As Dword, ByVal value As Double) Export As Long
'------------------------------------------------------------------------------
'Purpose  : Set all elements of an array to value
'
'Prereq.  : -
'Parameter: Value - Value to which all elements are set
'Returns  : True / False
'Note     : -
'
'   Author: Knuth Konrad 2020-11-11
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------

   Local l  As Long
   Local u  As Long
   Local vb As Dword

   Trace On

   l  = vbArrayLBound(psa, 1)
   u  = vbArrayUBound(psa, 1)
   vb = vbArrayFirstElem(psa)

   Dim vba(l To u) As Double At vb

   Try
      Mat vba() = Con(value)
   Catch
      Function = %S_False
      Exit Function
   End Try

   Function = %S_Ok

End Function
'==============================================================================

Function baSetInteger Alias "baSetInteger" (psa As Dword, ByVal value As Integer) Export As Long
'------------------------------------------------------------------------------
'Purpose  : Set all elements of an array to value
'
'Prereq.  : -
'Parameter: Value - Value to which all elements are set
'Returns  : True / False
'Note     : -
'
'   Author: Knuth Konrad 2020-11-11
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------

   Local l  As Long
   Local u  As Long
   Local vb As Dword

   Trace On

   l  = vbArrayLBound(psa, 1)
   u  = vbArrayUBound(psa, 1)
   vb = vbArrayFirstElem(psa)

   Dim vba(l To u) As Integer At vb

   Try
      Mat vba() = Con(value)
   Catch
      Function = %S_False
      Exit Function
   End Try

   Function = %S_Ok

End Function
'==============================================================================

Function baSetLong Alias "baSetLong" (psa As Dword, ByVal value As Long) Export As Long
'------------------------------------------------------------------------------
'Purpose  : Set all elements of an array to value
'
'Prereq.  : -
'Parameter: Value - Value to which all elements are set
'Returns  : True / False
'Note     : -
'
'   Author: Knuth Konrad 2020-11-11
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------

   Local l  As Long
   Local u  As Long
   Local vb As Dword

   Trace On

   l  = vbArrayLBound(psa, 1)
   u  = vbArrayUBound(psa, 1)
   vb = vbArrayFirstElem(psa)

   Dim vba(l To u) As Long At vb

   Try
      Mat vba() = Con(value)
   Catch
      Function = %S_False
      Exit Function
   End Try

   Function = %S_Ok

End Function
'==============================================================================

Function baSetSingle Alias "baSetSingle" (psa As Dword, ByVal value As Single) Export As Long
'------------------------------------------------------------------------------
'Purpose  : Set all elements of an array to value
'
'Prereq.  : -
'Parameter: Value - Value to which all elements are set
'Returns  : True / False
'Note     : -
'
'   Author: Knuth Konrad 2020-11-11
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------

   Local l  As Long
   Local u  As Long
   Local vb As Dword

   Trace On

   l  = vbArrayLBound(psa, 1)
   u  = vbArrayUBound(psa, 1)
   vb = vbArrayFirstElem(psa)

   Dim vba(l To u) As Single At vb

   Try
      Mat vba() = Con(value)
   Catch
      Function = %S_False
      Exit Function
   End Try

   Function = %S_Ok

End Function
'==============================================================================

'------------------------------------------------------------------------------
' Helper methods for baSort2Arrays()
'------------------------------------------------------------------------------

Function baArrayByteSet Alias "baArrayByteSet" (psa As Dword, ByVal lWhichArray As Long) Export As Long
'------------------------------------------------------------------------------
'Purpose  : Create a PB array from VB
'
'Prereq.  : -
'Parameter: psa - VB array
'           lWhichArray - set array 1 or 2?
'Returns  : True / False
'Note     : -
'
'   Author: Knuth Konrad 2021-03-08
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------

   Local l  As Long
   Local u  As Long
   Local vb As Dword

   Trace On

   l  = vbArrayLBound(psa, 1)
   u  = vbArrayUBound(psa, 1)
   vb = vbArrayFirstElem(psa)

   Select Case lWhichArray
   Case 1
      Try
         ReDim gaByte1(l To u) As Byte At vb
      Catch
         Trace Print "Err: " & Format$(Err) & ", " & Error$(Err)
         Function = %False
         Exit Function
      End Try
   Case 2
      Try
         ReDim gaByte1(l To u) As Byte At vb
      Catch
         Trace Print "Err: " & Format$(Err) & ", " & Error$(Err)
         Function = %False
         Exit Function
      End Try
   Case Else
      Function = %False
   End Select

   Function = %True

End Function
'==============================================================================

Function baArrayCurrencySet Alias "baArrayCurrencySet" (psa As Dword, ByVal lWhichArray As Long) Export As Long
'------------------------------------------------------------------------------
'Purpose  : Create a PB array from VB
'
'Prereq.  : -
'Parameter: psa - VB array
'           lWhichArray - set array 1 or 2?
'Returns  : True / False
'Note     : -
'
'   Author: Knuth Konrad 2021-03-08
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------

   Local l  As Long
   Local u  As Long
   Local vb As Dword

   Trace On

   l  = vbArrayLBound(psa, 1)
   u  = vbArrayUBound(psa, 1)
   vb = vbArrayFirstElem(psa)

   Select Case lWhichArray
   Case 1
      Try
         ReDim gaCurrency1(l To u) As Byte At vb
      Catch
         Function = %False
         Exit Function
      End Try
   Case 2
      Try
         ReDim gaCurrency2(l To u) As Byte At vb
      Catch
         Function = %False
         Exit Function
      End Try
   Case Else
      Function = %False
   End Select

   Function = %True

End Function
'==============================================================================

Function baArrayDoubleSet Alias "baArrayDoubleSet" (psa As Dword, ByVal lWhichArray As Long) Export As Long
'------------------------------------------------------------------------------
'Purpose  : Create a PB array from VB
'
'Prereq.  : -
'Parameter: psa - VB array
'           lWhichArray - set array 1 or 2?
'Returns  : True / False
'Note     : -
'
'   Author: Knuth Konrad 2021-03-08
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------

   Local l  As Long
   Local u  As Long
   Local vb As Dword

   Trace On

   l  = vbArrayLBound(psa, 1)
   u  = vbArrayUBound(psa, 1)
   vb = vbArrayFirstElem(psa)

   Select Case lWhichArray
   Case 1
      Try
         ReDim gaDouble1(l To u) As Byte At vb
      Catch
         Function = %False
         Exit Function
      End Try
   Case 2
      Try
         ReDim gaDouble2(l To u) As Byte At vb
      Catch
         Function = %False
         Exit Function
      End Try
   Case Else
      Function = %False
   End Select

   Function = %True

End Function
'==============================================================================

Function baArrayIntegerSet Alias "baArrayIntegerSet" (psa As Dword, ByVal lWhichArray As Long) Export As Long
'------------------------------------------------------------------------------
'Purpose  : Create a PB array from VB
'
'Prereq.  : -
'Parameter: psa - VB array
'           lWhichArray - set array 1 or 2?
'Returns  : True / False
'Note     : -
'
'   Author: Knuth Konrad 2021-03-08
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------

   Local l  As Long
   Local u  As Long
   Local vb As Dword

   Trace On

   l  = vbArrayLBound(psa, 1)
   u  = vbArrayUBound(psa, 1)
   vb = vbArrayFirstElem(psa)

   Select Case lWhichArray
   Case 1
      Try
         ReDim gaInteger1(l To u) As Byte At vb
      Catch
         Function = %False
         Exit Function
      End Try
   Case 2
      Try
         ReDim gaInteger2(l To u) As Byte At vb
      Catch
         Function = %False
         Exit Function
      End Try
   Case Else
      Function = %False
   End Select

   Function = %True

End Function
'==============================================================================

Function baArrayLongSet Alias "baArrayLongSet" (psa As Dword, ByVal lWhichArray As Long) Export As Long
'------------------------------------------------------------------------------
'Purpose  : Create a PB array from VB
'
'Prereq.  : -
'Parameter: psa - VB array
'           lWhichArray - set array 1 or 2?
'Returns  : True / False
'Note     : -
'
'   Author: Knuth Konrad 2021-03-08
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------

   Local l  As Long
   Local u  As Long
   Local vb As Dword

   Trace On

   l  = vbArrayLBound(psa, 1)
   u  = vbArrayUBound(psa, 1)
   vb = vbArrayFirstElem(psa)

   Select Case lWhichArray
   Case 1
      Try
         ReDim gaLong1(l To u) As Byte At vb
      Catch
         Function = %False
         Exit Function
      End Try
   Case 2
      Try
         ReDim gaLong2(l To u) As Byte At vb
      Catch
         Function = %False
         Exit Function
      End Try
   Case Else
      Function = %False
   End Select

   Function = %True

End Function
'==============================================================================

Function baArraySingleSet Alias "baArraySingleSet" (psa As Dword, ByVal lWhichArray As Long) Export As Long
'------------------------------------------------------------------------------
'Purpose  : Create a PB array from VB
'
'Prereq.  : -
'Parameter: psa - VB array
'           lWhichArray - set array 1 or 2?
'Returns  : True / False
'Note     : -
'
'   Author: Knuth Konrad 2021-03-08
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------

   Local l  As Long
   Local u  As Long
   Local vb As Dword

   Trace On

   l  = vbArrayLBound(psa, 1)
   u  = vbArrayUBound(psa, 1)
   vb = vbArrayFirstElem(psa)

   Select Case lWhichArray
   Case 1
      Try
         ReDim gaSingle1(l To u) As Byte At vb
      Catch
         Function = %False
         Exit Function
      End Try
   Case 2
      Try
         ReDim gaSingle2(l To u) As Byte At vb
      Catch
         Function = %False
         Exit Function
      End Try
   Case Else
      Function = %False
   End Select

   Function = %True

End Function
'==============================================================================

Function baArrayStringSet Alias "baArrayStringSet" (psa As Dword, ByVal lWhichArray As Long) Export As Long
'------------------------------------------------------------------------------
'Purpose  : Create a PB array from VB
'
'Prereq.  : -
'Parameter: psa - VB array
'           lWhichArray - set array 1 or 2?
'Returns  : True / False
'Note     : -
'
'   Author: Knuth Konrad 2021-03-08
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------

   Local l  As Long
   Local u  As Long
   Local vb As Dword

   Trace On

   l  = vbArrayLBound(psa, 1)
   u  = vbArrayUBound(psa, 1)
   vb = vbArrayFirstElem(psa)

   Select Case lWhichArray
   Case 1
      Try
         ReDim gaString1(l To u) As Byte At vb
      Catch
         Trace Print "Err: " & Format$(Err) & ", " & Error$(Err)
         Function = %False
         Exit Function
      End Try
   Case 2
      Try
         ReDim gaString2(l To u) As Byte At vb
      Catch
         Trace Print "Err: " & Format$(Err) & ", " & Error$(Err)
         Function = %False
         Exit Function
      End Try
   Case Else
      Function = %False
   End Select

   Function = %True

End Function
'==============================================================================

Function baSort2Arrays Alias "baSort2Arrays" (ByVal lArrayType1 As Long, _
   ByVal lArrayType2 As Long, Optional ByVal lDescending As Long) Export As Long
'------------------------------------------------------------------------------
'Purpose  : Sorts/reorders 2 arrays. The first array (ps1) will be sorted.
'           The second array however (ps2) will have its array elements arranged
'           as if they "stick" to the first element, i.e. when ps1 is sorted and
'           ps1(1) has become ps1(5) thereafter, ps2(1) now also will be ps2(5)
'
'Prereq.  : -
'Parameter: lArrayType1, lArrayType2 - Data types of the VB arrays
'Returns  : True / False
'Note     : -
'
'   Author: Knuth Konrad 2021-03-08
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------


'%DATA_TYPE_BYTE = 0
'%DATA_TYPE_CURRENCY = 1
'%DATA_TYPE_DOUBLE = 2
'%DATA_TYPE_INTEGER = 3
'%DATA_TYPE_LONG = 4
'%DATA_TYPE_SINGLE = 5
'%DATA_TYPE_STRING = 6


   ' ** ToDo:
   ' - How to pass back the sorted arrays?
   ' Check that array2's size is >= array1's size
   ' - Don't erase array contents in the baArrayXXXSet methods, as that may erase user input

   Local l1, l2  As Long
   Local u1, u2  As Long
   Local vb1, vb2 As Dword

   Trace On

   Trace Print "Array types: " & Format$(lArrayType1) & ", " & Format$(lArrayType2)

   Select Case lArrayType1
   Case %DATA_TYPE_BYTE

      Select Case lArrayType2
      Case %DATA_TYPE_BYTE
         If IsFalse(lDescending) Then
            Array Sort gaByte1(), TagArray gaByte2()
         Else
            Array Sort gaByte1(), TagArray gaByte2(), Descend
         End If

      Case %DATA_TYPE_CURRENCY
         If IsFalse(lDescending) Then
            Array Sort gaByte1(), TagArray gaDouble2()
         Else
            Array Sort gaByte1(), TagArray gaDouble2(), Descend
         End If

      Case %DATA_TYPE_DOUBLE
      Case %DATA_TYPE_INTEGER
      Case %DATA_TYPE_LONG
      Case %DATA_TYPE_SINGLE
      Case %DATA_TYPE_STRING
         If IsFalse(lDescending) Then
            Array Sort gaByte1(), TagArray gaString2()
         Else
            Array Sort gaByte1(), TagArray gaString2(), Descend
         End If

      Case Else
         Function = 1
         Exit Function
      End Select  '// Case lArrayType2

   Case %DATA_TYPE_CURRENCY
   Case %DATA_TYPE_DOUBLE
   Case %DATA_TYPE_INTEGER
   Case %DATA_TYPE_LONG
   Case %DATA_TYPE_SINGLE
   Case %DATA_TYPE_STRING
   Case Else
      Function = 1
      Exit Function
   End Select  '// Case lArrayType1

   Function = %S_Ok

End Function
'==============================================================================


'----------------------------------------------------------------------------
Rem </MKVBDEC>
'----------------------------------------------------------------------------
