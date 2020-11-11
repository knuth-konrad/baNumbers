Attribute VB_Name = "PCtrl"
'------------------------------------------------------------------------------
'Purpose  : Control Enhancements und Tools
'
'Prereq.  : -
'Note     : Tools für non-instrinc Controls (Listview, Treeview etc.)
'           ausgelagert nach PCtrlCC
'
'   Author: Knuth Konrad 19.07.2007
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
Option Explicit
DefLng A-Z
'------------------------------------------------------------------------------
'*** Constants ***
'------------------------------------------------------------------------------
Private Const GWL_EXSTYLE As Long = (-20)
Private Const GWL_STYLE As Long = (-16)
Private Const WS_BORDER  As Long = &H800000
Private Const WS_CAPTION  As Long = &HC00000
Private Const WS_DLGFRAME  As Long = &H400000
Private Const WS_MAXIMIZEBOX  As Long = &H10000
Private Const WS_MINIMIZEBOX  As Long = &H20000
Private Const WS_THICKFRAME  As Long = &H40000
Private Const WS_EX_APPWINDOW As Long = &H40000
Private Const WS_VSCROLL As Long = &H200000
Private Const WS_HSCROLL As Long = &H100000

Private Const SW_HIDE As Long = 0
Private Const SW_SHOW As Long = 5

' Window position
Private Const SWP_FRAMECHANGED As Long = &H20
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4

' Combo Box messages
'%CB_GETEDITSEL            = &H140
'%CB_LIMITTEXT             = &H141
'%CB_SETEDITSEL            = &H142
'%CB_ADDSTRING             = &H143
'%CB_DELETESTRING          = &H144
'%CB_DIR                   = &H145
'%CB_GETCOUNT              = &H146
'%CB_GETCURSEL             = &H147
'%CB_GETLBTEXT             = &H148
'%CB_GETLBTEXTLEN          = &H149
'%CB_INSERTSTRING          = &H14A
'%CB_RESETCONTENT          = &H14B
'%CB_FINDSTRING            = &H14C
'%CB_SETCURSEL             = &H14E
'%CB_SHOWDROPDOWN          = &H14F
'%CB_GETITEMDATA           = &H150
'%CB_SETITEMDATA           = &H151
'%CB_GETDROPPEDCONTROLRECT = &H152
'%CB_SETITEMHEIGHT         = &H153
'%CB_GETITEMHEIGHT         = &H154
'%CB_SETEXTENDEDUI         = &H155
'%CB_GETEXTENDEDUI         = &H156
'%CB_GETDROPPEDSTATE       = &H157
'%CB_FINDSTRINGEXACT       = &H158
'%CB_SETLOCALE             = &H159
'%CB_GETLOCALE             = &H15A
'%CB_GETTOPINDEX           = &H15B
'%CB_SETTOPINDEX           = &H15C
'%CB_GETHORIZONTALEXTENT   = &H15D
'%CB_SETHORIZONTALEXTENT   = &H15E
'%CB_GETDROPPEDWIDTH       = &H15F
'%CB_SETDROPPEDWIDTH       = &H160
'%CB_INITSTORAGE           = &H161
'%CB_MULTIPLEADDSTRING     = &H163
'%CB_GETCOMBOBOXINFO       = &H164
'%CB_MSGMAX                = &H165  ' depends on Windows version
Private Const CB_ADDSTRING As Long = &H143
Private Const CB_FINDSTRING As Long = &H14C
Private Const CB_FINDSTRINGEXACT As Long = &H158
Private Const CB_GETCOUNT As Long = &H146
Private Const CB_GETCURSEL As Long = &H147
Private Const CB_GETLBTEXT As Long = &H148
Private Const CB_GETLBTEXTLEN As Long = &H149
Private Const CB_SELECTSTRING As Long = &H14D
Private Const CB_SETCURSEL As Long = &H14E
Private Const CB_SETITEMDATA As Long = &H151
Private Const CB_ERR As Long = (-1)


' Listbox messages
'%LB_ADDSTRING          = &H180
'%LB_INSERTSTRING       = &H181
'%LB_DELETESTRING       = &H182
'%LB_SELITEMRANGEEX     = &H183
'%LB_RESETCONTENT       = &H184
'%LB_SETSEL             = &H185
'%LB_SETCURSEL          = &H186
'%LB_GETSEL             = &H187
'%LB_GETCURSEL          = &H188
'%LB_GETTEXT            = &H189
'%LB_GETTEXTLEN         = &H18A
'%LB_GETCOUNT           = &H18B
'%LB_SELECTSTRING       = &H18C
'%LB_DIR                = &H18D
'%LB_GETTOPINDEX        = &H18E
'%LB_FINDSTRING         = &H18F
'%LB_GETSELCOUNT        = &H190
'%LB_GETSELITEMS        = &H191
'%LB_SETTABSTOPS        = &H192
'%LB_GETHORIZONTALEXTENT= &H193
'%LB_SETHORIZONTALEXTENT= &H194
'%LB_SETCOLUMNWIDTH     = &H195
'%LB_ADDFILE            = &H196
'%LB_SETTOPINDEX        = &H197
'%LB_GETITEMRECT        = &H198
'%LB_GETITEMDATA        = &H199
'%LB_SETITEMDATA        = &H19A
'%LB_SELITEMRANGE       = &H19B
'%LB_SETANCHORINDEX     = &H19C
'%LB_GETANCHORINDEX     = &H19D
'%LB_SETCARETINDEX      = &H19E
'%LB_GETCARETINDEX      = &H19F
'%LB_SETITEMHEIGHT      = &H1A0
'%LB_GETITEMHEIGHT      = &H1A1
'%LB_FINDSTRINGEXACT    = &H1A2
'%LB_SETLOCALE          = &H1A5
'%LB_GETLOCALE          = &H1A6
'%LB_SETCOUNT           = &H1A7
'%LB_INITSTORAGE        = &H1A8
'%LB_ITEMFROMPOINT      = &H1A9
'%LB_MULTIPLEADDSTRING  = &H1B1
'%LB_GETLISTBOXINFO     = &H1B2
'%LB_MSGMAX             = &H1B3  ' depends on Windows version
Private Const LB_FINDSTRINGEXACT As Long = &H1A2
Private Const LB_ADDSTRING As Long = &H180
Private Const LB_SETITEMDATA As Long = &H19A
Private Const LB_SETSEL As Long = &H185
Private Const LB_GETCURSEL As Long = &H188
Private Const LB_GETTEXTLEN As Long = &H18A
Private Const LB_GETTEXT As Long = &H189
Private Const LB_GETITEMRECT  As Long = &H198
Private Const LB_SETHORIZONTALEXTENT As Long = &H194

Private Const LB_ERR As Long = (-1)

'Edit control styles
Private Const ES_NUMBER As Long = &H2000

'* Window Redraw
' %WM_SETREDRAW        = &H000B???
Private Const WM_SETREDRAW As Long = &HB

'* DrawText
' Der Text wird am unterem Rand ausgerichtet (Nur in Verbindung mit DT_SINGLELINE)
Private Const DT_BOTTOM As Long = &H8

' Die Funktion füllt die RECT-Struktur mit den Koordinaten, die für das
' Zeichnen des Textes benötigt werden, zeichnet den Text aber nicht
Private Const DT_CALCRECT As Long = &H400

' Der Text wird horizontal zentriert
Private Const DT_CENTER As Long = &H1

' Zeichnet den Text, wie es eine Textbox tun würde, nur teilweise sichtbare
' Zeilen bei zu knapp berechneten übergebenen Koordinaten werden nicht gezeichnet
Private Const DT_EDITCONTROL As Long = &H2000

' Der übergebene Puffer mit dem zu zeichnenden Text wird mit dem Text
' gefüllt, der nicht in angegebenen Koordinaten dargestellt werden kann (Nur
' in Verbindung mit DT_MODIFYSTRING)
Private Const DT_END_ELLIPSIS As Long = &H8000
 
' Zeichnet TAB-Zeichen mit 8 Leerzeichen (Nicht in Verbindung mit
' DT_WORD_ELLIPSIS, DT_PATH_ELLIPSIS und DT_END_ELLIPSIS)
Private Const DT_EXPANDTABS As Long = &H40
 
' Fügt jeder Zeile die Höhe der Externen Führung hinzu
Private Const DT_EXTERNALLEADING As Long = &H200
 
' (Windows 2000/XP) Zeichnet die auf einem &-Zeichen folgenden Unterstriche
' nicht, die &-Zeichen werden aber wie bisher ausgeblendet
Private Const DT_HIDEPREFIX As Long = &H100000
 
' Benutzt die Systemschriftart um die benötigten Koordinaten zu berechnen
Private Const DT_INTERNAL As Long = &H1000
 
' Der Text wird links ausgerichtet
Private Const DT_LEFT As Long = &H0
 
' Der übergebene Puffer wird mit den Wörtern oder Zeichen gefüllt die aus
' Platzmangel nicht gezeichnet werden konnten (Nur in Verbindung mit
' DT_END_ELLIPSIS oder DT_PATH_ELLIPSIS)
Private Const DT_MODIFYSTRING As Long = &H10000
 
' Der Text wird an den angegebenen Koordinaten nicht abgeschnitten falls der
' Platz nicht ausreichend ist
Private Const DT_NOCLIP As Long = &H100
 
' (Windows 98, ME, NT, 2000, XP) Es werden Doppelzeichen für Zeilenumbrüche
' verwendet (nicht in Verbindung mit DT_WORDBREAK)
Private Const DT_NOFULLWIDTHCHARBREAK As Long = &H80000
 
' Zeichnet die auf einem &-Zeichen folgendenden Unterstriche nicht, die
' &-Zeichen werden normal angezeigt
Private Const DT_NOPREFIX As Long = &H800
 
' Ersetzt Buchstaben in der Mitte des übergebenen Strings damit er in den
' Koordinaten gezeichnet werden kann (nur in Verbindung mit DT_MODIFYSTRING)
Private Const DT_PATH_ELLIPSIS As Long = &H4000
 
' (Windows 2000, XP) Zeichnet bei Verwendung von &-Zeichen nur die Unterstriche
Private Const DT_PREFIXONLY As Long = &H200000
 
' Der Text wird rechts ausgerichtet
Private Const DT_RIGHT As Long = &H2
 
' Der Text wird von Rechts nach Links gezeichnet wenn ein Hebräischer Font gesetzt ist
Private Const DT_RTLREADING As Long = &H20000
 
' Der Text wird in einer einzelnen Zeile gezeichnet, VBCrLf-Zeichen werden ignoriert
Private Const DT_SINGLELINE As Long = &H20
 
' Setzt die Anzahl der Leerzeichen für ein Tab-Zeichen, die Bits 8 bis 15
' erwarten die Anzahl der Leerstellen
Private Const DT_TABSTOP As Long = &H80
 
' Der Text wird oben ausgerichtet
Private Const DT_TOP As Long = &H0
 
' Der Text wird vertikal zentriert (nur in Verbindung mit DT_SINGLELINE)
Private Const DT_VCENTER As Long = &H4
 
' Ist eine Zeile zu lange für die angegebenen Koordinaten, so wird zwischen
' den benötigten Wörtern ein VBCrLf-Zeichen eingefügt
Private Const DT_WORDBREAK As Long = &H10
 
' Der übergebene Puffer wird mit den Wörtern gefüllt die in den angegebenen
' Koordinaten aus Platzgründen nicht dargestellt werden konnten
Private Const DT_WORD_ELLIPSIS As Long = &H40000

' InitCommonControls
Private Const ICC_USEREX_CLASSES = &H200
'------------------------------------------------------------------------------
'*** Enumeration/TYPEs ***
'------------------------------------------------------------------------------
Private Type ApiRECT
   rctLeft As Long
   rctTop As Long
   rctRight As Long
   rctBottom As Long
End Type

' InitCommonControls
Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

' Textbox Scrollbars
Public Enum eTxtScrollbar
   scrollNone = 0
   scrollHorizontal
   scrollVertical
   scrollBoth
End Enum
'------------------------------------------------------------------------------
'*** Declares ***
'------------------------------------------------------------------------------
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
   (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
   (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
   ByVal hWndInsertAfter As Long, _
   ByVal x As Long, ByVal y As Long, _
   ByVal cx As Long, _
   ByVal cy As Long, _
   ByVal wFlags As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
   (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function GetAsyncKeyState Lib "user32" _
   (ByVal vKey As Long) As Integer
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function InvalidateRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As ApiRECT, _
   ByVal bErase As Long) As Long
Private Declare Function InvalidateClientRect Lib "user32" Alias "InvalidateRect" (ByVal hWnd As Long, _
   lpRect As Any, ByVal bErase As Long) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As ApiRECT) As Long

Private Declare Function DrawText Lib "user32" _
   Alias "DrawTextA" ( _
   ByVal hDC As Long, _
   ByVal lpStr As String, _
   ByVal nCount As Long, _
   lpRect As ApiRECT, _
   ByVal wFormat As Long) As Long
  
Private Declare Function GetLastError Lib "kernel32" () As Long

' InitCommonControls
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" _
   (iccex As tagInitCommonControlsEx) As Boolean
'------------------------------------------------------------------------------
'*** Variables ***
'------------------------------------------------------------------------------
'==============================================================================

Sub MarkText(frm As Form)
'------------------------------------------------------------------------------
'Purpose  : Selektiert den gesammten Text einer Textbox
'
'Prereq.  : -
'Parameter: frm   - Form auf der die Textbox liegt
'Note     : Aufruf erfolgt typischerweise im GotFocus-Event der Textbox
'
'   Author: Knuth Konrad 17.08.1999
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   
   On Error Resume Next
   
   With frm
      If TypeOf .ActiveControl Is TextBox Then
         If Not CBool(GetAsyncKeyState(vbKeyLButton) Or _
            GetAsyncKeyState(vbKeyMButton) Or _
            GetAsyncKeyState(vbKeyRButton)) Then
         
            .ActiveControl.SelStart = 0
            .ActiveControl.SelLength = Len(.ActiveControl.Text)
         End If
      End If
   End With
   
   Err.Clear
   On Error GoTo 0
   
End Sub
'==============================================================================

Sub MarkCombo(frm As Form)
'------------------------------------------------------------------------------
'Purpose  : Selektiert den gesammten Text einer ComboBox
'
'Prereq.  : -
'Parameter: frm   -  Form auf der die Textbox liegt
'Note     : Aufruf erfolgt typischerweise im GotFocus-Event der Combobox
'
'   Author: Knuth Konrad 17.08.1999
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   
   On Error Resume Next
   
   With frm
      If TypeOf .ActiveControl Is ComboBox Then
         .ActiveControl.SelStart = 0
         .ActiveControl.SelLength = Len(.ActiveControl.Text)
      End If
   End With
   
   Err.Clear
   On Error GoTo 0
   
End Sub
'==============================================================================

Public Sub ComboAutoComplete(ByRef SourceCtl As VB.ComboBox, ByRef KeyAscii As Integer, ByRef LeftOffPos As Long)
'------------------------------------------------------------------------------
'Purpose  : -
'
'Prereq.  : -
'Parameter: -
'Note     : Example of how to call it (note that Combo1's style should be set to 0 - Dropdown Combo):
'           Private Sub Combo1_KeyPress(KeyAscii As Integer)
'           Static iLeftOff As Long
'           ComboAutoComplete Combo1, KeyAscii, iLeftOff
'           End Sub
'
'   Author: lebb(?) 24.04.2017
'   Source: http://www.xtremevbtalk.com/archive/index.php/t-90541.html
'  Changed: -
'------------------------------------------------------------------------------
   Dim iStart As Long, lListIndex As Long
   Dim sSearchKey As String
   
   With SourceCtl
      
      'If text entered so far matches item(s) in the list, use autocomplete
      Select Case Chr$(KeyAscii)
      
      Case vbBack
      'Let backspace characters process as usual; otherwise try to match text
      
      Case Else
         If Chr$(KeyAscii) <> vbBack Then
            .SelText = Chr$(KeyAscii)
         
            iStart = .SelStart
         
            If LeftOffPos <> 0 Then
               .SelStart = LeftOffPos
               iStart = LeftOffPos
            End If
         
         sSearchKey = CStr(Left$(.Text, iStart))
         
   '      .ListIndex = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal CStr(Left$(.Text, iStart)))
   '
   '      If .ListIndex = -1 Then
   '         LeftOffPos = Len(sSearchKey)
   '      End If
         
         lListIndex = SendMessage(.hWnd, CB_FINDSTRING, -1, ByVal CStr(Left$(.Text, iStart)))
         Call SendMessage(.hWnd, CB_SETCURSEL, lListIndex, ByVal 0)
         
         If lListIndex = -1 Then
            LeftOffPos = Len(sSearchKey)
         End If
         
         .SelStart = iStart
         .SelLength = Len(.Text)
         LeftOffPos = 0
         
         KeyAscii = 0
         End If
      End Select
   End With
   
End Sub
'==============================================================================

Function GetOption(opts As Object) As Long
'------------------------------------------------------------------------------
'Purpose  : Liefert den Index eines Optionbutton-ControlArrays zurück
'           dessen Value True ist
'
'Prereq.  : -
'Parameter: opts  -  OptionButton-ControlArray
'Returns  : Index des Controls dessen Value True ist oder -1 bei Fehler/keiner Auswahl
'Note     : -
'
'   Author: Bruce McKinney - Hardcore Visual Basic 5
'   Source: -
'  Changed: 19.08.1999, Knuth Konrad
'           Fehlerbehandlung hinzugefügt.
'------------------------------------------------------------------------------
   Dim opt As OptionButton
   
   On Error GoTo GetOptionFail
   
   ' Annehmen das nichts ausgewählt wurde
   GetOption = -1
   
   For Each opt In opts
      If opt.Value Then
         GetOption = opt.Index
         Exit Function
      End If
   Next
   
GetOptionExit:
   On Error GoTo 0
   Exit Function
   
GetOptionFail:
   Err.Clear
   GetOption = -1
   Resume GetOptionExit
   
End Function
'==============================================================================

Public Function GetComboIndex(ByVal cbo As ComboBox, ByVal sPattern As String, _
   Optional bolTrim As Boolean = False, Optional lStart As Long = -1) As Long
'------------------------------------------------------------------------------
'Purpose  : Liefert den Index einer Combobox für einen bestimmten Eintrag
'
'Prereq.  : -
'Parameter: sPattern -  Eintrag (Text) nach dem gesucht werden soll
'           cbo      -  zu durchsuchende ComboBox
'           bolTrim  -  Führende und nachfolgende Leerzeichen mit abschneiden
'           lStart   -  ComboBox-Item ab dem die Suche durchgeführt werden soll
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 22.05.2000
'   Source: -
'  Changed: 10.07.2000
'           - lStart als optionalen Parameter hinzugefügt
'------------------------------------------------------------------------------
   
   If bolTrim Then
      sPattern = Trim$(sPattern)
   End If
   
   GetComboIndex = SendMessage(cbo.hWnd, CB_FINDSTRINGEXACT, lStart, _
      ByVal sPattern)
   
End Function
'==============================================================================

Public Function ComboGetListCount(ByVal cbo As ComboBox) As Long
'------------------------------------------------------------------------------
'Purpose  : Return the number of ListItems in a combobox
'
'Prereq.  : -
'Parameter: cbo   - Combobox control
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 25.06.2018
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   
   ComboGetListCount = SendMessage(cbo.hWnd, CB_GETCOUNT, ByVal 0, ByVal 0)
   
End Function
'==============================================================================

Public Sub ComboAddItemData(ByVal cbo As ComboBox, ByVal sText As String, _
   Optional ByVal lItemData As Variant)
'------------------------------------------------------------------------------
'Purpose  : Fügt ein Item samt Itemdata zu einer Combobox hinzu
'
'Prereq.  : -
'Parameter: cbo         -  Combobox
'           sText       -  Hinzuzufügender Text
'           lItemData   -  Itemdata zu diesem Text
'Note     : Ermöglicht das Einfügen von Text *und* ItemData zu einer sortierten
'           Listbox
'
'   Author: Knuth Konrad 12.03.2002
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Dim lIndex As Long
   Dim lData As Long
   
   lIndex = SendMessage(cbo.hWnd, CB_ADDSTRING, 0, ByVal sText)
   
   If Not IsMissing(lItemData) Then
      lData = CLng(lItemData)
      lIndex = SendMessage(cbo.hWnd, CB_SETITEMDATA, lIndex, ByVal lData)
   End If
   
End Sub
'==============================================================================

Public Function ComboAddItemDataEx(ByVal cbo As ComboBox, ByVal sText As String, _
   Optional ByVal lItemData As Variant) As Long
'------------------------------------------------------------------------------
'Purpose  : Fügt ein Item samt Itemdata zu einer Combobox hinzu
'
'Prereq.  : -
'Parameter: cbo         -  Combobox
'           sText       -  Hinzuzufügender Text
'           lItemData   -  Itemdata zu diesem Text
'Returns  : Index des hinzugefügten Items
'Note     : Ermöglicht das Einfügen von Text *und* ItemData zu einer sortierten
'           Combobox
'
'   Author: Knuth Konrad 12.03.2002
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Dim lIndex As Long, lRetval As Long, lData As Long
   
   lIndex = SendMessage(cbo.hWnd, CB_ADDSTRING, 0, ByVal sText)
   
   If Not IsMissing(lItemData) Then
      lData = CLng(lItemData)
      lRetval = SendMessage(cbo.hWnd, CB_SETITEMDATA, lIndex, ByVal lData)
   End If
   
   If lRetval <> CB_ERR Then
      ComboAddItemDataEx = lIndex
   Else
      ComboAddItemDataEx = CB_ERR
   End If
   
End Function
'==============================================================================

Public Function ComboSetSelection(ByVal cbo As ComboBox, ByVal sText As String, Optional ByVal lDefaultIndex As Long = CB_ERR, _
   Optional ByVal lStartFrom As Long = -1) As Long
'------------------------------------------------------------------------------
'Purpose  : Selektiert einen Eintrag in einer Combobox
'
'Prereq.  : -
'Parameter: cbo            -  Combobox
'           sText          -  zu selektierender Eintrag
'           lDefaultIndex  -  von -1 (keine Auswahl) abweichende Standardauswahl wenn Eintrag nicht gefunden wird
'           lStartFrom  -  Suche ab Element lStartForm starten
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 18.04.2012
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Dim lRet As Long
   
   lRet = SendMessage(cbo.hWnd, CB_SELECTSTRING, lStartFrom, ByVal sText)
   
   If lRet = CB_ERR Then
      ComboSetSelection = lDefaultIndex
   Else
      ComboSetSelection = lRet
   End If
   
End Function
'==============================================================================

Public Function ComboGetItemData(ByVal cbo As ComboBox) As Long
'------------------------------------------------------------------------------
'Purpose  : Ermittelt den Wert von ItemData des ausgewählten Eintrags einer Combobox
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 07.03.2008
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   
   ComboGetItemData = cbo.ItemData(cbo.ListIndex)
   
End Function
'==============================================================================

Public Function ComboGetItem(ByVal cbo As ComboBox) As String
'------------------------------------------------------------------------------
'Purpose  : Ermittelt den ausgewählten Wert des ausgewählten Eintrags einer Combobox
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 07.03.2008
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Dim sTemp As String, lTemp As Long, lListIndex As Long
   
   ' Stringlänge ermitteln
   lListIndex = SendMessage(cbo.hWnd, CB_GETCURSEL, ByVal 0, ByVal 0)
   lTemp = SendMessage(cbo.hWnd, CB_GETLBTEXTLEN, lListIndex, ByVal 0)
   
   sTemp = Space$(lTemp + 1)
   
   If SendMessage(cbo.hWnd, CB_GETLBTEXT, lListIndex, ByVal sTemp) > 0 Then
      ComboGetItem = Left$(sTemp, lTemp)
   End If
   
End Function
'==============================================================================

Public Function GetListBoxIndex(ByVal lst As ListBox, ByVal sPattern As String, _
   Optional bolTrim As Boolean = False, Optional lStart As Long = -1) As Long
'------------------------------------------------------------------------------
'Purpose  : Liefert den Index einer Listbox für einen bestimmten Eintrag
'
'Prereq.  : -
'Parameter: sPattern -  Eintrag (Text) nach dem gesucht werden soll
'Returns  : lst      -  zu durchsuchende ComboBox
'Note     : -
'
'   Author: Knuth Konrad 10.07.2000
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   
   If bolTrim Then
      sPattern = Trim$(sPattern)
   End If
   
   GetListBoxIndex = SendMessage(lst.hWnd, LB_FINDSTRINGEXACT, lStart, _
      ByVal sPattern)
   
End Function
'==============================================================================

Public Sub ListBoxAddItemData(ByVal lst As ListBox, ByVal sText As String, _
   Optional ByVal lItemData As Variant, Optional ByVal bolCheck As Boolean = False)
'------------------------------------------------------------------------------
'Purpose  : Fügt ein Item samt Itemdata zu einer Listbox hinzu
'
'Prereq.  : -
'Parameter: lst         -  Listbox
'           sText       -  Hinzuzufügender Text
'           lItemData   -  Itemdata zu diesem Text
'           bolCheck    -  Bei einer Listbox mit dem Style "1 - Kontrollkästchen",
'                          Checkbox setzen ja/nein
'Note     : Ermöglicht das Einfügen von Text *und* ItemData zu einer sortierten
'           Listbox
'
'   Author: Knuth Konrad 12.03.2002
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Dim lIndex As Long
   Dim lData As Long
   Dim lRetval As Long
   
   lIndex = SendMessage(lst.hWnd, LB_ADDSTRING, 0, ByVal sText)
   If Not IsMissing(lItemData) Then
      lData = CLng(lItemData)
      lRetval = SendMessage(lst.hWnd, LB_SETITEMDATA, lIndex, ByVal lData)
   End If
   
   If bolCheck = True And lst.Style = vbListBoxCheckbox Then
      lst.Selected(lIndex) = True
   End If
   
End Sub
'==============================================================================

Public Function ListBoxAddItemDataEx(ByVal lst As ListBox, ByVal sText As String, _
   Optional ByVal lItemData As Variant, Optional ByVal bolCheck As Boolean = False) As Long
'------------------------------------------------------------------------------
'Purpose  : Fügt ein Item samt Itemdata zu einer Listbox hinzu
'
'Prereq.  : -
'Parameter: lst         -  Listbox
'           sText       -  Hinzuzufügender Text
'           lItemData   -  Itemdata zu diesem Text
'           bolCheck    -  Bei einer Listbox mit dem Style "1 - Kontrollkästchen",
'                          Checkbox setzen ja/nein
'Returns  : Index des hinzugefügten Items
'Note     : -
'
'   Author: Knuth Konrad 30.06.2016
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Dim lIndex As Long, lRetval As Long, lData As Long
   
   lIndex = SendMessage(lst.hWnd, LB_ADDSTRING, 0, ByVal sText)
   If Not IsMissing(lItemData) Then
      lData = CLng(lItemData)
      lRetval = SendMessage(lst.hWnd, LB_SETITEMDATA, lIndex, ByVal lData)
   End If
   
   If bolCheck = True And lst.Style = vbListBoxCheckbox Then
      lst.Selected(lIndex) = True
   End If
   
   If lRetval <> LB_ERR Then
      ListBoxAddItemDataEx = lIndex
   Else
      ListBoxAddItemDataEx = LB_ERR
   End If
   
End Function
'==============================================================================

Public Function ListBoxGetItemData(ByVal lst As ListBox) As Long
'------------------------------------------------------------------------------
'Purpose  : Ermittelt den Wert von ItemData des ausgewählten Eintrags einer Listbox
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 07.03.2008
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   
   ListBoxGetItemData = lst.ItemData(lst.ListIndex)
   
End Function
'==============================================================================

Public Sub ListBoxSetItemData(ByVal lst As ListBox, ByVal lListIndex As Long, _
   Optional ByVal lItemData As Variant)
'------------------------------------------------------------------------------
'Purpose  : Sets the ItemData property of a Listbox
'
'Prereq.  : -
'Parameter: lst         - Listbox control
'           lListIndex  - ListIndex of ListItem to add ItemData
'           lItemData   - Data to set
'Note     : -
'
'   Author: Knuth Konrad 06.07.2016
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Dim lData As Long
   Dim lRetval As Long
   
   If Not IsMissing(lItemData) Then
      lData = CLng(lItemData)
      lRetval = SendMessage(lst.hWnd, LB_SETITEMDATA, lListIndex, ByVal lData)
   End If
   
End Sub
'==============================================================================

Public Function ListBoxGetItem(ByVal lst As ListBox) As String
'------------------------------------------------------------------------------
'Purpose  : Ermittelt den ausgewählten Wert des ausgewählten Eintrags einer Listbox
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 07.03.2008
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Dim sTemp As String, lTemp As Long, lListIndex As Long
   
   ' Stringlänge ermitteln
   lListIndex = SendMessage(lst.hWnd, LB_GETCURSEL, ByVal 0, ByVal 0)
   lTemp = SendMessage(lst.hWnd, LB_GETTEXTLEN, lListIndex, ByVal 0)
   
   sTemp = Space$(lTemp + 1)
   
   If SendMessage(lst.hWnd, LB_GETTEXT, lListIndex, ByVal sTemp) > 0 Then
      ListBoxGetItem = Left$(sTemp, lTemp)
   End If
   
End Function
'==============================================================================

Public Sub ListBoxAdjustWidth(ByVal lst As ListBox, ByVal frm As Form, _
   Optional ByVal dblAddedSpace As Double = 2)
'------------------------------------------------------------------------------
'Purpose  : Erweitert Listbox Width auf die Größe des größten Eintrags
'
'Prereq.  : -
'Parameter: -
'Note     : -
'
'   Author: Knuth Konrad 30.01.2014
'   Source: -
'  Changed: 06.09.2018
'           - Account for different fonts in Listbox and Form
'           02.10.2018
'           - Make sure there's a valid Form object
'------------------------------------------------------------------------------
   Dim lTemp As Long, lLen As Long, lListIndex As Long
   Dim i As Long
   Dim dblLen As Double, dblTemp As Double
   Dim f As Form
   
   ' Stringlänge des längsten Listboxeintrags ermitteln
   ' Font in Listbox und Form identisch?
   If CBool(lst.Font = frm.Font) = False Then
      
      Set f = frm
      f.Font = lst.Font
   
      dblLen = 0
      For i = 0 To lst.ListCount - 1
         
         dblTemp = FormTextWidth(f, lst.List(i))
         
         If dblTemp > dblLen Then
            dblLen = dblTemp
         End If
      
      Next i
   
      Set f = Nothing
   
   Else     '// If CBool(lst.Font = frm.Font) = False
      
      dblLen = 0
      For i = 0 To lst.ListCount - 1
         
         dblTemp = FormTextWidth(frm, lst.List(i))
         
         If dblTemp > dblLen Then
            dblLen = dblTemp
         End If
      
      Next i
   
   End If   '// If CBool(lst.Font = frm.Font) = False
   
   lTemp = SendMessage(lst.hWnd, LB_SETHORIZONTALEXTENT, dblLen + dblAddedSpace, ByVal 0)
   
End Sub
'==============================================================================

Private Function FormTextWidth(ByVal frm As Form, ByVal sText As String) As Double
'------------------------------------------------------------------------------
'Purpose  : Determine the width of a text in pixel
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 25.06.2018
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Dim dblLen As Double
   
   dblLen = frm.TextWidth(sText)
   
   ' Evtl. Twips in Pixel umrechnen
   If frm.ScaleMode = vbTwips Then
      dblLen = dblLen / Screen.TwipsPerPixelX
   End If
   
   FormTextWidth = dblLen
   
End Function
'==============================================================================

Public Function SetWindowMinBox(ByVal frm As Object, _
   ByVal bolRemove As Boolean) As Boolean
'------------------------------------------------------------------------------
'Purpose  : Setzt für ein Fenster bestimmte Attribute wie MinButton, MaxButton,
'           BorderStyle etc.
'
'Prereq.  : -
'Parameter: frm         - Formular das manipuliert werden soll
'           bolRemove   - True = MinimizeBox entfernen
'                         False = MinimizeBox hinzufügen
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 23.05.2000
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Dim lRetval As Long
   
   lRetval = GetWindowLong(frm.hWnd, GWL_STYLE)
   If lRetval > 0 Then
      If bolRemove Then
      '-> MinimizeBox entfernen
         lRetval = lRetval - WS_MINIMIZEBOX
         lRetval = SetWindowLong(frm.hWnd, GWL_STYLE, lRetval)
      Else
      '-> MinimizeBox hinzufügen
         lRetval = lRetval Or WS_MINIMIZEBOX
         lRetval = SetWindowLong(frm.hWnd, GWL_STYLE, lRetval)
      End If
   Else
      lRetval = 0
   End If
   
   SetWindowMinBox = CBool(lRetval)
   
   frm.Refresh
   DoEvents
   
End Function
'==============================================================================

Public Function SetWindowMaxBox(ByVal frm As Object, _
   ByVal bolRemove As Boolean) As Boolean
'------------------------------------------------------------------------------
'Purpose  : Setzt für ein Fenster bestimmte Attribute wie MinButton, MaxButton,
'           BorderStyle etc.
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 23.05.2000
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Dim lRetval As Long
   
   lRetval = GetWindowLong(frm.hWnd, GWL_STYLE)
   If lRetval > 0 Then
      If bolRemove Then
      '-> MaximizeBox entfernen
         lRetval = lRetval - WS_MAXIMIZEBOX
         lRetval = SetWindowLong(frm.hWnd, GWL_STYLE, lRetval)
      Else
      '-> MaximizeBox hinzufügen
         lRetval = lRetval Or WS_MAXIMIZEBOX
         lRetval = SetWindowLong(frm.hWnd, GWL_STYLE, lRetval)
      End If
   Else
      lRetval = 0
   End If
   
   SetWindowMaxBox = CBool(lRetval)
   
   frm.Refresh
   DoEvents
   
End Function
'==============================================================================

Public Function ShowInWindow10Taskbar(ByVal frm As Object, _
   ByVal bolRemove As Boolean) As Boolean
'------------------------------------------------------------------------------
'Purpose  : Setzt für ein Fenster bestimmte Attribute wie MinButton, MaxButton,
'           BorderStyle etc.
'
'Prereq.  : -
'Parameter: frm         - Formular das manipuliert werden soll
'           bolRemove   - True = Von Taskbar entfernen
'                         False = Zu Taskbar hinzufügen
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 15.08.2016
'   Source: http://stackoverflow.com/questions/30809532/how-to-correctly-have-modeless-form-appear-in-taskbar
'           http://stackoverflow.com/questions/8746301/force-modal-form-to-be-shown-in-taskbar
'  Changed: -
'------------------------------------------------------------------------------
   Dim lRetval As Long
   
   'Call SetWindowLong(hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) Or WS_EX_APPWINDOW)
   
   lRetval = GetWindowLong(frm.hWnd, GWL_EXSTYLE)
   If lRetval > 0 Then
      If bolRemove Then
      '-> MinimizeBox entfernen
         lRetval = lRetval And WS_EX_APPWINDOW
         lRetval = SetWindowLong(frm.hWnd, GWL_EXSTYLE, lRetval)
      Else
      '-> MinimizeBox hinzufügen
         lRetval = lRetval Or WS_EX_APPWINDOW
         lRetval = SetWindowLong(frm.hWnd, GWL_EXSTYLE, lRetval)
      End If
   Else
      lRetval = 0
   End If
   
   ShowInWindow10Taskbar = CBool(lRetval)
   
   frm.Hide
   frm.Show
   frm.Refresh
   DoEvents
   
End Function
'==============================================================================

Public Sub SwitchEnable(ByVal obj As Control)
'------------------------------------------------------------------------------
'Purpose  : Switch bei einem Control die Eigenschaft Enabled
'
'Prereq.  : -
'Parameter: obj   -  Control dessen Enabled-Property geändert werden soll
'Note     : -
'
'   Author: Knuth Konrad 07.06.2000
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   
   On Error Resume Next
   
   obj.Enabled = Not obj.Enabled
   
   Err.Clear
   On Error GoTo 0
   
End Sub
'==============================================================================

Public Function TxtSetNumeric(ByVal txt As TextBox) As Long
'------------------------------------------------------------------------------
'Purpose  : Setzt Textbox auf Style ES_NUMERIC
'
'Prereq.  : -
'Parameter: txt   -  Textbox-Control
'Returns  : 0     -  Fehler
'           <>0   -  Erfolg (Vorheriger Window Style)
'Note     : -
'
'   Author: Knuth Konrad 18.08.2003
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Dim lRet As Long
   
   lRet = GetWindowLong(txt.hWnd, GWL_STYLE)
   TxtSetNumeric = SetWindowLong(txt.hWnd, GWL_STYLE, lRet Or ES_NUMBER)
   
End Function
'==============================================================================

Public Function TxtIsEmpty(ByVal txt As TextBox) As Boolean
'------------------------------------------------------------------------------
'Purpose  : Hat eine Textbox einen Inhalt?
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author: Knuth Konrad 16.12.2005
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   
   If Len(Trim$(txt.Text)) < 1 Then
      TxtIsEmpty = True
   Else
      TxtIsEmpty = False
   End If
   
End Function
'==============================================================================

Public Sub TxtSetScrollbars(ByVal txt As TextBox, ByVal eScrollbars As eTxtScrollbar)
'------------------------------------------------------------------------------
'Purpose  : Sets the scrollbar property of a textbox at runtime
'
'Prereq.  : -
'Parameter: -
'Note     : -
'
'   Author: "strongm"
'   Source: https://www.tek-tips.com/viewthread.cfm?qid=177192
'  Changed: Select type of scrollbar(s)
'------------------------------------------------------------------------------
   Dim lCurrentStyle As Long
   
   lCurrentStyle = GetWindowLong(txt.hWnd, GWL_STYLE)
   
   Select Case eScrollbars

      Case scrollNone

       ' Make invsible if visible by changing window style
       SetWindowLong txt.hWnd, GWL_STYLE, lCurrentStyle Xor (WS_VSCROLL Or WS_HSCROLL)
      
      Case scrollHorizontal
         
         ' Make visible by changing window style
         SetWindowLong txt.hWnd, GWL_STYLE, lCurrentStyle Or WS_HSCROLL

      Case scrollVertical
         
         ' Make visible by changing window style
         SetWindowLong txt.hWnd, GWL_STYLE, lCurrentStyle Or WS_VSCROLL

      Case scrollBoth
         
         ' Make visible by changing window style
         SetWindowLong txt.hWnd, GWL_STYLE, lCurrentStyle Or (WS_VSCROLL And WS_HSCROLL)

   End Select
      
'   If Visible Then
'       ' Make visible by changing window style
'       SetWindowLong txt.hWnd, GWL_STYLE, lCurrentStyle Or WS_VSCROLL
'   ElseIf lCurrentStyle And WS_VSCROLL Then
'       ' Make invsible if visible by changing window style
'       SetWindowLong txt.hWnd, GWL_STYLE, lCurrentStyle Xor WS_VSCROLL
'   End If
   
   ' Scrollbar style is cached, so we need to do SetWindowPos to activate the style change
   ' Without this line the status of the scrollbar will NOT change
   SetWindowPos txt.hWnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED + SWP_NOMOVE + SWP_NOSIZE + SWP_NOZORDER

End Sub
'==============================================================================

Public Sub SelectText()
  
If (TypeOf Screen.ActiveControl Is TextBox) Then
   If Not CBool(GetAsyncKeyState(vbKeyLButton) Or _
      GetAsyncKeyState(vbKeyMButton) Or _
      GetAsyncKeyState(vbKeyRButton)) Then

      With Screen.ActiveControl
         .SelStart = 0
         .SelLength = Len(.Text)
      End With
   End If
End If

End Sub
'==============================================================================

Public Function CtrlUBound(ByVal ctrl As Object) As Long
'------------------------------------------------------------------------------
'Purpose  : Ermittelt die obere Dimensionsgrenze eines Control-Arrays
'
'Prereq.  : -
'Parameter: ctrl  - Zu prüfendes Control-Array
'Returns  : UBound des Arrays oder -1 falls ctrl keine Control (Array) ist
'Note     : -
'
'   Author: Knuth Konrad 08.01.2008
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Dim i As Long, sTemp As String
   
   On Error GoTo CtrlUBoundErrHandler
   
   ' In einer Schleife immer auf das Property Name zugreifen
   ' bis ein Fehler ausgelöst wird = kein weiteres Control
   ' mehr im Array vorhanden.
   Do
      sTemp = ctrl(i).Name
      i = i + 1
   Loop Until i > 255
   
CtrlUBoundErrExit:
   On Error GoTo 0
   Exit Function
   
CtrlUBoundErrHandler:
   
   If i > 0 Then
      CtrlUBound = i - 1
   Else
      CtrlUBound = -1
   End If
   
   Resume CtrlUBoundErrExit
   
End Function
'==============================================================================

Public Function CtrlRedrawDisable(ByVal hWnd As Long) As Long
'------------------------------------------------------------------------------
'Purpose  : Disables window redraw of hWnd
'
'Prereq.  : -
'Parameter: -
'Note     : This message can be useful if an application must add several items to a list box.
'           The application can call this message with wParam set to FALSE, add the items, and
'           then call the message again with wParam set to TRUE. Finally, the application can
'           call the InvalidateRect function to cause the list box to be repainted.
'
'   Author: Knuth Konrad 17.06.2013
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   
   CtrlRedrawDisable = SendMessage(hWnd, WM_SETREDRAW, 0, ByVal 0)
   
End Function
'==============================================================================

Public Function CtrlRedrawEnable(ByVal hWnd As Long) As Long
'------------------------------------------------------------------------------
'Purpose  : Enables window redraw of hWnd
'
'Prereq.  : -
'Parameter: -
'Note     : This message can be useful if an application must add several items to a list box.
'           The application can call this message with wParam set to FALSE, add the items, and
'           then call the message again with wParam set to TRUE. Finally, the application can
'           call the InvalidateRect function to cause the list box to be repainted.
'
'   Author: Knuth Konrad 17.06.2013
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Dim lRetval As Long, udtRect As ApiRECT
   
   lRetval = SendMessage(hWnd, WM_SETREDRAW, 1, ByVal 0)
   Call GetWindowRect(hWnd, udtRect)
   CtrlRedrawEnable = lRetval Or InvalidateRect(hWnd, udtRect, 1)
   
End Function
'==============================================================================

Public Sub LabelShortenPathText(ByVal sPath As String, _
   ByVal lbl As Label, ByVal frm As Form)
'------------------------------------------------------------------------------
'Purpose  : -
'
'Prereq.  : -
'Parameter: -
'Note     : -
'
'   Author: Knuth Konrad 13.12.2013
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Dim udtRect As ApiRECT
   Dim hDC As Long, lRet As Long, lErr As Long
   Dim txt As TextBox
   
   ' Create a temporary Textbox on top of the label in order to
   ' measure the label's dimensions
   Set txt = frm.Controls.Add("VB.TextBox", "txt")
   
   With txt
      .Move lbl.Left, lbl.Top, lbl.Width, lbl.Height
   End With
   
   Call GetWindowRect(txt.hWnd, udtRect)
   
   frm.Controls.Remove "txt"
   
   ' All info gathered: draw the (path) text
   lbl.Visible = False
   hDC = GetDC(frm.hWnd)
   lRet = DrawText(hDC, sPath, LenB(sPath), udtRect, DT_PATH_ELLIPSIS Or DT_LEFT Or DT_MODIFYSTRING Or DT_SINGLELINE)
   
   If lRet = 0 Then
      lErr = GetLastError()
   End If
   
End Sub
'==============================================================================

Public Function InitCommonControlsVB() As Boolean
'------------------------------------------------------------------------------
'Purpose  : -
'
'Prereq.  : -
'Parameter: -
'Returns  : -
'Note     : -
'
'   Author:  Steve McMahon
'   Source: http://www.vbaccelerator.com/home/VB/Code/Libraries/XP_Visual_Styles/Using_XP_Visual_Styles_in_VB/article.asp
'  Changed: -
'------------------------------------------------------------------------------
      
   On Error Resume Next
   
   Dim iccex As tagInitCommonControlsEx
   ' Ensure CC available:
   
   With iccex
      .lngSize = LenB(iccex)
      .lngICC = ICC_USEREX_CLASSES
   End With
   
   InitCommonControlsEx iccex
   
   InitCommonControlsVB = (Err.Number = 0)
   
   On Error GoTo 0
   
End Function
'==============================================================================

