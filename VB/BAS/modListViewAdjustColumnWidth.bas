Attribute VB_Name = "modListViewAdjustColumnWidth"
'------------------------------------------------------------------------------
'Purpose  : Autosize Listview columns
'
'Prereq.  : -
'Note     : -
'
'   Author: Harald M. Genauck
'   Source: http://www.aboutvb.de/khw/artikel/khwlistviewadjustcolumnwidth.htm
'  Changed: 14.07.2016
'           - Prevent runtime error if Listview has no headers/columns
'------------------------------------------------------------------------------
Option Explicit
DefLng A-Z
'------------------------------------------------------------------------------
'*** Constants ***
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'*** Enumeration/TYPEs ***
'------------------------------------------------------------------------------
Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
'------------------------------------------------------------------------------
'*** Declares ***
'------------------------------------------------------------------------------
Private Declare Function GetClientRect Lib "USER32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function LockWindowUpdate Lib "USER32" (ByVal hwndLock As Long) As Long
Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'------------------------------------------------------------------------------
'*** Variables ***
'------------------------------------------------------------------------------
'==============================================================================

Sub ListViewAdjustColumnWidth(ListView As ListView, Optional Position As Integer, _
   Optional IncludeHeaders As Boolean, Optional ByVal LastColumnFillSize As Boolean)
'------------------------------------------------------------------------------
'Purpose  : -
'
'Prereq.  : -
'Parameter: -
'Note     : -
'
'   Author: Harald M. Genauck
'   Source: -
'  Changed: 14.07.2016
'           - Prevent runtime error if Listview has no headers/columns
'------------------------------------------------------------------------------
   Dim i As Integer
   Dim nListItem As ListItem
   Dim nKey As String
   Dim nColumn As Integer
   Dim nPosition As Integer
   Dim nSmallIconsSet As Boolean
   Dim nRect As RECT
   Dim nWidth As Single
   
   Const LVM_SETCOLUMNWIDTH As Long = &H101E
   Const LVSCW_AUTOSIZE As Long = -1&
       
   Const PROCEDURE_NAME As String = "modListViewAdjustColumnWidth:ListViewAdjustColumnWidth->"
   'Call gobjLog.AppTrace(PROCEDURE_NAME, ListView, Position, IncludeHeaders, LastColumnFillSize)
   
   With ListView
      
      LockWindowUpdate .hWnd
      
      If IncludeHeaders And .ColumnHeaders.Count > 0 Then
         
         If .SmallIcons Is Nothing Then
             Set .SmallIcons = .ColumnHeaderIcons
             nSmallIconsSet = True
         End If
          
          Select Case Position
              Case 1 To .ColumnHeaders.Count
                  nKey = CStr(Now)
                  If zHasIcon(.ColumnHeaders(1)) Then
                      Set nListItem = .ListItems.Add(1, nKey, .ColumnHeaders(1).Text & "           ")
                  Else
                      Set nListItem = .ListItems.Add(1, nKey, .ColumnHeaders(1).Text & "  ")
                  End If
                  nPosition = .ColumnHeaders(1).Position
                  If nPosition = Position Then
                      nColumn = 0
                  End If
                  For i = 2 To .ColumnHeaders.Count
                      If zHasIcon(.ColumnHeaders(i)) Then
                          nListItem.ListSubItems.Add , , .ColumnHeaders(i).Text & "         "
                      Else
                          nListItem.ListSubItems.Add , , .ColumnHeaders(i).Text
                      End If
                      nPosition = .ColumnHeaders(i).Position
                      If nPosition = Position Then
                          nColumn = i - 1
                      End If
                  Next
                  SendMessage .hWnd, LVM_SETCOLUMNWIDTH, nColumn, LVSCW_AUTOSIZE
              Case Else
                  nKey = CStr(Now)
                  If zHasIcon(.ColumnHeaders(1)) Then
                      Set nListItem = .ListItems.Add(1, nKey, .ColumnHeaders(1).Text & "           ")
                  Else
                      Set nListItem = .ListItems.Add(1, nKey, .ColumnHeaders(1).Text & "  ")
                  End If
                  SendMessage .hWnd, LVM_SETCOLUMNWIDTH, 0, LVSCW_AUTOSIZE
                  For i = 2 To .ColumnHeaders.Count
                      If zHasIcon(.ColumnHeaders(i)) Then
                          nListItem.ListSubItems.Add , , .ColumnHeaders(i).Text & "         "
                      Else
                          nListItem.ListSubItems.Add , , .ColumnHeaders(i).Text
                      End If
                      SendMessage .hWnd, LVM_SETCOLUMNWIDTH, i - 1, LVSCW_AUTOSIZE
                  Next
          End Select
          .ListItems.Remove nKey
          If nSmallIconsSet Then
              Set .SmallIcons = Nothing
          End If
      Else     '// If IncludeHeaders
          
          Select Case Position
              Case 1 To .ColumnHeaders.Count
                  nPosition = .ColumnHeaders(Position).Position
                  If nPosition = Position Then
                      SendMessage .hWnd, LVM_SETCOLUMNWIDTH, Position - 1, LVSCW_AUTOSIZE
                  Else
                      SendMessage .hWnd, LVM_SETCOLUMNWIDTH, nPosition - 1, LVSCW_AUTOSIZE
                  End If
              Case Else
                  For i = 0 To .ColumnHeaders.Count
                      SendMessage .hWnd, LVM_SETCOLUMNWIDTH, i, LVSCW_AUTOSIZE
                  Next
          End Select
      
      End If   '// If IncludeHeaders
      
      If LastColumnFillSize Then
          For i = 1 To .ColumnHeaders.Count - 1
              nWidth = nWidth + .ColumnHeaders(i).Width
          Next 'i
          GetClientRect .hWnd, nRect
          nRect.Right = nRect.Right * Screen.TwipsPerPixelX
          If nRect.Right > nWidth Then
              .ColumnHeaders(.ColumnHeaders.Count).Width = nRect.Right - nWidth
          End If
      End If
      
      .Refresh
      
   End With
   
   LockWindowUpdate 0&
   
End Sub
'==============================================================================

Public Sub ListViewLastColumnFillSize(ListView As ListView)
'------------------------------------------------------------------------------
'Purpose  : Stretches the last column of a ListView to fill the remaining spaces
'
'Prereq.  : -
'Parameter: -
'Note     : -
'
'   Author: Knuth Konrad 25.06.2018
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
   Dim i As Integer
   Dim nRect As RECT
   Dim nWidth As Single
       
   Const PROCEDURE_NAME As String = "modListViewAdjustColumnWidth:ListViewLastColumnFillSize->"
   'Call gobjLog.AppTrace(PROCEDURE_NAME, ListView)
   
   With ListView
       For i = 1 To .ColumnHeaders.Count - 1
           nWidth = nWidth + .ColumnHeaders(i).Width
       Next 'i
       GetClientRect .hWnd, nRect
       nRect.Right = (nRect.Right - 1) * Screen.TwipsPerPixelX
       If nRect.Right > nWidth Then
           .ColumnHeaders(.ColumnHeaders.Count).Width = nRect.Right - nWidth
       End If
   End With
   
End Sub
'==============================================================================

Private Function zHasIcon(ColumnHeader As ColumnHeader) As Boolean
   
   Const PROCEDURE_NAME As String = "modListViewAdjustColumnWidth:zHasIcon->"
   'Call gobjLog.AppTrace(PROCEDURE_NAME, ColumnHeader)
   
   Dim nIcon As Variant
   
   nIcon = ColumnHeader.Icon
   If nIcon = "0" Then
   Else
       zHasIcon = CBool(Len(nIcon))
   End If
   
End Function
'==============================================================================
