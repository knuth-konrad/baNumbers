VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9300
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13245
   LinkTopic       =   "Form1"
   ScaleHeight     =   9300
   ScaleWidth      =   13245
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ListBox List1 
      Height          =   8835
      Left            =   9360
      TabIndex        =   22
      Top             =   240
      Width           =   3735
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   8895
      Left            =   3480
      TabIndex        =   18
      Top             =   240
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   15690
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Clear results"
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   8760
      Value           =   1  'Aktiviert
      Width           =   2895
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Integer to DWord"
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   8160
      Width           =   3015
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Integer to Word"
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   7680
      Width           =   3015
   End
   Begin VB.CommandButton Command6 
      Caption         =   "TypeOfVariant"
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   7200
      Width           =   3015
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Swap"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   6720
      Width           =   3015
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Format number"
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   6240
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Speed comparison"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   5760
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Array duplicate"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   5280
      Width           =   3015
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data type"
      Height          =   3375
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   3015
      Begin VB.TextBox txtElements 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   600
         TabIndex        =   21
         Text            =   "1000"
         Top             =   2880
         Width           =   2055
      End
      Begin VB.OptionButton optData 
         Caption         =   "Date"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   19
         Top             =   2280
         Width           =   1815
      End
      Begin VB.OptionButton optData 
         Caption         =   "Integer"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   9
         Top             =   480
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton optData 
         Caption         =   "Long"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   8
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton optData 
         Caption         =   "Single"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   7
         Top             =   1200
         Width           =   1815
      End
      Begin VB.OptionButton optData 
         Caption         =   "Double"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   6
         Top             =   1560
         Width           =   1815
      End
      Begin VB.OptionButton optData 
         Caption         =   "Currency"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   5
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "# of array elements"
         Height          =   195
         Left            =   600
         TabIndex        =   20
         Top             =   2640
         Width           =   1350
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Method"
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   3135
      Begin VB.OptionButton optMethod 
         Caption         =   "Sort"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optMethod 
         Caption         =   "Median"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Array Sort / Median of Array"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   4800
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z

Public Enum eDataType
   dtInteger = 0
   dtLong
   dtSingle
   dtDouble
   dtCurrency
   dtDate
End Enum

Const ARRAY_MAX As Long = 1000

Private Sub Command1_Click()
   
   Dim j As Long
   
   ' Number of array elements
   Dim lArrayMax As Long
   lArrayMax = Val(Me.txtElements.Text)
   
   ReDim i(1 To lArrayMax) As Integer
   ReDim l(1 To lArrayMax) As Long
   ReDim s(1 To lArrayMax) As Single
   ReDim d(1 To lArrayMax) As Double, dtm(1 To lArrayMax) As Date
   ReDim c(1 To lArrayMax) As Currency
   
   Dim vnt As Variant, oItem As ListItem
   
'   DebugProfileReset
'   DebugProfile ("Before Loop")
'   For Index = 0 To 10000
'       DebugProfile ("Before Call To Foo")
'       Foo
'       DebugProfile ("Before Call To Bar")
'       Bar
'       DebugProfile ("Before Call To Baz")
'       Baz
'   Next Index
'   DebugProfile ("After Loop")
'   DebugProfileStop
   
   Clearcontrols
   
   Dim pp As New cPerformanceProfile
   pp.DebugProfileReset
   
   Randomize Timer
   
   SetupListview 0, "Unsorted", "Sorted", "Median"
   
   Me.List1.AddItem pp.DebugProfile("Creating random array", ecbtLoopBegin)
   
   For j = 1 To lArrayMax
   
      Select Case GetOption(optData())
      Case eDataType.dtInteger
         i(j) = Int((1000 - 0 + 1) * Rnd + 0)
         Set oItem = Me.ListView1.ListItems.Add(, , Str$(i(j)))
      Case eDataType.dtLong
         ' Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
         l(j) = Int((1000 - 0 + 1) * Rnd + 0)
         Set oItem = Me.ListView1.ListItems.Add(, , Str$(l(j)))
      Case eDataType.dtSingle
         s(j) = Int((1000 - 0 + 1) * Rnd + 0) + Rnd
         Set oItem = Me.ListView1.ListItems.Add(, , Str$(s(j)))
      Case eDataType.dtDouble
         d(j) = Int((1000 - 0 + 1) * Rnd + 0) + Rnd
         Set oItem = Me.ListView1.ListItems.Add(, , Str$(d(j)))
      Case eDataType.dtCurrency
         c(j) = Int((1000 - 0 + 1) * Rnd + 0) + Rnd
         Set oItem = Me.ListView1.ListItems.Add(, , Str$(c(j)))
      Case eDataType.dtDate
         ' Date is actually a double in VB
         dtm(j) = Int((1000 - 0 + 1) * Rnd + 0) + Rnd
         d(j) = dtm(j)
         Set oItem = Me.ListView1.ListItems.Add(, , dtm(j))
      End Select
   
   Next j
   
   Me.List1.AddItem pp.DebugProfile("End creating random array", ecbtLoopEnd)
   
   
   ' Sort the array
   Me.List1.AddItem pp.DebugProfile("Sorting array", ecbtStatementBegin)
   
   Select Case GetOption(optData())
   Case eDataType.dtInteger
      baSortInt i()
   Case eDataType.dtLong
      baSortLong l()
   Case eDataType.dtSingle
      baSortSingle s()
   Case eDataType.dtDouble
      baSortDouble d()
   Case eDataType.dtCurrency
      baSortCurrency c()
   Case eDataType.dtDate
      baSortDouble d()
   End Select
   
   Me.List1.AddItem pp.DebugProfile("Done sorting array", ecbtStatementEnd)
   
   ' Compute the median
   Me.List1.AddItem pp.DebugProfile("Computing median", ecbtStatementBegin)
   
   Select Case GetOption(optData())
   Case eDataType.dtInteger
      vnt = CInt(baMedianInt(i()))
   Case eDataType.dtLong
      vnt = CLng(baMedianLong(l()))
   Case eDataType.dtSingle
      vnt = CSng(baMedianSingle(s()))
   Case eDataType.dtDouble
      vnt = CDbl(baMedianDouble(d()))
   Case eDataType.dtCurrency
      vnt = CCur(baMedianCurrency(c()))
   Case eDataType.dtDate
      vnt = CDate(CDbl(baMedianDouble(d())))
   End Select
   
   Me.List1.AddItem pp.DebugProfile("Done computing median", ecbtStatementEnd)
      
   ' Display the sorted array
   Me.List1.AddItem pp.DebugProfile("Displaying sorted array", ecbtLoopBegin)
   
   For j = 1 To lArrayMax
   
      Select Case GetOption(optData())
      Case eDataType.dtInteger
         Me.ListView1.ListItems(j).ListSubItems.Add Text:=Str$(i(j))
      Case eDataType.dtLong
         Me.ListView1.ListItems(j).ListSubItems.Add Text:=Str$(l(j))
      Case eDataType.dtSingle
         Me.ListView1.ListItems(j).ListSubItems.Add Text:=Str$(s(j))
      Case eDataType.dtDouble
         Me.ListView1.ListItems(j).ListSubItems.Add Text:=Str$(d(j))
      Case eDataType.dtCurrency
         Me.ListView1.ListItems(j).ListSubItems.Add Text:=Str$(c(j))
      Case eDataType.dtDate
         Me.ListView1.ListItems(j).ListSubItems.Add Text:=CDate(d(j))
      End Select
   
   Next j

   Me.ListView1.ListItems(1).ListSubItems.Add Text:=Str$(vnt)
   
   Me.List1.AddItem pp.DebugProfile("End displaying sorted array", ecbtLoopEnd)
   
   ListViewAdjustColumnWidth Me.ListView1, , True
   ListBoxAdjustWidth Me.List1, Me
   
End Sub

Private Sub Command2_Click()
   
   ' Number of array elements
   Dim lArrayMax As Long
   lArrayMax = Val(Me.txtElements.Text)
   
   Dim j As Long
   Dim oItem As ListItem
   
   ReDim i(1 To lArrayMax) As Integer, i1(1 To lArrayMax) As Integer
   ReDim l(1 To lArrayMax) As Long, l1(1 To lArrayMax) As Long
   ReDim s(1 To lArrayMax) As Single, s1(1 To lArrayMax) As Single
   ReDim d(1 To lArrayMax) As Double, d1(1 To lArrayMax) As Double
   ReDim c(1 To lArrayMax) As Currency, c1(1 To lArrayMax) As Currency
   
   Randomize Timer
   
'Sub SetupListview(ByVal lLvwRows As Long, ParamArray vntColHdrs() As Variant)
   
   SetupListview 0, "Integer", "", "Long", "", "Single", "", "Double", "", "Currency", ""
   
   For j = 1 To lArrayMax
   
      i(j) = Int((1000 - 0 + 1) * Rnd + 0)
      l(j) = Int((1000 - 0 + 1) * Rnd + 0)
      s(j) = Int((1000 - 0 + 1) * Rnd + 0) + Rnd
      d(j) = Int((1000 - 0 + 1) * Rnd + 0) + Rnd
      c(j) = Int((1000 - 0 + 1) * Rnd + 0) + Rnd
      
      ' List1.AddItem Str$(i(j)) & vbTab & Str$(l(j)) & vbTab & Str$(s(j)) & vbTab & Str$(d(j)) & vbTab & Str$(c(j))
      
      With Me.ListView1
         Set oItem = .ListItems.Add(Text:=Str$(i(j)))
      End With
      
   Next j
   
   ' List1.AddItem String$(10, "-") & "  Magic: ArraySet  " & String$(10, "-")
   
   baDuplicateIntegerArray i(), i1()
   baDuplicateLongArray l(), l1()
   baDuplicateSingleArray s(), s1()
   baDuplicateDoubleArray d(), d1()
   baDuplicateCurrencyArray c(), c1()
   
   For j = 1 To lArrayMax
      
      'List1.AddItem Str$(i1(j)) & vbTab & Str$(l1(j)) & vbTab & Str$(s1(j)) & vbTab & Str$(d1(j)) & vbTab & Str$(c1(j))
      
   Next j
   
   'List1.AddItem String$(10, "-") & "  Ende  " & String$(10, "-")
   
End Sub

Private Sub Command3_Click()
   
   ' Number of array elements
   Dim lArrayMax As Long
   lArrayMax = Val(Me.txtElements.Text)
   
   Dim j As Long
   
   ReDim i(0 To lArrayMax) As Integer, i1(0 To lArrayMax) As Integer
   ReDim l(0 To lArrayMax) As Long, l1(0 To lArrayMax) As Long
   ReDim s(0 To lArrayMax) As Single, s1(0 To lArrayMax) As Single
   ReDim d(0 To lArrayMax) As Double, d1(0 To lArrayMax) As Double
   ReDim c(0 To lArrayMax) As Currency, c1(0 To lArrayMax) As Currency
   
   Dim vnt As Variant, vnt1 As Variant
   
   Randomize Timer
   
'   List1.AddItem String$(10, "-") & " Start " & String$(10, "-")
   
   For j = 0 To lArrayMax
   
      Select Case GetOption(optData())
      Case eDataType.dtInteger
         i(j) = Int((1000 - 0 + 1) * Rnd + 0)
         'List1.AddItem Str$(i(j))
      Case eDataType.dtLong
         ' Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
         l(j) = Int((1000 - 0 + 1) * Rnd + 0)
         'List1.AddItem Str$(l(j))
      Case eDataType.dtSingle
         s(j) = Int((1000 - 0 + 1) * Rnd + 0) + Rnd
         'List1.AddItem Str$(s(j))
      Case eDataType.dtDouble
         d(j) = Int((1000 - 0 + 1) * Rnd + 0) + Rnd
         'List1.AddItem Str$(d(j))
      Case eDataType.dtCurrency
         c(j) = Int((1000 - 0 + 1) * Rnd + 0) + Rnd
         'List1.AddItem Str$(c(j))
      End Select
   
   Next j
   
   Select Case GetOption(optMethod())
   Case 0
   '  Sort
      'List1.AddItem String$(10, "-") & "  Magic: Sort  " & String$(10, "-")
      Select Case GetOption(optData())
      Case eDataType.dtInteger
         ' baNumbers.baDuplicateIntegerArray i(), i1()
         baSortInt i()
      Case eDataType.dtLong
         ' baNumbers.baDuplicateLongArray l(), l1()
         baSortLong l()
      Case eDataType.dtSingle
         ' baNumbers.baDuplicateSingleArray s(), s1()
         baSortSingle s()
      Case eDataType.dtDouble
         ' baNumbers.baDuplicateDoubleArray d(), d1()
         baSortDouble d()
      Case eDataType.dtCurrency
         ' baNumbers.baDuplicateCurrencyArray c(), c1()
         baSortCurrency c()
      End Select
   Case 1
      'List1.AddItem String$(10, "-") & "  Magic: Median  " & String$(10, "-")
      Select Case GetOption(optData())
      Case eDataType.dtInteger
         vnt = baMedianInt(i())
      Case eDataType.dtLong
         vnt = baMedianLong(l())
      Case eDataType.dtSingle
         vnt = baMedianSingle(s())
      Case eDataType.dtDouble
         vnt = baMedianDouble(d())
      Case eDataType.dtCurrency
         vnt = baMedianCurrency(c())
      End Select
      'List1.AddItem Str$(vnt)
   End Select
   
   If optMethod(0).Value = True Then
      For j = 0 To lArrayMax
      
         Select Case GetOption(optData())
         Case eDataType.dtInteger
            'List1.AddItem Str$(i(j))
         Case eDataType.dtLong
            'List1.AddItem Str$(l(j))
         Case eDataType.dtSingle
            'List1.AddItem Str$(s(j))
         Case eDataType.dtDouble
            'List1.AddItem Str$(d(j))
         Case eDataType.dtCurrency
            'List1.AddItem Str$(c(j))
         End Select
      
      Next j
   End If
   
   'List1.AddItem String$(10, "-") & "  Ende  " & String$(10, "-")
   
End Sub

Private Sub Command4_Click()
   
   Dim cur As Currency, sResult As String
   
   cur = Val(InputBox("Number to format", "Format number", "1234.56"))
   
   sResult = baFormatNumber(cur, LANG_FRENCH, SUBLANG_FRENCH_SWISS)
   
   MsgBox "Swiss: " & sResult
   
   sResult = baFormatNumber(cur, LANG_ENGLISH, SUBLANG_ENGLISH_US)
   
   MsgBox "English (US): " & sResult
   
End Sub

Private Sub Command5_Click()

   Dim cur1 As Currency, cur2 As Currency
   Dim d1 As Double, d2 As Double
   Dim i1 As Integer, i2 As Integer
   Dim l1 As Long, l2 As Long
   Dim s1 As Single, s2 As Single
   Dim lRet As Long
   
   'List1.AddItem "Swap ->"
   
   'List1.AddItem "Currency"
   cur1 = 1.23
   'List1.AddItem "d1: " & CStr(cur1)
   cur2 = 4.56
   'List1.AddItem "d1: " & CStr(cur2)
   
   lRet = baSwapCurrency(cur1, cur2)
   
   'List1.AddItem "dlRet " & CStr(lRet)
   'List1.AddItem "cur1: " & CStr(cur1)
   'List1.AddItem "cur2: " & CStr(cur2)
   
   'List1.AddItem "Double"
   d1 = 1.23
   'List1.AddItem "d1: " & CStr(d1)
   d2 = 4.56
   'List1.AddItem "d2: " & CStr(d2)
   
   lRet = baSwapDouble(d1, 2)
   
   'List1.AddItem "dlRet " & CStr(lRet)
   'List1.AddItem "d1: " & CStr(d1)
   'List1.AddItem "d2: " & CStr(d2)
   
   'List1.AddItem "Integer"
   i1 = 1
   'List1.AddItem "i1: " & CStr(i1)
   i2 = 2
   'List1.AddItem "i2: " & CStr(i2)
   
   lRet = baSwapInteger(i1, i2)
   
   'List1.AddItem "dlRet " & CStr(lRet)
   'List1.AddItem "i1: " & CStr(i1)
   'List1.AddItem "i2: " & CStr(i2)
   
   'List1.AddItem "Long"
   l1 = 100
   'List1.AddItem "l1: " & CStr(l1)
   l2 = 200
   'List1.AddItem "l2: " & CStr(l2)
   
   lRet = baSwapLong(l1, l2)
   
   'List1.AddItem "dlRet " & CStr(lRet)
   'List1.AddItem "l1: " & CStr(l1)
   'List1.AddItem "l2: " & CStr(l2)
   
   'List1.AddItem "Single"
   s1 = 1.2345
   'List1.AddItem "s1: " & CStr(s1)
   s2 = 6.0987
   'List1.AddItem "s2: " & CStr(s2)
   
   lRet = baSwapSingle(s1, s2)
   
   'List1.AddItem "dlRet " & CStr(lRet)
   'List1.AddItem "s1: " & CStr(s1)
   'List1.AddItem "s2: " & CStr(s2)

End Sub

Private Sub Command6_Click()

   Dim vnt As Variant, vnt1 As Variant
   Dim b As Byte
   Dim i As Integer
   Dim l As Long
   Dim f As Single
   Dim d As Double
   Dim c As Currency
   Dim dtm As Date
   Dim s As String
   
   b = 1
   i = 2
   l = 3
   f = 1.23
   d = 4.56
   c = 7.89
   dtm = Now
   s = "String"
   Dim a(1 To 2) As Integer
   
   
   'List1.AddItem "TypeOfVariant->"
   
   vnt = b
   'List1.AddItem "vnt = b: " & CStr(TypeOfVariant(vnt))
   
   vnt = i
   'List1.AddItem "vnt = i: " & CStr(TypeOfVariant(vnt))
   
   vnt = l
   'List1.AddItem "vnt = l: " & CStr(TypeOfVariant(vnt))
   
   vnt = f
   'List1.AddItem "vnt = f: " & CStr(TypeOfVariant(vnt))
   
   vnt = d
   'List1.AddItem "vnt = d: " & CStr(TypeOfVariant(vnt))
   
   vnt = c
   'List1.AddItem "vnt = : " & CStr(TypeOfVariant(vnt))
   
   vnt = dtm
   'List1.AddItem "vnt = dtm: " & CStr(TypeOfVariant(vnt))
   
   vnt = vnt1
   'List1.AddItem "vnt = vnt1: " & CStr(TypeOfVariant(vnt))
   
   vnt = s
   'List1.AddItem "vnt = s: " & CStr(TypeOfVariant(vnt))
   
   vnt = a
   'List1.AddItem "vnt = a: " & CStr(TypeOfVariant(vnt))

End Sub

Private Sub Command7_Click()
   
   Dim i As Integer, l As Long
   
   i = Val(InputBox("Integer value"))
   l = Int2Wrd(i)
   
   MsgBox CStr(i) & " is " & CStr(l), vbOKOnly
   
End Sub

Private Sub Command8_Click()
   
   Dim i As Integer, c As Currency
   
   i = Val(InputBox("Integer value"))
   c = Int2DWrd(i)
   
   MsgBox CStr(i) & " is " & CStr(c), vbOKOnly
   
End Sub

Sub Clearcontrols()

   With Me
      If .Check1.Value = vbChecked Then
         .List1.Clear
         .ListView1.ListItems.Clear
         .ListView1.ColumnHeaders.Clear
      End If
   End With

End Sub

Sub SetupListview(ByVal lLvwRows As Long, ParamArray vntColHdrs() As Variant)

   Dim i As Long
   
   ' Set up the Listview control
   With Me.ListView1
      .View = lvwReport
      .ListItems.Clear
      ' Column (headers)
      .ColumnHeaders.Clear
      For i = LBound(vntColHdrs) To UBound(vntColHdrs)
         .ColumnHeaders.Add Text:=vntColHdrs(i)
      Next i
   End With
   
End Sub

Private Sub Form_Load()
   
   ' Initialize the controls
   
   Me.txtElements.Text = CStr(ARRAY_MAX)
   SetupListview 0, ""
   
End Sub
