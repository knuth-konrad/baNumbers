VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5904
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   12036
   LinkTopic       =   "Form1"
   ScaleHeight     =   5904
   ScaleWidth      =   12036
   StartUpPosition =   3  'Windows-Standard
   Begin MSComctlLib.ListView ListView1 
      Height          =   5415
      Left            =   3480
      TabIndex        =   13
      Top             =   240
      Width           =   8295
      _ExtentX        =   14626
      _ExtentY        =   9546
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
   Begin VB.CommandButton cmdTypeOfVariant 
      Caption         =   "TypeOfVariant"
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   5280
      Width           =   3015
   End
   Begin VB.CommandButton cmdFormatNumber 
      Caption         =   "Format number"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   4800
      Width           =   3015
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data type"
      Height          =   3855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3015
      Begin VB.OptionButton optData 
         Caption         =   "Byte"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.TextBox txtElements 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   600
         TabIndex        =   9
         Text            =   "1000"
         Top             =   3240
         Width           =   2055
      End
      Begin VB.OptionButton optData 
         Caption         =   "Date"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   7
         Top             =   2640
         Width           =   1815
      End
      Begin VB.OptionButton optData 
         Caption         =   "Integer"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   2
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton optData 
         Caption         =   "Long"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   3
         Top             =   1200
         Width           =   1815
      End
      Begin VB.OptionButton optData 
         Caption         =   "Single"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   4
         Top             =   1560
         Width           =   1815
      End
      Begin VB.OptionButton optData 
         Caption         =   "Double"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   5
         Top             =   1920
         Width           =   1815
      End
      Begin VB.OptionButton optData 
         Caption         =   "Currency"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   6
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "# of array elements"
         Height          =   195
         Left            =   600
         TabIndex        =   8
         Top             =   3000
         Width           =   1350
      End
   End
   Begin VB.CommandButton cmdArraySortMedian 
      Caption         =   "Array Sort / Median of Array"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   4320
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
   dtByte = 0
   dtInteger
   dtLong
   dtSingle
   dtDouble
   dtCurrency
   dtDate
End Enum

Const ARRAY_MAX As Long = 1000

Private Sub cmdArraySortMedian_Click()
   
   Dim j As Long
   
   ' Number of array elements
   Dim lArrayMax As Long
   lArrayMax = Val(Me.txtElements.Text)
   
   ReDim b(1 To lArrayMax) As Byte, b1(1 To lArrayMax) As Byte, b2(1 To lArrayMax) As Byte
   ReDim i(1 To lArrayMax) As Integer, i1(1 To lArrayMax) As Integer, i2(1 To lArrayMax) As Integer
   ReDim l(1 To lArrayMax) As Long, l1(1 To lArrayMax) As Long, l2(1 To lArrayMax) As Long
   ReDim s(1 To lArrayMax) As Single, s1(1 To lArrayMax) As Single, s2(1 To lArrayMax) As Single
   ReDim d(1 To lArrayMax) As Double, d1(1 To lArrayMax) As Double, d2(1 To lArrayMax) As Double
   ReDim dtm(1 To lArrayMax) As Date, dtm1(1 To lArrayMax) As Date, dtm2(1 To lArrayMax) As Date
   ReDim c(1 To lArrayMax) As Currency, c1(1 To lArrayMax) As Currency, c2(1 To lArrayMax) As Currency
   
   Dim vnt As Variant, oItem As ListItem
   
   Clearcontrols
   
   Randomize Timer
   
   SetupListview "Unsorted", "Sorted ASC", "Sorted DESC", "Median"
   
   For j = 1 To lArrayMax
   
      Select Case GetOption(optData())
      Case eDataType.dtByte
         b(j) = Int((255 - 0 + 1) * Rnd + 0)
         Set oItem = Me.ListView1.ListItems.Add(, , Str$(b(j)))
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
   
   ' For demonstration purpose, create 2 duplicates of the original array
   ' and sort 1 copy ascending and 1 copy descending.
   ' Typically you want the original array to be sorted, so this step isn't required.
   Select Case GetOption(optData())
   Case eDataType.dtByte
      b1() = b(): b2() = b()
      baSortByte b1()
      baSortByte b2(), True
   
   Case eDataType.dtInteger
      i1() = i(): i2() = i()
      baSortInteger i1()
      baSortInteger i2(), True
   
   Case eDataType.dtLong
      l1() = l(): l2() = l()
      baSortLong l1()
      baSortLong l2(), True
   
   Case eDataType.dtSingle
      s1() = s(): s2() = s()
      baSortSingle s1()
      baSortSingle s2(), True
   
   Case eDataType.dtDouble
      d1() = d(): d2() = d()
      baSortDouble d1()
      baSortDouble d2(), True
   
   Case eDataType.dtCurrency
      c1() = c(): c2() = c()
      baSortCurrency c1()
      baSortCurrency c2(), True
   
   Case eDataType.dtDate
   ' In VB6/VBA, a Date datatype is basically a
   ' Double, with the pre-decimal place representing the date part
   ' and the fraction the time part
      d1() = d(): d2() = d()
      baSortDouble d1()
      baSortDouble d2(), True
   End Select
   
   ' Compute the median
   Select Case GetOption(optData())
   Case eDataType.dtByte
      vnt = CByte(baMedianByte(b()))
   Case eDataType.dtInteger
      vnt = CInt(baMedianInteger(i()))
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
      
   ' Display the sorted array
   For j = 1 To lArrayMax
   
      Select Case GetOption(optData())
      Case eDataType.dtByte
         Me.ListView1.ListItems(j).ListSubItems.Add Text:=Str$(b1(j))
         Me.ListView1.ListItems(j).ListSubItems.Add Text:=Str$(b2(j))
      
      Case eDataType.dtInteger
         Me.ListView1.ListItems(j).ListSubItems.Add Text:=Str$(i1(j))
         Me.ListView1.ListItems(j).ListSubItems.Add Text:=Str$(i2(j))
      
      Case eDataType.dtLong
         Me.ListView1.ListItems(j).ListSubItems.Add Text:=Str$(l1(j))
         Me.ListView1.ListItems(j).ListSubItems.Add Text:=Str$(l2(j))
      
      Case eDataType.dtSingle
         Me.ListView1.ListItems(j).ListSubItems.Add Text:=Str$(s1(j))
         Me.ListView1.ListItems(j).ListSubItems.Add Text:=Str$(s2(j))
      
      Case eDataType.dtDouble
         Me.ListView1.ListItems(j).ListSubItems.Add Text:=Str$(d1(j))
         Me.ListView1.ListItems(j).ListSubItems.Add Text:=Str$(d2(j))
      
      Case eDataType.dtCurrency
         Me.ListView1.ListItems(j).ListSubItems.Add Text:=Str$(c1(j))
         Me.ListView1.ListItems(j).ListSubItems.Add Text:=Str$(c2(j))
      
      Case eDataType.dtDate
         Me.ListView1.ListItems(j).ListSubItems.Add Text:=CDate(d1(j))
         Me.ListView1.ListItems(j).ListSubItems.Add Text:=CDate(d2(j))
      End Select
   
   Next j

   ' Finally, show the computed median
   Me.ListView1.ListItems(1).ListSubItems.Add Text:=Str$(vnt)
   
   ListViewAdjustColumnWidth Me.ListView1, , True

End Sub

Private Sub cmdFormatNumber_Click()
   
   Dim cur As Currency, sResult As String
   
   Clearcontrols
   
   SetupListview "French/French-Swiss", "English/English-US", "German/German-Germany"
   
   cur = Val(InputBox("Number to format", "Format number", "1234.56"))
   
   Me.ListView1.ListItems.Add Text:=baFormatNumber(cur, LANG_FRENCH, SUBLANG_FRENCH_SWISS)
   Me.ListView1.ListItems(1).ListSubItems.Add Text:=baFormatNumber(cur, LANG_ENGLISH, SUBLANG_ENGLISH_US)
   Me.ListView1.ListItems(1).ListSubItems.Add Text:=baFormatNumber(cur, LANG_GERMAN, SUBLANG_GERMAN)
   
   Me.ListView1.ListItems.Add Text:=baFormatNumber(-cur, LANG_FRENCH, SUBLANG_FRENCH_SWISS)
   Me.ListView1.ListItems(2).ListSubItems.Add Text:=baFormatNumber(-cur, LANG_ENGLISH, SUBLANG_ENGLISH_US)
   Me.ListView1.ListItems(2).ListSubItems.Add Text:=baFormatNumber(-cur, LANG_GERMAN, SUBLANG_GERMAN)
   
   ListViewAdjustColumnWidth Me.ListView1, , True

End Sub

Private Sub cmdTypeOfVariant_Click()

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
   s = "Hello World!"
   Dim a(1 To 2) As Integer
   
   Clearcontrols
   'SetupListview "", "Byte", "Integer", "Long", "Single", "Double", "Currency", "Date", "String"
   SetupListview "Value", "Org. data type", "Variant type", ""
   
   Dim oItem As ListItem
   
   vnt = b
   Set oItem = Me.ListView1.ListItems.Add(Text:=CStr(b))
   Me.ListView1.ListItems(1).ListSubItems.Add Text:="Byte"
   Me.ListView1.ListItems(1).ListSubItems.Add Text:=CStr(TypeOfVariant(vnt))
   Me.ListView1.ListItems(1).ListSubItems.Add Text:=baTypeOfVariantToString(TypeOfVariant(vnt))
   
   vnt = i
   Set oItem = Me.ListView1.ListItems.Add(Text:=CStr(i))
   Me.ListView1.ListItems(2).ListSubItems.Add Text:="Integer"
   Me.ListView1.ListItems(2).ListSubItems.Add Text:=CStr(TypeOfVariant(vnt))
   Me.ListView1.ListItems(2).ListSubItems.Add Text:=baTypeOfVariantToString(TypeOfVariant(vnt))
   
   vnt = l
   Set oItem = Me.ListView1.ListItems.Add(Text:=CStr(l))
   Me.ListView1.ListItems(3).ListSubItems.Add Text:="Long"
   Me.ListView1.ListItems(3).ListSubItems.Add Text:=CStr(TypeOfVariant(vnt))
   Me.ListView1.ListItems(3).ListSubItems.Add Text:=baTypeOfVariantToString(TypeOfVariant(vnt))
   
   vnt = f
   Set oItem = Me.ListView1.ListItems.Add(Text:=CStr(f))
   Me.ListView1.ListItems(4).ListSubItems.Add Text:="Single"
   Me.ListView1.ListItems(4).ListSubItems.Add Text:=CStr(TypeOfVariant(vnt))
   Me.ListView1.ListItems(4).ListSubItems.Add Text:=baTypeOfVariantToString(TypeOfVariant(vnt))
   
   vnt = d
   Set oItem = Me.ListView1.ListItems.Add(Text:=CStr(d))
   Me.ListView1.ListItems(5).ListSubItems.Add Text:="Double"
   Me.ListView1.ListItems(5).ListSubItems.Add Text:=CStr(TypeOfVariant(vnt))
   Me.ListView1.ListItems(5).ListSubItems.Add Text:=baTypeOfVariantToString(TypeOfVariant(vnt))
   
   vnt = c
   Set oItem = Me.ListView1.ListItems.Add(Text:=CStr(c))
   Me.ListView1.ListItems(6).ListSubItems.Add Text:="Currency"
   Me.ListView1.ListItems(6).ListSubItems.Add Text:=CStr(TypeOfVariant(vnt))
   Me.ListView1.ListItems(6).ListSubItems.Add Text:=baTypeOfVariantToString(TypeOfVariant(vnt))
   
   vnt = dtm
   Set oItem = Me.ListView1.ListItems.Add(Text:=CStr(dtm))
   Me.ListView1.ListItems(7).ListSubItems.Add Text:="Date"
   Me.ListView1.ListItems(7).ListSubItems.Add Text:=CStr(TypeOfVariant(vnt))
   Me.ListView1.ListItems(7).ListSubItems.Add Text:=baTypeOfVariantToString(TypeOfVariant(vnt))
   
   vnt = vnt1
   Set oItem = Me.ListView1.ListItems.Add(Text:=CStr(vnt))
   Me.ListView1.ListItems(8).ListSubItems.Add Text:="Variant"
   Me.ListView1.ListItems(8).ListSubItems.Add Text:=CStr(TypeOfVariant(vnt))
   Me.ListView1.ListItems(8).ListSubItems.Add Text:=baTypeOfVariantToString(TypeOfVariant(vnt))
   
   vnt = s
   Set oItem = Me.ListView1.ListItems.Add(Text:=s)
   Me.ListView1.ListItems(9).ListSubItems.Add Text:="String"
   Me.ListView1.ListItems(9).ListSubItems.Add Text:=CStr(TypeOfVariant(vnt))
   Me.ListView1.ListItems(9).ListSubItems.Add Text:=baTypeOfVariantToString(TypeOfVariant(vnt))
   
   vnt = a
   Set oItem = Me.ListView1.ListItems.Add(Text:=CStr(a(1)) & ", " & CStr(a(2)))
   Me.ListView1.ListItems(10).ListSubItems.Add Text:="Array"
   Dim e As eVariantType
   e = TypeOfVariant(vnt)
   Me.ListView1.ListItems(10).ListSubItems.Add Text:=CStr(e)
   Me.ListView1.ListItems(10).ListSubItems.Add Text:=baTypeOfVariantToString(e)

   ListViewAdjustColumnWidth Me.ListView1, , True

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
      .ListView1.ListItems.Clear
      .ListView1.ColumnHeaders.Clear
   End With

End Sub

Sub SetupListview(ParamArray vntColHdrs() As Variant)

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
   SetupListview ""
   
   
   ReDim acctnum(1 To 5) As Byte
   ReDim users$(1 To 5)

   users$(1) = "zz"
   users$(2) = "aa"
   users$(3) = "yy"
   users$(4) = "bb"
   users$(5) = "xx"

   acctnum(1) = 78
   acctnum(2) = 98
   acctnum(3) = 45
   acctnum(4) = 32
   acctnum(5) = 1
   
   Debug.Print "baArrayStringSet: " & baArrayStringSet(users(), 2)
   Debug.Print "baArrayByteSet: " & baArrayByteSet(acctnum(), 1)
   Debug.Print String$(3, "-")
   
   Dim i As Long
   For i = 1 To 5
      Debug.Print i; " AcctNum: "; acctnum(i); " Users: "; users(i)
   Next i
   Debug.Print String$(3, "-")
   
   Debug.Print "baSort2Arrays: " & CStr(baSort2Arrays(eArrayDataType.adtByte, eArrayDataType.adtString))
   Debug.Print String$(3, "-")
   
   For i = 1 To 5
      Debug.Print i; " AcctNum: "; acctnum(i); " Users: "; users(i)
   Next i
   Debug.Print String$(3, "-")
   
   
End Sub
