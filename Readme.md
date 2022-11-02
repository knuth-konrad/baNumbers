# baNumbers - Arrays / Numbers helper

baNumbers is a _(mainly targeted at VB6/VBA)_ array and numbers helper library.

The DLL is written in PowerBASIC _(PBWIN 6.04)_, the source of the DLL is the file [baNumbers.bas](.\baNumbers.bas), located in the root of the repository.

The folder [VB](.\VB) contains a VB6 demonstration project. The VB prototypes _(aka 'Declarations')_ are located in the file [VB\BAS\baNumbers.bas](.\VB\BAS\baNumbers.bas). You'll find some other VB helper functions in this folder, too.

---

## Arrays

### Sorting

Each of the following methods takes an array of the respective data type as input and sorts the array data in ascending order. Please note that the array is passed ```ByRef```, i.e. the original array will be altered.

```vb
Sub baSortByte(a() As Byte)
Sub baSortCurrency(a() As Currency)
Sub baSortDouble(a() As Double)
Sub baSortInteger(a() As Integer)
Sub baSortLong(a() As Long)
Sub baSortSingle(a() As Single)
```

---

### Median

Each of the following methods takes an array of the respective data type as input and computes and returns the median of the array contents.

```vb
Function baMedianByte(a() As Byte) As Byte
Function baMedianCurrency(a() As Currency) As Currency
Function baMedianDouble(a() As Double) As Double
Function baMedianInteger(a() As Integer) As Integer
Function baMedianLong(a() As Long) As Long
Function baMedianSingle(a() As Single) As Single
```

---

### Set all elements to (value)

Each of the following methods takes an array of the respective data type as input and sets all array elements to the provided value. _Please note:_ the ```Currency``` data type is not supported.

```vb
Function baSetByte(a() As Byte, ByVal value As Byte) As Boolean
Function baSetInteger(a() As Integer, ByVal value As Integer) As Boolean
Function baSetLong(a() As Long, ByVal value As Long) As Boolean
Function baSetSingle (a() As Single, ByVal value As Single) As Boolean
Function baSetDouble (a() As Double, ByVal value As Double) As Boolean
```

---

### Fill a Currency array with random numbers

This method fills a ```Currency``` array with random numbers within 0 &lt;= x &lt;= 1.

```vb
Sub baRndArray (a() As Currency)
```

### Fill a Long array with random numbers

This method fills a ```Long``` array with random numbers within lLower &lt;= x &lt;= lUpper.

```vb
Sub baRndRangeArray (a() As Long, ByVal lLower As Long, ByVal lUpper As Long)
```

---

### Sort two arrays

Sorts/reorders 2 arrays. The first array _a1()_ will be sorted. The second array _a2(_) however will have its array elements arranged as if they _"stick"_ to the first array, i.e. when _a1()_ is sorted and element _a1(1)_ has become element _a1(5)_ thereafter, _a2(1)_ now also will be _a2(5)_. The arrays may be of different data types.

_Please note_: the second array **must** have _at least_ the same number of elements than the first array.

#### Parameters

- eDataType1, eDataType2  
Data type of the respective array

- bolDescending  
If ```True```, sort array 1 descending order.

```vb
Enum eArrayDataType
   adtByte = 0
   adtCurrency = 1
   adtDouble = 2
   adtInteger = 3
   adtLong = 4
   adtSingle = 5
   adtString = 6
End Enum
```

```vb
Function baSort2Arrays Lib "baNumbers.dll" (ByVal eDataType1 As eArrayDataType, eDataType2 As eArrayDataType, Optional ByVal bolDescending As Boolean = False) As Long
```

The function returns the following error codes:

- 0 = Success
- 1 = Invalid data type passed in ```eDataType```

---

### Search arrays

Each of the following methods takes an array of the respective data type as input and sorts the array data in ascending order. Please note that the array is passed ```ByRef```, i.e. the original array will be altered.

```vb
Sub baSortByte(a() As Byte)
Sub baSortCurrency(a() As Currency)
Sub baSortDouble(a() As Double)
Sub baSortInteger(a() As Integer)
Sub baSortLong(a() As Long)
Sub baSortSingle(a() As Single)
```

---

## Numbers

### Fraction

Each of the following methods returns the fractional part of the floating point number.

```vb
Function baFracCur(ByVal curValue As Currency) As Currency
Function baFracDouble(ByVal dblValue As Double) As Double
Function baFracSingle(ByVal fValue As Single) As Single
```

---

### Swapping

Each of the following methods swaps the value of two variables.

```vb
Function baSwapByte(ByRef v1 As Byte, ByRef v2 As Byte) As Boolean
Function baSwapCurrency(ByRef v1 As Currency, ByRef v2 As Currency) As Boolean
Function baSwapDouble(ByRef v1 As Double, ByRef v2 As Double) As Boolean
Function baSwapInteger(ByRef v1 As Integer, ByRef v2 As Integer) As Boolean
Function baSwapLong(ByRef v1 As Long, ByRef v2 As Long) As Boolean
Function baSwapSingle(ByRef v1 As Single, ByRef v2 As Single) As Boolean
```

---

### Signed to unsigned integer

The following methods return the unsigned value of a (negative) signed integer value, i.e. ```Integer``` to ```Word``` etc.
_Please note:_ due to VB's data type limitation, some values are returned as ```Currency``` to avoid overflow errors.

```vb
Function Int2Wrd(ByVal iValue As Integer) As Long
Function Int2DWrd(ByVal iValue As Integer) As Currency
Function Lng2DWrd(ByVal lValue As Long) As Currency
Function Lng2Quad(ByVal lValue As Long) As Currency
```

---

### (Locale) Formatting

Returns the string representation of number, formatted according to the _locale_ settings.

```vb
Function baFormatNumber (ByVal curNumber As Currency, _
   ByVal wLangLocale As Long, ByVal wSubLangLocale As Long) As String
```

Returns the string representation of a number, formatted according to _LANGID_.

```vb
Function baFormatNumberEx(ByVal curNumber As Currency, _
   ByVal dwLangID As Long) As String
```

---

## Variant

### Variant subtype

Determines the specific ```Variant``` subtype, e.g.

```vb
Dim l As Long, v As Variant
v = l
' Prints "3" = eVariantType.vtLongIntegerSigned
Debug.Print TypeOfVariant(v)
```

Please note that arrays will report as ```vtArray``` of ```vt(DataType)```. e.g.

```vb
Dim i(1 To 2) As Integer, v As Variant
v = i
' Prints "8194" = eVariantType.vtArray Or eVariantType.vtIntegerSigned
Debug.Print TypeOfVariant(v)
```

```vb
Function TypeOfVariant (ByVal vnt As Variant) As eVariantType
```

### Variant subtype constant to string

Return the string representation of the variant subtype

```vb
Function baTypeOfVariantToString(ByVal eValue As eVariantType) As String
```
