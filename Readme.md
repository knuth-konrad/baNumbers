# baNumbers - Arrays / Numbers helper

baNumbers is a _(mainly targeted at VB6/VBA)_ array and numbers helper library.

---

## Arrays

### Sorting

Each of the following methods takes an array of the respective data type as input and sorts the array data in ascending order. Please note that the array is passed ```ByRef```, i.e. the original array will be altered.

```vb
Sub baSortCurrency(a() As Currency)
Sub baSortDouble(a() As Double)
Sub baSortInteger(a() As Integer)
Sub baSortLong(a() As Long)
Sub baSortSingle(a() As Single)
```

### Duplicating

Each of the following methods takes two arrays of the respective data type as input and duplicates the data of array _a()_ to array _b()_. Please note that _both_ arrays are passed ```ByRef```, i.e. data of array _b()_ will be overwritten with data from array _a()_.

```vb
Function baDuplicateCurrencyArray(a() As Currency, b() As Currency) As Boolean
Function baDuplicateDoubleArray(a() As Double, b() As Double) As Boolean
Function baDuplicateIntegerArray(a() As Integer, b() As Integer) As Boolean
Function baDuplicateLongArray(a() As Long, b() As Long) As Boolean
Function baDuplicateSingleArray(a() As Single, b() As Single) As Boolean
```

### Median

Each of the following methods takes an array of the respective data type as input and computes and returns the median of the array contents.

```vb
Function baMedianCurrency(a() As Currency) As Currency
Function baMedianDouble(a() As Double) As Double
Function baMedianInteger(a() As Integer) As Integer
Function baMedianLong(a() As Long) As Long
Function baMedianSingle(a() As Single) As Single
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

### Swapping

Each of the following methods swaps the value of two variables.

```vb
Function baSwapCurrency(ByRef v1 As Currency, ByRef v2 As Currency) As Boolean
Function baSwapDouble(ByRef v1 As Double, ByRef v2 As Double) As Boolean
Function baSwapInteger(ByRef v1 As Integer, ByRef v2 As Integer) As Boolean
Function baSwapLong(ByRef v1 As Long, ByRef v2 As Long) As Boolean
Function baSwapSingle(ByRef v1 As Single, ByRef v2 As Single) As Boolean
```

### (Locale) Formatting

Returns the string representation of number, formatted according to the _locale_ settings.

```vb
Function baFormatNumber (ByVal curNumber As Currency, _
   ByVal wLangLocale As Long, ByVal wSubLangLocale As Long) As String
```

Returns the string representation of number, formatted according to _LANGID_.

```vb
Function baFormatNumberEx(ByVal curNumber As Currency, _
   ByVal dwLangID As Long) As String
```
