<div align="center">

## String Compare Logical Demo


</div>

### Description

The code is used to demonstrate the usage of API function StrCmpLogicalW (Shlwapi.dll). It will help you to sort a string array by comparing them logically. The result will sort the result logically, eg, "music9.dat" is lower than "music10.dat".
 
### More Info
 
For StrQSortLogical: the input is a string array=strArray() As String, lower index=lowerB As Long, upper index=upperB As Long)

For StrSwap: Both are string variables

For StrCmpLogical: Both are string variable

User may need to provide a list of string to be sorted before using the function.

StrQSortLogical: no return value, but the function modify the string array directly

StrSwap: no return value, but both strings are swapped.

StrCmpLogical return 0 if both strings are equal, -1 if string 1 lower than string 2, 1 if string 1 &gt; string 2

The code is written for VB6 only because it make used of undeocumented function VarPtr and StrPtr function, which is not supported under .NET.

The function may malfunction if empty array is passed into it.

Be wary in using the API declaration directly, because the function StrCmpLogicalW run by assuming its input as NULL-terminated-UNICODE characters. If you passed VB6 strings (which is UNICODE by nature), VB6 SMARTly convert them into ANSI code, thus the result might be wrong and unpredictable. So, you may used the wrapper function instead.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Nordin Rahman](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/nordin-rahman.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/nordin-rahman-string-compare-logical-demo__1-61858/archive/master.zip)

### API Declarations

```
' to compare two string logically
' return:
'  0 if same
'  1 if string 1 is larger than string 2
' -1 if string 1 is smaller than string 2
' use string pointer to both strings as input
' for example: string1="A1.txt", string2="A10.txt"
' StrCmpLogicalP(ByVal StrPtr(string1), ByVal StrPtr(string2))
' return -1, because "A1.txt" is smaller than "A10.txt" logically
Public Declare Function StrCmpLogicalP Lib "Shlwapi.dll" Alias "StrCmpLogicalW" ( _
  ByVal ptr1 As Long, _
  ByVal ptr2 As Long _
) As Long
' use in copying variable operation
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
  Destination As Any, _
  Source As Any, _
  ByVal Length As Long _
)
```


### Source Code

```
' wrapper to StrCmpLogicalP
' given input two string
Public Function StrCmpLogical(str1 As String, str2 As String) As Long
  StrCmpLogical = StrCmpLogicalP(ByVal StrPtr(str1), ByVal StrPtr(str2))
End Function
' wrapper to two string swap
Public Sub StrSwap(str1 As String, str2 As String)
  Dim ptr As Long
  CopyMemory ptr, ByVal VarPtr(str1), 4
  CopyMemory ByVal VarPtr(str1), ByVal VarPtr(str2), 4
  CopyMemory ByVal VarPtr(str2), ptr, 4
End Sub
' to sort logically an array of string, strArray,
' starting from lower index, lowerB,
' end at upper index, upperB
Public Sub StrQSortLogical(strArray() As String, lowerB As Long, upperB As Long)
 Dim i As Long
 Dim j As Long
 Dim X As Long
 Dim Y As Long
 i = lowerB
 j = upperB
 X = StrPtr(strArray((lowerB + upperB) / 2))
 Do While (i <= j)
  Do While (StrCmpLogicalP(ByVal StrPtr(strArray(i)), ByVal X) < 0 And i < upperB)
   i = i + 1
  Loop
  Do While (StrCmpLogicalP(ByVal X, ByVal StrPtr(strArray(j))) < 0 And j > lowerB)
   j = j - 1
  Loop
  ' The Actual swapping is here
  If (i <= j) Then
   CopyMemory Y, ByVal VarPtr(strArray(i)), 4
   CopyMemory ByVal VarPtr(strArray(i)), ByVal VarPtr(strArray(j)), 4
   CopyMemory ByVal VarPtr(strArray(j)), Y, 4
   i = i + 1
   j = j - 1
  End If
 Loop
 If (lowerB < j) Then Call StrQSortLogical(strArray, lowerB, j)
 If (i < upperB) Then Call StrQSortLogical(strArray, i, upperB)
End Sub
```

