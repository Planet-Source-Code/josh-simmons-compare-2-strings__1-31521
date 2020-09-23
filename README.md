<div align="center">

## Compare 2 Strings


</div>

### Description

This VERY short function compares 2 strings and returns a number (that can be converted into percentage if multiplied by 100) that represents how closely related 2 strings are. For instance "ABCDE" and "ABCDF" would return say.... .8 (80%). Great for suggesting fixes for spelling errors et cetera. Feel free to use, abuse, and manipulate the code however you want, I'm sure it's not original, but I know it can be helpful :).
 
### More Info
 
CompareTXT(String1 as String, String2 as String)

String1 and String2 are the only paramaters you need to send to the function, those are the two strings you wish to compare.

% value that represents exactly how similar two strings are.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Josh Simmons](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/josh-simmons.md)
**Level**          |Intermediate
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/josh-simmons-compare-2-strings__1-31521/archive/master.zip)





### Source Code

```
Public Function CompareTXT(String1 As String, String2 As String) As Single
Dim i, y, x As Integer
Dim a, b As String
String1 = UCase(String1)  'take this out if you
String2 = UCase(String2)  'want it to be case
              'sensitive
If String1 = String2 Then CompareTXT = 1: Exit Function
              'if the strings are
              'the same, don't
              'bother to waste time
              'and space on working
              'them out :).
If Len(String1) > Len(String2) Then x = Len(String1)
If Len(String2) > Len(String1) Then x = Len(String2)
If Len(String1) = Len(String2) Then x = Len(String1)
              'find out the length
              'of the longest string
For i = 1 To x
  a = Mid(String1, i, 1) 'get 1 character from
  b = Mid(String2, i, 1) 'each string and compare
  If a = b Then y = y + 1 'the characters
Next
CompareTXT = y / x
End Function
```

