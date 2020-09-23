<div align="center">

## Yet Another Small Word Wrap Function


</div>

### Description

Due to the astonishing popularity of someone else's recent submission and me feeling bored...
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[jeremyxtz](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jeremyxtz.md)
**Level**          |Beginner
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jeremyxtz-yet-another-small-word-wrap-function__1-62303/archive/master.zip)





### Source Code

```
'note if a word is longer than maxchars the word won't be split
Private Sub Form_Load()
MsgBox WrappedString("fish goes to market in the happy place and stays away in time so it does", 17)
MsgBox WrappedString("fish goes to market in the happy place and stays away in time so it does", 17)
End Sub
Function WrappedString(st As String, maxchars As Integer) As String
Dim c As Integer, i As Integer, lastc As Integer, lastspace As Integer
For i = 1 To Len(st)
c = c + 1
If Mid(st, i, 1) = " " Then lastspace = i: lastc = c
If c > maxchars And lastc <> 0 Then Mid(st, lastspace, 1) = Chr(10): c = maxchars - lastc + 1
Next
WrappedString = Replace(st, Chr(10), vbCrLf)
End Function
'same function but using a byte array. One more line but its the
'efficient way to handle strings
Function WrappedStringB(st As String, maxchars As Integer) As String
Dim stb() As Byte, c As Integer, i As Integer, lastspace As Integer, lastc As Integer
stb = st
For i = 0 To UBound(stb) Step 2
c = c + 1
If (stb(i)) = 32 Then lastspace = i: lastc = c
If c > maxchars And lastc <> 0 Then stb(lastspace) = 10: c = maxchars - lastc + 1
Next
WrappedStringB = Replace(stb, Chr(10), vbCrLf)
End Function
```

