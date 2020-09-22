<div align="center">

## Format Bytes


</div>

### Description

Format bytes into 3 digit kb, mg, gb, etc...

Will appear as #.## $$, ##.# $$, or ### $$
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ben White](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ben-white.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Libraries](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/libraries__4-35.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ben-white-format-bytes__4-7686/archive/master.zip)





### Source Code

```
' General Logic
' If bytes < 1000 Then ### Bytes
' ElseIf bytes < (2^10 * 10) Then #.## KB
' ElseIf bytes < (2^10 * 100) Then ##.# KB
' ElseIf bytes < (2^10 * 1000) Then ### KB
' ElseIf bytes < (2^20 * 10) Then #.## MB
' ElseIf bytes < (2^20 * 100) Then ##.# MB
' ElseIf bytes < (2^20 * 1000) Then ### MB
' etc...
Function FormatBytes(ByVal Bytes)
	Dim strType, x, y, z
	strType = Array("KB","MB","GB","TB","PB","EB","ZB","YB")
	If Bytes < 0 Then Exit Function
	If Bytes < 1000 Then
		FormatBytes = Bytes & " Bytes"
		Exit Function
	End If
	For x = 1 to 8
		z = 2^(x*10)
		For y = 1 to 3
			If Bytes < z*(10^y) Then
				FormatBytes = FormatNumber(Bytes/z, 3-y) & " " & strType(x-1)
				Exit Function
			End If
		Next
	Next
End Function
```

