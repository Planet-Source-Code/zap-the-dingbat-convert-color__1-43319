<div align="center">

## Convert Color


</div>

### Description

Convert color codes

RGB to VB

RGB to HTML

VB to RGB

VB to HTML

HTML to RGB

HTML to VB
 
### More Info
 
colConvertType: type of convertion

strColor: string containing color value. if rgb then spesify r,g,b as 3 values

string containing new color value.

If RGB then returns 1 string delimited by ","


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Zap The Dingbat](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/zap-the-dingbat.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0, VBA MS Access, VBA MS Excel
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/zap-the-dingbat-convert-color__1-43319/archive/master.zip)





### Source Code

```
Public Enum colConvertType
  colRGBtoVB = 0
  colRGBtoHTML = 1
  colVBtoRGB = 2
  colVBtoHTML = 3
  colHTMLtoRGB = 4
  colHTMLtoVB = 5
End Enum
Function convertColor(convert As colConvertType, ParamArray strColor()) As String
  Select Case convert
    Case colRGBtoHTML
      convertColor = Left("00" & Hex(strColor(0)), 2)
      convertColor = Left("00" & convertColor & Hex(strColor(1)), 2)
      convertColor = Left("00" & convertColor & Hex(strColor(2)), 2)
    Case colRGBtoVB
      convertColor = RGB(strColor(0), strColor(1), strColor(2))
    Case colVBtoRGB
      convertColor = Right("000000" & Hex(strColor(0)), 6)
      r = CByte("&h" & Mid(convertColor, 5, 2))
      g = CByte("&h" & Mid(convertColor, 3, 2))
      b = CByte("&h" & Mid(convertColor, 1, 2))
      convertColor = r & "," & g & "," & b
    Case colVBtoHTML
      convertColor = Right("000000" & Hex(strColor(0)), 6)
      r = Mid(convertColor, 5, 2)
      g = Mid(convertColor, 3, 2)
      b = Mid(convertColor, 1, 2)
      convertColor = r & g & b
    Case colHTMLtoRGB
      convertColor = Right("000000" & Hex(strColor(0)), 6)
      r = CByte("&h" & Mid(convertColor, 1, 2))
      g = CByte("&h" & Mid(convertColor, 3, 2))
      b = CByte("&h" & Mid(convertColor, 5, 2))
      convertColor = r & "," & g & "," & b
    Case colHTMLtoVB
      convertColor = Right("000000" & Hex(strColor(0)), 6)
      r = CByte("&h" & Mid(convertColor, 1, 2))
      g = CByte("&h" & Mid(convertColor, 3, 2))
      b = CByte("&h" & Mid(convertColor, 5, 2))
      convertColor = RGB(r, g, b)
  End Select
End Function
```

