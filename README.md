<div align="center">

## Determine IDE/Debugging Status


</div>

### Description

This function will return whether you are running your program or DLL from within the IDE, or compiled. I use it as part of my DLL's like active document DLL's to setup information that would normally be supplied from the outside.
 
### More Info
 
Returns True if you are running inside the VB 5.0 or 6.0 IDE.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[L\. F\. Carpenter](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/l-f-carpenter.md)
**Level**          |Intermediate
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/l-f-carpenter-determine-ide-debugging-status__1-11615/archive/master.zip)

### API Declarations

```
Private Declare Function GetModuleFileName Lib "kernel32" _
 Alias "GetModuleFileNameA" _
 ( _
  ByVal hModule As Long, _
  ByVal lpFileName As String, _
  ByVal nSize As Long _
 ) As Long
```


### Source Code

```

Public Function InVBDesignEnvironment() As Boolean
 Dim strFileName As String
 Dim lngCount As Long
 strFileName = String(255, 0)
 lngCount = GetModuleFileName(App.hInstance, strFileName, 255)
 strFileName = Left(strFileName, lngCount)
 InVBDesignEnvironment = False
 If UCase(Right(strFileName, 7)) = "VB5.EXE" Then
  InVBDesignEnvironment = True
 ElseIf UCase(Right(strFileName, 7)) = "VB6.EXE" Then
  InVBDesignEnvironment = True
 End If
End Function
```

