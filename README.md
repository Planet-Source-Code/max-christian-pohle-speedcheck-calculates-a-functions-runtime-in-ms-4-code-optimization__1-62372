<div align="center">

## speedcheck \- calculates a functions runtime in ms \(4 code\-optimization\)


</div>

### Description

Just one line of code is needed to test a functions speed -&gt; With this class-module.

It creates an object which is automatically terminated together with the function you check and it uses debug.print to let you know how long your pc had been busy (in ms) with that function.

Sure you can save the value of timer() and read it before your function reaches exit/end function, but its harder to remove this before you release your app... So try this!
 
### More Info
 
Dim A as New &lt;Object&gt; is slower than Dim A as &lt;Object&gt;: Set A = New &lt;Object&gt;


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Max Christian Pohle](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/max-christian-pohle.md)
**Level**          |Advanced
**User Rating**    |5.0 (25 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/max-christian-pohle-speedcheck-calculates-a-functions-runtime-in-ms-4-code-optimization__1-62372/archive/master.zip)

### API Declarations

```
'Create a new Class-Module for this
'and call it "SpeedCheck":
Option Explicit
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private StartPoint As Long
Public Property Get CurRunTime() As Long
  RunTime = GetTickCount - StartPoint
End Property
Private Sub Class_Initialize()
  StartPoint = GetTickCount
End Sub
Private Sub Class_Terminate()
  Debug.Print "- - - - - Your function needed: " &amp; GetTickCount - StartPoint &amp; " ms"
End Sub
```


### Source Code

```
'This is just an example to use it...
'Its as easy as possible-
'therefor I do not use option explicit here
'and I work with a slow variant-datatype :-)
Private Sub Form_Load()
  Dim A As SpeedCheck: Set A = New SpeedCheck
  For I = 0 To 1000
    Debug.Print "Debug-Print is very slow!"
  Next I
End Sub
```

