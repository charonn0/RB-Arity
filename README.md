# RB-Arity
A **proof-of-concept** dynamic delegate invoker for Realbasic and Xojo. 

```Arity(ærɨti): In logic, mathematics, and computer science, the arity of a function or operation is the number of arguments or operands the function or operation accepts. ```

## Example
This example changes the titlebar caption on a window:
```vbnet
  Dim SetWindowText As Arity.Invoker = Arity.GetProcAddress("User32", "SetWindowTextA")
  Dim caption As CString = "Caption set!"
  SetWindowText.Invoke(MyWindow.Handle, caption)
```

## Features
* Invoke external functions using library and function names that are decided upon at runtime instead of compile-time.
* Invoke external functions with a variable number of arguments.

## Limitations and caveats
* Windows only. (Other targets probably doable, but not implemented.)
* External functions may receive 0-6 arguments. Passing more than 6 arguments will raise an exception.
* Argument and return values must be convertable to Ptrs 
* No typechecking
