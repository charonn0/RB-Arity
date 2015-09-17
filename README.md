# RB-Arity
A **proof-of-concept** dynamic delegate invoker for Realbasic and Xojo. 

##Features
* Invoke external functions using library and function names that are decided upon at runtime instead of compile-time.
* Invoke external functions with a variable number of arguments.

##Limitations and caveats
* Windows only. (Other targets probably doable, but not implemented.)
* External functions may receive 0-6 arguments. Passing more than 6 arguments will raise an exception.
* Argument and return values must be convertable to Ptrs 
* No typechecking
