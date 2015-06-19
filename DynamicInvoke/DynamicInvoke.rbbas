#tag Module
Protected Module DynamicInvoke
	#tag DelegateDeclaration, Flags = &h21
		Private Delegate Function ArityF0() As Ptr
	#tag EndDelegateDeclaration

	#tag DelegateDeclaration, Flags = &h21
		Private Delegate Function ArityF1(Argument1 As Ptr) As Ptr
	#tag EndDelegateDeclaration

	#tag DelegateDeclaration, Flags = &h21
		Private Delegate Function ArityF2(Argument1 As Ptr, Argument2 As Ptr) As Ptr
	#tag EndDelegateDeclaration

	#tag DelegateDeclaration, Flags = &h21
		Private Delegate Function ArityF3(Argument1 As Ptr, Argument2 As Ptr, Argument3 As Ptr) As Ptr
	#tag EndDelegateDeclaration

	#tag DelegateDeclaration, Flags = &h21
		Private Delegate Function ArityF4(Argument1 As Ptr, Argument2 As Ptr, Argument3 As Ptr, Argument4 As Ptr) As Ptr
	#tag EndDelegateDeclaration

	#tag DelegateDeclaration, Flags = &h21
		Private Delegate Function ArityF5(Argument1 As Ptr, Argument2 As Ptr, Argument3 As Ptr, Argument4 As Ptr, Argument5 As Ptr) As Ptr
	#tag EndDelegateDeclaration

	#tag DelegateDeclaration, Flags = &h21
		Private Delegate Function ArityF6(Argument1 As Ptr, Argument2 As Ptr, Argument3 As Ptr, Argument4 As Ptr, Argument5 As Ptr, Argument6 As Ptr) As Ptr
	#tag EndDelegateDeclaration

	#tag DelegateDeclaration, Flags = &h21
		Private Delegate Sub ArityS0()
	#tag EndDelegateDeclaration

	#tag DelegateDeclaration, Flags = &h21
		Private Delegate Sub ArityS1(Argument1 As Ptr)
	#tag EndDelegateDeclaration

	#tag DelegateDeclaration, Flags = &h21
		Private Delegate Sub ArityS2(Argument1 As Ptr, Argument2 As Ptr)
	#tag EndDelegateDeclaration

	#tag DelegateDeclaration, Flags = &h21
		Private Delegate Sub ArityS3(Argument1 As Ptr, Argument2 As Ptr, Argument3 As Ptr)
	#tag EndDelegateDeclaration

	#tag DelegateDeclaration, Flags = &h21
		Private Delegate Sub ArityS4(Argument1 As Ptr, Argument2 As Ptr, Argument3 As Ptr, Argument4 As Ptr)
	#tag EndDelegateDeclaration

	#tag DelegateDeclaration, Flags = &h21
		Private Delegate Sub ArityS5(Argument1 As Ptr, Argument2 As Ptr, Argument3 As Ptr, Argument4 As Ptr, Argument5 As Ptr)
	#tag EndDelegateDeclaration

	#tag DelegateDeclaration, Flags = &h21
		Private Delegate Sub ArityS6(Argument1 As Ptr, Argument2 As Ptr, Argument3 As Ptr, Argument4 As Ptr, Argument5 As Ptr, Argument6 As Ptr)
	#tag EndDelegateDeclaration

	#tag Method, Flags = &h1
		Protected Function GetProcAddress(LibName As String, ProcName As String) As DynamicInvoke.Invoker
		  Dim libmb, procmb As MemoryBlock
		  libmb = New MemoryBlock(LibName.Len + 1)
		  libmb.CString(0) = LibName
		  procmb = New MemoryBlock(ProcName.Len + 1)
		  procmb.CString(0) = ProcName
		  Dim hmod As Integer = LoadLibraryA(libmb)
		  If hmod = 0 Then Return Nil
		  Dim proc As Ptr = _GetProcAddress(hmod, procmb)
		  If proc <> Nil Then Return New DynamicInvoke.Invoker(proc)
		End Function
	#tag EndMethod

	#tag ExternalMethod, Flags = &h21
		Private Declare Function LoadLibraryA Lib "Kernel32" (LibName As Ptr) As Integer
	#tag EndExternalMethod

	#tag Method, Flags = &h21
		Private Function MarshalToPtr(DataValue As Variant) As Ptr
		  Dim ValueType As Integer = VarType(DataValue)
		  Select Case ValueType
		    
		  Case Variant.TypeNil ' Sometimes Nil is an error; sometimes not
		    Return Nil
		    
		  Case Variant.TypeBoolean
		    If DataValue.BooleanValue Then
		      Return MarshalToPtr(1)
		    Else
		      Return MarshalToPtr(0)
		    End If
		    
		  Case Variant.TypePtr, Variant.TypeInteger
		    Return DataValue.PtrValue
		    
		  Case Variant.TypeWString, Variant.TypeCString
		    Return DataValue.PtrValue
		    
		  Case Variant.TypeObject
		    Select Case DataValue
		    Case IsA MemoryBlock
		      Return DataValue.PtrValue
		    Else
		      Raise New UnsupportedFormatException
		    End Select
		  Else
		    Raise New UnsupportedFormatException
		  End Select
		End Function
	#tag EndMethod

	#tag ExternalMethod, Flags = &h21
		Private Declare Function _GetProcAddress Lib "Kernel32" Alias "GetProcAddress" (hModule As Integer, ProcName As Ptr) As Ptr
	#tag EndExternalMethod


	#tag ViewBehavior
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			InheritedFrom="Object"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			InheritedFrom="Object"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			InheritedFrom="Object"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			InheritedFrom="Object"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			InheritedFrom="Object"
		#tag EndViewProperty
	#tag EndViewBehavior
End Module
#tag EndModule
