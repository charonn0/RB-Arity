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

	#tag ExternalMethod, Flags = &h21
		Private Soft Declare Function FreeLibrary Lib "Kernel32" (hModule As Integer) As Boolean
	#tag EndExternalMethod

	#tag ExternalMethod, Flags = &h21
		Private Soft Declare Function GetModuleFileNameW Lib "Kernel32" (hModule As Integer, Filename As Ptr, Size As Integer) As Integer
	#tag EndExternalMethod

	#tag ExternalMethod, Flags = &h21
		Private Soft Declare Function GetModuleHandleW Lib "Kernel32" (LibName As WString) As Integer
	#tag EndExternalMethod

	#tag Method, Flags = &h1
		Protected Function GetProcAddress(LibName As String, Ordinal As Integer) As DynamicInvoke.Invoker
		  #If Not TargetWin32 Then
		    Raise New PlatformNotSupportedException
		  #endif
		  If Procedures = Nil Then Procedures = New Dictionary
		  Dim Procedure As DynamicInvoke.Invoker = Procedures.Lookup(LibName + Str(Ordinal), Nil)
		  If Procedure = Nil Then
		    Dim hModule As DynamicInvoke.Library = Library.LoadLibrary(LibName)
		    If hModule = Nil Then Return Nil
		    Dim proc As Ptr = GetProcAddress_(hModule.ModuleID, Ptr(Ordinal))
		    If proc <> Nil Then
		      Procedure = New DynamicInvoke.Invoker(proc, Str(Ordinal), hModule)
		      Procedures.Value(LibName + Str(Ordinal)) = Procedure
		    End If
		  End If
		  Return Procedure
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetProcAddress(LibName As String, ProcName As String) As DynamicInvoke.Invoker
		  #If Not TargetWin32 Then
		    Raise New PlatformNotSupportedException
		  #endif
		  If Procedures = Nil Then Procedures = New Dictionary
		  Dim Procedure As DynamicInvoke.Invoker = Procedures.Lookup(LibName + ProcName, Nil)
		  If Procedure = Nil Then
		    Dim hModule As DynamicInvoke.Library = Library.LoadLibrary(LibName)
		    If hModule = Nil Then Return Nil
		    Dim procn As New MemoryBlock(ProcName.Len + 1)
		    procn.CString(0) = ProcName
		    Dim proc As Ptr = GetProcAddress_(hModule.ModuleID, Procn)
		    If proc <> Nil Then
		      Procedure = New DynamicInvoke.Invoker(proc, ProcName, hModule)
		      Procedures.Value(LibName + ProcName) = Procedure
		    End If
		  End If
		  Return Procedure
		End Function
	#tag EndMethod

	#tag ExternalMethod, Flags = &h21
		Private Soft Declare Function GetProcAddress_ Lib "Kernel32" Alias "GetProcAddress" (hModule As Integer, ProcName As Ptr) As Ptr
	#tag EndExternalMethod

	#tag Method, Flags = &h1
		Protected Function IsFunctionAvailable(LibName As String, ProcName As String) As Boolean
		  Return GetProcAddress(LibName, ProcName) <> Nil
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function IsOrdinalAvailable(LibName As String, Ordinal As Integer) As Boolean
		  Return GetProcAddress(LibName, Ordinal) <> Nil
		End Function
	#tag EndMethod

	#tag ExternalMethod, Flags = &h21
		Private Soft Declare Function LoadLibraryW Lib "Kernel32" (LibName As WString) As Integer
	#tag EndExternalMethod


	#tag Property, Flags = &h21
		Private Procedures As Dictionary
	#tag EndProperty


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
