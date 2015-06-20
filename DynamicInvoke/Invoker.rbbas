#tag Class
Protected Class Invoker
	#tag Method, Flags = &h0
		Function ArgType(Index As Integer) As Integer
		  If UBound(mArgTypes) >= Index Then Return mArgTypes(Index)
		  Return -1
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ArgType(Index As Integer, Assigns NewVarType As Integer)
		  If UBound(mArgTypes) < Index Then ReDim mArgTypes(Index)
		  mArgTypes(Index) = NewVarType
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(Proc As Ptr, ProcName As String, LibModule As Library)
		  Procedure = Proc
		  hModule = LibModule
		  mName = ProcName
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Invoke(ParamArray Args() As Variant)
		  Dim pArgs() As Ptr
		  For i As Integer = 0 To UBound(Args)
		    Dim t As Integer = -1
		    If UBound(mArgTypes) >= i Then t = mArgTypes(i)
		    pArgs.Append(MarshalToPtr(Args(i)))
		  Next
		  Select Case UBound(Args)
		  Case -1
		    Dim proc As New ArityS0(Procedure)
		    proc.Invoke()
		    
		  Case 0
		    Dim proc As New ArityS1(Procedure)
		    proc.Invoke(pArgs(0))
		    
		  Case 1
		    Dim proc As New ArityS2(Procedure)
		    proc.Invoke(pArgs(0), pArgs(1))
		    
		  Case 2
		    Dim proc As New ArityS3(Procedure)
		    proc.Invoke(pArgs(0), pArgs(1), pArgs(2))
		    
		  Case 3
		    Dim proc As New ArityS4(Procedure)
		    proc.Invoke(pArgs(0), pArgs(1), pArgs(2), pArgs(3))
		    
		  Case 4
		    Dim proc As New ArityS5(Procedure)
		    proc.Invoke(pArgs(0), pArgs(1), pArgs(2), pArgs(3), pArgs(4))
		    
		  Case 5
		    Dim proc As New ArityS6(Procedure)
		    proc.Invoke(pArgs(0), pArgs(1), pArgs(2), pArgs(3), pArgs(4), pArgs(5))
		  Else
		    Raise New TooManyArgumentsException
		  End Select
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Invoke(ParamArray Args() As Variant) As Variant
		  Dim pArgs() As Ptr
		  For i As Integer = 0 To UBound(Args)
		    pArgs.Append(MarshalToPtr(Args(i)))
		  Next
		  Select Case UBound(Args)
		  Case -1
		    Dim proc As New ArityF0(Procedure)
		    Return proc.Invoke()
		    
		  Case 0
		    Dim proc As New ArityF1(Procedure)
		    Return proc.Invoke(pArgs(0))
		    
		  Case 1
		    Dim proc As New ArityF2(Procedure)
		    Return proc.Invoke(pArgs(0), pArgs(1))
		    
		  Case 2
		    Dim proc As New ArityF3(Procedure)
		    Return proc.Invoke(pArgs(0), pArgs(1), pArgs(2))
		    
		  Case 3
		    Dim proc As New ArityF4(Procedure)
		    Return proc.Invoke(pArgs(0), pArgs(1), pArgs(2), pArgs(3))
		    
		  Case 4
		    Dim proc As New ArityF5(Procedure)
		    Return proc.Invoke(pArgs(0), pArgs(1), pArgs(2), pArgs(3), pArgs(4))
		    
		  Case 5
		    Dim proc As New ArityF6(Procedure)
		    Return proc.Invoke(pArgs(0), pArgs(1), pArgs(2), pArgs(3), pArgs(4), pArgs(5))
		    
		  Else
		    Raise New TooManyArgumentsException
		  End Select
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Library() As String
		  Return hModule.Name
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Shared Function MarshalToPtr(DataValue As Variant, ValueType As Integer = -1) As Ptr
		  If ValueType = -1 Then ValueType = VarType(DataValue)
		  Select Case ValueType
		    
		  Case Variant.TypeNil
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

	#tag Method, Flags = &h0
		Function Name() As String
		  Return mName
		End Function
	#tag EndMethod


	#tag Property, Flags = &h1
		Protected hModule As Library
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected mArgTypes() As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mName As String
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected Procedure As Ptr
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
End Class
#tag EndClass
