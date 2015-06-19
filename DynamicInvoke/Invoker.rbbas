#tag Class
Protected Class Invoker
	#tag Method, Flags = &h0
		Sub Constructor(Proc As Ptr)
		  Procedure = Proc
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Invoke(ParamArray Args() As Variant)
		  Dim pArgs() As Ptr
		  For i As Integer = 0 To UBound(Args)
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
		    
		  End Select
		End Function
	#tag EndMethod


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
