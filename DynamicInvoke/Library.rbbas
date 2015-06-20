#tag Class
Private Class Library
	#tag Method, Flags = &h1
		Protected Sub Constructor(hMod As Integer)
		  hModule = hMod
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		 Shared Function LoadLibrary(LibName As String) As Library
		  If Modules = Nil Then Modules = New Dictionary
		  Dim h As Integer = LoadLibraryW(LibName)
		  If h = 0 Then Return Nil
		  Modules.Value(LibName) = h
		  Return New Library(h)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ModuleID() As Integer
		  Return hModule
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Name() As String
		  For Each hmod As String In Modules.Keys
		    If Modules.Value(hmod) = hModule Then Return hmod
		  Next
		End Function
	#tag EndMethod


	#tag Property, Flags = &h1
		Protected hModule As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private Shared Modules As Dictionary
	#tag EndProperty


End Class
#tag EndClass
