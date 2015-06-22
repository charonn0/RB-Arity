#tag Class
Protected Class Library
	#tag Method, Flags = &h1
		Protected Sub Constructor(hMod As Integer)
		  hModule = hMod
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function File() As FolderItem
		  #If Not TargetWin32 Then
		    Raise New PlatformNotSupportedException
		  #endif
		  Dim path As New MemoryBlock(1024)
		  Dim sz As Integer = GetModuleFileNameW(hModule, path, path.Size)
		  If sz > 0 Then Return GetFolderItem(path.WString(0), FolderItem.PathTypeAbsolute)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		 Shared Function FreeLibrary(LibName As String) As Boolean
		  #If Not TargetWin32 Then
		    Raise New PlatformNotSupportedException
		  #endif
		  If Modules = Nil Then Modules = New Dictionary
		  If Modules.HasKey(LibName) And FreeLibrary(LibName) Then
		    Modules.Remove(LibName)
		    Return True
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		 Shared Function LoadLibrary(LibName As String) As DynamicInvoke.Library
		  #If Not TargetWin32 Then
		    Raise New PlatformNotSupportedException
		  #endif
		  If Modules = Nil Then Modules = New Dictionary
		  If Modules.HasKey(LibName) Then Return Modules.Value(LibName)
		  Dim h As Integer = GetModuleHandleW(LibName)
		  Dim hMod As DynamicInvoke.Library
		  If h = 0 Then h = LoadLibraryW(LibName)
		  If h = 0 Then Return Nil
		  hMod = New Library(h)
		  Modules.Value(LibName) = hMod
		  Return hMod
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ModuleID() As Integer
		  Return hModule
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Name() As String
		  Dim f As FolderItem = Me.File
		  If f <> Nil Then Return f.Name
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Operator_Compare(OtherLib As DynamicInvoke.Library) As Integer
		  If OtherLib Is Nil Then Return 1
		  Return Sign(Me.hModule - OtherLib.hModule)
		End Function
	#tag EndMethod


	#tag Property, Flags = &h1
		Protected hModule As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private Shared Modules As Dictionary
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
