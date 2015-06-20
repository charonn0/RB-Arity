#tag Class
Private Class Library
	#tag Method, Flags = &h1
		Protected Sub Constructor(hMod As Integer)
		  hModule = hMod
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function File() As FolderItem
		  Dim path As New MemoryBlock(1024)
		  Dim sz As Integer = GetModuleFileNameW(hModule, path, path.Size)
		  If sz > 0 Then Return GetFolderItem(path.WString(0), FolderItem.PathTypeAbsolute)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		 Shared Function FreeLibrary(LibName As String) As Boolean
		  If Modules = Nil Then Modules = New Dictionary
		  If Modules.HasKey(LibName) And FreeLibrary(LibName) Then
		    Modules.Remove(LibName)
		    Return True
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		 Shared Function LoadLibrary(LibName As String) As Library
		  If Modules = Nil Then Modules = New Dictionary
		  If Modules.HasKey(LibName) Then Return Modules.Value(LibName)
		  Dim h As Integer = GetModuleHandleW(LibName)
		  Dim hMod As Library
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


	#tag Property, Flags = &h1
		Protected hModule As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private Shared Modules As Dictionary
	#tag EndProperty


End Class
#tag EndClass
