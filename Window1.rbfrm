#tag Window
Begin Window Window1
   BackColor       =   &hFFFFFF
   Backdrop        =   ""
   CloseButton     =   True
   Composite       =   False
   Frame           =   0
   FullScreen      =   False
   HasBackColor    =   False
   Height          =   1.1e+2
   ImplicitInstance=   True
   LiveResize      =   True
   MacProcID       =   0
   MaxHeight       =   32000
   MaximizeButton  =   False
   MaxWidth        =   32000
   MenuBar         =   529260543
   MenuBarVisible  =   True
   MinHeight       =   64
   MinimizeButton  =   True
   MinWidth        =   64
   Placement       =   0
   Resizeable      =   True
   Title           =   "Untitled"
   Visible         =   True
   Width           =   2.82e+2
   Begin PushButton PushButton1
      AutoDeactivate  =   True
      Bold            =   ""
      ButtonStyle     =   0
      Cancel          =   ""
      Caption         =   "MessageBoxW"
      Default         =   ""
      Enabled         =   True
      Height          =   22
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   ""
      Left            =   20
      LockBottom      =   ""
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   ""
      LockTop         =   True
      Scope           =   0
      TabIndex        =   0
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0
      TextUnit        =   0
      Top             =   22
      Underline       =   ""
      Visible         =   True
      Width           =   90
   End
   Begin PushButton PushButton2
      AutoDeactivate  =   True
      Bold            =   ""
      ButtonStyle     =   0
      Cancel          =   ""
      Caption         =   "SetWindowTextA"
      Default         =   ""
      Enabled         =   True
      Height          =   22
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   ""
      Left            =   112
      LockBottom      =   ""
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   ""
      LockTop         =   True
      Scope           =   0
      TabIndex        =   1
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0
      TextUnit        =   0
      Top             =   22
      Underline       =   ""
      Visible         =   True
      Width           =   108
   End
End
#tag EndWindow

#tag WindowCode
	#tag Event
		Sub Open()
		  'If DynamicInvoke.IsFunctionAvailable("libcurl", "curl_global_init") Then
		  'Dim GetLastError As DynamicInvoke.Invoker = DynamicInvoke.GetProcAddress("Kernel32", "GetLastError")
		  'Dim err As Integer
		  'If DynamicInvoke.IsOrdinalAvailable("libcurl", 21) Then
		  'Dim curl_global_init As DynamicInvoke.Invoker = DynamicInvoke.GetProcAddress("libcurl", 21)
		  'curl_global_init.Invoke(3)
		  'Break
		  '
		  'End If
		  'err = GetLastError.Invoke()
		  'Break
		  'End If
		  'Break
		End Sub
	#tag EndEvent


#tag EndWindowCode

#tag Events PushButton1
	#tag Event
		Sub Action()
		  Dim MessageBox As DynamicInvoke.Invoker = DynamicInvoke.GetProcAddress("User32", "MessageBoxW")
		  Const MB_ICONINFORMATION  = &h00000040
		  Const MB_RETRYCANCEL = &h00000005
		  
		  Dim s1 As String = ConvertEncoding(MessageBox.ProcedureName, Encodings.UTF16)
		  Dim s2 As String = ConvertEncoding(MessageBox.Library, Encodings.UTF16)
		  
		  Dim value As Integer = MessageBox.Invoke(Self.Handle, s1, s2, MB_ICONINFORMATION Or MB_RETRYCANCEL)
		  If value = 4 Then
		    MessageBox.Invoke(Self.Handle, "You clicked Retry!", "Retry!", MB_ICONINFORMATION)
		  Else
		    MessageBox.Invoke(Self.Handle, "You clicked Cancel!", "Cancel!", MB_ICONINFORMATION)
		  End If
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events PushButton2
	#tag Event
		Sub Action()
		  Dim SetWindowText As DynamicInvoke.Invoker = DynamicInvoke.GetProcAddress("User32", "SetWindowTextA")
		  Dim caption As CString = "Caption set!"
		  SetWindowText.Invoke(Self.Handle, caption)
		  
		End Sub
	#tag EndEvent
#tag EndEvents
