Attribute VB_Name = "WindowScannerModule"
'This module contains this program's core procedures.
Option Explicit

'The Microsoft Windows API constants, functions and structures used by this program.
Private Type LUID
   LowPart As Long
   HighPart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
   pLuid As LUID
   Attributes As Long
End Type

Public Type POINTAPI
   x As Long
   y As Long
End Type

Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type TOKEN_PRIVILEGES
   PrivilegeCount As Long
   Privileges(1) As LUID_AND_ATTRIBUTES
End Type

Public Const ES_PASSWORD As Long = &H20&
Public Const GWL_STYLE As Long = -16
Public Const HWND_BOTTOM As Long = 1
Public Const HWND_NOTOPMOST As Long = -2
Public Const HWND_TOP As Long = 0
Public Const HWND_TOPMOST As Long = -1
Public Const SW_HIDE As Long = 0
Public Const SW_RESTORE As Long = 9
Public Const SW_SHOWMAXIMIZED As Long = 3
Public Const SW_SHOWMINIMIZED As Long = 2
Public Const SW_SHOWNA As Long = 8
Public Const SWP_DRAWFRAME As Long = &H20&
Public Const SWP_NOACTIVATE As Long = &H10&
Public Const SWP_NOMOVE As Long = &H2&
Public Const SWP_SHOWWINDOW As Long = &H40&
Public Const WM_CLOSE As Long = &H10&
Public Const WM_SETTEXT As Long = &HC&
Public Const WS_CHILD As Long = &H40000000
Public Const WS_GROUP As Long = &H20000
Public Const WS_POPUP As Long = &H80000000
Public Const WS_TABSTOP As Long = &H10000
Private Const EM_GETPASSWORDCHAR As Long = &HD2&
Private Const EM_SETPASSWORDCHAR As Long = &HCC&
Private Const ERROR_IO_PENDING As Long = 997
Private Const ERROR_NOT_ALL_ASSIGNED As Long = 1300
Private Const ERROR_SUCCESS As Long = 0
Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000&
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200&
Private Const GCL_HMODULE As Long = -16
Private Const PROCESS_ALL_ACCESS As Long = &H1F0FFF
Private Const PROCESS_QUERY_INFORMATION As Long = &H400&
Private Const SE_DEBUG_NAME As String = "SeDebugPrivilege"
Private Const SE_PRIVILEGE_DISABLED As Long = &H0&
Private Const SE_PRIVILEGE_ENABLED As Long = &H2&
Private Const TOKEN_ALL_ACCESS As Long = &HFF&
Private Const WM_GETTEXT As Long = &HD&
Private Const WM_GETTEXTLENGTH As Long = &HE&

Public Declare Function BringWindowToTop Lib "User32.dll" (ByVal hwnd As Long) As Long
Public Declare Function EnableWindow Lib "User32.dll" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public Declare Function EnumWindows Lib "User32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function GetDesktopWindow Lib "User32.dll" () As Long
Public Declare Function GetParent Lib "User32.dll" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowLongA Lib "User32.dll" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowRect Lib "User32.dll" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function IsIconic Lib "User32.dll" (ByVal hwnd As Long) As Long
Public Declare Function IsWindow Lib "User32.dll" (ByVal hwnd As Long) As Long
Public Declare Function IsWindowEnabled Lib "User32.dll" (ByVal hwnd As Long) As Long
Public Declare Function IsWindowUnicode Lib "User32.dll" (ByVal hwnd As Long) As Long
Public Declare Function IsWindowVisible Lib "User32.dll" (ByVal hwnd As Long) As Long
Public Declare Function MoveWindow Lib "User32.dll" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function ScreenToClient Lib "User32.dll" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function SendMessageW Lib "User32.dll" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetParent Lib "User32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowLongA Lib "User32.dll" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "User32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowWindow Lib "User32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Sub Sleep Lib "Kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare Function AdjustTokenPrivileges Lib "Advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare Function CloseHandle Lib "Kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function EnumChildWindows Lib "User32.dll" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function FormatMessageA Lib "Kernel32.dll" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function GetClassLongA Lib "User32.dll" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetClassNameA Lib "User32.dll" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetCurrentProcess Lib "Kernel32.dll" () As Long
Private Declare Function GetModuleFileNameExA Lib "Psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function GetProcessImageFileNameW Lib "Psapi.dll" (ByVal hProcess As Long, ByVal lpImageFileName As Long, ByVal nSize As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "User32.dll" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function LookupPrivilegeValueA Lib "Advapi32.dll" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function OpenProcessToken Lib "Advapi32.dll" (ByVal ProcessH As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function PostMessageA Lib "User32.dll" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function RealGetWindowClassA Lib "User32.dll" (ByVal hwnd As Long, ByVal pszType As String, ByVal cchType As Long) As Long
Private Declare Function UpdateWindow Lib "User32.dll" (ByVal hwnd As Long) As Long
Private Declare Function WaitMessage Lib "User32.dll" () As Long

'The constants used by this program.
Public Const NO_HANDLE As Long = 0        'Indicates no handle.
Private Const MAX_PATH As Long = 260       'Defines the maximum number of characters allowed for a file path.
Private Const MAX_STRING As Long = 65535   'Defines the maximum number of characters used for a string buffer.

'This structure defines the thread, module and process information of a window.
Public Type WindowProcessStr
   ModuleH As Long         'Defines the handle of the module that created the window.
   ModulePath As String    'Defines the path of the module that created the window.
   ThreadId As Long        'Defines the window's thread id.
   ProcessH As Long        'Defines the handle of the process to which the window belongs.
   ProcessId As Long       'Defines the id of the process to which the window belongs.
   ProcessPath As String   'Defines the path of the process' executable to which the window belongs.
End Type

'This structure defines the properties of a window.
Public Type WindowStr
   ClassName As String   'Defines the window's class.
   Enabled As Boolean    'Indicates whether the window is enabled.
   Handle As Long        'Defines the window's handle.
   Parent As Long        'Defines the window's parent.
   Text As String        'Defines the window's text.
   Visible As Boolean    'Inicates whether the window is visible.
End Type

Public Windows() As WindowStr   'Contains all open windows found.
'This procedure checks whether an error has occurred during the most recent Windows API call.
Public Function CheckForError(Optional ReturnValue As Long = 0, Optional ResetSuppression As Boolean = False, Optional Ignored As Long = ERROR_SUCCESS) As Long
Dim Description As String
Dim ErrorCode As Long
Dim Length As Long
Dim Message As String
Static SuppressAPIErrors As Boolean

   ErrorCode = Err.LastDllError
   Err.Clear
   
   On Error GoTo ErrorTrap
   
   If ResetSuppression Then SuppressAPIErrors = False
   
   If Not (ErrorCode = ERROR_SUCCESS Or ErrorCode = Ignored) Then
      Description = String$(MAX_STRING, vbNullChar)
      Length = FormatMessageA(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, CLng(0), ErrorCode, CLng(0), Description, Len(Description), CLng(0))
      If Length = 0 Then
         Description = "No description."
      ElseIf Length > 0 Then
         Description = Left$(Description, Length - 1)
      End If
     
      Message = "API error code: " & CStr(ErrorCode) & " - " & Description
      Message = Message & "Return value: " & CStr(ReturnValue) & vbCr
      Message = Message & "Continue displaying API error messages?"
      If Not SuppressAPIErrors Then SuppressAPIErrors = (MsgBox(Message, vbYesNo Or vbExclamation) = vbNo)
   End If
   
EndRoutine:
   CheckForError = ReturnValue
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure returns the specified window's base class.
Public Function GetWindowBaseClass(WindowH As Long) As String
On Error GoTo ErrorTrap
Dim Length As Long
Dim WindowBaseClass As String

   WindowBaseClass = String$(MAX_STRING, vbNullChar)
   Length = CheckForError(RealGetWindowClassA(WindowH, ByVal WindowBaseClass, Len(WindowBaseClass)))
   WindowBaseClass = Left$(WindowBaseClass, Length)
   
EndRoutine:
   GetWindowBaseClass = WindowBaseClass
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function




'This procedure returns the specified window's class.
Public Function GetWindowClass(WindowH As Long) As String
On Error GoTo ErrorTrap
Dim Length As Long
Dim WindowClass As String

   WindowClass = String$(MAX_STRING, vbNullChar)
   Length = CheckForError(GetClassNameA(WindowH, ByVal WindowClass, Len(WindowClass)))
   WindowClass = Left$(WindowClass, Length)
   
EndRoutine:
   GetWindowClass = WindowClass
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure collects the specified window's information and adds it to the list of active windows found.
Public Sub GetWindowInformation(WindowH As Long)
On Error GoTo ErrorTrap
   ReDim Preserve Windows(LBound(Windows()) To UBound(Windows()) + 1) As WindowStr
   
   With Windows(UBound(Windows()))
      .ClassName = GetWindowClass(WindowH)
      .Enabled = CBool(CheckForError(IsWindowEnabled(WindowH)))
      .Handle = WindowH
      .Parent = CheckForError(GetParent(WindowH))
      .Text = GetWindowText(WindowH)
      .Visible = CBool(CheckForError(IsWindowVisible(WindowH)))
   End With
   
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure returns the specified window's process information.
Public Function GetWindowProcess(WindowH As Long) As WindowProcessStr
On Error GoTo ErrorTrap
Dim Length As Long
Dim WindowProcess As WindowProcessStr

   With WindowProcess
      .ModuleH = CheckForError(GetClassLongA(WindowH, GCL_HMODULE))
      .ThreadId = CheckForError(GetWindowThreadProcessId(WindowH, .ProcessId))
      
      .ProcessH = CheckForError(OpenProcess(PROCESS_ALL_ACCESS, CLng(False), .ProcessId))
      If Not .ProcessH = NO_HANDLE Then
         .ModulePath = String$(MAX_PATH, vbNullChar)
         Length = CheckForError(GetModuleFileNameExA(.ProcessH, .ModuleH, .ModulePath, Len(.ModulePath)))
         CheckForError CloseHandle(.ProcessH)
      End If
      
      .ProcessH = CheckForError(OpenProcess(PROCESS_QUERY_INFORMATION, CLng(False), .ProcessId))
      If Not .ProcessH = NO_HANDLE Then
         .ProcessPath = String$(MAX_PATH, vbNullChar)
         Length = CheckForError(GetProcessImageFileNameW(.ProcessH, StrPtr(.ProcessPath), Len(.ProcessPath)))
         .ProcessPath = Left$(.ProcessPath, Length)
         CheckForError CloseHandle(.ProcessH)
      End If
   End With
   
EndRoutine:
   GetWindowProcess = WindowProcess
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure returns the specified window's text.
Public Function GetWindowText(WindowH As Long) As String
On Error GoTo ErrorTrap
Dim Length As Long
Dim PasswordCharacter As Long
Dim WindowText As String

   If WindowHasStyle(WindowH, ES_PASSWORD) Then
      PasswordCharacter = CheckForError(SendMessageW(WindowH, EM_GETPASSWORDCHAR, CLng(0), CLng(0)))
      If Not PasswordCharacter = 0 Then
         CheckForError PostMessageA(WindowH, EM_SETPASSWORDCHAR, CLng(0), CLng(0))
         Sleep CLng(1000)
      End If
   End If
   
   WindowText = String$(CheckForError(SendMessageW(WindowH, WM_GETTEXTLENGTH, CLng(0), CLng(0))) + 1, vbNullChar)
   Length = CheckForError(SendMessageW(WindowH, WM_GETTEXT, Len(WindowText), StrPtr(WindowText)))
   
   If Not PasswordCharacter = 0 Then
      CheckForError PostMessageA(WindowH, EM_SETPASSWORDCHAR, PasswordCharacter, CLng(0))
   End If
   
   WindowText = Left$(WindowText, Length)
   
EndRoutine:
   GetWindowText = WindowText
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure handles any child windows that are found.
Private Function HandleChildWindows(ByVal hwnd As Long, ByVal lParam As Long) As Long
On Error GoTo ErrorTrap

   GetWindowInformation hwnd
   
EndRoutine:
   HandleChildWindows = CLng(True) 'Indicates to continue enumerating child windows.
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure handles any errors that occur.
Public Sub HandleError()
Dim Description As String
Dim ErrorCode As Long

   ErrorCode = Err.Number
   Description = Err.Description
   
   On Error Resume Next
   MsgBox "Error: " & CStr(ErrorCode) & vbCr & Description, vbExclamation
End Sub

'This procedure handles any top level windows that are found.
Public Function HandleWindows(ByVal hwnd As Long, ByVal lParam As Long) As Long
On Error GoTo ErrorTrap
   
   GetWindowInformation hwnd
   CheckForError EnumChildWindows(hwnd, AddressOf HandleChildWindows, CLng(0))
   
EndRoutine:
   HandleWindows = CLng(True)
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function


'This procedure is executed when this program is started.
Private Sub Main()
On Error GoTo ErrorTrap

   CheckForError , ResetSuppression:=True
   ReDim Windows(0 To 0) As WindowStr
   SetDebugPrivilege Disabled:=False
   
   WindowScannerWindow.Show
   Do While DoEvents()
      CheckForError WaitMessage()
   Loop
   
   SetDebugPrivilege Disabled:=True
   
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure returns information about this program.
Public Function ProgramInformation() As String
On Error GoTo ErrorTrap
Dim Information As String

   With App
      Information = .Title & " v" & CStr(.Major) & "." & CStr(.Minor) & CStr(.Revision) & " - by: " & App.CompanyName
   End With

EndRoutine:
   ProgramInformation = Information
   Exit Function

ErrorTrap:
   HandleError
   Resume EndRoutine
End Function


'This procedure indicates whether the specified handle refers to a window.
Public Function RefersToWindow(WindowH As Long) As Boolean
On Error GoTo ErrorTrap
Dim HIsWindow As Boolean
Dim Message As String

   HIsWindow = CBool(CheckForError(IsWindow(WindowH)))
   If Not HIsWindow Then
      Message = "The selected object is not a window." & vbCr
      Message = Message & "This could be due to the following causes:" & vbCr
      Message = Message & "The window no longer exists," & vbCr
      Message = Message & "or its handle has changed."
      MsgBox Message, vbInformation
   End If
   
EndRoutine:
   RefersToWindow = HIsWindow
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure refreshes the specified window and any parent windows.
Public Sub RefreshWindow(ByVal WindowH As Long)
On Error GoTo ErrorTrap

   Do
      CheckForError UpdateWindow(WindowH)
      WindowH = CheckForError(GetParent(WindowH))
      DoEvents
   Loop Until WindowH = NO_HANDLE
   
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure enables/disables the debug privilege.
Private Sub SetDebugPrivilege(Disabled As Boolean)
On Error GoTo ErrorTrap
Dim Length As Long
Dim NewPrivileges As TOKEN_PRIVILEGES
Dim PreviousPrivileges As TOKEN_PRIVILEGES
Dim PrivilegeId As LUID
Dim ReturnValue As Long
Dim TokenH As Long

   ReturnValue = CheckForError(OpenProcessToken(GetCurrentProcess(), TOKEN_ALL_ACCESS, TokenH))
   If Not ReturnValue = 0 Then
      ReturnValue = CheckForError(LookupPrivilegeValueA(vbNullString, SE_DEBUG_NAME, PrivilegeId), , Ignored:=ERROR_IO_PENDING)
      If Not ReturnValue = 0 Then
         NewPrivileges.Privileges(0).pLuid = PrivilegeId
         NewPrivileges.PrivilegeCount = CLng(1)
         
         If Disabled Then
            NewPrivileges.Privileges(0).Attributes = SE_PRIVILEGE_DISABLED
            CheckForError AdjustTokenPrivileges(TokenH, CLng(False), NewPrivileges, Len(NewPrivileges), PreviousPrivileges, Length), , Ignored:=ERROR_NOT_ALL_ASSIGNED
         ElseIf Not Disabled Then
            NewPrivileges.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
            CheckForError AdjustTokenPrivileges(TokenH, CLng(False), NewPrivileges, Len(NewPrivileges), PreviousPrivileges, Length)
         End If
      End If
      CheckForError CloseHandle(TokenH)
   End If
   
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure indicates whether a window has the specified style.
Public Function WindowHasStyle(WindowH As Long, Style As Long) As Boolean
On Error GoTo ErrorTrap

   WindowHasStyle = (CheckForError(GetWindowLongA(WindowH, GWL_STYLE) And Style) = Style)

EndRoutine:
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

