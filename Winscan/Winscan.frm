VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form WindowScannerWindow 
   ClientHeight    =   4464
   ClientLeft      =   60
   ClientTop       =   636
   ClientWidth     =   4692
   ClipControls    =   0   'False
   Icon            =   "Winscan.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   18.6
   ScaleMode       =   4  'Character
   ScaleWidth      =   39.1
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox IgnoreAmpersandsBox 
      Caption         =   "&Ignore ampersands."
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      ToolTipText     =   "Some window classes do not display (all) ampersands in their text."
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CheckBox LookForParentWindowsBox 
      Caption         =   "&Look for parent windows."
      Height          =   255
      Left            =   240
      TabIndex        =   3
      ToolTipText     =   "Also lists any child windows belonging to the windows matching the specified search criteria."
      Top             =   1320
      Width           =   2295
   End
   Begin VB.CommandButton SearchButton 
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      ToolTipText     =   "Starts looking for windows matching the specified search text and/or class."
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CheckBox WholePhrasesOnlyBox 
      Caption         =   "&Whole phrases only."
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      ToolTipText     =   "The entire text and/or class name must match the specified text and/or class name."
      Top             =   960
      Width           =   1935
   End
   Begin VB.CheckBox CaseSensitiveBox 
      Caption         =   "&Case sensitive."
      Height          =   255
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "The case of the text and/or class names must match the specified search text and/or class name."
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox WindowClassBox 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      ToolTipText     =   "Enter the class of windows to look for in this field."
      Top             =   600
      Width           =   3735
   End
   Begin VB.TextBox WindowTextBox 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      ToolTipText     =   "Enter the text to look for in windows in this field."
      Top             =   120
      Width           =   3735
   End
   Begin MSFlexGridLib.MSFlexGrid SearchResultsTable 
      Height          =   2175
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Double click to perform an action on the selected search result."
      Top             =   2160
      Width           =   4335
      _ExtentX        =   7641
      _ExtentY        =   3831
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
   End
   Begin VB.Label Label 
      Caption         =   "Class:"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label 
      Caption         =   "Text:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   495
   End
   Begin VB.Menu ProgramMainMenu 
      Caption         =   "&Program"
      Begin VB.Menu HelpMenu 
         Caption         =   "&Help"
         Shortcut        =   ^H
      End
      Begin VB.Menu InformationMenu 
         Caption         =   "&Information"
         Shortcut        =   ^J
      End
      Begin VB.Menu ProgramSeparator1Menu 
         Caption         =   "-"
      End
      Begin VB.Menu QuitMenu 
         Caption         =   "&Quit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu SearchResultsMainMenu 
      Caption         =   "S&earch-results"
      Begin VB.Menu CopyMenu 
         Caption         =   "&Copy."
         Shortcut        =   ^C
      End
      Begin VB.Menu FindNextMatchMenu 
         Caption         =   "Find &next match."
         Shortcut        =   {F3}
      End
      Begin VB.Menu FindTextMenu 
         Caption         =   "&Find text."
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu WindowMainMenu 
      Caption         =   "&Window"
      Begin VB.Menu CloseWindowMenu 
         Caption         =   "&Close window."
         Shortcut        =   ^D
      End
      Begin VB.Menu EnableDisableWindowMenu 
         Caption         =   ""
         Shortcut        =   {F1}
      End
      Begin VB.Menu ShowHideWindowMenu 
         Caption         =   ""
         Shortcut        =   {F2}
      End
      Begin VB.Menu WindowSeparator1Menu 
         Caption         =   "-"
      End
      Begin VB.Menu ChangePositionDimensionsMenu 
         Caption         =   "Change &dimensions/position."
         Shortcut        =   ^P
      End
      Begin VB.Menu ChangeParentMenu 
         Caption         =   "Change &parent."
         Shortcut        =   ^Q
      End
      Begin VB.Menu ChangeStateMainMenu 
         Caption         =   "Change &state."
         Begin VB.Menu ChangeStateMenu 
            Caption         =   "&Maximize."
            Index           =   0
            Shortcut        =   ^M
         End
         Begin VB.Menu ChangeStateMenu 
            Caption         =   "M&inimize."
            Index           =   1
            Shortcut        =   ^N
         End
         Begin VB.Menu ChangeStateMenu 
            Caption         =   "&Restore."
            Index           =   2
            Shortcut        =   ^R
         End
      End
      Begin VB.Menu ChangeTextMenu 
         Caption         =   "Change &text."
         Shortcut        =   ^T
      End
      Begin VB.Menu ChangeZOrderMainMenu 
         Caption         =   "Change &z-order."
         Begin VB.Menu ChangeZOrderMenu 
            Caption         =   "&Bottom."
            Index           =   0
            Shortcut        =   ^B
         End
         Begin VB.Menu ChangeZOrderMenu 
            Caption         =   "&Middle (Below permant top windows.)"
            Index           =   1
            Shortcut        =   ^O
         End
         Begin VB.Menu ChangeZOrderMenu 
            Caption         =   "&Top."
            Index           =   2
            Shortcut        =   ^U
         End
         Begin VB.Menu ChangeZOrderMenu 
            Caption         =   "To&p (Permanently.)"
            Index           =   3
            Shortcut        =   ^W
         End
      End
      Begin VB.Menu FlashWindowMenu 
         Caption         =   "&Flash window."
         Shortcut        =   ^G
      End
      Begin VB.Menu WindowSeparator2Menu 
         Caption         =   "-"
      End
      Begin VB.Menu GetBaseClassInformationMenu 
         Caption         =   "&Get Base Class Information"
         Shortcut        =   ^E
      End
      Begin VB.Menu GetProcessInformationMenu 
         Caption         =   "&Get process information."
         Shortcut        =   ^I
      End
   End
   Begin VB.Menu ExcludeMainMenu 
      Caption         =   "E&xclude"
      Begin VB.Menu ExcludeMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "WindowScannerWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This module contains this program's interface.
Option Explicit

'This enumeration lists the window properties which can be excluded from the search results.
Private Enum ExcludableE
   ExcludeNone = -1    'Exclude no windows.
   ExcludeChild        'Child windows.
   ExcludeParent       'Parent windows.
   ExcludeDisabled     'Disabled windows.
   ExcludeEnabled      'Enabled windows.
   ExcludeHidden       'Hidden windows.
   ExcludeVisible      'Visible windows.
   ExcludeNonGroup     'Non-Group windows.
   ExcludeGroup        'Group windows.
   ExcludeNonPopup     'Non-popup windows.
   ExcludePopup        'Popup windows.
   ExcludeNonTabStop   'Non-TabStop windows.
   ExcludeTabStop      'TabStop windows.
   ExcludeNonUnicode   'Non-Unicode windows.
   ExcludeUnicode      'Unicode windows.
End Enum

Private Matches() As Long   'Contains the indexes of the windows that match the specified search criteria.

'This procedure builds elements of this window's interface.
Public Sub BuildInterface()
On Error GoTo ErrorTrap
Dim ExcludeMenus() As Variant
Dim Index As Long

   Me.Width = Screen.Width / 2
   Me.Height = Screen.Height / 2
   Me.Caption = ProgramInformation()
   
   With SearchResultsTable
      .Row = 0
      .Col = 0: .Text = "Handle:"
      .Col = 1: .Text = "Text:"
      .Col = 2: .Text = "Class:"
      .Col = 3: .Text = "Parent:"
   End With
   
   ExcludeMenus = Array("Child", "Parent", "Disabled", "Enabled", "Hidden", "Visible", "Non-group", "Group", "Non-popup", "Popup", "Non-tab stop", "Tab stop", "Non-unicode", "Unicode")
   For Index = LBound(ExcludeMenus()) To UBound(ExcludeMenus())
      If Index > ExcludeMenu.UBound Then Load ExcludeMenu(Index)
      ExcludeMenu(Index).Caption = "&" & ExcludeMenus(Index) & " windows."
   Next Index
   
   ToggleExcludedProperty ExcludeNone
    
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure changes the specified window's parent window.
Private Sub ChangeParent(Index As Long)
On Error GoTo ErrorTrap
Dim NewParent As String
Dim Styles As Long

   If RefersToWindow(Windows(Index).Handle) Then
      NewParent = InputBox$("New parent window:", , CStr(Windows(Index).Parent))
      
      If Not StrPtr(NewParent) = 0 Then
         CheckForError SetParent(Windows(Index).Handle, CLng(Val(NewParent)))
         
         Styles = CheckForError(GetWindowLongA(Windows(Index).Handle, GWL_STYLE))
         If CLng(Val(NewParent)) = NO_HANDLE Then
            If WindowHasStyle(Windows(Index).Handle, WS_CHILD) Then Styles = Styles Xor WS_CHILD
         Else
            Styles = Styles Or WS_CHILD
         End If
         CheckForError SetWindowLongA(Windows(Index).Handle, GWL_STYLE, Styles)
         
         RefreshWindow Windows(Index).Handle
         UpdateSearchResults
      End If
   Else
      DisplaySearchResult Index
   End If
   
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure checks whether the specified window has an excluded property.
Private Function IsExcluded(Index As Long) As Boolean
On Error GoTo ErrorTrap
Dim Exclusion As ExcludableE
Dim Excluded As Boolean

   Excluded = False
   For Exclusion = ExcludeMenu.LBound To ExcludeMenu.UBound
      If ExcludeMenu(Exclusion).Checked Then
         Select Case Exclusion
            Case ExcludeChild
               Excluded = (Not Windows(Index).Parent = NO_HANDLE)
            Case ExcludeParent
               Excluded = (Windows(Index).Parent = NO_HANDLE)
            Case ExcludeDisabled
               Excluded = Not Windows(Index).Enabled
            Case ExcludeEnabled
               Excluded = Windows(Index).Enabled
            Case ExcludeHidden
               Excluded = Not Windows(Index).Visible
            Case ExcludeVisible
               Excluded = Windows(Index).Visible
            Case ExcludeNonGroup
               Excluded = Not WindowHasStyle(Windows(Index).Handle, WS_GROUP)
            Case ExcludeGroup
               Excluded = WindowHasStyle(Windows(Index).Handle, WS_GROUP)
            Case ExcludeNonPopup
               Excluded = Not WindowHasStyle(Windows(Index).Handle, WS_POPUP)
            Case ExcludePopup
               Excluded = WindowHasStyle(Windows(Index).Handle, WS_POPUP)
            Case ExcludeNonTabStop
               Excluded = Not WindowHasStyle(Windows(Index).Handle, WS_TABSTOP)
            Case ExcludeTabStop
               Excluded = WindowHasStyle(Windows(Index).Handle, WS_TABSTOP)
            Case ExcludeNonUnicode
               Excluded = Not CBool(CheckForError(IsWindowUnicode(Windows(Index).Handle)))
            Case ExcludeUnicode
               Excluded = CBool(CheckForError(IsWindowUnicode(Windows(Index).Handle)))
         End Select
      End If
      If Excluded Then Exit For
   Next Exclusion
   
EndRoutine:
   IsExcluded = Excluded
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure gives the command to change the selected window's parent window.
Private Sub ChangeParentMenu_Click()
On Error GoTo ErrorTrap

   ChangeParent Matches(SearchResultsTable.Row - 1)

EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure sets the specified window's dimensions and position.
Private Sub ChangePositionDimensions(Index As Long)
On Error GoTo ErrorTrap
Dim Coordinate As POINTAPI
Dim Dimensions As RECT
Dim NewXywh() As String
Dim Xywh As String
   
   If RefersToWindow(Windows(Index).Handle) Then
      CheckForError GetWindowRect(Windows(Index).Handle, Dimensions)
      
      ReDim NewXywh(0 To 3) As String
      
      With Dimensions
         NewXywh(2) = CStr(.Right - .Left)
         NewXywh(3) = CStr(.Bottom - .Top)
         
         If Not Windows(Index).Parent = NO_HANDLE Then
            Coordinate.x = .Left
            Coordinate.y = .Top
            CheckForError ScreenToClient(Windows(Index).Parent, Coordinate)
            .Left = CStr(Coordinate.x)
            .Top = CStr(Coordinate.y)
         End If
      
         NewXywh(0) = CStr(.Left)
         NewXywh(1) = CStr(.Top)
      End With
      
      Xywh = InputBox$("New dimensions and position (x, y, width, height):", , Join(NewXywh(), ","))
      If Not Xywh = vbNullString Then
         NewXywh() = Split(Replace(Xywh, " ", vbNullString), ",")
         
         CheckForError MoveWindow(Windows(Index).Handle, CLng(Val(NewXywh(0))), CLng(Val(NewXywh(1))), CLng(Val(NewXywh(2))), CLng(Val(NewXywh(3))), CLng(True))
         RefreshWindow Windows(Index).Handle
         UpdateSearchResults
      End If
   Else
     DisplaySearchResult Index
   End If
    
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the command to set the selected window's dimensions and position.
Private Sub ChangePositionDimensionsMenu_Click()
On Error GoTo ErrorTrap

   ChangePositionDimensions Matches(SearchResultsTable.Row - 1)

EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure changes the specified window's state to the specified new state.
Private Sub ChangeState(Index As Long, NewState As Long)
On Error GoTo ErrorTrap

   If RefersToWindow(Windows(Index).Handle) Then
      CheckForError ShowWindow(Windows(Index).Handle, NewState)
      RefreshWindow Windows(Index).Handle
      UpdateSearchResults
   Else
      DisplaySearchResult Index
   End If
       
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the command to maximize, minimize, or, restore the selected window.
Private Sub ChangeStateMenu_Click(Index As Integer)
On Error GoTo ErrorTrap
Dim NewState As Long
   
   Select Case Index
      Case 0
         NewState = SW_SHOWMAXIMIZED
      Case 1
         NewState = SW_SHOWMINIMIZED
      Case 2
         NewState = SW_RESTORE
   End Select
   
   ChangeState Matches(SearchResultsTable.Row - 1), NewState
   
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure changes the specified window's text.
Private Sub ChangeText(Index As Long)
On Error GoTo ErrorTrap
Dim WindowText As String
 
   If RefersToWindow(Windows(Index).Handle) Then
      WindowText = GetWindowText(Windows(Index).Handle)
      WindowText = InputBox$("New window text:", , WindowText)
      If Not StrPtr(WindowText) = 0 Then
         CheckForError SendMessageW(Windows(Index).Handle, WM_SETTEXT, CLng(0), StrPtr(WindowText))
      
         RefreshWindow Windows(Index).Handle
         UpdateSearchResults
      End If
   Else
      DisplaySearchResult Index
   End If
    
   
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the command to change the selected window's text.
Private Sub ChangeTextMenu_Click()
On Error GoTo ErrorTrap

   ChangeText Matches(SearchResultsTable.Row - 1)

EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure changes the specified window's z-order to specified new z-order.
Private Sub ChangeZorder(Index As Long, NewZOrder As Long)
On Error GoTo ErrorTrap
Dim Dimensions As RECT

   If RefersToWindow(Windows(Index).Handle) Then
      CheckForError GetWindowRect(Windows(Index).Handle, Dimensions)
      
      With Dimensions
         CheckForError SetWindowPos(Windows(Index).Handle, NewZOrder, CLng(0), CLng(0), .Right - .Left, .Bottom - .Top, SWP_DRAWFRAME Or SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_SHOWWINDOW)
      End With
      
      RefreshWindow Windows(Index).Handle
      UpdateSearchResults
   Else
      DisplaySearchResult Index
   End If
      
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the command to change the selected window's z-order.
Private Sub ChangeZOrderMenu_Click(Index As Integer)
On Error GoTo ErrorTrap
Dim NewZOrder As Long
 
   Select Case NewZOrder
      Case 0
         NewZOrder = HWND_BOTTOM
      Case 1
         NewZOrder = HWND_NOTOPMOST
      Case 2
         NewZOrder = HWND_TOP
      Case 3
         NewZOrder = HWND_TOPMOST
   End Select
   
   ChangeZorder Matches(SearchResultsTable.Row - 1), NewZOrder
   
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the command to close the selected window.
Private Sub CloseWindowMenu_Click()
On Error GoTo ErrorTrap

   CloseWindow Matches(SearchResultsTable.Row - 1)
   
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure closes the specified window.
Private Sub CloseWindow(Index As Long)
On Error GoTo ErrorTrap

   If RefersToWindow(Windows(Index).Handle) Then
      CheckForError SendMessageW(Windows(Index).Handle, WM_CLOSE, CLng(0), CLng(0))
      If Not Windows(Index).Parent = 0 Then RefreshWindow Windows(Index).Parent
      UpdateSearchResults
   Else
      DisplaySearchResult Index
   End If
   
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure copies the selected window information to the clipboard.
Private Sub CopyMenu_Click()
On Error GoTo ErrorTrap

   Clipboard.Clear
   Clipboard.SetText SearchResultsTable.Text, vbCFText

EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure displays the specified search result at the selected row.
Private Sub DisplaySearchResult(Index As Long)
On Error GoTo ErrorTrap
Dim Column As Long
Dim CurrentColumn As Long

   With SearchResultsTable
      CurrentColumn = .Col
      .Col = 0: .Text = Windows(Index).Handle
      If Windows(Index).Parent = 0 Then .CellAlignment = flexAlignLeftCenter Else .CellAlignment = flexAlignRightCenter
      .Col = 1: .CellAlignment = flexAlignLeftCenter: .Text = Windows(Index).Text
      .Col = 2: .CellAlignment = flexAlignLeftCenter: .Text = Windows(Index).ClassName
      .Col = 3: .CellAlignment = flexAlignRightCenter: .Text = Windows(Index).Parent
      For Column = 0 To 3
         .Col = Column
         If CBool(CheckForError(IsWindow(Windows(Index).Handle))) Then
            If WindowHasStyle(Windows(Index).Handle, WS_POPUP) Then .CellBackColor = vbCyan Else .CellBackColor = vbWhite
            If Windows(Index).Enabled Then .CellForeColor = vbBlack Else .CellForeColor = vbRed
            .CellFontBold = Windows(Index).Visible
            .CellFontItalic = WindowHasStyle(Windows(Index).Handle, ES_PASSWORD)
         ElseIf Not CBool(CheckForError(IsWindow(Windows(Index).Handle))) Then
            .CellBackColor = vbWhite
            .CellForeColor = vbYellow
            .CellFontBold = True
         End If
     Next Column
     .Col = CurrentColumn
   End With

EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure displays the search results containing the specified text and class.
Private Sub DisplaySearchResults(SearchText As String, SearchClass As String)
On Error GoTo ErrorTrap
Dim Index As Long
Dim WindowClass As String
Dim WindowParent As Long
Dim WindowText As String

   ReDim Matches(0 To 0) As Long

   Screen.MousePointer = vbHourglass
   With SearchResultsTable
      .Rows = 1
      For Index = LBound(Windows()) To UBound(Windows())
         If Windows(Index).Handle > 0 Then
            If LookForParentWindowsBox.Value = vbChecked Then
               WindowParent = Windows(Index).Handle
               Do Until Match(GetWindowClass(WindowParent), SearchClass) And Match(GetWindowText(WindowParent), SearchText)
                  If CheckForError(GetParent(WindowParent)) = 0 Then Exit Do
                  WindowParent = CheckForError(GetParent(WindowParent))
               Loop
               
               WindowClass = GetWindowClass(WindowParent)
               WindowText = GetWindowText(WindowParent)
            ElseIf LookForParentWindowsBox.Value = vbUnchecked Then
               WindowClass = Windows(Index).ClassName
               WindowText = Windows(Index).Text
            End If
            
            If Match(WindowClass, SearchClass) Then
               If Match(WindowText, SearchText) Then
                  If Not IsExcluded(Index) Then
                     .Rows = .Rows + 1
                     .Row = .Rows - 1
                     DisplaySearchResult Index
                     Matches(UBound(Matches())) = Index
                     ReDim Preserve Matches(LBound(Matches()) To UBound(Matches()) + 1) As Long
                  End If
               End If
            End If
         End If
      Next Index
      .Col = 0
      If .Rows > 1 Then .Row = 1 Else .Row = 0
   End With
   
EndRoutine:
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure enables or disables the specified window.
Private Sub EnableDisableWindow(Index As Long)
On Error GoTo ErrorTrap

   If RefersToWindow(Windows(Index).Handle) Then
      CheckForError EnableWindow(Windows(Index).Handle, CLng(Not Windows(Index).Enabled))
      RefreshWindow Windows(Index).Handle
      UpdateSearchResults
   Else
      DisplaySearchResult Index
   End If
   
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the command to enable or disable the selected window.
Private Sub EnableDisableWindowMenu_Click()
On Error GoTo ErrorTrap

   EnableDisableWindow Matches(SearchResultsTable.Row - 1)
   
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure searches the search results for the specified text.
Private Sub FindInSearchResults(SearchText As String, FindNext As Boolean, ByRef FoundColumn As Long, ByRef FoundRow As Long)
On Error GoTo ErrorTrap
Dim Column As Long
Dim Matched As Boolean
Dim Row As Long
Dim StartColumn As Long

   Screen.MousePointer = vbHourglass
   With SearchResultsTable
      .SetFocus
      FoundColumn = -1
      FoundRow = -1
      If FindNext Then StartColumn = .Col + 1 Else StartColumn = .Col
    
      For Row = .Row To .Rows - 1
         For Column = StartColumn To .Cols - 1
            Matched = Match(UCase$(.TextMatrix(Row, Column)), UCase$(SearchText))
            If Matched Then
               FoundColumn = Column
               FoundRow = Row
               .Col = FoundColumn
               .Row = FoundRow
               .TopRow = FoundRow
               Screen.MousePointer = vbDefault
               Exit For
            End If
         Next Column
         If Matched Then Exit For
         StartColumn = 0
      Next Row
   End With
   
EndRoutine:
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the command to toggle the exclusion state of the selected window property.
Private Sub ExcludeMenu_Click(Index As Integer)
On Error GoTo ErrorTrap

   ToggleExcludedProperty CLng(Index)
   
   DisplaySearchResults WindowTextBox.Text, WindowClassBox.Text

EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the command to request the user to specify a search text.
Private Sub FindNextMatchMenu_Click()
On Error GoTo ErrorTrap

   FindText FindNext:=True

EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the command to search the search results for the specified text.
Private Sub FindText(FindNext As Boolean)
On Error GoTo ErrorTrap
Dim FoundColumn As Long
Dim FoundRow As Long
Dim NewSearchText As String
Static SearchText As String

   If (SearchText = vbNullString) Or (Not FindNext) Then
      NewSearchText = InputBox$("Find:", , SearchText)
      If Not NewSearchText = vbNullString Then SearchText = NewSearchText
   End If
   
   If Not SearchText = vbNullString Then
      FindInSearchResults SearchText, FindNext, FoundColumn, FoundRow
      If FoundColumn < 0 And FoundRow < 0 Then
         If MsgBox("Could not find """ & SearchText & ".""" & vbCrLf & "Search again from start?", vbYesNo Or vbQuestion) = vbYes Then
            With SearchResultsTable
               .Col = 0
               If .Rows > 1 Then .Row = 1 Else .Row = 0
            End With
            FindInSearchResults SearchText, FindNext, FoundColumn, FoundRow
            If FoundColumn < 0 And FoundRow < 0 Then MsgBox "Could not find """ & SearchText & ".""", vbInformation
         End If
      End If
   End If
   
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure gives the command to request the user to specify a search text.
Private Sub FindTextMenu_Click()
On Error GoTo ErrorTrap
   
   FindText FindNext:=False

EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure flashes the specified window.
Private Sub FlashWindow(Index As Long)
On Error GoTo ErrorTrap
Dim Flash As Long
Dim WindowH As Long
Dim WindowParent As Long

   If RefersToWindow(Windows(Index).Handle) Then
      WindowH = Windows(Index).Handle
      Do
         CheckForError EnableWindow(WindowH, CLng(True))
         CheckForError ShowWindow(WindowH, SW_SHOWNA)
         CheckForError BringWindowToTop(WindowH)
         If CBool(CheckForError(IsIconic(WindowH))) Then CheckForError ShowWindow(WindowH, SW_RESTORE)
         
         WindowParent = CheckForError(GetParent(WindowH))
         If WindowParent = 0 Then Exit Do
         WindowH = WindowParent
      Loop
      
      Screen.MousePointer = vbHourglass
      CheckForError ShowWindow(Windows(Index).Handle, SW_SHOWNA)
      For Flash = 0 To 9
         CheckForError ShowWindow(Windows(Index).Handle, SW_HIDE)
         DoEvents: Sleep CLng(250)
         CheckForError ShowWindow(Windows(Index).Handle, SW_SHOWNA)
         DoEvents: Sleep CLng(250)
      Next Flash
      UpdateSearchResults
   Else
      DisplaySearchResult Index
   End If
       
EndRoutine:
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure displays the selected window's base class information.
Private Sub GetBaseClassInformation(Index As Long)
On Error GoTo ErrorTrap
Dim WindowH As Long

   If RefersToWindow(Windows(Index).Handle) Then
      MsgBox "Base class: " & GetWindowBaseClass(Windows(Index).Handle), vbOKOnly Or vbInformation, App.Title
   Else
      DisplaySearchResult Index
   End If
   
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the command to flash the selected window.
Private Sub FlashWindowMenu_Click()
On Error GoTo ErrorTrap
   
   FlashWindow Matches(SearchResultsTable.Row - 1)

EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure initializes this window and its interface elements.
Private Sub Form_Load()
On Error GoTo ErrorTrap
   
   Erase Matches()
   
   BuildInterface
   
   SearchForWindows

EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure adjusts the interface contained by this window to its new size.
Private Sub Form_Resize()
On Error Resume Next
Dim Column As Long

   With SearchResultsTable
      .Height = Me.ScaleHeight - .Top - 1
      .Width = Me.ScaleWidth - 2
      For Column = 0 To .Cols - 1
         .ColWidth(Column) = Me.Width / 4.4
      Next Column
   End With
   
   SearchButton.Left = Me.ScaleWidth - SearchButton.Width - 2
   WindowClassBox.Width = Me.ScaleWidth - WindowClassBox.Left - 2
   WindowTextBox.Width = Me.ScaleWidth - WindowTextBox.Left - 2
End Sub

'This procedure gives the command to display the selected window's base class information.
Private Sub GetBaseClassInformationMenu_Click()
On Error GoTo ErrorTrap
   
   GetBaseClassInformation Matches(SearchResultsTable.Row - 1)

EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure displays the specified window's process information.
Private Sub GetProcessInformation(Index As Long)
On Error GoTo ErrorTrap
Dim Message As String
Dim WindowProcess As WindowProcessStr

   If RefersToWindow(Windows(Index).Handle) Then
      WindowProcess = GetWindowProcess(Windows(Index).Handle)
      
      With WindowProcess
         Message = "Process: " & CStr(.ProcessH) & " - " & .ProcessPath & vbCr
         Message = Message & "Process id: " & CStr(.ProcessId) & vbCr
         Message = Message & "Thread id: " & CStr(.ThreadId) & vbCr
         Message = Message & "Module: " & CStr(.ModuleH) & " - " & .ModulePath
      End With
      
      MsgBox Message, vbInformation, App.Title & " - " & CStr(Windows(Index).Handle)
   Else
      DisplaySearchResult Index
   End If
   
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the command to display the specified window's process information.
Private Sub GetProcessInformationMenu_Click()
On Error GoTo ErrorTrap

   GetProcessInformation Matches(SearchResultsTable.Row - 1)

EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the command to open this program's help file.
Private Sub HelpMenu_Click()
On Error GoTo ErrorTrap
Dim HelpPath As String

   HelpPath = App.Path
   If Not Right$(HelpPath, 1) = "\" Then HelpPath = HelpPath & "\"
   HelpPath = HelpPath & "Winscan.hta"
   
   If Dir$(HelpPath, vbArchive Or vbNormal) = vbNullString Then
      MsgBox "Could not find the help file.", vbExclamation
   Else
      Shell "Mshta.exe """ & HelpPath & """", vbMaximizedFocus
   End If
   
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure displays information about this program.
Private Sub InformationMenu_Click()
On Error GoTo ErrorTrap
   
   MsgBox App.Comments, vbInformation

EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure compares the specified texts using the options selected by the user.
Private Function Match(ByVal CompareText As String, ByVal SearchText As String) As Boolean
On Error GoTo ErrorTrap
Dim Result As Boolean

   Result = False
   
   If CaseSensitiveBox.Value = vbUnchecked Then
      CompareText = UCase$(CompareText)
      SearchText = UCase$(SearchText)
   End If
    
   If IgnoreAmpersandsBox.Value = vbChecked Then
      CompareText = Replace(CompareText, "&", vbNullString)
      SearchText = Replace(SearchText, "&", vbNullString)
   End If
     
   If SearchText = vbNullString Then
      Result = True
   ElseIf Not SearchText = vbNullString Then
      If WholePhrasesOnlyBox.Value = vbChecked Then
         Result = (SearchText = CompareText)
      ElseIf WholePhrasesOnlyBox.Value = vbUnchecked Then
         Result = (InStr(CompareText, SearchText) > 0)
      End If
   End If

EndRoutine:
   Match = Result
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function


'This procedure closes this window.
Private Sub QuitMenu_Click()
On Error GoTo ErrorTrap
   
   Unload Me

EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure gives the command to start searching for any active windows that meet the specified search criteria.
Private Sub SearchButton_Click()
On Error GoTo ErrorTrap
   
   SearchForWindows

EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure starts searching for any active windows that meet the specified search criteria.
Private Sub SearchForWindows()
On Error GoTo ErrorTrap
   
   Screen.MousePointer = vbHourglass
   
   CheckForError , ResetSuppression:=True
   ReDim Windows(0 To 0) As WindowStr
   
   GetWindowInformation CheckForError(GetDesktopWindow())
   CheckForError EnumWindows(AddressOf HandleWindows, CLng(0))
   DisplaySearchResults WindowTextBox.Text, WindowClassBox.Text
   UpdateMenus
   
EndRoutine:
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure displays the action menu for the selected window when the user double clicks the search results.
Private Sub SearchResultsTable_DblClick()
On Error GoTo ErrorTrap
   
   If SearchResultsTable.Row > 0 Then PopupMenu WindowMainMenu

EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure shows or hides the specified window.
Private Sub ShowHideWindow(Index As Long)
On Error GoTo ErrorTrap

   If RefersToWindow(Windows(Index).Handle) Then
      If Windows(Index).Visible Then
         CheckForError ShowWindow(Windows(Index).Handle, SW_HIDE)
      ElseIf Not Windows(Index).Visible Then
         CheckForError ShowWindow(Windows(Index).Handle, SW_SHOWNA)
      End If
      
      RefreshWindow Windows(Index).Handle
      UpdateSearchResults
   Else
      DisplaySearchResult Index
   End If
  
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the command to update the action menu when a window is selected.
Private Sub SearchResultsTable_EnterCell()
On Error GoTo ErrorTrap
   
   UpdateMenus

EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the command to show or hide the selected window.
Private Sub ShowHideWindowMenu_Click()
On Error GoTo ErrorTrap
   
   ShowHideWindow Matches(SearchResultsTable.Row - 1)

EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure toggles the exclusion state of the specified window property.
Private Sub ToggleExcludedProperty(PropertyIndex As ExcludableE)
On Error GoTo ErrorTrap
Dim Index As Long
Dim OtherIndex As Long

   If PropertyIndex = ExcludeNone Then
      For Index = ExcludeMenu.LBound To ExcludeMenu.UBound
         ExcludeMenu(Index).Checked = False
      Next Index
   Else
      If (PropertyIndex Mod 2) = 0 Then OtherIndex = PropertyIndex + 1 Else OtherIndex = PropertyIndex - 1
      
      ExcludeMenu(PropertyIndex).Checked = Not ExcludeMenu(PropertyIndex).Checked
      ExcludeMenu(OtherIndex).Checked = False
   End If

EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub



'This procedure updates the action menu.
Private Sub UpdateMenus()
On Error GoTo ErrorTrap

   With SearchResultsTable
      SearchResultsMainMenu.Enabled = (.Row > 0)
      WindowMainMenu.Enabled = (.Row > 0)
      
      If .Row > 0 Then
         If Windows(Matches(.Row - 1)).Enabled Then EnableDisableWindowMenu.Caption = "&Disable" Else EnableDisableWindowMenu.Caption = "&Enable"
         If Windows(Matches(.Row - 1)).Visible Then ShowHideWindowMenu.Caption = "&Hide" Else ShowHideWindowMenu.Caption = "Sh&ow"
         
         EnableDisableWindowMenu.Caption = EnableDisableWindowMenu.Caption & " window."
         ShowHideWindowMenu.Caption = ShowHideWindowMenu.Caption & " window."
      End If
   End With

EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure updates the search results.
Private Sub UpdateSearchResults()
On Error GoTo ErrorTrap
Dim CurrentColumn As Long
Dim CurrentRow As Long
Dim Row As Long

   Screen.MousePointer = vbHourglass
   With SearchResultsTable
      CurrentColumn = .Col
      CurrentRow = .Row
      For Row = 1 To .Rows - 1
         If CBool(CheckForError(IsWindow(Windows(Matches(Row - 1)).Handle))) Then
            Windows(Matches(Row - 1)).ClassName = GetWindowClass(Windows(Matches(Row - 1)).Handle)
            Windows(Matches(Row - 1)).Enabled = CBool(CheckForError(IsWindowEnabled(Windows(Matches(Row - 1)).Handle)))
            Windows(Matches(Row - 1)).Parent = CheckForError(GetParent(Windows(Matches(Row - 1)).Handle))
            Windows(Matches(Row - 1)).Text = GetWindowText(Windows(Matches(Row - 1)).Handle)
            Windows(Matches(Row - 1)).Visible = CBool(CheckForError(IsWindowVisible(Windows(Matches(Row - 1)).Handle)))
        End If
        .Row = Row
        DisplaySearchResult Matches(Row - 1)
     Next Row
     .Col = CurrentColumn
     .Row = CurrentRow
   End With

EndRoutine:
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure resets the API error message suppression when the user opens the window menu.
Private Sub WindowMainMenu_Click()
On Error GoTo ErrorTrap
   
   CheckForError , ResetSuppression:=True

EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


