VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4
Private Const NIM_ADD = &H0
Private Const NIM_DELETE = &H2

Private Const WM_KEYDOWN = &H100
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206

Private Type NOTIFYICONDATA
    cbSize           As Long
    hwnd             As Long
    uID              As Long
    uFlags           As Long
    uCallbackMessage As Long
    hIcon            As Long
    szTip            As String * 64
End Type

Public Event MouseMove()
Public Event LButtonDown()
Public Event LButtonUp()
Public Event LButtonDblClk()
Public Event RButtonDown()
Public Event RButtonUp()
Public Event RButtonDblClk()

Dim NID As NOTIFYICONDATA, Created As Boolean
Public Function Create(ByVal hWindow As Long, ByVal hIcon As Long, Optional ByVal ToolTip As String) As Boolean
    If hWindow <> 0 And hIcon <> 0 Then
        If Created Then Remove
        With NID
            .hwnd = hWindow
            .uID = hIcon
            .uFlags = NIF_ICON Or NIF_MESSAGE
            If ToolTip <> "" Then
                .uFlags = .uFlags Or NIF_TIP
                .szTip = ToolTip & vbNullChar
            End If
            .uCallbackMessage = WM_RBUTTONDOWN
            .hIcon = hIcon
            .cbSize = Len(NID)
        End With
        If Shell_NotifyIcon(NIM_ADD, NID) Then
            Create = True
            Created = True
        End If
    End If
End Function
Public Sub HandleEvent(ByVal X As Single)
    If Created Then
        Select Case X / Screen.TwipsPerPixelX
        Case WM_MOUSEMOVE: RaiseEvent MouseMove
        Case WM_LBUTTONUP: RaiseEvent LButtonUp
        Case WM_LBUTTONDOWN: RaiseEvent LButtonDown
        Case WM_LBUTTONDBLCLK: RaiseEvent LButtonDblClk
        Case WM_RBUTTONDOWN: RaiseEvent RButtonDown
        Case WM_RBUTTONUP: RaiseEvent RButtonUp
        Case WM_RBUTTONDBLCLK: RaiseEvent RButtonDblClk
        Case Else
            Debug.Print Hex(X / Screen.TwipsPerPixelX)
        End Select
    End If
End Sub
Public Function Remove() As Boolean
    If Created Then
        Remove = Shell_NotifyIcon(NIM_DELETE, NID)
        Created = False
    End If
End Function

'Class Event

Private Sub Class_Terminate()
    Remove
End Sub
