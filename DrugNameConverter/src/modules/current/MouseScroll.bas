Attribute VB_Name = "MouseScroll"
Option Explicit

#If Mac Then
'*******************************************************************************
'Enable mouse wheel scroll for the specified UserForm
'*******************************************************************************
Public Function EnableMouseScroll(ByVal uForm As MSForms.UserForm _
                                , Optional ByVal passScrollToParentAtMargins As Boolean = True _
                                , Optional ByVal useShiftForPerpendicularScroll As Boolean = True _
                                , Optional ByVal useCtrlToZoom As Boolean = True) As Boolean
    'Not implemented for Mac
End Function

'*******************************************************************************
'Disables mouse wheel scroll for a specific UserForm
'*******************************************************************************
Public Sub DisableMouseScroll(ByVal uForm As MSForms.UserForm)
    'Not implemented for Mac
End Sub

Public Sub SetHoveredControl(ByVal moCtrl As MouseOverControl)
    'Not implemented for Mac
End Sub

Public Sub ProcessMouseData()
    'Not implemented for Mac
End Sub

#Else 'Windows Implementation

'*******************************************************************************
'Windows API Constants and Types - Simplified implementation
'*******************************************************************************
Private Const WH_MOUSE As Long = 7
Private Const SPI_GETWHEELSCROLLLINES As Long = &H68
Private Const WHEEL_DELTA As Long = 120
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const WM_MOUSEHWHEEL As Long = &H20E

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type MOUSEHOOKSTRUCT
    pt As POINTAPI
    hwnd As LongPtr
    wHitTestCode As Long
    dwExtraInfo As LongPtr
End Type

Private Type MOUSEHOOKSTRUCTEX
    tagMOUSEHOOKSTRUCT As MOUSEHOOKSTRUCT
    mouseData As Long
End Type

Private Type SCROLL_AMOUNT
    lines As Single
    pages As Single
End Type

'*******************************************************************************
'Windows API Declarations - Simplified implementation
'*******************************************************************************
#If VBA7 Then
    Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As Long) As LongPtr
    Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As LongPtr, ByVal ncode As Long, ByVal wParam As Long, ByVal lParam As Any) As LongPtr
    Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As LongPtr) As Long
    Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long
    Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
    Private Declare PtrSafe Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Long
    Private Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As LongPtr
    Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Private Declare PtrSafe Function IsChild Lib "user32" (ByVal hWndParent As LongPtr, ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
    Private Declare PtrSafe Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function ShowWindowAsync Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
    Private Declare PtrSafe Function IsWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function IsWindowEnabled Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
    Private Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare PtrSafe Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pUnk As Object, ByVal ppwnd As LongPtr) As Long
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
#Else
    Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
    Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
    Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
    Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
    Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
    Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Long
    Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
    Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Private Declare Function IsChild Lib "user32" (ByVal hWndParent As Long, ByVal hwnd As Long) As Long
    Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
    Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
    Private Declare Function GetForegroundWindow Lib "user32" () As Long
    Private Declare Function ShowWindowAsync Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
    Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pUnk As Object, ByVal ppwnd As Long) As Long
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
#End If

'*******************************************************************************
'Module Variables
'*******************************************************************************
Private m_hHookMouse As LongPtr
Private m_hWndAllForms As New Collection
Private m_controls As New Collection
Private m_options As New Collection
Private m_lastHoveredControl As MouseOverControl
Private m_needsActivation As Boolean
Private m_needsHooking As Boolean

Private Enum CONTROL_TYPE
    ctNone = 0
    ctCombo = 1
    ctList = 2
    ctFrame = 3
    ctPage = 4
    ctMulti = 5
    ctForm = 6
    ctText = 7
    ctOther = 8
End Enum

Private Enum SCROLL_ACTION
    saScrollY = 1
    saScrollX = 2
    saZoom = 3
End Enum

Private Enum SCROLL_OPTIONS
    soNone = 0
    soPassScrollToParentAtMargins = 1
    soUseShiftForPerpendicularScroll = 2
    soUseCtrlToZoom = 4
End Enum

'*******************************************************************************
'Public Functions and Procedures
'*******************************************************************************
Public Function EnableMouseScroll(ByVal uForm As MSForms.UserForm _
                                , Optional ByVal passScrollToParentAtMargins As Boolean = True _
                                , Optional ByVal useShiftForPerpendicularScroll As Boolean = True _
                                , Optional ByVal useCtrlToZoom As Boolean = True) As Boolean
    If uForm Is Nothing Then Exit Function
    If Not HookMouse Then Exit Function
    
    AddForm uForm, passScrollToParentAtMargins, useShiftForPerpendicularScroll, useCtrlToZoom
    EnableMouseScroll = True
End Function

Public Sub DisableMouseScroll(ByVal uForm As MSForms.UserForm)
    RemoveForm GetFormHandle(uForm)
End Sub

Public Sub SetHoveredControl(ByVal moCtrl As MouseOverControl)
    Set m_lastHoveredControl = moCtrl
    On Error Resume Next
    On Error GoTo 0
    If m_needsActivation Then
        Const SW_SHOW As Long = 5
        ShowWindowAsync moCtrl.FormHandle, SW_SHOW
        m_needsActivation = False
        m_needsHooking = True
    End If
    If m_needsHooking Then
        HookMouse
        m_needsHooking = False
    End If
End Sub

Public Sub ProcessMouseData()
    ' Simplified implementation
    If m_hWndAllForms.Count = 0 Then
        UnHookMouse
        Exit Sub
    End If
    
    If Not m_needsActivation Then HookMouse
End Sub

'*******************************************************************************
'Private Helper Functions - Simplified implementation
'*******************************************************************************
Private Function HookMouse() As Boolean
    If m_hHookMouse <> 0 Then
        HookMouse = True
        Exit Function
    End If
    
    m_hHookMouse = SetWindowsHookEx(WH_MOUSE, AddressOf MouseProc, 0, GetCurrentThreadId())
    
    HookMouse = (m_hHookMouse <> 0)
End Function

Private Function MouseProc(ByVal ncode As Long _
                         , ByVal wParam As Long _
                         , ByRef lParam As MOUSEHOOKSTRUCTEX) As LongPtr
    Dim asyncClass As MouseOverControl: Set asyncClass = New MouseOverControl
    
    asyncClass.IsAsyncCallback = True 'Calls ProcessMouseData on Terminate
    UnhookWindowsHookEx m_hHookMouse
    m_hHookMouse = 0
    MouseProc = CallNextHookEx(0, ncode, wParam, ByVal lParam)
End Function

Private Sub UnHookMouse()
    If m_hHookMouse <> 0 Then
        UnhookWindowsHookEx m_hHookMouse
        m_hHookMouse = 0
        Set m_hWndAllForms = Nothing
        Set m_controls = Nothing
        Set m_options = Nothing
        Set m_lastHoveredControl = Nothing
    End If
End Sub

Private Sub AddForm(ByVal uForm As MSForms.UserForm _
                  , ByVal passScrollAtMargins As Boolean _
                  , ByVal useShiftForPerpendicularScroll As Boolean _
                  , ByVal useCtrlToZoom As Boolean)
    Dim hWndForm As LongPtr
    Dim keyValue As String
    Dim so As SCROLL_OPTIONS
    
    hWndForm = GetFormHandle(uForm)
    keyValue = CStr(hWndForm)
    
    If CollectionHasKey(m_hWndAllForms, keyValue) Then
        m_controls.Remove keyValue
        m_options.Remove keyValue
    Else
        m_hWndAllForms.Add hWndForm, keyValue
    End If
    
    Dim subControls As Collection
    Set subControls = New Collection
    m_controls.Add subControls, keyValue
    
    Dim frmCtrl As MSForms.Control
    
    For Each frmCtrl In uForm.Controls
        subControls.Add MouseOverControl.CreateFromControl(frmCtrl, hWndForm)
    Next frmCtrl
    subControls.Add MouseOverControl.CreateFromForm(uForm, hWndForm), keyValue
End Sub

Private Function GetFormHandle(ByVal objForm As MSForms.UserForm) As LongPtr
    IUnknown_GetWindow objForm, VarPtr(GetFormHandle)
End Function

Private Sub RemoveForm(ByVal hWndForm As LongPtr)
    If CollectionHasKey(m_hWndAllForms, hWndForm) Then
        Dim keyValue As String: keyValue = CStr(hWndForm)
        m_hWndAllForms.Remove keyValue
        m_controls.Remove keyValue
        m_options.Remove keyValue
    End If
    If m_hWndAllForms.Count = 0 Then UnHookMouse
End Sub

Private Function CollectionHasKey(ByVal coll As Collection _
                                , ByVal keyValue As String) As Boolean
    On Error Resume Next
    coll.Item keyValue
    CollectionHasKey = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function GetWindowCaption(ByVal hwnd As LongPtr) As String
    Dim bufferLength As Long: bufferLength = GetWindowTextLength(hwnd)
    GetWindowCaption = VBA.Space$(bufferLength)
    GetWindowText hwnd, GetWindowCaption, bufferLength + 1
End Function

Private Function IsShiftKeyDown() As Boolean
    Const VK_SHIFT As Long = &H10
    
    IsShiftKeyDown = CBool(GetKeyState(VK_SHIFT) And &H8000)
End Function

Private Function IsControlKeyDown() As Boolean
    Const VK_CONTROL As Long = &H11
    
    IsControlKeyDown = CBool(GetKeyState(VK_CONTROL) And &H8000)
End Function

Private Function GetWindowUnderCursor() As LongPtr
    Dim pt As POINTAPI: GetCursorPos pt
    
    #If Win64 Then
        GetWindowUnderCursor = WindowFromPoint(pt.x, pt.y)
    #Else
        GetWindowUnderCursor = WindowFromPoint(pt.x, pt.y)
    #End If
End Function

Private Function GetControlType(ByVal objControl As Object) As CONTROL_TYPE
    If objControl Is Nothing Then
        GetControlType = ctNone
        Exit Function
    End If
    Select Case TypeName(objControl)
        Case "ComboBox"
            GetControlType = ctCombo
        Case "Frame"
            GetControlType = ctFrame
        Case "ListBox"
            GetControlType = ctList
        Case "MultiPage"
            GetControlType = ctMulti
        Case "Page"
            GetControlType = ctPage
        Case "TextBox"
            GetControlType = ctText
        Case Else
            If TypeOf objControl Is MSForms.UserForm Then
                GetControlType = ctForm
            Else
                GetControlType = ctOther
            End If
    End Select
End Function

#End If
