Attribute VB_Name = "modFasterDoEvents"
Option Explicit

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type Msg
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    Time As Long
    pt As POINTAPI
End Type

Public M As Msg

Public Const WM_SYSCOMMAND As Long = &H112
Public Const WM_CLOSE As Long = &H10
Public Const WM_DESTROY As Long = &H2
Public Const PM_NOREMOVE As Long = &H0

Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long



