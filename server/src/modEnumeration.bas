Attribute VB_Name = "modEnumeration"

' Packets sent by authentication server to game server
Public Enum AuthPackets
    ASetPlayerLoginToken
    ASetUsergroup
End Enum

' Packets sent by server to client
Public Enum ServerPackets
    SAlertMsg = 1
    SSetPlayerLoginToken
    ' Make sure SMSG_COUNT is below everything else
    SMSG_COUNT
End Enum

' Packets sent by client to server
Public Enum ClientPackets
    CNewAccount = 1
    CAuthLogin
    CForgotPassword
    ' Make sure CMSG_COUNT is below everything else
    CMSG_COUNT
End Enum

Public Enum User
    Free = 1
    Vip1
    Vip2
    Vip3
    Vip4
    Vip5
    Tutor
    Banned
End Enum
