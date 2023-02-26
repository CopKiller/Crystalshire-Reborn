Attribute VB_Name = "modEnumeration"

' Packets sent by authentication server to game server
Public Enum AuthPackets
    ASetPlayerLoginToken
End Enum

' Packets sent by server to client
Public Enum ServerPackets
    SAlertMsg = 1
    SSetPlayerLoginToken
    SLoginOk
    SNewCharClasses
    ' Make sure SMSG_COUNT is below everything else
    SMSG_COUNT
End Enum

' Packets sent by client to server
Public Enum ClientPackets
    CNewAccount = 1
    CAuthLogin
    CAuthAddChar
    ' Make sure CMSG_COUNT is below everything else
    CMSG_COUNT
End Enum

' Menu
Public Enum MenuCount
    menuMain = 1
    menuLogin
    menuRegister
    menuCredits
    menuClass
    menuNewChar
End Enum

' Stats used by Players, Npcs and Classes
Public Enum Stats
    Strength = 1
    Endurance
    Intelligence
    Agility
    Willpower
    ' Make sure Stat_Count is below everything else
    Stat_Count
End Enum

' Vitals used by Players, Npcs and Classes
Public Enum Vitals
    HP = 1
    MP
    ' Make sure Vital_Count is below everything else
    Vital_Count
End Enum

' Equipment used by Players
Public Enum Equipment
    Weapon = 1
    Armor
    Helmet
    Shield
    Legs
    Boots
    Amulet
    RingLeft
    RingRight
    ' Make sure Equipment_Count is below everything else
    Equipment_Count
End Enum
