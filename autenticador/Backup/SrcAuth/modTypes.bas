Attribute VB_Name = "modTypes"
Option Explicit

Public Const MAX_INV As Long = 35
Public Const MAX_BANK As Long = 100

Public Inv(1 To MAX_PLAYERS) As InvRec

Public Bank(1 To MAX_PLAYERS) As BankRec


Public Type PlayerInvRec
    Num As Long
    value As Long
    Bound As Byte
End Type

Public Type InvRec
    Login As String * ACCOUNT_LENGTH
    Item(1 To MAX_INV) As PlayerInvRec
End Type

Public Type BankRec
    Login As String * ACCOUNT_LENGTH
    Item(1 To MAX_BANK) As PlayerInvRec
End Type

Public Type PlayerSpellRec
    Spell As Long
    Uses As Long
End Type

Public Type HotbarRec
    Slot As Long
    sType As Byte
End Type

Public Type EquipmentRec
    Num As Long
    Bound As Byte
End Type

Public Type PlayerQuestRec
    Status As Long
    ActualTask As Long
    CurrentCount As Long    'Used to handle the Amount property
End Type

Public Type PlayerRec
    ' Saved local vars
    Login As String * ACCOUNT_LENGTH
    Password As String * NAME_LENGTH
    mail As String * EMAIL_LENGTH

    ' char
    Name As String * ACCOUNT_LENGTH
    Sex As Byte
    Class As Byte
    Sprite As Integer
    Level As Integer
    EXP As Long
    Access As Byte
    PK As Byte

    ' Vitals
    Vital(1 To Vitals.Vital_Count - 1) As Long

    ' Stats
    Stat(1 To Stats.Stat_Count - 1) As Long
    POINTS As Long

    ' Worn equipment
    Equipment(1 To Equipment.Equipment_Count - 1) As EquipmentRec

    Spell(1 To MAX_PLAYER_SPELLS) As PlayerSpellRec

    ' Hotbar
    Hotbar(1 To MAX_HOTBAR) As HotbarRec

    ' Position
    Map As Integer
    x As Byte
    Y As Byte
    dir As Byte

    ' Variables
    Variable(1 To MAX_BYTE) As Long

    ' Tutorial
    TutorialState As Byte

    ' Banned
    isBanned As Byte
    isMuted As Byte

    ' Guild
    Guild_MembroID As Byte
    Guild_ID As Integer

    ' Spell CD Save in Playerrec
    SpellCD(1 To MAX_PLAYER_SPELLS) As Integer

    ' Give one the serial number
    Serial(1 To MAX_SERIAL_NUMBER) As Integer

    ' Premium
    Premium As Byte
    StartPremium As String
    DaysPremium As Long

    ' Quests
    PlayerQuest(1 To MAX_QUESTS) As PlayerQuestRec
    
    ' Protect Items
    ProtectDrop As Byte
End Type
