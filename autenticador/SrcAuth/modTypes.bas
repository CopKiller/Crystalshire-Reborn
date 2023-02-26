Attribute VB_Name = "modTypes"
Option Explicit

Public Class() As ClassRec

Private Type ClassRec
    Name As String * NAME_LENGTH
    Stat(1 To Stats.Stat_Count - 1) As Byte
    MaleSprite() As Long
    FemaleSprite() As Long
    
    MaxHP As Long
    MaxMP As Long
    
    START_MAP As Integer
    START_X As Integer
    START_Y As Integer

    startItemCount As Long
    StartItem() As Long
    StartValue() As Long

    startSpellCount As Long
    StartSpell() As Long
End Type

Public Type PlayerInvRec
    Num As Integer
    value As Long
    Bound As Byte
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
    Num As Integer
    Bound As Byte
End Type

Private Type TaskTimerRec
    Active As Byte            ' Is Active?
    TimerType As Byte         ' 0=Days; 1=Hours; 2=Minutes; 3=Seconds.
    Timer As Currency             ' Time with /\

    Teleport As Byte          ' Teleport cannot end task in time.
    MapNum As Integer         ' Map Number to teleport /\
    ResetType As Byte         ' 0=Resetar Task ; 1=Resetar Quest.
    x As Byte
    Y As Byte
    
    Msg As String * TASK_DEFEAT_LENGTH
End Type

Public Type PlayerQuestRec
    Status As Byte
    ActualTask As Byte
    CurrentCount As Long 'Used to handle the Amount property
    Data As String * 19 ' Salva o now que tem 19 dígitos, pra usar como comparação na hora de iniciar novamente a quest
    
    TaskTimer As TaskTimerRec
End Type

Private Type SerialOptionRec
    Hash As Long
    YearBirthday As Integer
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
    
    ' Bolsa
    Inv(1 To MAX_INV) As PlayerInvRec
    
    ' Banco
    Bank(1 To MAX_BANK) As PlayerInvRec

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
    Serial(1 To MAX_SERIAL_NUMBER) As SerialOptionRec
    
    ' Premium
    Premium As Byte
    StartPremium As String * DATA_LENGTH
    DaysPremium As Long
    
    ' Quests
    PlayerQuest(1 To MAX_QUESTS) As PlayerQuestRec
    
    ' Protect Items
    ProtectDrop As Byte
    
    ' BirthDay date
    BirthDay As Date '00/00/0000'
    
    ' CheckIn Diary
    CheckIn As Byte
    LastCheckIn As Date '00/00/0000'
    
    ' Gold
    Gold As Long
End Type

