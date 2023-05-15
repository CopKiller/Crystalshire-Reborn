Attribute VB_Name = "modTypes"
Option Explicit

' Public data structures
Public Map As MapRec

Public Bank As BankRec
Public TempTile() As TempTileRec
Public Player(1 To MAX_PLAYERS) As PlayerRec
Public Class() As ClassRec
Public Item(1 To MAX_ITEMS) As ItemRec
Public NPC(1 To MAX_NPCS) As NpcRec
Public MapItem(1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAP_NPCS) As MapNpcRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Spell(1 To MAX_SPELLS) As SpellRec
Public Resource(1 To MAX_RESOURCES) As ResourceRec
Public Animation(1 To MAX_ANIMATIONS) As AnimationRec
Public Conv(1 To MAX_CONVS) As ConvWrapperRec
Public ActionMsg(1 To MAX_BYTE) As ActionMsgRec
Public Blood(1 To MAX_BYTE) As BloodRec
Public AnimInstance(1 To MAX_BYTE) As AnimInstanceRec
Public Party As PartyRec
Public Autotile() As AutotileRec
Public Options As OptionsRec

Private Type OptionsRec
    Music As Byte
    sound As Byte
    NoAuto As Byte
    Render As Byte
    Username As String
    Password As String
    SaveUser As Long
    SavePass As Long
    channelState(0 To Channel_Count - 1) As Byte
    PlayIntro As Byte
    Resolution As Byte
    Fullscreen As Byte
    ' ChangeControls
    Correr As Byte
    Atacar As Byte
    Hotbar(1 To MAX_HOTBAR) As Byte
    Bolsa As Byte
    Magias As Byte
    Quests As Byte
    Guild As Byte
    Personagem As Byte
    PegarItem As Byte
    Chat As Byte
    Options As Byte
    Up As Byte
    Down As Byte
    Left As Byte
    Right As Byte
    UsarSetas As Byte
    Target As Byte

    ItemName As Byte
    ItemAnimation As Byte

    Reconnect As Byte

    FPSConection As Byte

    ' Non Saved in Vars
    ' Temporario pro auto reconnect !
    TmpLogin As String
    TmpPassword As String
    Debug As Byte
End Type

Public Type PartyRec
    Leader As Long
    Member(1 To MAX_PARTY_MEMBERS) As Long
    MemberCount As Long
End Type

Public Type PlayerInvRec
    Num As Long
    Value As Long
    bound As Byte
    Frame As Byte    ' Client Side Only
End Type

Public Type PlayerSpellRec
    Spell As Long
    Uses As Long
End Type

Private Type BankRec
    Item(1 To MAX_BANK) As PlayerInvRec
End Type

Public Type PlayerQuestRec
    Status As Byte
    ActualTask As Byte
    CurrentCount As Long    'Used to handle the Amount property
    data As String * 19    ' Salva o now que tem 19 dígitos, pra usar como comparação na hora de iniciar novamente a quest

    TaskTimer As TaskTimerRec
End Type

Private Type PlayerRec
    ' General
    Name As String
    Class As Long
    Sprite As Long
    Level As Byte
    EXP As Long
    Access As Byte
    PK As Byte
    ' Vitals
    Vital(1 To Vitals.Vital_Count - 1) As Long
    MaxVital(1 To Vitals.Vital_Count - 1) As Long
    ' Stats
    Stat(1 To Stats.Stat_Count - 1) As Byte
    POINTS As Long
    ' Worn equipment
    Equipment(1 To Equipment.Equipment_Count - 1) As Long
    ' Position
    Map As Long
    X As Byte
    Y As Byte
    dir As Byte
    ' Variables
    Variable(1 To MAX_BYTE) As Long
    ' Client use only
    xOffset As Integer
    yOffset As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Currency
    MapGetTimer As Currency
    step As Byte
    Anim As Long
    AnimTimer As Currency
    LastMoving As Currency

    ' Guild
    Guild_MembroID As Byte
    Guild_ID As Integer
    Guild_Icon As Byte
    Premium As Byte

    ' Quest
    PlayerQuest(1 To MAX_QUESTS) As PlayerQuestRec

    ' Balao animado
    StatusFrame As Long
    StatusNum(1 To (status_count - 1)) As StatusRec

    'Stun
    StunTimer As Currency
    StunDuration As Long

    ' Golds
    Gold As Long
End Type

Private Type EventCommandRec
Type As Byte
    text As String
    Colour As Long
    channel As Byte
    TargetType As Byte
    Target As Long
    X As Long
    Y As Long
End Type

Public Type EventPageRec
    chkPlayerVar As Byte
    chkSelfSwitch As Byte
    chkHasItem As Byte

    PlayerVarNum As Long
    SelfSwitchNum As Long
    HasItemNum As Long

    PlayerVariable As Long

    GraphicType As Byte
    Graphic As Long
    GraphicX As Long
    GraphicY As Long

    MoveType As Byte
    MoveSpeed As Byte
    MoveFreq As Byte

    WalkAnim As Byte
    StepAnim As Byte
    DirFix As Byte
    WalkThrough As Byte

    Priority As Byte
    Trigger As Byte

    CommandCount As Long
    Commands() As EventCommandRec
End Type

Public Type EventRec
    Name As String
    X As Long
    Y As Long
    pageCount As Long
    EventPage() As EventPageRec
End Type

Private Type MapDataRec
    Name As String
    Music As String
    Moral As Byte

    Up As Long
    Down As Long
    Left As Long
    Right As Long

    BootMap As Long
    BootX As Byte
    BootY As Byte

    MaxX As Byte
    MaxY As Byte

    BossNpc As Long

    Panorama As Byte

    Weather As Byte
    WeatherIntensity As Byte

    Fog As Byte
    FogSpeed As Byte
    FogOpacity As Byte

    Red As Byte
    Green As Byte
    Blue As Byte
    Alpha As Byte
    Size As Byte

    Sun As Byte

    DayNight As Byte

    NPC(1 To MAX_MAP_NPCS) As Long
End Type

Private Type TileDataRec
    X As Long
    Y As Long
    tileSet As Long
End Type

Public Type TileRec
    Layer(1 To MapLayer.Layer_Count - 1) As TileDataRec
    Autotile(1 To MapLayer.Layer_Count - 1) As Byte

Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4 As Long
    Data5 As Long
    DirBlock As Byte
End Type

Private Type MapTileRec
    EventCount As Long
    Tile() As TileRec
    Events() As EventRec
End Type

Private Type MapRec
    MapData As MapDataRec
    TileData As MapTileRec
End Type

Private Type ClassRec
    Name As String * NAME_LENGTH
    Stat(1 To Stats.Stat_Count - 1) As Byte
    MaleSprite() As Long
    FemaleSprite() As Long
    ' For client use
    Vital(1 To Vitals.Vital_Count - 1) As Long
End Type

Public Type ItemRec
    Name As String * NAME_LENGTH
    Desc As String * DESC_LENGTH
    sound As String * NAME_LENGTH

    Pic As Long

Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    ClassReq As Byte
    AccessReq As Byte
    LevelReq As Integer
    Mastery As Byte
    price As Long
    Add_Stat(1 To Stats.Stat_Count - 1) As Long
    Rarity As Byte
    Speed As Long
    Handed As Long
    BindType As Byte
    Stat_Req(1 To Stats.Stat_Count - 1) As Long
    Animation As Integer
    Paperdoll As Integer

    ' consume
    AddHP As Long
    AddMP As Long
    AddEXP As Long
    CastSpell As Long
    instaCast As Byte

    ' food
    HPorSP As Long
    FoodPerTick As Long
    FoodTickCount As Long
    FoodInterval As Currency

    ' requirements
    proficiency As Byte
    Stackable As Byte
    GiveSpellNum As Integer
    BlockChance As Byte

    'Percent Atributes
    Stat_Percent(1 To Stats.Stat_Count - 1) As Byte
    Data2_Percent As Byte

    AtributeBase As Byte
    DropDead As Byte
    DropDeadChance As Byte
End Type

Private Type MapItemRec
    playerName As String
    Num As Long
    Value As Long
    Frame As Byte
    X As Byte
    Y As Byte
    bound As Byte
    Gravity As Integer
    yOffset As Integer
    xOffset As Integer
End Type

Public Type NpcRec
    Name As String * NAME_LENGTH
    AttackSay As String * 100
    sound As String * NAME_LENGTH
    Sprite As Integer
    Behaviour As Byte
    Range As Byte
    Stat(1 To Stats.Stat_Count - 1) As Long
    Animation As Integer
    Level As Long
    Conv As Long
    ' Npc drops
    DropChance(1 To MAX_NPC_DROPS) As Double
    DropItem(1 To MAX_NPC_DROPS) As Integer
    DropItemValue(1 To MAX_NPC_DROPS) As Long
    ' Casting
    Spell(1 To MAX_NPC_SPELLS) As Long
    'spawn variavel
    SpawnSecs As Long
    RndSpawn As Byte
    SpawnSecsMin As Long
    'exp variavel
    EXP As Long
    RandExp As Byte
    Percent_5 As Byte
    Percent_10 As Byte
    Percent_20 As Byte
    Shadow As Byte
    Balao As Byte
    BlockChance As Byte
End Type

Private Type MapNpcRec
    Num As Long
    Target As Long
    TargetType As Byte
    Vital(1 To Vitals.Vital_Count - 1) As Long
    Map As Long
    X As Byte
    Y As Byte
    dir As Byte
    ' Client use only
    xOffset As Long
    yOffset As Long
    Moving As Byte
    Attacking As Byte
    AttackTimer As Currency
    step As Byte
    Anim As Long
    AnimTimer As Currency
    StatusFrame As Long
    StunDuration As Long
    Dead As Byte
End Type

Public Type TradeItemRec
    Item As Long
    ItemValue As Long
    CostItem As Long
    CostValue As Long
    Frame As Byte
End Type

Private Type ShopRec
    Name As String * NAME_LENGTH
    BuyRate As Long
    TradeItem(1 To MAX_TRADES) As TradeItemRec
End Type

Public Type SpellRec
    Name As String * NAME_LENGTH
    Desc As String * DESC_LENGTH
    sound As String * NAME_LENGTH

Type As Byte
    MPCost As Long
    LevelReq As Long
    AccessReq As Long
    ClassReq As Long
    CastTime As Long
    CDTime As Long
    Icon As Long
    Map As Long
    X As Long
    Y As Long
    dir As Byte
    Vital As Long
    Duration As Long
    Interval As Long
    Range As Byte
    IsAoE As Boolean
    AoE As Long
    CastAnim As Long
    SpellAnim As Long
    StunDuration As Long
    ' ranking
    UniqueIndex As Long
    NextRank As Long
    NextUses As Long
    CanRun As Byte
End Type

Private Type TempTileRec
    ' doors... obviously
    DoorOpen As Byte
    DoorFrame As Byte
    DoorTimer As Currency
    DoorAnimate As Byte    ' 0 = nothing| 1 = opening | 2 = closing
    ' fading appear tiles
    isFading(1 To MapLayer.Layer_Count - 1) As Boolean
    fadeAlpha(1 To MapLayer.Layer_Count - 1) As Long
    FadeTimer(1 To MapLayer.Layer_Count - 1) As Long
    FadeDir(1 To MapLayer.Layer_Count - 1) As Byte
End Type

Public Type MapResourceRec
    X As Long
    Y As Long
    ResourceState As Byte
End Type

Private Type ResourceRec
    Name As String * NAME_LENGTH
    SuccessMessage As String * NAME_LENGTH
    EmptyMessage As String * NAME_LENGTH
    sound As String * NAME_LENGTH
    ResourceType As Byte
    ResourceImage As Long
    ExhaustedImage As Long
    ItemReward As Long
    ToolRequired As Long
    health As Long
    RespawnTime As Long
    WalkThrough As Boolean
    Animation As Long
    Shadow As Byte
End Type

Private Type ActionMsgRec
    message As String
    Created As Currency

Type As Long
    Color As Long
    Scroll As Long
    X As Long
    Y As Long
    Timer As Currency
    Alpha As Long
End Type

Private Type BloodRec
    Sprite As Long
    Timer As Currency
    X As Long
    Y As Long
    Alpha As Byte
End Type

Private Type AnimationRec
    Name As String * NAME_LENGTH
    sound As String * NAME_LENGTH
    Sprite(0 To 1) As Long
    Frames(0 To 1) As Long
    LoopCount(0 To 1) As Long
    looptime(0 To 1) As Long
End Type

Private Type AnimInstanceRec
    Animation As Long
    X As Long
    Y As Long
    ' used for locking to players/npcs
    lockindex As Long
    LockType As Byte
    isCasting As Byte
    ' timing
    Timer(0 To 1) As Currency
    ' rendering check
    Used(0 To 1) As Boolean
    ' counting the loop
    LoopIndex(0 To 1) As Long
    frameIndex(0 To 1) As Long
End Type

Public Type HotbarRec
    Slot As Long
    sType As Byte
End Type

Public Type PointRec
    X As Long
    Y As Long
End Type

Public Type QuarterTileRec
    QuarterTile(1 To 4) As PointRec
    renderState As Byte
    srcX(1 To 4) As Long
    srcY(1 To 4) As Long
End Type

Public Type AutotileRec
    Layer(1 To MapLayer.Layer_Count - 1) As QuarterTileRec
End Type

Public Type ConvRec
    Conv As String
    rText(1 To 4) As String
    rTarget(1 To 4) As Long
    Event As Long
    Data1 As Long
    Data2 As Long
    Data3 As Long
End Type

Private Type ConvWrapperRec
    Name As String * NAME_LENGTH
    chatCount As Long
    Conv() As ConvRec
End Type

Public Type ChatBubbleRec
    Msg As String
    Colour As Long
    Target As Long
    TargetType As Byte
    Timer As Currency
    Active As Boolean
End Type

Public Type TextColourRec
    text As String
    Colour As Long
End Type

Public Type GeomRec
    top As Long
    Left As Long
    Height As Long
    Width As Long
End Type
