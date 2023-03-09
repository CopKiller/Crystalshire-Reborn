Attribute VB_Name = "modTypes"
Option Explicit

' Public data structures
Public Map(1 To MAX_MAPS) As MapRec

Public TempTile(1 To MAX_MAPS) As TempTileRec
Public PlayersOnMap(1 To MAX_MAPS) As Long
Public Player(1 To MAX_PLAYERS) As PlayerRec
Public TempPlayer(1 To MAX_PLAYERS) As TempPlayerRec
Public Class() As ClassRec
Public Item(1 To MAX_ITEMS) As ItemRec
Public NPC(1 To MAX_NPCS) As NpcRec
Public MapItem(1 To MAX_MAPS, 1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAPS) As MapNpcDataRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Spell(1 To MAX_SPELLS) As SpellRec
Public Resource(1 To MAX_RESOURCES) As ResourceRec
Public Animation(1 To MAX_ANIMATIONS) As AnimationRec
Public Party(1 To MAX_PARTYS) As PartyRec
Public Options As OptionsRec

Public Type PlayerInvRec
    Num As Integer
    Value As Long
    Bound As Byte
End Type

Private Type OptionsRec
    MOTD As String
    PartyBonus As Integer
    START_MAP As Integer
    START_X As Byte
    START_Y As Byte
    GAME_NAME As String
    GAME_WEBSITE As String
    DAYNIGHT As Byte
    PREMIUMEXP As Integer
    PREMIUMDROP As Byte
    LOTTERYBONUS As Integer
End Type

Public Type PartyRec
    Leader As Long
    Member(1 To MAX_PARTY_MEMBERS) As Long
    MemberCount As Long
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
    X As Byte
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
    
    BirthDay As Date '00/00/0000'
    
    ' CheckIn Diary
    CheckIn As Byte
    LastCheckIn As Date '00/00/0000'
    
    ' Gold
    Gold As Long
End Type

Public Type SpellBufferRec
    Spell As Integer
    Timer As Currency
    target As Byte
    tType As Byte
End Type

Public Type DoTRec
    Used As Boolean
    Spell As Long
    Timer As Currency
    Caster As Long
    StartTime As Long
End Type

Public Type TempPlayerRec
    ' Non saved local vars
    ConnectedTime As Currency
    Buffer As clsBuffer
    InGame As Boolean
    AttackTimer As Currency
    DataTimer As Currency
    DataBytes As Long
    DataPackets As Long
    TargetType As Byte
    target As Byte
    GettingMap As Byte
    InShop As Long
    StunTimer As Currency
    StunDuration As Long
    InBank As Boolean
    inEvent As Boolean
    eventNum As Long
    pageNum As Long
    commandNum As Long
    ' trade
    TradeRequest As Long
    InTrade As Long
    TradeOffer(1 To MAX_INV) As PlayerInvRec
    TradeGold As Long
    AcceptTrade As Boolean
    ' dot/hot
    DoT(1 To MAX_DOTS) As DoTRec
    HoT(1 To MAX_DOTS) As DoTRec
    ' spell buffer
    spellBuffer As SpellBufferRec
    ' regen
    stopRegen As Boolean
    stopRegenTimer As Currency
    ' party
    inParty As Long
    partyInvite As Long
    ' chat
    inChatWith As Long
    curChat As Long
    c_mapNum As Long
    c_mapNpcNum As Long
    ' food
    foodItem(1 To Vitals.Vital_Count - 1) As Long
    foodTick(1 To Vitals.Vital_Count - 1) As Long
    foodTimer(1 To Vitals.Vital_Count - 1) As Currency

    guildInvite As Long
    
    StatusNum(1 To (Status_Count - 1)) As StatusRec
    
    AFKTimer As Currency
    
    ConjuntoID As Integer
    Bonus As BonusRec
End Type

Private Type TempEventRec
    X As Long
    Y As Long
    SelfSwitch As Byte
End Type

Private Type EventCommandRec
Type As Byte
    Text As String
    colour As Long
    Channel As Byte
    TargetType As Byte
    target As Long
End Type

Private Type EventPageRec
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

Private Type EventRec
    Name As String
    X As Long
    Y As Long
    PageCount As Long
    EventPage() As EventPageRec
End Type

Private Type MapDataRec
    Name As String
    Music As String
    Moral As Byte

    Up As Long
    Down As Long
    left As Long
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
    
    DAYNIGHT As Byte

    NPC(1 To MAX_MAP_NPCS) As Long
End Type

Private Type TileDataRec
    X As Long
    Y As Long
    Tileset As Long
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

Private Type ItemRec
    Name As String * NAME_LENGTH
    Desc As String * DESC_LENGTH
    Sound As String * NAME_LENGTH

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
    Num As Long
    Value As Long
    X As Byte
    Y As Byte
    ' ownership + despawn
    PlayerName As String
    playerTimer As Currency
    canDespawn As Boolean
    despawnTimer As Currency
    Bound As Byte
End Type

Private Type NpcRec
    Name As String * NAME_LENGTH
    AttackSay As String * 100
    Sound As String * NAME_LENGTH
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
    target As Long
    TargetType As Byte
    Vital(1 To Vitals.Vital_Count - 1) As Long
    X As Byte
    Y As Byte
    dir As Byte
    ' For server use only
    SpawnWait As Currency
    AttackTimer As Currency
    StunDuration As Long
    StunTimer As Currency
    ' regen
    stopRegen As Boolean
    stopRegenTimer As Currency
    ' dot/hot
    DoT(1 To MAX_DOTS) As DoTRec
    HoT(1 To MAX_DOTS) As DoTRec
    ' chat
    c_lastDir As Byte
    c_inChatWith As Long
    ' spell casting
    spellBuffer As SpellBufferRec
    SpellCD(1 To MAX_NPC_SPELLS) As Long
    
    ' Dead and spawn
    ActionMsgSpawn As Long
    SecondsToSpawn As Long
    tmpNum As Byte
    Dead As Byte
End Type

Private Type MapNpcDataRec
    NPC(1 To MAX_MAP_NPCS) As MapNpcRec
End Type

Private Type TradeItemRec
    Item As Long
    ItemValue As Long
    costitem As Long
    costvalue As Long
    Frame As Byte
End Type

Private Type ShopRec
    Name As String * NAME_LENGTH
    BuyRate As Long
    TradeItem(1 To MAX_TRADES) As TradeItemRec
End Type

Private Type SpellRec
    Name As String * NAME_LENGTH
    Desc As String * DESC_LENGTH
    Sound As String * NAME_LENGTH

Type As Byte
    mpCost As Long
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
    DoorOpen() As Byte
    DoorTimer As Currency
End Type

Private Type TempMapDataRec
    NPC() As MapNpcRec
End Type

Private Type ResourceRec
    Name As String * NAME_LENGTH
    SuccessMessage As String * NAME_LENGTH
    EmptyMessage As String * NAME_LENGTH
    Sound As String * NAME_LENGTH

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

Private Type AnimationRec
    Name As String * NAME_LENGTH
    Sound As String * NAME_LENGTH

    Sprite(0 To 1) As Long
    Frames(0 To 1) As Long
    LoopCount(0 To 1) As Long
    LoopTime(0 To 1) As Long
End Type

Private Type ConvRec
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
