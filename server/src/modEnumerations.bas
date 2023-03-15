Attribute VB_Name = "modEnumerations"
Option Explicit

' The order of the packets must match with the client's packet enumeration

' Packets sent by server to client
Public Enum ServerPackets
    SAlertMsg = 1
    SSetPlayerLoginToken
    SLoginOk
    SNewCharClasses
    SClassesData
    SInGame
    SPlayerInv
    SPlayerInvUpdate
    SPlayerWornEq
    SPlayerHp
    SPlayerMp
    SSendMapHpMp
    SPlayerStats
    SPlayerData
    SPlayerMove
    SNpcMove
    SPlayerDir
    SNpcDir
    SPlayerXY
    SPlayerXYMap
    SAttack
    SNpcAttack
    SCheckForMap
    SMapData
    SMapItemData
    SMapNpcData
    SMapDone
    SGlobalMsg
    SAdminMsg
    SPlayerMsg
    SMapMsg
    SSpawnItem
    SItemEditor
    SUpdateItem
    SREditor
    SSpawnNpc
    SNpcDead
    SNpcEditor
    SUpdateNpc
    SMapKey
    SEditMap
    SShopEditor
    SUpdateShop
    SSpellEditor
    SUpdateSpell
    SSpells
    SLeft
    SResourceCache
    SResourceEditor
    SUpdateResource
    SSendPing
    SDoorAnimation
    SActionMsg
    SPlayerEXP
    SBlood
    SAnimationEditor
    SUpdateAnimation
    SAnimation
    SMapNpcVitals
    SCooldown
    SClearSpellBuffer
    SSayMsg
    SOpenShop
    SResetShopAction
    SStunned
    SMapWornEq
    STrade
    SCloseTrade
    STradeUpdate
    STradeStatus
    STarget
    SHotbar
    SHighIndex
    SSound
    STradeRequest
    SPartyInvite
    SPartyUpdate
    SPartyVitals
    SChatUpdate
    SConvEditor
    SUpdateConv
    SStartTutorial
    SChatBubble
    SCancelAnimation
    SPlayerVariables
    SEvent
    SBank
    SPlayerBankUpdate
    SGuildWindow
    SUpdateGuild
    SGuildInvite
    SSerialWindow
    SUpdateSerial
    SSerialEditor
    SPlayerDPremium
    SPremiumEditor
    SQuestEditor
    SUpdateQuest
    SPlayerQuest
    SQuestMessage
    SQuestCancel
    SStatus
    SClientTime
    SMessage
    SConjuntoEditor
    SUpdateConjunto
    SUpdateConjuntoWindow
    SSendDayReward
    SCheckItemCRC
    SCheckNpcCRC
    SLotteryWindow
    SGoldUpdate
    SLotteryInfo
    SEventMsg
    ' Make sure SMSG_COUNT is below everything else
    SMSG_COUNT
End Enum

' Packets sent by client to server
Public Enum ClientPackets
    CNewAccount = 1
    CAuthLogin
    CAuthAddChar
    CLogin
    CSayMsg
    CEmoteMsg
    CBroadcastMsg
    CGuildMsg
    CPartyMsg
    CPlayerMsg
    CPlayerMove
    CPlayerDir
    CUseItem
    CAttack
    CUseStatPoint
    CPlayerInfoRequest
    CWarpMeTo
    CWarpToMe
    CWarpTo
    CSetSprite
    CGetStats
    CRequestNewMap
    CMapData
    CNeedMap
    CMapGetItem
    CMapDropItem
    CMapRespawn
    CMapReport
    CKickPlayer
    CBanPlayer
    CRequestEditMap
    CRequestEditItem
    CSaveItem
    CRequestEditNpc
    CSaveNpc
    CRequestEditShop
    CSaveShop
    CRequestEditSpell
    CSaveSpell
    CSetAccess
    CWhosOnline
    CSetMotd
    CTarget
    CSpells
    CCast
    CQuit
    CSwapInvSlots
    CRequestEditResource
    CSaveResource
    CCheckPing
    CUnequip
    CRequestPlayerData
    CRequestItems
    CRequestNPCS
    CRequestResources
    CSpawnItem
    CRequestEditAnimation
    CSaveAnimation
    CRequestAnimations
    CRequestSpells
    CRequestShops
    CRequestLevelUp
    CForgetSpell
    CCloseShop
    CBuyItem
    CSellItem
    CChangeBankSlots
    CDepositItem
    CWithdrawItem
    CCloseBank
    CAdminWarp
    CTradeRequest
    CAcceptTrade
    CDeclineTrade
    CTradeItem
    CTradeGold
    CUntradeItem
    CHotbarChange
    CHotbarUse
    CSwapSpellSlots
    CAcceptTradeRequest
    CDeclineTradeRequest
    CPartyRequest
    CAcceptParty
    CDeclineParty
    CPartyLeave
    CChatOption
    CRequestEditConv
    CSaveConv
    CRequestConvs
    CFinishTutorial
    CCriarGuild
    CGuildInvite
    CGuildInviteResposta
    CSaveGuild
    CGuildKick
    CGuildDestroy
    CLeaveGuild
    CGuildPromote
    CRequestEditSerial
    CSaveSerial
    CRequestSerial
    CSendSerial
    CRequestEditPremium
    CChangePremium
    CRemovePremium
    CRequestEditQuest
    CSaveQuest
    CRequestQuests
    CPlayerHandleQuest
    CQuestLogUpdate
    CStatus
    CRequestEditConjunto
    CSaveConjunto
    CRequestConjuntos
    CCheckIn
    CSendBet
    ' Make sure CMSG_COUNT is below everything else
    CMSG_COUNT
End Enum

Public HandleDataSub(CMSG_COUNT) As Long

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

' Layers in a map
Public Enum MapLayer
    Ground = 1
    Mask
    Mask2
    Fringe
    Fringe2
    ' Make sure Layer_Count is below everything else
    Layer_Count
End Enum

' Sound entities
Public Enum SoundEntity
    seAnimation = 1
    seItem
    seNpc
    seResource
    seSpell
    ' Make sure SoundEntity_Count is below everything else
    SoundEntity_Count
End Enum

Public Enum Status
    Important = 1    ' !
    Question      ' ?
    Music         ' (6
    Love          ' <3
    Angry         ' Bravo
    Exhausted     ' Exausto
    Confused      ' Confuso
    Typing        ' Digitando
    Idea          ' Solução
    Afk           ' Inativo
    Flashed       ' Cegado

    Status_Count
End Enum

' Event Types
Public Enum EventType
    ' Message
    evAddText = 1
    evShowText
    evShowChatBubble
    evShowChoices
    evInputNumber
    ' Game Progression
    evPlayerVar
    evEventSwitch
    ' Flow Control
    evIfElse
    evExitProcess
    ' Player
    evChangeGold
    evChangeItems
    evChangeHP
    evChangeMP
    evChangeEXP
    evChangeLevel
    evChangeSkills
    evChangeClass
    evChangeSprite
    evChangeSex
    ' Movement
    evWarpPlayer
    evScrollMap
    ' Character
    evShowAnimation
    evShowEmoticon
    ' Screen Controls
    evFadeout
    evFadein
    evTintScreen
    evFlashScreen
    evShakeScreen
    ' Music and Sounds
    evPlayBGM
    evFadeoutBGM
    evPlayBGS
    evFadeoutBGS
    evPlaySound
    evStopSound
End Enum
