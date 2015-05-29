Attribute VB_Name = "modEnumerations"
Option Explicit

' The order of the packets must match with the server's packet enumeration

' Packets sent by server to client
Public Enum ServerPackets
    SAlertMsg = 1
    SLoginOk
    SNewCharClasses
    SClassesData
    SInGame
    SPlayerInv
    SPlayerInvUpdate
    SPlayerWornEq
    SPlayerHp
    SPlayerMp
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
    SBank
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
    SSpawnEvent
    SEventMove
    SEventDir
    SEventChat
    SEventStart
    SEventEnd
    SPlayBGM
    SPlaySound
    SFadeoutBGM
    SStopSound
    SSwitchesAndVariables
    SMapEventData
    SChatBubble
    SSpecialEffect
    SHouseConfigs
    SBuyHouse
    SMax
    SVisit
    SFurniture
    SMailBox
    SMailUnread
    SSelChar
    SFriends
    SQuestEditor
    SUpdateQuest
    SPlayerQuest
    SQuestMessage
    SZoneEdit
    SServerInfo
    SPlayers
    SAccounts
    SAdmin
    SMaps
    SBans
    SServerOpts
    SHouseEdit
    SEditPlayer
    SPic
    SHoldPlayer
    SGameOpts
    SPetEditor
    SUpdatePet
    SPetMove
    SPetDir
    SPetVital
    SClearPetSpellBuffer
    SRandomDungeonEditor
    SUpdateRandomDungeon
    SRandomDungeonMap
    SProjectileEditor
    SUpdateProjectile
    SMapProjectile
    ' Make sure SMSG_COUNT is below everything else
    SMSG_COUNT
End Enum

' Packets sent by client to server
Public Enum ClientPackets
    CNewAccount = 1
    CDelAccount
    CLogin
    CAddChar
    CUseChar
    CSayMsg
    CEmoteMsg
    CBroadcastMsg
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
    CBanList
    CBanDestroy
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
    CSearch
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
    CEventChatReply
    CEvent
    CSwitchesAndVariables
    CRequestSwitchesAndVariables
    CBuyHouse
    CVisit
    CAcceptVisit
    CPlaceFurniture
    CDeleteMail
    CSendMail
    CReadMail
    CTakeMailItem
    CRestartServer
    CNewMap
    CDelChar
    CEditFriend
    CRequestEditQuest
    CSaveQuest
    CRequestQuests
    CPlayerHandleQuest
    CQuestLogUpdate
    CRequestEditZone
    CSaveZones
    CAuthLogin
    CAuthPing
    CReAuth
    CAlertMsg
    CAuthNews
    CNotification
    CNewAcc
    CActivate
    CNewCode
    CEditAccountLogin
    CEditAccount
    CAdmin
    CServerOpts
    CSaveServerOpt
    CRequestEditHouse
    CSaveHouses
    CEditPlayer
    CSavePlayer
    CEventTouch
    CGameOpts
    CSaveGameOpt
    CMitigation
    CRequestEditPet
    CSavePet
    CRequestPets
    CPetMove
    csetbehaviour
    CReleasePet
    CPetSpell
    CPetUseStatPoint
    CRequestEditRandomDungeon
    CSaveRandomDungeon
    CRequestRandomDungeon
    CRequestEditProjectiles
    CSaveProjectile
    CRequestProjectiles
    CClearProjectile
    ' Make sure CMSG_COUNT is below everything else
    CMSG_COUNT
End Enum

Public HandleDataSub(SMSG_COUNT) As Long

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

' Layers in a map
Public Enum ExMapLayer
    Mask3 = 1
    Mask4
    Mask5
    Fringe3
    Fringe4
    Fringe5
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

Public Enum MoveRouteOpts
    MoveUp = 1
    MoveDown
    MoveLeft
    MoveRight
    MoveRandom
    MoveTowardsPlayer
    MoveAwayFromPlayer
    StepForward
    StepBack
    Wait100ms
    Wait500ms
    Wait1000ms
    TurnUp
    TurnDown
    TurnLeft
    TurnRight
    Turn90Right
    Turn90Left
    Turn180
    TurnRandom
    TurnTowardPlayer
    TurnAwayFromPlayer
    SetSpeed8xSlower
    SetSpeed4xSlower
    SetSpeed2xSlower
    SetSpeedNormal
    SetSpeed2xFaster
    SetSpeed4xFaster
    SetFreqLowest
    SetFreqLower
    SetFreqNormal
    SetFreqHigher
    SetFreqHighest
    WalkingAnimOn
    WalkingAnimOff
    DirFixOn
    DirFixOff
    WalkThroughOn
    WalkThroughOff
    PositionBelowPlayer
    PositionWithPlayer
    PositionAbovePlayer
    ChangeGraphic
End Enum

' Event Types
Public Enum EventType
    ' Message
    evAddText = 1
    evShowText
    evShowChoices
    ' Game Progression
    evPlayerVar
    evPlayerSwitch
    evSelfSwitch
    ' Flow Control
    evCondition
    evExitProcess
    ' Player
    evChangeItems
    evRestoreHP
    evRestoreMP
    evLevelUp
    evChangeLevel
    evChangeSkills
    evChangeClass
    evChangeSprite
    evChangeSex
    evChangePK
    ' Movement
    evWarpPlayer
    evSetMoveRoute
    ' Character
    evPlayAnimation
    ' Music and Sounds
    evPlayBGM
    evFadeoutBGM
    evPlaySound
    evStopSound
    'Etc...
    evCustomScript
    evSetAccess
    'Shop/Bank
    evOpenBank
    evOpenShop
    'New
    evGiveExp
    evShowChatBubble
    evLabel
    evGotoLabel
    evSpawnNpc
    evFadeIn
    evFadeOut
    evFlashWhite
    evSetFog
    evSetWeather
    evSetTint
    evWait
    evOpenMail
    evBeginQuest
    evEndQuest
    evQuestTask
    evShowPicture
    evHidePicture
    evWaitMovement
    evHoldPlayer
    evReleasePlayer
End Enum

Public Enum SpriteEnum
    Body = 1
    Shirt
    Pants
    Hair
    Shoes
    Sprite_Count
End Enum

Public Enum FaceEnum
    Hair = 1
    Head
    Eyes
    EyeBrows
    Ears
    Mouth
    Nose
    Shirt
    Etc
    Face_Count
End Enum
