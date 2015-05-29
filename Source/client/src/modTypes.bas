Attribute VB_Name = "modTypes"
Option Explicit

' Public data structures
Public Map As MapRec
Public Bank As BankRec
Public TempTile() As TempTileRec
Public Player() As PlayerRec
Public Class() As ClassRec
Public Item() As ItemRec
Public Npc() As NpcRec
Public MapItem(1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAP_NPCS) As MapNpcRec
Public Shop() As ShopRec
Public spell() As SpellRec
Public Resource() As ResourceRec
Public Animation() As AnimationRec
Public Switches(1 To MAX_SWITCHES) As String
Public Variables(1 To MAX_VARIABLES) As String
Public MapSounds() As MapSoundRec
Public MapSoundCount As Long
Public WeatherParticle(1 To MAX_WEATHER_PARTICLES) As WeatherParticleRec
Public Autotile() As AutotileRec
Public HouseConfig() As HouseRec
Public MapZones() As ZoneRec
Public ZoneNPC() As ZoneNPCRec
Public House() As HouseRec
Public Projectiles(1 To MAX_PROJECTILES) As ProjectileRec
Public MapProjectiles(1 To MAX_PROJECTILES) As MapProjectileRec

Public Pictures(1 To 10) As PictureRec
Private Type PictureRec
    pic As Long
    type As Long
    XOffset As Long
    YOffset As Long
End Type

Public Bans() As BanRec
Public BanCount As Long

Private Type BanRec
    BanReason As String * 150
    BanName As String * ACCOUNT_LENGTH
    IPAddress As String * 16
    BanChar As String * NAME_LENGTH
End Type

' client-side stuff
Public ActionMsg(1 To MAX_BYTE) As ActionMsgRec
Public Blood(1 To MAX_BYTE) As BloodRec
Public AnimInstance(1 To MAX_BYTE) As AnimInstanceRec
Public Party As PartyRec

Public ServerCount As Long
Public Servers() As ServerRec

' options
Public Options As OptionsRec

Private Type ServerRec
    Game_Name As String
    SavePass As Byte
    Password As String * NAME_LENGTH
    Username As String * ACCOUNT_LENGTH
    ip As String
    port As Long
End Type

' Type recs
Private Type OptionsRec
    Music As Byte
    sound As Byte
    Debug As Byte
    Render As Byte
    DefaultServer As Byte
    HideServerList As Byte
    ClicktoWalk As Byte
    FullScreen As Byte
    GfxMode As Byte
    MenuMusic As String
End Type

Public Type PartyRec
    Leader As Long
    Member(1 To MAX_PARTY_MEMBERS) As Long
    MemberCount As Long
End Type

Public Type PlayerInvRec
    Num As Long
    Value As Long
End Type

Private Type BankRec
    Item(1 To MAX_BANK) As PlayerInvRec
End Type

Private Type SpellAnim
    Spellnum As Long
    Timer As Long
    FramePointer As Long
End Type

Private Type PlayerRec
    ' General
    Name As String
    Class As Long
    Sprite(1 To SpriteEnum.Sprite_Count - 1) As Long
    Face(1 To FaceEnum.Face_Count - 1) As Long
    Level As Byte
    Exp As Long
    Access As Byte
    PK As Byte
    Sex As Byte
    ' Vitals
    Vital(1 To Vitals.Vital_Count - 1) As Long
    MaxVital(1 To Vitals.Vital_Count - 1) As Long
    ' Stats
    stat(1 To Stats.Stat_Count - 1) As Byte
    Points As Long
    ' Worn equipment
    Equipment(1 To Equipment.Equipment_Count - 1) As Long
    ' Position
    Map As Long
    X As Byte
    Y As Byte
    dir As Byte
    ' Client use only
    XOffset As Integer
    YOffset As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    MapGetTimer As Long
    Step As Byte
    EventTimer As Long
    InHouse As Long
    PlayerQuest(1 To 250) As PlayerQuestRec
    Pet As PlayerPetRec
End Type

Private Type TileDataRec
    X As Long
    Y As Long
    Tileset As Long
End Type

Public Type ConditionalBranchRec
    Condition As Long
    Data1 As Long
    data2 As Long
    Data3 As Long
    CommandList As Long
    ElseCommandList As Long
End Type

Public Type MoveRouteRec
    Index As Long
    Data1 As Long
    data2 As Long
    Data3 As Long
    Data4 As Long
    Data5 As Long
    Data6 As Long
End Type

Public Type EventCommandRec
    Index As Long
    Text1 As String
    Text2 As String
    Text3 As String
    Text4 As String
    Text5 As String
    Data1 As Long
    data2 As Long
    Data3 As Long
    Data4 As Long
    Data5 As Long
    Data6 As Long
    ConditionalBranch As ConditionalBranchRec
    MoveRouteCount As Long
    MoveRoute() As MoveRouteRec
End Type

Public Type CommandListRec
    CommandCount As Long
    ParentList As Long
    Commands() As EventCommandRec
End Type

Public Type EventPageRec
    'These are condition variables that decide if the event even appears to the player.
    chkVariable As Long
    VariableIndex As Long
    VariableCondition As Long
    VariableCompare As Long
    chkSwitch As Long
    SwitchIndex As Long
    SwitchCompare As Long
    chkHasItem As Long
    HasItemIndex As Long
    HasItemAmount As Long
    chkSelfSwitch As Long
    SelfSwitchIndex As Long
    SelfSwitchCompare As Long
    'End Conditions
    'Handles the Event Sprite
    GraphicType As Byte
    Graphic As Long
    GraphicX As Long
    GraphicY As Long
    GraphicX2 As Long
    GraphicY2 As Long
    'Handles Movement - Move Routes to come soon.
    MoveType As Byte
    MoveSpeed As Byte
    MoveFreq As Byte
    MoveRouteCount As Long
    MoveRoute() As MoveRouteRec
    IgnoreMoveRoute As Long
    RepeatMoveRoute As Long
    'Guidelines for the event
    WalkAnim As Byte
    DirFix As Byte
    WalkThrough As Byte
    ShowName As Byte
    'Trigger for the event
    Trigger As Byte
    'Commands for the event
    CommandListCount As Long
    CommandList() As CommandListRec
    Position As Byte
    questnum As Integer
    'Client Needed Only
    X As Long
    Y As Long
End Type

Public Type EventRec
    Name As String
    Global As Long
    pageCount As Long
    Pages() As EventPageRec
    X As Long
    Y As Long
End Type

Public Type TileRec
    Layer(1 To MapLayer.Layer_Count - 1) As TileDataRec
    Autotile(1 To MapLayer.Layer_Count - 1) As Byte
    type As Byte
    Data1 As Long
    data2 As Long
    Data3 As Long
    Data4 As String
    DirBlock As Byte
End Type

Public Type ExTileRec
    Layer(1 To ExMapLayer.Layer_Count - 1) As TileDataRec
    Autotile(1 To ExMapLayer.Layer_Count - 1) As Byte
End Type

Private Type MapEventRec
    Name As String
    dir As Long
    X As Long
    Y As Long
    GraphicType As Long
    GraphicX As Long
    GraphicY As Long
    GraphicX2 As Long
    GraphicY2 As Long
    GraphicNum As Long
    Moving As Long
    MovementSpeed As Long
    Position As Long
    XOffset As Long
    YOffset As Long
    Step As Long
    Visible As Long
    WalkAnim As Long
    DirFix As Long
    ShowDir As Long
    WalkThrough As Long
    ShowName As Long
    questnum As Long
End Type

Private Type MapRec
    Name As String * NAME_LENGTH
    Music As String * MUSIC_LENGTH
    BGS As String * MUSIC_LENGTH
    Revision As Long
    Moral As Byte
    Up As Long
    Down As Long
    Left As Long
    Right As Long
    BootMap As Long
    BootX As Byte
    BootY As Byte
    Weather As Long
    WeatherIntensity As Long
    Fog As Long
    FogSpeed As Long
    FogOpacity As Long
    Red As Long
    Green As Long
    Blue As Long
    Alpha As Long
    MaxX As Byte
    MaxY As Byte
    Tile() As TileRec
    exTile() As ExTileRec
    Npc(1 To MAX_MAP_NPCS) As Long
    NpcSpawnType(1 To MAX_MAP_NPCS) As Long
    EventCount As Long
    Events() As EventRec
    'Client Side Only -- Temporary
    CurrentEvents As Long
    MapEvents() As MapEventRec
End Type

Public CopyEvent As EventRec
Public CopyEventPage As EventPageRec

Private Type FacePartsRec
    FHeads() As Long
    FHair() As Long
    FEyes() As Long
    FEyebrows() As Long
    FEars() As Long
    FMouth() As Long
    FNose() As Long
    FCloth() As Long
    FEtc() As Long
    FFace() As Long
End Type

Private Type ClassRec
    Name As String * NAME_LENGTH
    stat(1 To Stats.Stat_Count - 1) As Byte
    ' For client use
    Vital(1 To Vitals.Vital_Count - 1) As Long
    MaleFaceParts As FacePartsRec
    FemaleFaceParts As FacePartsRec
End Type

Private Type ItemRec
    Name As String * NAME_LENGTH
    Desc As String * 255
    sound As String * NAME_LENGTH
    pic As Long
    type As Byte
    Data1 As Long
    data2 As Long
    Data3 As Long
    classReq As Long
    AccessReq As Long
    LevelReq As Long
    Mastery As Byte
    Price As Long
    Add_Stat(1 To Stats.Stat_Count - 1) As Byte
    Rarity As Byte
    speed As Long
    Handed As Long
    BindType As Byte
    Stat_Req(1 To Stats.Stat_Count - 1) As Byte
    Animation As Long
    Paperdoll As Long
    AddHP As Long
    AddMP As Long
    AddEXP As Long
    CastSpell As Long
    instaCast As Byte
    Stackable As Byte
    FurnitureWidth As Long
    FurnitureHeight As Long
    FurnitureBlocks(0 To 3, 0 To 3) As Long
    FurnitureFringe(0 To 3, 0 To 3) As Long
End Type

Private Type MapItemRec
    playerName As String
    Num As Long
    Value As Long
    Frame As Byte
    X As Byte
    Y As Byte
End Type

Private Type NpcRec
    Name As String * NAME_LENGTH
    AttackSay As String * 100
    sound As String * NAME_LENGTH
    Sprite As Long
    SpawnSecs As Long
    Behaviour As Byte
    Range As Byte
    DropChance As Long
    DropItem As Long
    DropItemValue As Long
    stat(1 To Stats.Stat_Count - 1) As Byte
    HP As Long
    Exp As Long
    Animation As Long
    Damage As Long
    Level As Long
    DropChances(1 To MAX_NPC_DROPS) As Byte
    DropItems(1 To MAX_NPC_DROPS) As Byte
    DropItemValues(1 To MAX_NPC_DROPS) As Integer
    ItemBehaviour As Byte
    Projectile As Long
End Type

Private Type MapNpcRec
    Num As Long
    target As Long
    targetType As Byte
    Vital(1 To Vitals.Vital_Count - 1) As Long
    Map As Long
    X As Byte
    Y As Byte
    dir As Byte
    ' Client use only
    XOffset As Long
    YOffset As Long
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    Step As Byte
End Type

Private Type TradeItemRec
    Item As Long
    ItemValue As Long
    CostItem As Long
    CostValue As Long
End Type

Private Type ShopRec
    Name As String * NAME_LENGTH
    BuyRate As Long
    TradeItem(1 To MAX_TRADES) As TradeItemRec
End Type

Private Type SpellRec
    Name As String * NAME_LENGTH
    Desc As String * 255
    sound As String * NAME_LENGTH
    type As Byte
    MPCost As Long
    LevelReq As Long
    AccessReq As Long
    classReq As Long
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
    Pet As Long
    IsProjectile As Boolean
    Projectile As Long
End Type

Private Type TempTileRec
    DoorOpen As Byte
    DoorFrame As Byte
    DoorTimer As Long
    DoorAnimate As Byte ' 0 = nothing| 1 = opening | 2 = closing
End Type

Public Type MapResourceRec
    X As Long
    Y As Long
    ResourceState As Byte
End Type

Public Type Vector
    X As Long
    Y As Long
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
    Health As Long
    RespawnTime As Long
    WalkThrough As Boolean
    Animation As Long
End Type

Private Type ActionMsgRec
    Message As String
    Created As Long
    type As Long
    color As Long
    Scroll As Long
    X As Long
    Y As Long
    Timer As Long
End Type

Private Type BloodRec
    Sprite As Long
    Timer As Long
    X As Long
    Y As Long
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
    LockZone As Long
    ' timing
    Timer(0 To 1) As Long
    ' rendering check
    Used(0 To 1) As Boolean
    ' counting the loop
    LoopIndex(0 To 1) As Long
    frameIndex(0 To 1) As Long
End Type

Public Type HotbarRec
    slot As Long
    sType As Byte
End Type

Public Type ButtonRec
    filename As String
    state As Byte
End Type

Public Type EventListRec
    CommandList As Long
    CommandNum As Long
End Type

Public Type MapSoundRec
    X As Long
    Y As Long
    SoundHandle As Long
    InUse As Boolean
    Channel As Long
End Type

Public Type WeatherParticleRec
    type As Long
    X As Long
    Y As Long
    Velocity As Long
    InUse As Long
End Type

'Auto tiles "/
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
    ExLayer(1 To ExMapLayer.Layer_Count - 1) As QuarterTileRec
End Type

Public Type ChatBubbleRec
    Msg As String
    colour As Long
    target As Long
    targetType As Byte
    Timer As Long
    active As Boolean
End Type

Private Type HouseRec
    ConfigName As String
    BaseMap As Long
    X As Long
    Y As Long
    MaxFurniture As Long
    Price As Long
End Type

'For Being in a house or someone elses house...
Public FurnitureCount As Long
Public FurnitureHouse As Long
Private Type FurnitureRec
    ItemNum As Long
    X As Long
    Y As Long
End Type
Public Furniture() As FurnitureRec

Public Mail() As MailRec
Public MailCount As Long
Public InboxListScroll As Long
Public Type MailRec
    Index As Long
    From As String
    Unread As Long
    Body As String
    ItemNum As Long
    ItemVal As Long
    Date As String
End Type

Public Type ZoneRec
    Name As String * NAME_LENGTH
    Maps() As Long
    MapCount As Long
    NPCs(1 To MAX_MAP_NPCS * 2) As Long
    Weather(1 To 5) As Byte
    WeatherIntensity As Byte
End Type

Public Type ZoneNPCRec
    Npc(1 To MAX_MAP_NPCS * 2) As MapNpcRec
End Type

Public Type ProjectileRec
    Name As String * NAME_LENGTH
    Sprite As Long
    Range As Byte
    speed As Long
    Damage As Long
End Type

Public Type MapProjectileRec
    ProjectileNum As Long
    Owner As Long
    OwnerType As Byte
    X As Long
    Y As Long
    dir As Byte
    Range As Long
    TravelTime As Long
    Timer As Long
End Type
