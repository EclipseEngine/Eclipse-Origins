Attribute VB_Name = "modTypes"
Option Explicit

' Public data structures
Public Map() As MapRec
Public TempEventMap() As GlobalEventsRec
Public MapCache() As Cache
Public temptile() As TempTileRec
Public PlayersOnMap() As Long
Public ResourceCache() As ResourceCacheRec
Public Player() As PlayerRec
Public Bank() As BankRec
Public TempPlayer() As TempPlayerRec
Public Class() As ClassRec
Public Item() As ItemRec
Public Npc() As NpcRec
Public MapItem() As MapItemRec
Public MapNpc() As MapDataRec
Public Shop() As ShopRec
Public Spell() As SpellRec
Public Resource() As ResourceRec
Public Animation() As AnimationRec
Public Party(1 To MAX_PARTYS) As PartyRec
Public Options As OptionsRec
Public Switches() As String
Public Variables() As String
Public MapBlocks() As MapBlockRec
Public HouseConfig() As HouseRec
Public MapZones() As ZoneRec
Public MonkeyPlayer As PlayerRec
Public ZoneNpc() As MapDataRec
Public account() As AccountRec
Public AccountCount As Long
Public Bans() As BanRec
Public BanCount As Long
Public NewOptions As NewOptionsRec
Public Projectiles(1 To MAX_PROJECTILES) As ProjectileRec
Public MapProjectiles() As MapProjectileRec

Private Type NewOptionsRec
    CombatMode As Long '0 for robin, 1 for new
    MaxLevel As Long
    MainMenuMusic As String
    ItemLoss As Byte
    ExpLoss As Byte
End Type

Private Type BanRec
    BanReason As String * 150
    BanName As String * ACCOUNT_LENGTH
    IPAddress As String * 16
    BanChar As String * NAME_LENGTH
End Type

Private Type AccountRec
    login As String * ACCOUNT_LENGTH
    pass As Long
    characters() As String * NAME_LENGTH
    ip As String
    access As Long
End Type


Private Type MoveRouteRec
    Index As Long
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4 As Long
    data5 As Long
    data6 As Long
End Type

Private Type GlobalEventRec
    x As Long
    y As Long
    Dir As Long
    Active As Long
    
    WalkingAnim As Long
    FixedDir As Long
    WalkThrough As Long
    ShowName As Long
    
    Position As Long
    
    GraphicType As Long
    GraphicNum As Long
    GraphicX As Long
    GraphicX2 As Long
    GraphicY As Long
    GraphicY2 As Long
    
    'Server Only Options
    MoveType As Long
    MoveSpeed As Long
    MoveFreq As Long
    MoveRouteCount As Long
    MoveRoute() As MoveRouteRec
    MoveRouteStep As Long
    
    RepeatMoveRoute As Long
    IgnoreIfCannotMove As Long
    
    MoveTimer As Long
    questnum As Long
    MoverouteComplete As Long
End Type

Public Type GlobalEventsRec
    EventCount As Long
    Events() As GlobalEventRec
End Type

Private Type OptionsRec
    Game_Name As String
    MOTD As String
    Port As Long
    Website As String
    SilentStartup As Long
    Key As String
    DataFolder As String
    UpdateURL As String
    StaffOnly As Long
    DisableRemoteRestart As Long
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

Private Type Cache
    data() As Byte
End Type

Private Type BankRec
    Item(1 To MAX_BANK) As PlayerInvRec
End Type

Public Type HotbarRec
    slot As Long
    sType As Byte
End Type

Private Type FurnitureRec
    ItemNum As Integer
    x As Integer
    y As Integer
End Type

Private Type PlayerHouseRec
    HouseIndex As Long
    FurnitureCount As Integer
    Furniture() As FurnitureRec
End Type

Public Type CharacterRec
    ' General
    Name As String * ACCOUNT_LENGTH
    Sex As Byte
    Class As Byte
    Sprite(1 To SpriteEnum.Sprite_Count - 1) As Integer
    Face(1 To FaceEnum.Face_Count - 1) As Integer
    Level As Byte
    Exp As Long
    access As Byte
    PK As Byte
    
    ' Vitals
    Vital(1 To Vitals.Vital_Count - 1) As Long
    
    ' Stats
    stat(1 To Stats.Stat_Count - 1) As Integer
    Points As Long
    
    ' Worn equipment
    Equipment(1 To Equipment.Equipment_Count - 1) As Integer
    
    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Integer
    
    ' Hotbar
    Hotbar(1 To MAX_HOTBAR) As HotbarRec
    
    ' Position
    Map As Long
    x As Byte
    y As Byte
    Dir As Byte
    
    Switches(0 To MAX_SWITCHES) As Byte
    Variables(0 To MAX_VARIABLES) As Long
    
    House As PlayerHouseRec
    
    InHouse As Long
    LastMap As Long
    LastX As Long
    LastY As Long
    
    Friends(1 To 25) As String * ACCOUNT_LENGTH
    PlayerQuest(1 To 250) As PlayerQuestRec
    Pet As PlayerPetRec
    
    'Instances
    InRandomDungeonNum As Long
    InRandomDungeonFloorNum As Long
    InInstance As Boolean
    InInstanceNum As Long
End Type


Public Type PlayerRec
    ' Account
    login As String * ACCOUNT_LENGTH
    Password As String * NAME_LENGTH
    characters() As CharacterRec
    ip As String
End Type

Public Type SpellBufferRec
    Spell As Long
    Timer As Long
    Target As Long
    TargetZone As Long
    tType As Byte
End Type

Public Type DoTRec
    Used As Boolean
    Spell As Long
    Timer As Long
    Caster As Long
    StartTime As Long
    'PET
    AttackerType As Long 'For Pets
End Type

Public Type ConditionalBranchRec
    Condition As Long
    Data1 As Long
    Data2 As Long
    Data3 As Long
    CommandList As Long
    ElseCommandList As Long
End Type

Private Type EventCommandRec
    Index As Byte
    Text1 As String
    Text2 As String
    Text3 As String
    Text4 As String
    Text5 As String
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4 As Long
    data5 As Long
    data6 As Long
    ConditionalBranch As ConditionalBranchRec
    MoveRouteCount As Long
    MoveRoute() As MoveRouteRec
End Type

Private Type CommandListRec
    CommandCount As Long
    ParentList As Long
    Commands() As EventCommandRec
End Type

Private Type EventPageRec
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
    WalkAnim As Long
    DirFix As Long
    WalkThrough As Long
    ShowName As Long
    
    'Trigger for the event
    Trigger As Byte
    
    'Commands for the event
    CommandListCount As Long
    CommandList() As CommandListRec
    
    Position As Byte
    
    questnum As Integer
    
    'For EventMap
    x As Long
    y As Long
End Type

Private Type EventRec
    Name As String
    Global As Byte
    PageCount As Long
    Pages() As EventPageRec
    x As Long
    y As Long
    'Self Switches re-set on restart.
    SelfSwitches(0 To 4) As Long
End Type

Public Type GlobalMapEvents
    eventID As Long
    pageID As Long
    x As Long
    y As Long
End Type

Private Type MapEventRec
    Dir As Long
    x As Long
    y As Long
    
    WalkingAnim As Long
    FixedDir As Long
    WalkThrough As Long
    ShowName As Long
    
    GraphicType As Long
    GraphicX As Long
    GraphicY As Long
    GraphicX2 As Long
    GraphicY2 As Long
    GraphicNum As Long
    
    movementspeed As Long
    Position As Long
    Visible As Long
    eventID As Long
    pageID As Long
    
    'Server Only Options
    MoveType As Long
    MoveSpeed As Long
    MoveFreq As Long
    MoveRouteCount As Long
    MoveRoute() As MoveRouteRec
    MoveRouteStep As Long
    
    RepeatMoveRoute As Long
    IgnoreIfCannotMove As Long
    questnum As Long
    
    MoveTimer As Long
    SelfSwitches(0 To 4) As Long
    MoverouteComplete As Long
End Type

Private Type EventMapRec
    CurrentEvents As Long
    EventPages() As MapEventRec
End Type

Private Type EventProcessingRec
    Active As Long
    CurList As Long
    CurSlot As Long
    eventID As Long
    pageID As Long
    WaitingForResponse As Long
    EventMovingID As Long
    EventMovingType As Long
    ActionTimer As Long
    ListLeftOff() As Long
End Type

Public Type TempPlayerRec
    ' Non saved local vars
    Buffer As clsBuffer
    InGame As Boolean
    AttackTimer As Long
    DataTimer As Long
    DataBytes As Long
    DataPackets As Long
    TargetType As Byte
    TargetZone As Byte
    Target As Long
    GettingMap As Byte
    SpellCD(1 To MAX_PLAYER_SPELLS) As Long
    InShop As Long
    StunTimer As Long
    StunDuration As Long
    InBank As Boolean
    ' trade
    TradeRequest As Long
    InTrade As Long
    TradeOffer(1 To MAX_INV) As PlayerInvRec
    AcceptTrade As Boolean
    ' dot/hot
    DoT(1 To MAX_DOTS) As DoTRec
    HoT(1 To MAX_DOTS) As DoTRec
    ' spell buffer
    spellBuffer As SpellBufferRec
    ' regen
    stopRegen As Boolean
    stopRegenTimer As Long
    ' party
    inParty As Long
    partyInvite As Long
    
    EventMap As EventMapRec
    EventProcessingCount As Long
    EventProcessing() As EventProcessingRec
    
    BuyHouseIndex As Long
    
    InvitationIndex As Long
    InvitationTimer As Long
    
    CurChar As Long
    
    'PET
    PetTarget As Long
    PetTargetType As Long
    PetTargetZone As Long
    PetBehavior As Long
    GoToX As Long
    GoToY As Long
    PetStunTimer As Long
    PetStunDuration As Long
    PetAttackTimer As Long
    PetSpellCD(1 To 4) As Long
    PetspellBuffer As SpellBufferRec
        ' dot/hot
    PetDoT(1 To MAX_DOTS) As DoTRec
    PetHoT(1 To MAX_DOTS) As DoTRec
        ' regen
    PetstopRegen As Boolean
    PetstopRegenTimer As Long
End Type

Private Type TileDataRec
    x As Long
    y As Long
    Tileset As Long
End Type

Public Type TileRec
    Layer(1 To MapLayer.Layer_Count - 1) As TileDataRec
    Autotile(1 To MapLayer.Layer_Count - 1) As Byte
    type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4 As String
    DirBlock As Byte
End Type

Public Type ExTileRec
    Layer(1 To ExMapLayer.Layer_Count - 1) As TileDataRec
    Autotile(1 To ExMapLayer.Layer_Count - 1) As Byte
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
    ExTile() As ExTileRec
    Npc(1 To MAX_MAP_NPCS) As Long
    NpcSpawnType(1 To MAX_MAP_NPCS) As Long
    EventCount As Long
    Events() As EventRec
End Type

Private Type FacePartsRec
    FHeads As String
    FHair As String
    FEyes As String
    FEyebrows As String
    FEars As String
    FMouth As String
    FNose As String
    FCloth As String
    FEtc As String
    FFace As String
End Type

Private Type ClassRec
    Name As String * NAME_LENGTH
    stat(1 To Stats.Stat_Count - 1) As Byte
    MaleSprite() As Long
    FemaleSprite() As Long
    
    startItemCount As Long
    StartItem() As Long
    StartValue() As Long
    
    startSpellCount As Long
    StartSpell() As Long
    MaleFaceParts As FacePartsRec
    FemaleFaceParts As FacePartsRec
End Type

Private Type ItemRec
    Name As String * NAME_LENGTH
    Desc As String * 255
    Sound As String * NAME_LENGTH
    
    Pic As Long

    type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    ClassReq As Long
    AccessReq As Long
    LevelReq As Long
    Mastery As Byte
    price As Long
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
    Num As Long
    Value As Long
    x As Byte
    y As Byte
    ' ownership + despawn
    playerName As String
    playerTimer As Long
    canDespawn As Boolean
    despawnTimer As Long
    PlayerDrop As Boolean
End Type

Private Type NpcRec
    Name As String * NAME_LENGTH
    AttackSay As String * 100
    Sound As String * NAME_LENGTH
    
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
    Target As Long
    TargetType As Byte
    Vital(1 To Vitals.Vital_Count - 1) As Long
    x As Byte
    y As Byte
    Dir As Byte
    ' For server use only
    SpawnWait As Long
    AttackTimer As Long
    StunDuration As Long
    StunTimer As Long
    ' regen
    stopRegen As Boolean
    stopRegenTimer As Long
    ' dot/hot
    DoT(1 To MAX_DOTS) As DoTRec
    HoT(1 To MAX_DOTS) As DoTRec
    Map As Long
    Inventory(1 To 20) As PlayerInvRec
End Type

Private Type TradeItemRec
    Item As Long
    itemvalue As Long
    costitem As Long
    costvalue As Long
End Type

Private Type ShopRec
    Name As String * NAME_LENGTH
    BuyRate As Long
    TradeItem(1 To MAX_TRADES) As TradeItemRec
End Type

Private Type SpellRec
    Name As String * NAME_LENGTH
    Desc As String * 255
    Sound As String * NAME_LENGTH
    
    type As Byte
    MPCost As Long
    LevelReq As Long
    AccessReq As Long
    ClassReq As Long
    CastTime As Long
    CDTime As Long
    Icon As Long
    Map As Long
    x As Long
    y As Long
    Dir As Byte
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
    DoorOpen() As Byte
    DoorTimer As Long
End Type

Private Type MapDataRec
    Npc() As MapNpcRec
End Type

Private Type MapResourceRec
    ResourceState As Byte
    ResourceTimer As Long
    x As Long
    y As Long
    cur_health As Long
End Type

Private Type ResourceCacheRec
    Resource_Count As Long
    ResourceData() As MapResourceRec
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
    Health As Long
    RespawnTime As Long
    WalkThrough As Boolean
    Animation As Long
End Type

Private Type AnimationRec
    Name As String * NAME_LENGTH
    Sound As String * NAME_LENGTH
    
    Sprite(0 To 1) As Long
    Frames(0 To 1) As Long
    LoopCount(0 To 1) As Long
    LoopTime(0 To 1) As Long
End Type

Public Type Vector
    x As Long
    y As Long
End Type

Public Type MapBlockRec
    Blocks() As Long
End Type

Private Type HouseRec
    ConfigName As String
    BaseMap As Long
    price As Long
    MaxFurniture As Long
    x As Long
    y As Long
End Type

Public Type ZoneRec
    Name As String * NAME_LENGTH
    Maps() As Long
    MapCount As Long
    NPCs(1 To MAX_MAP_NPCS * 2) As Long
    Weather(1 To 5) As Byte
    WeatherIntensity As Byte
    CurrentWeather As Byte
    WeatherTimer As Long
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
    x As Long
    y As Long
    Dir As Byte
    Timer As Long
End Type
