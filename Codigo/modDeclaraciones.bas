Attribute VB_Name = "modDeclaraciones"
'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************

''
' modDeclaraciones
'
' @remarks Declaraciones
' @author ^[GS]^
' @version 0.1.12
' @date 20081218

Option Explicit

Public Meteo                 As New clsMeteo
Public Audio                 As New clsAudio

Public Const MSGMod          As String = "Este mapa h� sido modificado." & vbCrLf & "Si no lo guardas perderas todos los cambios �Deseas guardarlo?"
Public Const MSGDang         As String = "CUIDADO! Este comando puede arruinar el mapa." & vbCrLf & "�Estas seguro que desea continuar?"
Public Const ENDL            As String * 2 = vbCrLf

'[Loopzer]
Public SeleccionIX           As Integer
Public SeleccionFX           As Integer
Public SeleccionIY           As Integer
Public SeleccionFY           As Integer
Public SeleccionAncho        As Integer
Public SeleccionAlto         As Integer
Public Seleccionando         As Boolean
Public SeleccionMap()        As MapBlock
Public DeSeleccionOX         As Integer
Public DeSeleccionOY         As Integer
Public DeSeleccionIX         As Integer
Public DeSeleccionFX         As Integer
Public DeSeleccionIY         As Integer
Public DeSeleccionFY         As Integer
Public DeSeleccionAncho      As Integer
Public DeSeleccionAlto       As Integer
Public DeSeleccionando       As Boolean
Public DeSeleccionMap()      As MapBlock
Public VerBlockeados         As Boolean
Public VerTriggers           As Boolean
Public VerGrilla             As Boolean ' grilla
Public VerCapa1              As Boolean
Public VerCapa2              As Boolean
Public VerCapa3              As Boolean
Public VerCapa4              As Boolean
Public VerTranslados         As Boolean
Public VerObjetos            As Boolean
Public VerNpcs               As Boolean
'[/Loopzer]

' Objeto de Translado
Public Cfg_TrOBJ             As Integer

'Path
Public inipath               As String
Public DirGraficos           As String
Public DirMidi               As String
Public DirIndex              As String
Public DirDats               As String
Public dirwavs               As String
Public bAutoGuardarMapa      As Byte
Public bAutoGuardarMapaCount As Byte
Public HotKeysAllow          As Boolean  ' Control Automatico de HotKeys
Public vMostrando            As Byte
Public WORK                  As Boolean
Public PATH_Save             As String
Public NumMap_Save           As Integer
Public NameMap_Save          As String

' DX Config
Public PantallaX             As Integer
Public PantallaY             As Integer

' [GS] 02/10/06
' Client Config
Public ClienteHeight         As Integer
Public ClienteWidth          As Integer

Public Type tSetupMods

    bDinamic    As Boolean
    byMemory    As Byte
    bUseVideo   As Boolean
    bNoMusic    As Boolean
    bNoSound    As Boolean

End Type

Public ClientSetup   As tSetupMods

Public SobreX        As Byte   ' Posicion X bajo el Cursor
Public SobreY        As Byte   ' Posicion Y bajo el Cursor

' Radar
Public MiRadarX      As Integer
Public MiRadarY      As Integer
Public bRefreshRadar As Boolean

Type SupData

    name As String
    Grh As Integer
    Width As Byte
    Height As Byte
    Block As Boolean
    Capa As Byte

End Type

Public MaxSup    As Integer

Public SupData() As SupData

Public Type NpcData

    name As String
    Body As Integer
    Head As Integer
    Heading As Byte

End Type

Public NumNPCs   As Long

'Public NumNPCsHOST As Integer
Public NpcData() As NpcData

Public Type ObjData

    name As String 'Nombre del obj
    ObjType As Integer 'Tipo enum que determina cuales son las caract del obj
    GrhIndex As Integer ' Indice del grafico que representa el obj
    GrhSecundario As Integer
    Info As String
    Ropaje As Integer 'Indice del grafico del ropaje
    WeaponAnim As Integer ' Apunta a una anim de armas
    ShieldAnim As Integer ' Apunta a una anim de escudo
    Texto As String

End Type

Public NumOBJs     As Integer
Public ObjData()   As ObjData
Public Conexion    As New Connection
Public prgRun      As Boolean
Public CurrentGrh  As Grh
Public Play        As Boolean
Public MapaCargado As Boolean
Public cFPS        As Long
Public dTiempoGT   As Double
Public dLastWalk   As Double

'Hold info about each map
Public Type MapInfo

    Music As String
    name As String
    MapVersion As Integer
    PK As Boolean
    MagiaSinEfecto As Byte
    InviSinEfecto As Byte
    ResuSinEfecto As Byte
    NoEncriptarMP As Byte
    Terreno As String
    Zona As String
    Restringir As String
    BackUp As Byte
    Changed As Byte ' flag for WorldEditor

End Type

'********** CONSTANTS ***********
'Heading Constants
Public Const NORTH       As Byte = 1
Public Const EAST        As Byte = 2
Public Const SOUTH       As Byte = 3
Public Const WEST        As Byte = 4

'Map sizes in tiles
Public Const XMaxMapSize As Byte = 100
Public Const XMinMapSize As Byte = 1
Public Const YMaxMapSize As Byte = 100
Public Const YMinMapSize As Byte = 1

'********** TYPES ***********
'Holds a local position
Public Type Position

    X As Integer
    Y As Integer

End Type

'Holds a world position
Public Type WorldPos

    Map As Integer
    X As Integer
    Y As Integer

End Type

' Cuerpos body.dat
Public Type tIndiceCuerpo

    Body(1 To 4) As Integer
    HeadOffsetX As Integer
    HeadOffsetY As Integer

End Type

' Lista de Cuerpos body.dat
Public Type tBodyData

    Walk(1 To 4) As Grh
    HeadOffset As Position

End Type

' body.dat
Public BodyData() As tBodyData
Public NumBodies  As Integer

'Lista de cabezas
Public Type tIndiceCabeza
    Head(1 To 4) As Integer
End Type

'Heads list
Public Type tHeadData
    Head(0 To 4) As Grh
End Type

Public HeadData() As tHeadData

'Hold info about a character
Public Type Char

    Active As Byte
    Heading As Byte
    Pos As Position

    Body As tBodyData
    Head As tHeadData
    
    Moving As Byte
    MoveOffset As Position
    
End Type

'Holds info about a object
Public Type Obj
    objindex As Integer
    Amount As Integer
End Type

Public Light_Count As Integer

Public Type Light

    X As Integer
    Y As Integer
    Active As Boolean 'Do we ignore this light?
    id As Long
    map_x As Long 'Coordinates
    map_y As Long
    Color As Long 'Start colour
    Range As Long
    RGBCOLOR As D3DCOLORVALUE

End Type

'Holds info about each tile position
Public Type MapBlock

    Graphic(1 To 4) As Grh
    CharIndex As Integer
    ObjGrh As Grh
    
    Trigger As Integer
    
    NPCIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
    
    light_index As Integer
    base_light(0 To 3) As Boolean 'Indica si el tile tiene luz propia.
    light_base_value(0 To 3) As Long 'Luz propia del tile.
    light_value(0 To 3) As Long 'Color de luz con el que esta siendo renderizado.
    
    particle_group_index As Integer
    parti_index As Integer
    AlturaPoligonos(0 To 3) As Byte

End Type

'********** Public VARS ***********
'Where the map borders are.. Set during load
Public MinXBorder                                                                                 As Byte
Public MaxXBorder                                                                                 As Byte
Public MinYBorder                                                                                 As Byte
Public MaxYBorder                                                                                 As Byte

'Object Constants
Public Const MAX_INVENORY_OBJS                                                                    As Integer = 10000

' Deshacer
Public Const maxDeshacer                                                                          As Integer = 10
Public MapData_Deshacer(1 To maxDeshacer, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
Type tDeshacerInfo
    Libre As Boolean
    Desc As String
End Type
Public MapData_Deshacer_Info(1 To maxDeshacer) As tDeshacerInfo

'********** Public ARRAYS ***********
Public MapData()                               As MapBlock 'Holds map data for current map
Public MapInfo                                 As MapInfo 'Holds map info for current map
Public CharList(1 To 10000)                    As Char 'Holds info about all characters on map
Public Lights()                                As Light

'Encabezado bmp
Type BITMAPFILEHEADER

    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long

End Type

'Info del encabezado del bmp
Type BITMAPINFOHEADER

    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long

End Type

Public gDespX              As Integer
Public gDespY              As Integer

'User status vars
Public CurMap              As Integer 'Current map loaded
Public UserIndex           As Integer
Global UserBody            As Integer
Global UserHead            As Integer
Public UserPos             As Position 'Holds current user pos
Public AddtoUserPos        As Position 'For moving user
Public UserCharIndex       As Integer
Public EngineRun           As Boolean
Public FramesPerSec        As Integer
Public FramesPerSecCounter As Long

'Main view size size in tiles
Public WindowTileWidth     As Integer
Public WindowTileHeight    As Integer

'Pixel offset of main view screen from 0,0
Public MainViewTop         As Integer
Public MainViewLeft        As Integer

'How many tiles the engine "looks ahead" when
'drawing the screen
Public TileBufferSize      As Integer

'Handle to where all the drawing is going to take place
Public DisplayFormhWnd     As Long

'Tile size in pixels
Public TilePixelHeight     As Integer
Public TilePixelWidth      As Integer

'Map editor variables
Public WalkMode            As Boolean

'Totals
Public NumMaps             As Integer 'Number of maps
Public Numheads            As Integer
Public NumGrhFiles         As Integer 'Number of bmps
Public MaxGrhs             As Integer 'Number of Grhs
Global NumChars            As Integer
Global LastChar            As Integer

'********** Direct X ***********
Public MainViewWidth       As Integer
Public MainViewHeight      As Integer

'********** OUTSIDE FUNCTIONS ***********

Public Const PATH_GRAPHICS As String = "Graficos"
Public Const PATH_INIT     As String = "Init"

'Good old BitBlt
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

'Sound stuff
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uRetrunLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'For Get and Write Var
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpfilename As String) As Long

'For KeyInput
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetTickCount Lib "kernel32" () As Long
