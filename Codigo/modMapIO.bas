Attribute VB_Name = "modMapIO"
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
' modMapIO
'
' @remarks Funciones Especificas al trabajo con Archivos de Mapas
' @author gshaxor@gmail.com
' @version 0.1.15
' @date 20060602

Option Explicit

'***************************
'Map format .CSM
'***************************
Private Type tMapHeader
    NumeroBloqueados As Long
    NumeroLayers(2 To 4) As Long
    NumeroTriggers As Long
    NumeroLuces As Long
    NumeroParticulas As Long
    NumeroNPCs As Long
    NumeroOBJs As Long
    NumeroTE As Long
End Type

Private Type tDatosBloqueados
    X As Integer
    Y As Integer
End Type

Private Type tDatosGrh
    X As Integer
    Y As Integer
    GrhIndex As Long
End Type

Private Type tDatosTrigger
    X As Integer
    Y As Integer
    Trigger As Integer
End Type

Private Type tDatosLuces
    X As Integer
    Y As Integer
    light_value(3) As Long
    base_light(0 To 3) As Boolean 'Indica si el tile tiene luz propia.
End Type

Private Type tDatosParticulas
    X As Integer
    Y As Integer
    Particula As Long
End Type

Private Type tDatosNPC
    X As Integer
    Y As Integer
    NPCIndex As Integer
End Type

Private Type tDatosObjs
    X As Integer
    Y As Integer
    objindex As Integer
    ObjAmmount As Integer
End Type

Private Type tDatosTE
    X As Integer
    Y As Integer
    DestM As Integer
    DestX As Integer
    DestY As Integer
End Type

Private Type tMapSize
    XMax As Integer
    XMin As Integer
    YMax As Integer
    YMin As Integer
End Type

Private Type tMapDat
    map_name As String
    battle_mode As Boolean
    backup_mode As Boolean
    restrict_mode As String
    music_number As String
    zone As String
    terrain As String
    ambient As String
    lvlMinimo As String
    SePuedeDomar As Boolean
    ResuSinEfecto As Boolean
    MagiaSinEfecto As Boolean
    InviSinEfecto As Boolean
    NoEncriptarMP As Boolean
    version As Long
End Type

Public MapSize As tMapSize
Private MapDat As tMapDat
'***************************
'Map format .CSM
'***************************

Private MapTitulo As String     ' GS > Almacena el titulo del mapa para el .dat

''
' Obtener el tamano de un archivo
'
' @param FileName Especifica el path del archivo
' @return   Nos devuelve el tamano

Public Function FileSize(ByVal FileName As String) As Long
    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************

    On Error GoTo FalloFile

    Dim nFileNum  As Integer

    Dim lFileSize As Long
    
    nFileNum = FreeFile
    Open FileName For Input As nFileNum
    lFileSize = LOF(nFileNum)
    Close nFileNum
    FileSize = lFileSize
    
    Exit Function
FalloFile:
    FileSize = -1

End Function

''
' Nos dice si existe el archivo/directorio
'
' @param file Especifica el path
' @param FileType Especifica el tipo de archivo/directorio
' @return   Nos devuelve verdadero o falso

Public Function FileExist(ByVal File As String, ByVal FileType As VbFileAttribute) As Boolean

    '*************************************************
    'Author: Unkwown
    'Last modified: 26/05/06
    '*************************************************
    If LenB(Dir(File, FileType)) = 0 Then
        FileExist = False
    Else
        FileExist = True

    End If

End Function

''
' Abre un Mapa
'
' @param Path Especifica el path del mapa

Public Sub AbrirMapa(ByVal Path As String, ByRef Buffer() As MapBlock, Optional ByVal SoloMap As Boolean = False)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************

    Call MapaV2_Cargar(Path, Buffer, SoloMap)

End Sub

''
' Guarda el Mapa
'
' @param Path Especifica el path del mapa

Public Sub GuardarMapa(Optional Path As String)
    '*************************************************
    'Author: Lorwik
    'Last modified: 01/11/08
    '*************************************************

    frmMain.Dialog.CancelError = True

    On Error GoTo ErrHandler

    If LenB(Path) = 0 Then
        Call frmMain.ObtenerNombreArchivo(True)
        Path = frmMain.Dialog.FileName

        If LenB(Path) = 0 Then Exit Sub

    End If

    If frmMain.Dialog.FilterIndex = 1 Then
        Call MapaV2_Guardar(Path)
    ElseIf frmMain.Dialog.FilterIndex = 2 Then
        Call Save_CSM(Path)
    End If

ErrHandler:

End Sub

''
' Nos pregunta donde guardar el mapa en caso de modificarlo
'
' @param Path Especifica si existiera un path donde guardar el mapa

Public Sub DeseaGuardarMapa(Optional Path As String)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************

    If MapInfo.Changed = 1 Then
        If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then
            Call GuardarMapa(Path)

        End If

    End If

End Sub

Public Sub NuevoMapa()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 21/05/06
    '*************************************************

    On Error Resume Next

    Dim loopc As Integer

    bAutoGuardarMapaCount = 0
    
    With frmMain
    
        '.mnuUtirialNuevoFormato.Checked = True
        .mnuReAbrirMapa.Enabled = False
        .TimAutoGuardarMapa.Enabled = False
        .lblMapVersion.Caption = 0

        MapaCargado = False

        For loopc = 0 To frmMain.MapPest.Count - 1
            .MapPest(loopc).Enabled = False
        Next

        .MousePointer = 11
    
    End With
    
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock

    For loopc = 1 To LastChar
        If CharList(loopc).Active = 1 Then Call EraseChar(loopc)
    Next loopc
    
    With MapInfo
    
        .MapVersion = 0
        .Name = "Nuevo Mapa"
        .Music = 0
        .PK = True
        .MagiaSinEfecto = 0
        .Terreno = "BOSQUE"
        .Zona = "CAMPO"
        .Restringir = "NO"
        .NoEncriptarMP = 0
    
    End With

    Call MapInfo_Actualizar

    bRefreshRadar = True ' Radar

    'Set changed flag
    MapInfo.Changed = 0
    frmMain.MousePointer = 0

    ' Vacio deshacer
    Call modEdicion.Deshacer_Clear

    MapaCargado = True

    frmMain.SetFocus

End Sub

Public Sub MapaV2_Guardar(ByVal SaveAs As String)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************

    On Error GoTo ErrorSave

    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim loopc       As Long
    Dim TempInt     As Integer
    Dim Y           As Long
    Dim X           As Long
    Dim ByFlags     As Byte
    Dim R           As Byte
    Dim G           As Byte
    Dim B           As Byte

    If FileExist(SaveAs, vbNormal) = True Then
        If MsgBox("�Desea sobrescribir " & SaveAs & "?", vbCritical + vbYesNo) = vbNo Then
            Exit Sub
        Else
            Kill SaveAs

        End If

    End If

    frmMain.MousePointer = 11

    ' y borramos el .inf tambien
    If FileExist(left$(SaveAs, Len(SaveAs) - 4) & ".inf", vbNormal) = True Then
        Kill left$(SaveAs, Len(SaveAs) - 4) & ".inf"

    End If

    'Open .map file
    FreeFileMap = FreeFile
    Open SaveAs For Binary As FreeFileMap
    Seek FreeFileMap, 1

    SaveAs = left$(SaveAs, Len(SaveAs) - 4)
    SaveAs = SaveAs & ".inf"

    'Open .inf file
    FreeFileInf = FreeFile
    Open SaveAs For Binary As FreeFileInf
    Seek FreeFileInf, 1

    'map Header
    
    ' Version del Mapa
    If frmMain.lblMapVersion.Caption < 32767 Then
        frmMain.lblMapVersion.Caption = frmMain.lblMapVersion + 1
        frmMapInfo.txtMapVersion = frmMain.lblMapVersion.Caption

    End If

    Put FreeFileMap, , CInt(frmMain.lblMapVersion.Caption)
    Put FreeFileMap, , MiCabecera
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    
    'inf Header
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    
    'Write .map file
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
            ByFlags = 0
                
            If MapData(X, Y).Blocked = 1 Then ByFlags = ByFlags Or 1
            If MapData(X, Y).Graphic(2).GrhIndex Then ByFlags = ByFlags Or 2
            If MapData(X, Y).Graphic(3).GrhIndex Then ByFlags = ByFlags Or 4
            If MapData(X, Y).Graphic(4).GrhIndex Then ByFlags = ByFlags Or 8
            If MapData(X, Y).Trigger Then ByFlags = ByFlags Or 16
            If MapData(X, Y).particle_group_index Then ByFlags = ByFlags Or 32
            If MapData(X, Y).light_index Then ByFlags = ByFlags Or 64
            If MapData(X, Y).AlturaPoligonos(0) Or MapData(X, Y).AlturaPoligonos(1) Or MapData(X, Y).AlturaPoligonos(2) Or MapData(X, Y).AlturaPoligonos(3) Then ByFlags = ByFlags Or 128
            Put FreeFileMap, , ByFlags
                
            Put FreeFileMap, , MapData(X, Y).Graphic(1).GrhIndex
                
            For loopc = 2 To 4

                If MapData(X, Y).Graphic(loopc).GrhIndex Then Put FreeFileMap, , MapData(X, Y).Graphic(loopc).GrhIndex
            Next loopc
                
            If MapData(X, Y).Trigger Then Put FreeFileMap, , MapData(X, Y).Trigger
                
            If MapData(X, Y).particle_group_index Then Put FreeFileMap, , MapData(X, Y).parti_index

            If MapData(X, Y).light_index Then
                Put FreeFileMap, , Lights(MapData(X, Y).light_index).Range
                R = Lights(MapData(X, Y).light_index).RGBCOLOR.R
                G = Lights(MapData(X, Y).light_index).RGBCOLOR.G
                B = Lights(MapData(X, Y).light_index).RGBCOLOR.B
                Put FreeFileMap, , R
                Put FreeFileMap, , G
                Put FreeFileMap, , B

            End If
                
            If MapData(X, Y).AlturaPoligonos(0) Or MapData(X, Y).AlturaPoligonos(1) Or MapData(X, Y).AlturaPoligonos(2) Or MapData(X, Y).AlturaPoligonos(3) Then
                Put FreeFileMap, , MapData(X, Y).AlturaPoligonos(0)
                Put FreeFileMap, , MapData(X, Y).AlturaPoligonos(1)
                Put FreeFileMap, , MapData(X, Y).AlturaPoligonos(2)
                Put FreeFileMap, , MapData(X, Y).AlturaPoligonos(3)
                    
                If MapData(X, Y).AlturaPoligonos(0) Then Put FreeFileMap, , MapData(X, Y).AlturaPoligonos(0)

                If MapData(X, Y).AlturaPoligonos(1) Then Put FreeFileMap, , MapData(X, Y).AlturaPoligonos(1)

                If MapData(X, Y).AlturaPoligonos(2) Then Put FreeFileMap, , MapData(X, Y).AlturaPoligonos(2)

                If MapData(X, Y).AlturaPoligonos(3) Then Put FreeFileMap, , MapData(X, Y).AlturaPoligonos(3)

            End If
                
            '.inf file
                
            ByFlags = 0
                
            If MapData(X, Y).TileExit.Map Then ByFlags = ByFlags Or 1
            If MapData(X, Y).NPCIndex Then ByFlags = ByFlags Or 2
            If MapData(X, Y).OBJInfo.objindex Then ByFlags = ByFlags Or 4
                
            Put FreeFileInf, , ByFlags
                
            If MapData(X, Y).TileExit.Map Then
                Put FreeFileInf, , MapData(X, Y).TileExit.Map
                Put FreeFileInf, , MapData(X, Y).TileExit.X
                Put FreeFileInf, , MapData(X, Y).TileExit.Y

            End If
                
            If MapData(X, Y).NPCIndex Then
                
                Put FreeFileInf, , CInt(MapData(X, Y).NPCIndex)

            End If
                
            If MapData(X, Y).OBJInfo.objindex Then
                Put FreeFileInf, , MapData(X, Y).OBJInfo.objindex
                Put FreeFileInf, , MapData(X, Y).OBJInfo.Amount

            End If
            
        Next X
    Next Y
    
    'Close .map file
    Close FreeFileMap
    
    'Close .inf file
    Close FreeFileInf

    Call Pestanas(SaveAs)

    'write .dat file
    SaveAs = left$(SaveAs, Len(SaveAs) - 4) & ".dat"
    MapInfo_Guardar SaveAs

    'Change mouse icon
    frmMain.MousePointer = 0
    MapInfo.Changed = 0

    Exit Sub

ErrorSave:
    MsgBox "Error en GuardarV2, nro. " & Err.Number & " - " & Err.Description

End Sub

Public Sub MapaV2_Cargar(ByVal Map As String, ByRef Buffer() As MapBlock, ByVal SoloMap As Boolean)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************

    On Error Resume Next

    Dim TempInt     As Integer
    Dim Body        As Integer
    Dim Head        As Integer
    Dim Heading     As Byte
    Dim Y           As Integer
    Dim X           As Integer
    Dim ByFlags     As Byte
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim TempLng     As Long
    Dim TempByte1   As Byte
    Dim TempByte2   As Byte
    Dim TempByte3   As Byte
           
    Call LightDestroyAll
    Call Particle_Group_Remove_All
    Call Map_ResetMontanita
    
    'Change mouse icon
    frmMain.MousePointer = 11
       
    'Open files
    FreeFileMap = FreeFile
    Open Map For Binary As FreeFileMap
    Seek FreeFileMap, 1
    
    If Not SoloMap Then
        Map = left$(Map, Len(Map) - 4)
        Map = Map & ".inf"
        
        FreeFileInf = FreeFile
        Open Map For Binary As FreeFileInf
        Seek FreeFileInf, 1

    End If
    
    'Cabecera map
    Get FreeFileMap, , MapInfo.MapVersion
    Get FreeFileMap, , MiCabecera
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
    
    If Not SoloMap Then
        'Cabecera inf
        Get FreeFileInf, , TempInt
        Get FreeFileInf, , TempInt
        Get FreeFileInf, , TempInt
        Get FreeFileInf, , TempInt
        Get FreeFileInf, , TempInt

    End If
    
    'Load arrays
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
    
            Get FreeFileMap, , ByFlags
            
            Buffer(X, Y).Blocked = (ByFlags And 1)
            
            Get FreeFileMap, , Buffer(X, Y).Graphic(1).GrhIndex
            Grh_Initialize Buffer(X, Y).Graphic(1), Buffer(X, Y).Graphic(1).GrhIndex
            
            'Layer 2 used?
            If ByFlags And 2 Then
                Get FreeFileMap, , Buffer(X, Y).Graphic(2).GrhIndex
                Grh_Initialize Buffer(X, Y).Graphic(2), Buffer(X, Y).Graphic(2).GrhIndex
            Else
                Buffer(X, Y).Graphic(2).GrhIndex = 0

            End If
                
            'Layer 3 used?
            If ByFlags And 4 Then
                Get FreeFileMap, , Buffer(X, Y).Graphic(3).GrhIndex
                Grh_Initialize Buffer(X, Y).Graphic(3), Buffer(X, Y).Graphic(3).GrhIndex
            Else
                Buffer(X, Y).Graphic(3).GrhIndex = 0

            End If
                
            'Layer 4 used?
            If ByFlags And 8 Then
                Get FreeFileMap, , Buffer(X, Y).Graphic(4).GrhIndex
                Grh_Initialize Buffer(X, Y).Graphic(4), Buffer(X, Y).Graphic(4).GrhIndex
            Else
                Buffer(X, Y).Graphic(4).GrhIndex = 0

            End If
             
            'Trigger used?
            If ByFlags And 16 Then
                Get FreeFileMap, , Buffer(X, Y).Trigger
            Else
                Buffer(X, Y).Trigger = 0

            End If
            
            If ByFlags And 32 Then
                Get FreeFileMap, , TempInt
                MapData(X, Y).particle_group_index = General_Particle_Create(TempInt, X, Y, -1)

            End If
            
            If ByFlags And 64 Then
                Get FreeFileMap, , TempLng
                Get FreeFileMap, , TempByte1
                Get FreeFileMap, , TempByte2
                Get FreeFileMap, , TempByte3
                Call LightSet(X, Y, True, TempLng, TempByte1, TempByte2, TempByte3)

            End If
            
            If ByFlags And 128 Then
                Get FreeFileMap, , MapData(X, Y).AlturaPoligonos(0)
                Get FreeFileMap, , MapData(X, Y).AlturaPoligonos(1)
                Get FreeFileMap, , MapData(X, Y).AlturaPoligonos(2)
                Get FreeFileMap, , MapData(X, Y).AlturaPoligonos(3)
                
                If MapData(X, Y).AlturaPoligonos(0) Then Get FreeFileMap, , MapData(X, Y).AlturaPoligonos(0)
                
                If MapData(X, Y).AlturaPoligonos(1) Then Get FreeFileMap, , MapData(X, Y).AlturaPoligonos(1)
                
                If MapData(X, Y).AlturaPoligonos(2) Then Get FreeFileMap, , MapData(X, Y).AlturaPoligonos(2)
                
                If MapData(X, Y).AlturaPoligonos(3) Then Get FreeFileMap, , MapData(X, Y).AlturaPoligonos(3)

            End If

            If Not SoloMap Then
                '.inf file
                Get FreeFileInf, , ByFlags
                
                If ByFlags And 1 Then
                    Get FreeFileInf, , Buffer(X, Y).TileExit.Map
                    Get FreeFileInf, , Buffer(X, Y).TileExit.X
                    Get FreeFileInf, , Buffer(X, Y).TileExit.Y

                End If
        
                If ByFlags And 2 Then
                    'Get and make NPC
                    Get FreeFileInf, , Buffer(X, Y).NPCIndex
        
                    If Buffer(X, Y).NPCIndex < 0 Then
                        Buffer(X, Y).NPCIndex = 0
                    Else
                        Body = NpcData(Buffer(X, Y).NPCIndex).Body
                        Head = NpcData(Buffer(X, Y).NPCIndex).Head
                        Heading = NpcData(Buffer(X, Y).NPCIndex).Heading
                        Call MakeChar(NextOpenChar(), Body, Head, Heading, X, Y)

                    End If

                End If
        
                If ByFlags And 4 Then
                    'Get and make Object
                    Get FreeFileInf, , Buffer(X, Y).OBJInfo.objindex
                    Get FreeFileInf, , Buffer(X, Y).OBJInfo.Amount

                    If Buffer(X, Y).OBJInfo.objindex > 0 Then
                        Grh_Initialize Buffer(X, Y).ObjGrh, ObjData(Buffer(X, Y).OBJInfo.objindex).GrhIndex

                    End If

                End If

            End If

        Next X
    Next Y
    
    'Close files
    Close FreeFileMap
    
    If Not SoloMap Then
        Close FreeFileInf
        
        Call Pestanas(Map)
        
        bRefreshRadar = True ' Radar
        
        Map = left$(Map, Len(Map) - 4) & ".dat"
        
        Call MapInfo_Cargar(Map)
        frmMain.lblMapVersion.Caption = MapInfo.MapVersion
        
        'Set changed flag
        MapInfo.Changed = 0
        
        ' Vacia el Deshacer
        Call modEdicion.Deshacer_Clear

    End If
    
    'Change mouse icon
    frmMain.MousePointer = 0
    MapaCargado = True

    frmMain.picRadar.Picture = LoadPicture(DirMinimapa & ReturnNumberFromString(Map) & ".bmp")

End Sub

'Solo se usa para el minimapa
Private Function ReturnNumberFromString(ByVal sString As String) As String  
   Dim i As Integer  
   For i = 1 To Len(sString)  
       If Mid(sString, i, 1) Like "[0-9]" Then  
           ReturnNumberFromString = ReturnNumberFromString + Mid(sString, i, 1)  
       End If  
   Next i  
End Function 

' *****************************************************************************
' MAPINFO *********************************************************************
' *****************************************************************************

''
' Guardar Informacion del Mapa (.dat)
'
' @param Archivo Especifica el Path del archivo .DAT

Public Sub MapInfo_Guardar(ByVal Archivo As String)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 28/05/06
    '*************************************************

    If LenB(MapTitulo) = 0 Then
        MapTitulo = NameMap_Save

    End If
    
    Set FileManager = New clsIniManager
    
    With FileManager
        Call .Initialize(Archivo)

        Call .ChangeValue(MapTitulo, "Name", MapInfo.Name)
        Call .ChangeValue(MapTitulo, "MusicNum", MapInfo.Music)
        Call .ChangeValue(MapTitulo, "MagiaSinefecto", Val(MapInfo.MagiaSinEfecto))
        Call .ChangeValue(MapTitulo, "InviSinEfecto", Val(MapInfo.InviSinEfecto))
        Call .ChangeValue(MapTitulo, "ResuSinEfecto", Val(MapInfo.ResuSinEfecto))
        Call .ChangeValue(MapTitulo, "NoEncriptarMP", Val(MapInfo.NoEncriptarMP))

        Call .ChangeValue(MapTitulo, "Terreno", MapInfo.Terreno)
        Call .ChangeValue(MapTitulo, "Zona", MapInfo.Zona)
        Call .ChangeValue(MapTitulo, "Restringir", MapInfo.Restringir)
        Call .ChangeValue(MapTitulo, "BackUp", Str(MapInfo.BackUp))

        If MapInfo.PK Then
            Call .ChangeValue(MapTitulo, "Pk", "0")
        Else
            Call .ChangeValue(MapTitulo, "Pk", "1")
        End If
    
    End With
    
    Set FileManager = Nothing
    
End Sub

''
' Abrir Informacion del Mapa (.dat)
'
' @param Archivo Especifica el Path del archivo .DAT

Public Sub MapInfo_Cargar(ByVal Archivo As String)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 02/06/06
    '*************************************************

    On Error Resume Next

    Dim Leer  As New clsIniManager
    Dim loopc As Integer
    Dim Path  As String

    MapTitulo = Empty
    
    Call Leer.Initialize(Archivo)

    For loopc = Len(Archivo) To 1 Step -1

        If mid(Archivo, loopc, 1) = "\" Then
            Path = left(Archivo, loopc)
            Exit For

        End If

    Next
    Archivo = Right(Archivo, Len(Archivo) - (Len(Path)))
    MapTitulo = UCase(left(Archivo, Len(Archivo) - 4))
    
    With MapInfo
    
        .Name = Leer.GetValue(MapTitulo, "Name")
        .Music = Leer.GetValue(MapTitulo, "MusicNum")
        .MagiaSinEfecto = Val(Leer.GetValue(MapTitulo, "MagiaSinEfecto"))
        .InviSinEfecto = Val(Leer.GetValue(MapTitulo, "InviSinEfecto"))
        .ResuSinEfecto = Val(Leer.GetValue(MapTitulo, "ResuSinEfecto"))
        .NoEncriptarMP = Val(Leer.GetValue(MapTitulo, "NoEncriptarMP"))
    
        If Val(Leer.GetValue(MapTitulo, "Pk")) = 0 Then
            MapInfo.PK = True
        Else
            MapInfo.PK = False

        End If
    
        .Terreno = Leer.GetValue(MapTitulo, "Terreno")
        .Zona = Leer.GetValue(MapTitulo, "Zona")
        .Restringir = Leer.GetValue(MapTitulo, "Restringir")
        .BackUp = Val(Leer.GetValue(MapTitulo, "BACKUP"))
    
    End With
    
    Call MapInfo_Actualizar
    
End Sub

''
' Actualiza el formulario de MapInfo
'

Public Sub MapInfo_Actualizar()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 02/06/06
    '*************************************************

    On Error Resume Next

    ' Mostrar en Formularios
    With frmMapInfo
        .txtMapNombre.Text = MapInfo.Name
        .txtMapMusica.Text = MapInfo.Music
        .txtMapTerreno.Text = MapInfo.Terreno
        .txtMapZona.Text = MapInfo.Zona
        .txtMapRestringir.Text = MapInfo.Restringir
        .chkMapBackup.Value = MapInfo.BackUp
        .chkMapMagiaSinEfecto.Value = MapInfo.MagiaSinEfecto
        .chkMapInviSinEfecto.Value = MapInfo.InviSinEfecto
        .chkMapResuSinEfecto.Value = MapInfo.ResuSinEfecto
        .chkMapNoEncriptarMP.Value = MapInfo.NoEncriptarMP
        .chkMapPK.Value = IIf(MapInfo.PK = True, 1, 0)
        .txtMapVersion = MapInfo.MapVersion
    End With
    
    With frmMain
        .lblMapNombre = MapInfo.Name
        .lblMapMusica = MapInfo.Music
    End With
    

End Sub

''
' Calcula la orden de Pestanas
'
' @param Map Especifica path del mapa

Public Sub Pestanas(ByVal Map As String)

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 28/05/06
    '*************************************************
    On Error Resume Next

    Dim loopc As Integer

    For loopc = Len(Map) To 1 Step -1

        If mid(Map, loopc, 1) = "\" Then
            PATH_Save = left(Map, loopc)
            Exit For

        End If

    Next
    Map = Right(Map, Len(Map) - (Len(PATH_Save)))

    For loopc = Len(left(Map, Len(Map) - 4)) To 1 Step -1

        If IsNumeric(mid(left(Map, Len(Map) - 4), loopc, 1)) = False Then
            NumMap_Save = Right(left(Map, Len(Map) - 4), Len(left(Map, Len(Map) - 4)) - loopc)
            NameMap_Save = left(Map, loopc)
            Exit For

        End If

    Next

    For loopc = (NumMap_Save - 4) To (NumMap_Save + 8)

        If FileExist(PATH_Save & NameMap_Save & loopc & ".map", vbArchive) = True Then
            frmMain.MapPest(loopc - NumMap_Save + 4).Visible = True
            frmMain.MapPest(loopc - NumMap_Save + 4).Enabled = True
            frmMain.MapPest(loopc - NumMap_Save + 4).Caption = NameMap_Save & loopc
        Else
            frmMain.MapPest(loopc - NumMap_Save + 4).Visible = False

        End If

    Next

End Sub

Public Sub CSMInfoSave()
    
    With MapDat
    
        .map_name = MapInfo.Name
        .music_number = MapInfo.Music
        .MagiaSinEfecto = MapInfo.MagiaSinEfecto
        .InviSinEfecto = MapInfo.InviSinEfecto
        .ResuSinEfecto = MapInfo.ResuSinEfecto
        .NoEncriptarMP = MapInfo.NoEncriptarMP
        .version = MapInfo.MapVersion
    
        If MapInfo.PK = True Then
            .battle_mode = True
        Else
            .battle_mode = False

        End If
    
        .terrain = MapInfo.Terreno
        .zone = MapInfo.Zona
        .restrict_mode = MapInfo.Restringir
        .backup_mode = MapInfo.BackUp
    
    End With
    
End Sub

Public Sub CSMInfoCargar()

    With MapInfo
    
        .Name = MapDat.map_name
        .Music = MapDat.music_number
        .MagiaSinEfecto = MapDat.MagiaSinEfecto
        .InviSinEfecto = MapDat.InviSinEfecto
        .ResuSinEfecto = MapDat.ResuSinEfecto
        .NoEncriptarMP = MapDat.NoEncriptarMP
        .MapVersion = MapDat.version
    
        If MapDat.battle_mode = True Then
            .PK = True
        Else
            .PK = False

        End If
    
        .Terreno = MapDat.terrain
        .Zona = MapDat.zone
        .Restringir = MapDat.restrict_mode
        .BackUp = MapDat.backup_mode
    
        Call MapInfo_Actualizar

    End With
    
End Sub

Sub Cargar_CSM(ByVal Map As String)

    On Error GoTo ErrorHandler

    Dim fh           As Integer
    Dim File         As Integer

    Dim MH           As tMapHeader
    
    Dim Blqs()       As tDatosBloqueados

    Dim L1()         As Long
    Dim L2()         As tDatosGrh
    Dim L3()         As tDatosGrh
    Dim L4()         As tDatosGrh

    Dim Triggers()   As tDatosTrigger
    Dim Luces()      As tDatosLuces
    Dim Particulas() As tDatosParticulas
    Dim Objetos()    As tDatosObjs
    Dim NPCs()       As tDatosNPC
    Dim TEs()        As tDatosTE

    Dim i            As Long
    Dim j            As Long

    DoEvents
    
    'Change mouse icon
    frmMain.MousePointer = 11
    
    fh = FreeFile
    Open Map For Binary Access Read As fh
    
        Get #fh, , MH
        Get #fh, , MapSize
        
        '�Queremos cargar un mapa de IAO 1.4?
        Get #fh, , MapDat
        
        ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize)
        ReDim L1(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize)
        
        Get #fh, , L1
        
        With MH
    
            If .NumeroBloqueados > 0 Then
                ReDim Blqs(1 To .NumeroBloqueados)
                Get #fh, , Blqs
    
                For i = 1 To .NumeroBloqueados
                    MapData(Blqs(i).X, Blqs(i).Y).Blocked = 1
                Next i
    
            End If
            
            If .NumeroLayers(2) > 0 Then
                ReDim L2(1 To .NumeroLayers(2))
                Get #fh, , L2
    
                For i = 1 To .NumeroLayers(2)
                    Call Grh_Initialize(MapData(L2(i).X, L2(i).Y).Graphic(2), L2(i).GrhIndex)
                Next i
    
            End If
            
            If .NumeroLayers(3) > 0 Then
                ReDim L3(1 To .NumeroLayers(3))
                Get #fh, , L3
    
                For i = 1 To .NumeroLayers(3)
                    Call Grh_Initialize(MapData(L3(i).X, L3(i).Y).Graphic(3), L3(i).GrhIndex)
                Next i
    
            End If
            
            If .NumeroLayers(4) > 0 Then
                ReDim L4(1 To .NumeroLayers(4))
                Get #fh, , L4
    
                For i = 1 To .NumeroLayers(4)
                    Call Grh_Initialize(MapData(L4(i).X, L4(i).Y).Graphic(4), L4(i).GrhIndex)
                Next i
    
            End If
            
            If .NumeroTriggers > 0 Then
                ReDim Triggers(1 To .NumeroTriggers)
                Get #fh, , Triggers
    
                For i = 1 To .NumeroTriggers
                    MapData(Triggers(i).X, Triggers(i).Y).Trigger = Triggers(i).Trigger
                Next i
    
            End If
            
            If .NumeroParticulas > 0 Then
                ReDim Particulas(1 To .NumeroParticulas)
                Get #fh, , Particulas
    
                For i = 1 To .NumeroParticulas
                    MapData(Particulas(i).X, Particulas(i).Y).particle_group_index = General_Particle_Create(Particulas(i).Particula, Particulas(i).X, Particulas(i).Y)
                Next i
    
            End If
                
            If .NumeroLuces > 0 Then
                ReDim Luces(1 To .NumeroLuces)
    
                Dim p As Byte
    
                Get #fh, , Luces
                'For i = 1 To .NumeroLuces
                'For p = 0 To 3
                'MapData(Luces(i).X, Luces(i).y).base_light(p) = Luces(i).base_light(p)
                'If MapData(Luces(i).X, Luces(i).y).base_light(p) Then _
                 MapData(Luces(i).X, Luces(i).y).light_value(p) = Luces(i).light_value(p)
    
                'Next p
                'Next i
            End If
                
            If .NumeroOBJs > 0 Then
                ReDim Objetos(1 To .NumeroOBJs)
                Get #fh, , Objetos
    
                For i = 1 To .NumeroOBJs
                    MapData(Objetos(i).X, Objetos(i).Y).OBJInfo.objindex = Objetos(i).objindex
                    MapData(Objetos(i).X, Objetos(i).Y).OBJInfo.Amount = Objetos(i).ObjAmmount
    
                    If MapData(Objetos(i).X, Objetos(i).Y).OBJInfo.objindex > NumOBJs Then
                        Call Grh_Initialize(MapData(Objetos(i).X, Objetos(i).Y).ObjGrh, 20299)
                    Else
                        Call Grh_Initialize(MapData(Objetos(i).X, Objetos(i).Y).ObjGrh, ObjData(MapData(Objetos(i).X, Objetos(i).Y).OBJInfo.objindex).GrhIndex)
    
                    End If
    
                Next i
    
            End If
                
            If .NumeroNPCs > 0 Then
                ReDim NPCs(1 To .NumeroNPCs)
                Get #fh, , NPCs
    
                For i = 1 To .NumeroNPCs
    
                    If NPCs(i).NPCIndex > 0 Then
                        MapData(NPCs(i).X, NPCs(i).Y).NPCIndex = NPCs(i).NPCIndex
                        Call MakeChar(NextOpenChar(), NpcData(NPCs(i).NPCIndex).Body, NpcData(NPCs(i).NPCIndex).Head, NpcData(NPCs(i).NPCIndex).Heading, NPCs(i).X, NPCs(i).Y)
    
                    End If
    
                Next i
    
            End If
    
            If .NumeroTE > 0 Then
                ReDim TEs(1 To .NumeroTE)
                Get #fh, , TEs
    
                For i = 1 To .NumeroTE
                    MapData(TEs(i).X, TEs(i).Y).TileExit.Map = TEs(i).DestM
                    MapData(TEs(i).X, TEs(i).Y).TileExit.X = TEs(i).DestX
                    MapData(TEs(i).X, TEs(i).Y).TileExit.Y = TEs(i).DestY
                Next i
    
            End If
            
        End With

    Close fh

    For j = YMinMapSize To YMaxMapSize
        For i = XMinMapSize To XMaxMapSize

            If L1(i, j) > 0 Then
                Call Grh_Initialize(MapData(i, j).Graphic(1), L1(i, j))

            End If

        Next i
    Next j

    '*******************************
    'Render lights
    'Light_Render_All
    '*******************************

    'Call DibujarMiniMapa ' Radar
    
    'MapInfo_Cargar Map
    frmMain.lblMapVersion.Caption = MapInfo.MapVersion
    
    Call Pestanas(Map)
    
    ' Vacia el Deshacer
    Call modEdicion.Deshacer_Clear
    
    'Change mouse icon
    frmMain.MousePointer = 0
    
    Call CSMInfoCargar
    
    'Set changed flag
    MapInfo.Changed = 0

    MapaCargado = True
    
ErrorHandler:

    If fh <> 0 Then Close fh
    
    File = FreeFile
    Open App.Path & "\Logs.txt" For Output As #File
        Print #File, Err.Description
    Close #File

End Sub

Public Function Save_CSM(ByVal MapRoute As String) As Boolean

    On Error GoTo ErrorHandler

    Dim fh           As Integer
    Dim MH           As tMapHeader
    
    Dim Blqs()       As tDatosBloqueados

    Dim L1()         As Long
    Dim L2()         As tDatosGrh
    Dim L3()         As tDatosGrh
    Dim L4()         As tDatosGrh

    Dim Triggers()   As tDatosTrigger
    Dim Luces()      As tDatosLuces
    Dim Particulas() As tDatosParticulas
    Dim Objetos()    As tDatosObjs
    Dim NPCs()       As tDatosNPC
    Dim TEs()        As tDatosTE

    Dim i            As Integer
    Dim j            As Integer

    If FileExist(MapRoute, vbNormal) = True Then
        If MsgBox("Desea sobrescribir " & MapRoute & "?", vbCritical + vbYesNo) = vbNo Then
            Exit Function
        Else
            Kill MapRoute

        End If

    End If

    frmMain.MousePointer = 11

    ReDim L1(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize)

    For j = YMinMapSize To YMaxMapSize
        For i = XMinMapSize To XMaxMapSize

            With MapData(i, j)

                If .Blocked Then
                    MH.NumeroBloqueados = MH.NumeroBloqueados + 1
                    ReDim Preserve Blqs(1 To MH.NumeroBloqueados)
                    Blqs(MH.NumeroBloqueados).X = i
                    Blqs(MH.NumeroBloqueados).Y = j

                End If
            
                L1(i, j) = .Graphic(1).GrhIndex
            
                If .Graphic(2).GrhIndex > 0 Then
                    MH.NumeroLayers(2) = MH.NumeroLayers(2) + 1
                    ReDim Preserve L2(1 To MH.NumeroLayers(2))
                    L2(MH.NumeroLayers(2)).X = i
                    L2(MH.NumeroLayers(2)).Y = j
                    L2(MH.NumeroLayers(2)).GrhIndex = .Graphic(2).GrhIndex

                End If
            
                If .Graphic(3).GrhIndex > 0 Then
                    MH.NumeroLayers(3) = MH.NumeroLayers(3) + 1
                    ReDim Preserve L3(1 To MH.NumeroLayers(3))
                    L3(MH.NumeroLayers(3)).X = i
                    L3(MH.NumeroLayers(3)).Y = j
                    L3(MH.NumeroLayers(3)).GrhIndex = .Graphic(3).GrhIndex

                End If
            
                If .Graphic(4).GrhIndex > 0 Then
                    MH.NumeroLayers(4) = MH.NumeroLayers(4) + 1
                    ReDim Preserve L4(1 To MH.NumeroLayers(4))
                    L4(MH.NumeroLayers(4)).X = i
                    L4(MH.NumeroLayers(4)).Y = j
                    L4(MH.NumeroLayers(4)).GrhIndex = .Graphic(4).GrhIndex

                End If
            
                If .Trigger > 0 Then
                    MH.NumeroTriggers = MH.NumeroTriggers + 1
                    ReDim Preserve Triggers(1 To MH.NumeroTriggers)
                    Triggers(MH.NumeroTriggers).X = i
                    Triggers(MH.NumeroTriggers).Y = j
                    Triggers(MH.NumeroTriggers).Trigger = .Trigger

                End If
            
                If .particle_group_index > 0 Then
                    MH.NumeroParticulas = MH.NumeroParticulas + 1
                    ReDim Preserve Particulas(1 To MH.NumeroParticulas)
                    Particulas(MH.NumeroParticulas).X = i
                    Particulas(MH.NumeroParticulas).Y = j
                    Particulas(MH.NumeroParticulas).Particula = CLng(particle_group_list(.particle_group_index).stream_type)

                End If
           
                'If .base_light(0) Or .base_light(1) _
                '        Or .base_light(2) Or .base_light(3) Then
                '    MH.NumeroLuces = MH.NumeroLuces + 1
                '    ReDim Preserve Luces(1 To MH.NumeroLuces)
                '    Dim p As Byte
                '    Luces(MH.NumeroLuces).X = i
                '    Luces(MH.NumeroLuces).y = j
                '    For p = 0 To 3
                '        Luces(MH.NumeroLuces).base_light(p) = .base_light(p)
                '        If .base_light(p) Then _
                '            Luces(MH.NumeroLuces).light_value(p) = .light_value(p)
                '    Next p
                'End If
            
                If .OBJInfo.objindex > 0 Then
                    MH.NumeroOBJs = MH.NumeroOBJs + 1
                    ReDim Preserve Objetos(1 To MH.NumeroOBJs)
                    Objetos(MH.NumeroOBJs).objindex = .OBJInfo.objindex
                    Objetos(MH.NumeroOBJs).ObjAmmount = .OBJInfo.Amount
                    Objetos(MH.NumeroOBJs).X = i
                    Objetos(MH.NumeroOBJs).Y = j

                End If
            
                If .NPCIndex > 0 Then
                    MH.NumeroNPCs = MH.NumeroNPCs + 1
                    ReDim Preserve NPCs(1 To MH.NumeroNPCs)
                    NPCs(MH.NumeroNPCs).NPCIndex = .NPCIndex
                    NPCs(MH.NumeroNPCs).X = i
                    NPCs(MH.NumeroNPCs).Y = j

                End If
            
                If .TileExit.Map > 0 Then
                    MH.NumeroTE = MH.NumeroTE + 1
                    ReDim Preserve TEs(1 To MH.NumeroTE)
                    TEs(MH.NumeroTE).DestM = .TileExit.Map
                    TEs(MH.NumeroTE).DestX = .TileExit.X
                    TEs(MH.NumeroTE).DestY = .TileExit.Y
                    TEs(MH.NumeroTE).X = i
                    TEs(MH.NumeroTE).Y = j

                End If

            End With

        Next i
    Next j

    Call CSMInfoSave
          
    fh = FreeFile
    Open MapRoute For Binary As fh
    
    Put #fh, , MH
    Put #fh, , MapSize
    Put #fh, , MapDat
    Put #fh, , L1

    With MH

        If .NumeroBloqueados > 0 Then Put #fh, , Blqs

        If .NumeroLayers(2) > 0 Then Put #fh, , L2

        If .NumeroLayers(3) > 0 Then Put #fh, , L3

        If .NumeroLayers(4) > 0 Then Put #fh, , L4

        If .NumeroTriggers > 0 Then Put #fh, , Triggers

        If .NumeroParticulas > 0 Then Put #fh, , Particulas

        If .NumeroLuces > 0 Then Put #fh, , Luces

        If .NumeroOBJs > 0 Then Put #fh, , Objetos

        If .NumeroNPCs > 0 Then Put #fh, , NPCs

        If .NumeroTE > 0 Then Put #fh, , TEs

    End With

    Close fh

    Call Pestanas(MapRoute)

    'Change mouse icon
    frmMain.MousePointer = 0
    MapInfo.Changed = 0

    Save_CSM = True

    Exit Function

ErrorHandler:

    If fh <> 0 Then Close fh

End Function

Public Sub Convertir(ByVal PathMAP As String, ByVal PathCSM As String, ByVal Desde As Integer, ByVal Hasta As Integer)

    Dim i As Long
    
    For i = Desde To Hasta
        
        'Si existe Mapa i.map, lo convertimos.
        If FileExist(PathMAP & "\Mapa" & i & ".map", vbNormal) Then
        
            Debug.Print "Abriendo " & PathMAP & "\Mapa" & i & ".map"
            Call AbrirMapa(PathMAP & "\Mapa" & i & ".map", MapData, False)
        
            Debug.Print "Guardando " & CurMap & " .CSM: " & PathCSM & "\Mapa" & i & ".csm"
            Call Save_CSM(PathCSM & "\Mapa" & i & ".csm")
        
        End If
        
    Next
    
End Sub
