Attribute VB_Name = "modEdicion"
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
' modEdicion
'
' @remarks Funciones de Edicion
' @author gshaxor@gmail.com
' @version 0.1.38
' @date 20061016

Option Explicit

''
' Vacia el Deshacer
'
Public Sub Deshacer_Clear()
'*************************************************
'Author: ^[GS]^
'Last modified: 15/10/06
'*************************************************
Dim i As Integer
' Vacio todos los campos afectados
For i = 1 To maxDeshacer
    MapData_Deshacer_Info(i).Libre = True
Next
' no hay que deshacer
frmMain.mnuDeshacer.Enabled = False
End Sub

''
' Agrega un Deshacer
'
Public Sub Deshacer_Add(ByRef Desc As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 16/10/06
'*************************************************
If Not frmMain.mnuUtilizarDeshacer.Checked Then Exit Sub

Dim i As Integer
Dim x As Integer
Dim y As Integer

' Desplazo todos los deshacer uno hacia atras
For i = maxDeshacer To 2 Step -1
    For x = XMinMapSize To XMaxMapSize
        For y = YMinMapSize To YMaxMapSize
            MapData_Deshacer(i, x, y) = MapData_Deshacer(i - 1, x, y)
        Next y
    Next x
    
    MapData_Deshacer_Info(i) = MapData_Deshacer_Info(i - 1)
Next i

' Guardo los valores
For x = XMinMapSize To XMaxMapSize
    For y = YMinMapSize To YMaxMapSize
        MapData_Deshacer(1, x, y) = MapData(x, y)
    Next y
Next x

MapData_Deshacer_Info(1).Desc = Desc
MapData_Deshacer_Info(1).Libre = False
frmMain.mnuDeshacer.Caption = "&Deshacer (Último: " & MapData_Deshacer_Info(1).Desc & ")"
frmMain.mnuDeshacer.Enabled = True
End Sub

''
' Deshacer un paso del Deshacer
'
Public Sub Deshacer_Recover()
'*************************************************
'Author: ^[GS]^
'Last modified: 15/10/06
'*************************************************
Dim i As Integer
Dim f As Integer
Dim j As Integer
Dim Body As Integer
Dim Head As Integer
Dim Heading As Byte

If Not MapData_Deshacer_Info(1).Libre Then
    ' Aplico deshacer
    For f = XMinMapSize To XMaxMapSize
        For j = YMinMapSize To YMaxMapSize
            If (MapData(f, j).NPCIndex <> 0 And MapData(f, j).NPCIndex <> MapData_Deshacer(1, f, j).NPCIndex) Or (MapData(f, j).NPCIndex <> 0 And MapData_Deshacer(1, f, j).NPCIndex = 0) Then
                ' Si hay un NPC, y en el deshacer es otro lo borramos
                ' (o) Si aun no NPC y en el deshacer no esta
                MapData(f, j).NPCIndex = 0
                Call EraseChar(MapData(f, j).CharIndex)
            End If
            
            If MapData_Deshacer(1, f, j).NPCIndex <> 0 And MapData(f, j).NPCIndex = 0 Then
                ' Si ahi un NPC en el deshacer y en el no esta lo hacemos
                Body = NpcData(MapData_Deshacer(1, f, j).NPCIndex).Body
                Head = NpcData(MapData_Deshacer(1, f, j).NPCIndex).Head
                Heading = NpcData(MapData_Deshacer(1, f, j).NPCIndex).Heading
                Call MakeChar(NextOpenChar(), Body, Head, Heading, f, j)
            Else
                MapData(f, j) = MapData_Deshacer(1, f, j)
            End If
        Next j
    Next f
    
    MapData_Deshacer_Info(1).Libre = True
    ' Desplazo todos los deshacer uno hacia adelante
    For i = 1 To maxDeshacer - 1
        For f = XMinMapSize To XMaxMapSize
            For j = YMinMapSize To YMaxMapSize
                MapData_Deshacer(i, f, j) = MapData_Deshacer(i + 1, f, j)
            Next j
        Next f
        MapData_Deshacer_Info(i) = MapData_Deshacer_Info(i + 1)
    Next i
    
    ' borro el ultimo
    MapData_Deshacer_Info(maxDeshacer).Libre = True
    ' ahi para deshacer?
    If MapData_Deshacer_Info(1).Libre = True Then
        frmMain.mnuDeshacer.Caption = "&Deshacer (No hay nada que deshacer)"
        frmMain.mnuDeshacer.Enabled = False
    Else
        frmMain.mnuDeshacer.Caption = "&Deshacer (Último: " & MapData_Deshacer_Info(1).Desc & ")"
        frmMain.mnuDeshacer.Enabled = True
    End If
Else
    MsgBox "No hay acciones para deshacer", vbInformation
End If
End Sub

''
' Manda una advertencia de Edicion Critica
'
' @return   Nos devuelve si acepta o no el cambio

Private Function EditWarning() As Boolean
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If MsgBox(MSGDang, vbExclamation + vbYesNo) = vbNo Then
    EditWarning = True
Else
    EditWarning = False
End If
End Function

''
' Bloquea los Bordes del Mapa
'

Public Sub Bloquear_Bordes()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Dim x As Integer
Dim y As Integer

If Not MapaCargado Then Exit Sub

modEdicion.Deshacer_Add "Bloquear los bordes" ' Hago deshacer

For x = XMinMapSize To XMaxMapSize
    For y = YMinMapSize To YMaxMapSize
        If x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
            MapData(x, y).Blocked = 1
        End If
    Next y
Next x

'Set changed flag
MapInfo.Changed = 1
End Sub

''
' Coloca la superficie seleccionada al azar en el mapa
'

Public Sub Superficie_Azar()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

On Error Resume Next
Dim tx As Integer
Dim tY As Integer
Dim Cuantos As Integer
Dim k As Integer

If Not MapaCargado Then Exit Sub

Cuantos = InputBox("Cuantos Grh se deben poner en este mapa?", "Poner Grh Al Azar", 0)

If Cuantos > 0 Then
    Call modEdicion.Deshacer_Add("Insertar Superficie al Azar")  ' Hago deshacer
    
    For k = 1 To Cuantos
        tx = RandomNumber(MinXBorder, MaxXBorder)
        tY = RandomNumber(MinYBorder, MaxYBorder)
        
        Call InsertarGrh(tx, tY, frmConfigSup.MOSAICO.Value = vbChecked, bAutoCompletarSuperficies, frmMain.cInsertarBloqueo.Value, False)
    Next k
End If

'Set changed flag
MapInfo.Changed = 1

End Sub

''
' Coloca la superficie seleccionada en todos los bordes
'

Public Sub Superficie_Bordes()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

Dim x As Integer
Dim y As Integer

If Not MapaCargado Then Exit Sub

modEdicion.Deshacer_Add "Insertar Superficie en todos los bordes" ' Hago deshacer

For x = XMinMapSize To XMaxMapSize
    For y = YMinMapSize To YMaxMapSize
        If x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
            Call InsertarGrh(x, y, frmConfigSup.MOSAICO.Value = vbChecked, False, frmMain.cInsertarBloqueo.Value, False)
            
             'Erase NPCs
            Call QuitarNpc(x, y, False)

            'Erase Objs
            Call QuitarObjeto(x, y, False)

            'Clear exits
            Call QuitarTileExit(x, y, False)
        End If
    Next y
Next x

'Set changed flag
MapInfo.Changed = 1

End Sub

''
' Coloca la misma superficie seleccionada en todo el mapa
'

Public Sub Superficie_Todo()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

If EditWarning Then Exit Sub

Dim x As Integer
Dim y As Integer

If Not MapaCargado Then Exit Sub

modEdicion.Deshacer_Add "Insertar Superficie en todo el mapa" ' Hago deshacer

For x = XMinMapSize To XMaxMapSize
    For y = YMinMapSize To YMaxMapSize
        Call InsertarGrh(x, y, frmConfigSup.MOSAICO.Value = vbChecked, False, MapData(x, y).Blocked, False)
    Next y
Next x

'Set changed flag
MapInfo.Changed = 1

End Sub

''
' Modifica los bloqueos de todo mapa
'
' @param Valor Especifica el estado de Bloqueo que se asignara


Public Sub Bloqueo_Todo(ByVal Valor As Byte)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

If EditWarning Then Exit Sub

Dim x As Integer
Dim y As Integer

If Not MapaCargado Then Exit Sub

modEdicion.Deshacer_Add "Bloquear todo el mapa" ' Hago deshacer

For x = XMinMapSize To XMaxMapSize
    For y = YMinMapSize To YMaxMapSize
        MapData(x, y).Blocked = Valor
    Next y
Next x

'Set changed flag
MapInfo.Changed = 1

End Sub

''
' Borra todo el Mapa menos los Triggers
'

Public Sub Borrar_Mapa()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

If EditWarning Then Exit Sub

Dim x As Integer
Dim y As Integer

If Not MapaCargado Then Exit Sub

modEdicion.Deshacer_Add "Borrar todo el mapa menos Triggers" ' Hago deshacer

For x = XMinMapSize To XMaxMapSize
    For y = YMinMapSize To YMaxMapSize
        MapData(x, y).Graphic(1).grhIndex = 1
        'Change blockes status
        MapData(x, y).Blocked = 0

        'Erase layer 2 and 3
        MapData(x, y).Graphic(2).grhIndex = 0
        MapData(x, y).Graphic(3).grhIndex = 0
        MapData(x, y).Graphic(4).grhIndex = 0

        'Erase NPCs
        Call QuitarNpc(x, y, False)

        'Erase Objs
        Call QuitarObjeto(x, y, False)

        'Clear exits
        Call QuitarTileExit(x, y, False)
        
        InitGrh MapData(x, y).Graphic(1), 1
    Next y
Next x

'Set changed flag
MapInfo.Changed = 1
End Sub

''
' Elimita los NPCs del mapa
'
' @param Hostiles Indica si elimita solo hostiles o solo npcs no hostiles

Public Sub Quitar_NPCs(ByVal Hostiles As Boolean)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

If EditWarning Then Exit Sub

modEdicion.Deshacer_Add "Quitar todos los NPCs" & IIf(Hostiles = True, " Hostiles", "No Hostiles") ' Hago deshacer

Dim x As Integer
Dim y As Integer

For x = XMinMapSize To XMaxMapSize
    For y = YMinMapSize To YMaxMapSize
        If MapData(x, y).NPCIndex > 0 Then
            If (Hostiles And NpcData(MapData(x, y).NPCIndex).Hostile) Or ((Hostiles = False) And (NpcData(MapData(x, y).NPCIndex).Hostile = False)) Then
                Call EraseChar(MapData(x, y).CharIndex)
                MapData(x, y).NPCIndex = 0
            End If
        End If
    Next y
Next x

bRefreshRadar = True ' Radar

'Set changed flag
MapInfo.Changed = 1
End Sub

''
' Elimita todos los Objetos del mapa
'

Public Sub Quitar_Objetos()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

If EditWarning Then Exit Sub

modEdicion.Deshacer_Add "Quitar todos los Objetos" ' Hago deshacer

Dim x As Integer
Dim y As Integer

For x = XMinMapSize To XMaxMapSize
    For y = YMinMapSize To YMaxMapSize
        Call QuitarObjeto(x, y, False)
    Next y
Next x

'Set changed flag
MapInfo.Changed = 1
End Sub

''
' Elimina todos los Triggers del mapa
'

Public Sub Quitar_Triggers()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

If EditWarning Then Exit Sub

modEdicion.Deshacer_Add "Quitar todos los Triggers" ' Hago deshacer

Dim x As Integer
Dim y As Integer

For x = XMinMapSize To XMaxMapSize
    For y = YMinMapSize To YMaxMapSize
        If MapData(x, y).Trigger > 0 Then
            MapData(x, y).Trigger = 0
        End If
    Next y
Next x

'Set changed flag
MapInfo.Changed = 1
End Sub

''
' Elimita todos los translados del mapa
'

Public Sub Quitar_Translados()
'*************************************************
'Author: ^[GS]^
'Last modified: 16/10/06
'*************************************************

If EditWarning Then Exit Sub

modEdicion.Deshacer_Add "Quitar todos los Translados" ' Hago deshacer

Dim x As Integer
Dim y As Integer

For x = XMinMapSize To XMaxMapSize
    For y = YMinMapSize To YMaxMapSize
        Call QuitarTileExit(x, y, False)
    Next y
Next x

'Set changed flag
MapInfo.Changed = 1

End Sub

''
' Elimita todo lo que se encuentre en los bordes del mapa
'

Public Sub Quitar_Bordes()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

If EditWarning Then Exit Sub

'*****************************************************************
'Clears a border in a room with current GRH
'*****************************************************************

Dim x As Integer
Dim y As Integer

If Not MapaCargado Then Exit Sub

modEdicion.Deshacer_Add "Quitar todos los Bordes" ' Hago deshacer

For x = XMinMapSize To XMaxMapSize
    For y = YMinMapSize To YMaxMapSize
        If x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
            MapData(x, y).Graphic(1).grhIndex = 1
            InitGrh MapData(x, y).Graphic(1), 1
            MapData(x, y).Blocked = 0
            
             'Erase NPCs
             Call QuitarNpc(x, y, False)

            'Erase Objs
            Call QuitarObjeto(x, y, False)

            'Clear exits
            Call QuitarTileExit(x, y, False)
            
            ' Triggers
            MapData(x, y).Trigger = 0
        End If
    Next y
Next x

'Set changed flag
MapInfo.Changed = 1

End Sub

''
' Elimita una capa completa del mapa
'
' @param Capa Especifica la capa


Public Sub Quitar_Capa(ByVal Capa As Byte)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

If EditWarning Then Exit Sub

'*****************************************************************
'Clears one layer
'*****************************************************************

Dim x As Integer
Dim y As Integer

If Not MapaCargado Then Exit Sub

modEdicion.Deshacer_Add "Quitar Capa " & Capa ' Hago deshacer

For x = XMinMapSize To XMaxMapSize
    For y = YMinMapSize To YMaxMapSize
        If Capa = 1 Then
            MapData(x, y).Graphic(Capa).grhIndex = 1
        Else
            MapData(x, y).Graphic(Capa).grhIndex = 0
        End If
    Next y
Next x

'Set changed flag
MapInfo.Changed = 1
End Sub

''
' Acciona la operacion al hacer doble click en una posicion del mapa
'
' @param tX Especifica la posicion X en el mapa
' @param tY Espeficica la posicion Y en el mapa

Sub DobleClick(tx As Integer, tY As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************
' Translados
Dim tTrans As WorldPos

tTrans = MapData(tx, tY).TileExit

If tTrans.Map > 0 Then
    If LenB(frmMain.Dialog.FileName) <> 0 Then
        If FileExist(PATH_Save & NameMap_Save & tTrans.Map & ".map", vbArchive) = True Then
            Call modMapIO.NuevoMapa
            frmMain.Dialog.FileName = PATH_Save & NameMap_Save & tTrans.Map & ".map"
            modMapIO.AbrirMapa frmMain.Dialog.FileName, MapData
            UserPos.x = tTrans.x
            UserPos.y = tTrans.y
            
            If WalkMode Then
                MoveCharbyPos UserCharIndex, UserPos.x, UserPos.y
                CharList(UserCharIndex).Heading = SOUTH
            End If
            
            frmMain.mnuReAbrirMapa.Enabled = True
        End If
    End If
End If
End Sub

''
' Realiza una operacion de edicion aislada sobre el mapa
'
' @param Button Indica el estado del Click del mouse
' @param tX Especifica la posicion X en el mapa
' @param tY Especifica la posicion Y en el mapa

Sub ClickEdit(ByVal Button As Integer, ByVal tx As Integer, ByVal tY As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

    Dim NPCIndex As Integer
    Dim objindex As Integer
    Dim Amount As Integer
    Dim x As Integer
    Dim y As Integer
    Dim Map As Integer
    
    If tY < YMinMapSize Or tY > YMaxMapSize Then Exit Sub
    If tx < XMinMapSize Or tx > XMaxMapSize Then Exit Sub
    
    ' Pasando sobre :P
    SobreY = tY
    SobreX = tx
    
    'Right
    If Button = vbRightButton Then
        Call GetMapData(tx, tY)
    'Left click
    ElseIf Button = vbLeftButton Then
        'Erase 2-3
        If frmMain.cQuitarEnTodasLasCapas.Value Then
            Call QuitarCapasMedias(tx, tY)
        'Borrar "esta" Capa
        ElseIf frmMain.cQuitarEnEstaCapa.Value Then
            Call QuitarEstaCapa(tx, tY)
        '************** Place grh
        ElseIf bSelectSup Then
            Call InsertarGrh(tx, tY, frmConfigSup.MOSAICO.Value = vbChecked, bAutoCompletarSuperficies, MapData(tx, tY).Blocked)
        '************** Place blocked tile
        ElseIf frmMain.cInsertarBloqueo.Value Then
            Call InsertarBloq(tx, tY)
        ElseIf frmMain.cQuitarBloqueo.Value Then
            Call QuitarBloq(tx, tY)
        '************** Place exit
        ElseIf frmMain.cInsertarTrans.Value Then
            Map = Val(frmMain.tTMapa.Text)
            x = Val(frmMain.tTX.Text)
            y = Val(frmMain.tTY.Text)
            
            If (Map < 0) Or (Map > NumMaps) Then
                MsgBox "Valor de Mapa invalido", vbCritical + vbOKOnly
                Exit Sub
            ElseIf (x < MinXBorder) Or (x > MaxXBorder) Then
                MsgBox "Valor de X invalido", vbCritical + vbOKOnly
                Exit Sub
            ElseIf (y < MinYBorder) Or (y > MaxYBorder) Then
                MsgBox "Valor de Y invalido", vbCritical + vbOKOnly
                Exit Sub
            End If
            
            If frmMain.cInsertarTransOBJ.Value Then _
                Call InsertarObjTranslado(tx, tY)
            
            If frmMain.cUnionManual.Value Then
                Call InsertarUnionManual(tx, tY, Map)
            Else
                Call InsertarTileExit(tx, tY, x, y, Map)
            End If
        ElseIf frmMain.cQuitarTrans.Value Then
            Call QuitarTileExit(tx, tY)
        '************** Place NPC
        ElseIf frmMain.cInsertarFunc(0).Value Then
            NPCIndex = Val(frmMain.cNumFunc(0).Text)
            
            Call InsertarNpc(tx, tY, NPCIndex)
        ElseIf frmMain.cInsertarFunc(1).Value Then
            NPCIndex = Val(frmMain.cNumFunc(1).Text)
                
            Call InsertarNpc(tx, tY, NPCIndex)
        ElseIf frmMain.cQuitarFunc(0).Value Or frmMain.cQuitarFunc(1).Value Then
            Call QuitarNpc(tx, tY)
        ' ***************** Control de Funcion de Objetos *****************
        ElseIf frmMain.cInsertarFunc(2).Value = True Then ' Insertar Objeto
            objindex = frmMain.cNumFunc(2).Text
            Amount = Val(frmMain.cCantFunc(2).Text)
            
            Call InsertarObjeto(tx, tY, objindex, Amount)
        ElseIf frmMain.cQuitarFunc(2).Value Then  ' Quitar Objeto
            Call QuitarObjeto(tx, tY)
        ' ***************** Control de Funcion de Triggers *****************
        ElseIf frmMain.cInsertarTrigger.Value Then ' Insertar Trigger
            Call InsertarTrigger(tx, tY, frmMain.lListado(4).ListIndex)
        ElseIf frmMain.cQuitarTrigger.Value Then  ' Quitar Trigger
            Call InsertarTrigger(tx, tY, 0)
        End If
    End If
End Sub

Public Sub GetMapData(ByVal x As Byte, ByVal y As Byte)
With MapData(x, y)
    ' Posicion
    frmMain.StatTxt.Text = frmMain.StatTxt.Text & ENDL & ENDL & "Posición " & x & "," & y
    
    ' Bloqueos
    If .Blocked = 1 Then frmMain.StatTxt.Text = frmMain.StatTxt.Text & " (BLOQ)"
    
    ' Translados
    If .TileExit.Map > 0 Then
        If frmMain.mnuAutoCapturarTranslados.Checked Then
            frmMain.tTMapa.Text = .TileExit.Map
            frmMain.tTX.Text = .TileExit.x
            frmMain.tTY = .TileExit.y
        End If
        
        frmMain.StatTxt.Text = frmMain.StatTxt.Text & " (Trans.: " & .TileExit.Map & "," & .TileExit.x & "," & .TileExit.y & ")"
    End If
    
    ' NPCs
    If .NPCIndex > 0 Then
        If NpcData(.NPCIndex).Hostile Then
            frmMain.StatTxt.Text = frmMain.StatTxt.Text & " (NPC-Hostil: " & .NPCIndex & " - " & NpcData(.NPCIndex).name & ")"
        Else
            frmMain.StatTxt.Text = frmMain.StatTxt.Text & " (NPC: " & .NPCIndex & " - " & NpcData(.NPCIndex).name & ")"
        End If
    End If
    
    ' OBJs
    If .OBJInfo.objindex > 0 Then
        frmMain.StatTxt.Text = frmMain.StatTxt.Text & " (Obj: " & .OBJInfo.objindex & " - " & ObjData(.OBJInfo.objindex).name & " - Cant.:" & .OBJInfo.Amount & ")"
    End If
    
    ' Capas
    frmMain.StatTxt.Text = frmMain.StatTxt.Text & ENDL & "Capa1: " & .Graphic(1).grhIndex & " - Capa2: " & .Graphic(2).grhIndex & " - Capa3: " & .Graphic(3).grhIndex & " - Capa4: " & .Graphic(4).grhIndex
    
    If frmMain.mnuAutoCapturarSuperficie.Checked And (Not bSelectSup) Then
        If .Graphic(4).grhIndex <> 0 Then
            frmMain.cCapas.Text = 4
            frmMain.cGrh.Text = .Graphic(4).grhIndex
        ElseIf .Graphic(3).grhIndex <> 0 Then
            frmMain.cCapas.Text = 3
            frmMain.cGrh.Text = .Graphic(3).grhIndex
        ElseIf .Graphic(2).grhIndex <> 0 Then
            frmMain.cCapas.Text = 2
            frmMain.cGrh.Text = .Graphic(2).grhIndex
        ElseIf .Graphic(1).grhIndex <> 0 Then
            frmMain.cCapas.Text = 1
            frmMain.cGrh.Text = .Graphic(1).grhIndex
        End If
    End If
    
    ' Limpieza
    If Len(frmMain.StatTxt.Text) > 4000 Then
        frmMain.StatTxt.Text = mid$(frmMain.StatTxt.Text, InStr(1000, frmMain.StatTxt.Text, ENDL & ENDL) + 4) '4 = len(ENDL & ENDL)
    End If
    
    frmMain.StatTxt.SelStart = Len(frmMain.StatTxt.Text)
End With
End Sub

Public Sub SelectTiles(ByVal Up As Boolean, ByVal tx As Integer, ByVal tY As Integer)
Dim x As Long
Dim y As Long

If (tx < XMinMapSize) Or (tY < YMinMapSize) Or (tx > XMaxMapSize) Or (tY > YMaxMapSize) Then Exit Sub

If MaxSelectX Then 'Si tenemos un max, tenemos el otro, y también tenemos los min
    For y = MinSelectY To MaxSelectY
        For x = MinSelectX To MaxSelectX
            MapData(x, y).Select = 0
        Next x
    Next y
End If

If GetAsyncKeyState(vbKeyShift) < 0 Then
    If Up And (MouseDownX = tx) And (MouseDownY = tY) Then 'Esto quiere decir que no hubo drag
        If MinSelectX = 0 Then
            MinSelectX = tx
            MaxSelectX = tx
            FirstSelectX = tx
            
            MinSelectY = tY 'Si no tenemos minX, tampoco hay minY
            MaxSelectY = tY
            FirstSelectY = tY
        Else 'Esto seria el segundo click
            If tx < FirstSelectX Then
                MinSelectX = tx
                MaxSelectX = FirstSelectX
            Else
                MinSelectX = FirstSelectX
                MaxSelectX = tx
            End If
            
            If tY < FirstSelectY Then
                MinSelectY = tY
                MaxSelectY = FirstSelectY
            Else
                MinSelectY = FirstSelectY
                MaxSelectY = tY
            End If
        End If
    ElseIf (MouseDownX <> tx) Or (MouseDownY <> tY) Then
        If MouseDownX < XMinMapSize Then MouseDownX = XMinMapSize
        If MouseDownX > XMaxMapSize Then MouseDownX = XMaxMapSize
        If MouseDownY < YMinMapSize Then MouseDownY = YMinMapSize
        If MouseDownY > YMaxMapSize Then MouseDownY = YMaxMapSize
        
        FirstSelectX = MouseDownX
        FirstSelectY = MouseDownY
            
        If tx > MouseDownX Then
            MinSelectX = MouseDownX
            MaxSelectX = tx
        Else
            MinSelectX = tx
            MaxSelectX = MouseDownX
        End If
        
        If tY > MouseDownY Then
            MinSelectY = MouseDownY
            MaxSelectY = tY
        Else
            MinSelectY = tY
            MaxSelectY = MouseDownY
        End If
    End If
    
    If MaxSelectX Then
        For y = MinSelectY To MaxSelectY
            For x = MinSelectX To MaxSelectX
                MapData(x, y).Select = 1
            Next x
        Next y
    End If
ElseIf Up Then
    MinSelectX = 0
    MaxSelectX = 0
    MinSelectY = 0
    MaxSelectY = 0
    FirstSelectX = 0
    FirstSelectY = 0
End If
End Sub

Public Sub AplicarBloqueos()
Dim x As Long
Dim y As Long

If MaxSelectX Then 'Si tenemos un max, tenemos el otro, y también tenemos los min
    Call modEdicion.Deshacer_Add("Bloquear selección")
    
    For y = MinSelectY To MaxSelectY
        For x = MinSelectX To MaxSelectX
            MapData(x, y).Blocked = 1
        Next x
    Next y
    
    MapInfo.Changed = 1
End If
End Sub

Public Sub AplicarSeleccionado()
Dim x As Long
Dim y As Long
Dim NPCIndex As Integer
Dim objindex As Integer
Dim Amount As Integer
Dim tx As Integer
Dim tY As Integer
Dim Map As Integer

If MaxSelectX Then 'Si tenemos un max, tenemos el otro, y también tenemos los min
    If frmMain.cInsertarTrans.Value Then
        Map = Val(frmMain.tTMapa.Text)
        tx = Val(frmMain.tTX.Text)
        tY = Val(frmMain.tTY.Text)
        
        If (Map < 0) Or (Map > NumMaps) Then
            MsgBox "Valor de Mapa invalido", vbCritical + vbOKOnly
            Exit Sub
        ElseIf (tx < MinXBorder) Or (tx > MaxXBorder) Then
            MsgBox "Valor de X invalido", vbCritical + vbOKOnly
            Exit Sub
        ElseIf (tY < MinYBorder) Or (tY > MaxYBorder) Then
            MsgBox "Valor de Y invalido", vbCritical + vbOKOnly
            Exit Sub
        End If
    End If
                
    For y = MinSelectY To MaxSelectY
        For x = MinSelectX To MaxSelectX
            If frmMain.cQuitarEnTodasLasCapas.Value Then
                Call QuitarCapasMedias(x, y, False)
            ElseIf frmMain.cQuitarEnEstaCapa.Value Then
                Call QuitarEstaCapa(x, y, False)
            ElseIf bSelectSup Then
                Call InsertarGrh(x, y, frmConfigSup.MOSAICO.Value = vbChecked, bAutoCompletarSuperficies, MapData(x, y).Blocked, False)
            ElseIf frmMain.cInsertarBloqueo.Value Then
                Call InsertarBloq(x, y, False)
            ElseIf frmMain.cQuitarBloqueo.Value Then
                Call QuitarBloq(x, y, False)
            ElseIf frmMain.cInsertarTrans.Value Then
                If frmMain.cInsertarTransOBJ.Value Then _
                    Call InsertarObjTranslado(x, y, False)
                
                If frmMain.cUnionManual.Value Then
                    Call InsertarUnionManual(x, y, Map, False)
                Else
                    Call InsertarTileExit(x, y, tx, tY, Map, False)
                End If
            ElseIf frmMain.cQuitarTrans.Value Then
                Call QuitarTileExit(x, y, False)
            ElseIf frmMain.cInsertarFunc(0).Value Then
                NPCIndex = Val(frmMain.cNumFunc(0).Text)
                
                Call InsertarNpc(x, y, NPCIndex, False)
            ElseIf frmMain.cInsertarFunc(1).Value Then
                NPCIndex = Val(frmMain.cNumFunc(1).Text)
                    
                Call InsertarNpc(x, y, NPCIndex, False)
            ElseIf frmMain.cQuitarFunc(0).Value Or frmMain.cQuitarFunc(1).Value Then
                Call QuitarNpc(x, y, False)
            ElseIf frmMain.cInsertarFunc(2).Value = True Then
                objindex = frmMain.cNumFunc(2).Text
                Amount = Val(frmMain.cCantFunc(2).Text)
                
                Call InsertarObjeto(x, y, objindex, Amount, False)
            ElseIf frmMain.cQuitarFunc(2).Value Then
                Call QuitarObjeto(x, y, False)
            ElseIf frmMain.cInsertarTrigger.Value Then
                Call InsertarTrigger(x, y, frmMain.lListado(4).ListIndex, False)
            ElseIf frmMain.cQuitarTrigger.Value Then
                Call InsertarTrigger(x, y, 0, False)
            End If
        Next x
    Next y
End If
End Sub

Public Sub QuitarCapasMedias(ByVal x As Byte, ByVal y As Byte, Optional ByVal ConDeshacer As Boolean = True)
Dim i As Byte
    
If ConDeshacer Then _
    Call modEdicion.Deshacer_Add("Quitar capas medias")
    
For i = 2 To 3
    MapData(x, y).Graphic(i).grhIndex = 0
Next i

MapInfo.Changed = 1
End Sub

Public Function QuitarEstaCapa(ByVal x As Byte, ByVal y As Byte, Optional ByVal ConDeshacer As Boolean = True) As Boolean
If MapData(x, y).Graphic(CurLayer).grhIndex <> 0 Then
    If ConDeshacer Then _
        Call modEdicion.Deshacer_Add("Quitar capa " & CurLayer)
    
    MapData(x, y).Graphic(CurLayer).grhIndex = 0
    MapInfo.Changed = 1
End If
End Function

Public Sub InsertarGrh(ByVal x As Byte, ByVal y As Byte, ByVal MOSAICO As Boolean, ByVal AutoCompletar As Boolean, ByVal Bloq As Boolean, Optional ByVal ConDeshacer As Boolean = True)
Dim grhIndex As Integer
Dim OffsetX As Long
Dim OffsetY As Long

If MOSAICO And AutoCompletar Then
    If ConDeshacer Then _
        Call modEdicion.Deshacer_Add("Insertar superficie autocompletada. Capa " & CurLayer)
        
    For OffsetX = 0 To mAncho - 1
        For OffsetY = 0 To MAlto - 1
            grhIndex = CurrentGrh(((x + OffsetX + DespX) Mod mAncho) + 1, ((y + OffsetY + DespY) Mod MAlto) + 1).grhIndex
                
            If Bloq Then
                MapData(x + OffsetX, y + OffsetY).Blocked = 1
            Else
                MapData(x + OffsetX, y + OffsetY).Blocked = 0
            End If
            
            InitGrh MapData(x + OffsetX, y + OffsetY).Graphic(CurLayer), grhIndex
        Next OffsetY
    Next OffsetX
    
    MapInfo.Changed = 1
Else
    If MOSAICO Then
        grhIndex = CurrentGrh(((x + DespX) Mod mAncho) + 1, ((y + DespY) Mod MAlto) + 1).grhIndex
    Else
        grhIndex = CurrentGrh(0).grhIndex
    End If
    
    With MapData(x, y)
        If .Graphic(CurLayer).grhIndex <> grhIndex Then
            If ConDeshacer Then _
                Call modEdicion.Deshacer_Add("Insertar superficie. Capa " & CurLayer)
                
            If Bloq Then
                .Blocked = 1
            Else
                .Blocked = 0
            End If
            
            InitGrh .Graphic(CurLayer), grhIndex
            
            MapInfo.Changed = 1
        End If
    End With
End If
End Sub

Public Sub InsertarBloq(ByVal x As Byte, ByVal y As Byte, Optional ByVal ConDeshacer As Boolean = True)
If MapData(x, y).Blocked <> 1 Then
    If ConDeshacer Then _
        Call modEdicion.Deshacer_Add("Insertar bloqueo")
        
    MapData(x, y).Blocked = 1
    MapInfo.Changed = 1 'Set changed flag
End If
End Sub

Public Sub QuitarBloq(ByVal x As Byte, ByVal y As Byte, Optional ByVal ConDeshacer As Boolean = True)
If MapData(x, y).Blocked <> 0 Then
    If ConDeshacer Then _
        Call modEdicion.Deshacer_Add("Quitar bloqueo")
        
    MapData(x, y).Blocked = 0
    MapInfo.Changed = 1 'Set changed flag
End If
End Sub

Public Sub InsertarObjTranslado(ByVal x As Byte, ByVal y As Byte, Optional ByVal ConDeshacer As Boolean = True)
With MapData(x, y)
    If (Cfg_TrOBJ > 0) And (Cfg_TrOBJ <= NumOBJs) Then
        If ObjData(Cfg_TrOBJ).ObjType = 19 Then
            If ConDeshacer Then _
                Call modEdicion.Deshacer_Add("Insertar Objeto de Translado")
            
            InitGrh .ObjGrh, ObjData(Cfg_TrOBJ).grhIndex
            .OBJInfo.objindex = Cfg_TrOBJ
            .OBJInfo.Amount = 1
            
            MapInfo.Changed = 1 'Set changed flag
        End If
    End If
End With
End Sub

Public Sub InsertarUnionManual(ByVal x As Byte, ByVal y As Byte, ByVal TargetMap As Integer, Optional ByVal ConDeshacer As Boolean = True)
With MapData(x, y).TileExit
    If ConDeshacer Then _
        Call modEdicion.Deshacer_Add("Insertar Translado de Union Manual")
    
    If x >= MaxXBorder Then ' 21 ' derecha
        .x = MinXBorder + 1
        .y = y
        .Map = TargetMap
    ElseIf x <= MinYBorder Then ' 9 ' izquierda
        .x = MaxXBorder - 1
        .y = y
        .Map = TargetMap
    End If
    
    If y >= MaxYBorder Then ' 94 '''' hacia abajo
        .x = x
        .y = MinYBorder + 1
        .Map = TargetMap
    ElseIf y <= MinYBorder Then ''' hacia arriba
        .x = x
        .y = MaxYBorder - 1
        .Map = TargetMap
    End If
    
    MapInfo.Changed = 1 'Set changed flag
End With
End Sub

Public Sub InsertarTileExit(ByVal x As Byte, ByVal y As Byte, ByVal TargetX As Byte, ByVal TargetY As Byte, ByVal TargetMap As Integer, Optional ByVal ConDeshacer As Boolean = True)
With MapData(x, y).TileExit
    If ConDeshacer Then _
        Call modEdicion.Deshacer_Add("Insertar Translado")
        
    .x = TargetX
    .y = TargetY
    .Map = TargetMap
    
    MapInfo.Changed = 1 'Set changed flag
End With
End Sub

Public Sub QuitarTileExit(ByVal x As Byte, ByVal y As Byte, Optional ByVal ConDeshacer As Boolean)
With MapData(x, y).TileExit
    If ConDeshacer Then _
        Call modEdicion.Deshacer_Add("Quitar Translado")
        
    .Map = 0
    .x = 0
    .y = 0
    
    MapInfo.Changed = 1 'Set changed flag
End With
End Sub

Public Sub InsertarNpc(ByVal x As Byte, ByVal y As Byte, ByVal NPCIndex As Integer, Optional ByVal ConDeshacer As Boolean = True)
Dim Body As Integer
Dim Head As Integer
Dim Heading As Integer

With MapData(x, y)
    If NPCIndex <> .NPCIndex Then
        If .NPCIndex > 0 Then _
            Call EraseChar(.CharIndex)
        
        If ConDeshacer Then _
            Call modEdicion.Deshacer_Add("Insertar NPC " & IIf(NpcData(NPCIndex).Hostile, "Hostil", "No Hostil"))
            
        .NPCIndex = NPCIndex
        
        Body = NpcData(NPCIndex).Body
        Head = NpcData(NPCIndex).Head
        Heading = NpcData(NPCIndex).Heading
        
        Call MakeChar(NextOpenChar(), Body, Head, Heading, x, y)
        MapInfo.Changed = 1 'Set changed flag
    End If
End With
End Sub

Public Sub QuitarNpc(ByVal x As Byte, ByVal y As Byte, Optional ByVal ConDeshacer As Boolean = True)
With MapData(x, y)
    If .NPCIndex > 0 Then
        If ConDeshacer Then _
            Call modEdicion.Deshacer_Add("Quitar NPC")
        
        .NPCIndex = 0
        Call EraseChar(.CharIndex)
        
        MapInfo.Changed = 1 'Set changed flag
    End If
End With
End Sub

Public Sub InsertarObjeto(ByVal x As Byte, ByVal y As Byte, ByVal objindex As Integer, ByVal Amount As Integer, Optional ByVal ConDeshacer As Boolean = True)
With MapData(x, y)
    If objindex > 0 Then
        If .OBJInfo.objindex <> objindex Or .OBJInfo.Amount <> Amount Then
            If ConDeshacer Then _
                Call modEdicion.Deshacer_Add("Insertar Objeto")
                
            .OBJInfo.objindex = objindex
            .OBJInfo.Amount = Amount
            
            Select Case ObjData(objindex).ObjType
                Case 4, 8, 10, 22 ' Arboles, Carteles, Foros, Yacimientos
                    .Graphic(3) = .ObjGrh
            End Select
            
            InitGrh .ObjGrh, ObjData(objindex).grhIndex
            
            MapInfo.Changed = 1 'Set changed flag
        End If
    End If
End With
End Sub

Public Sub QuitarObjeto(ByVal x As Byte, ByVal y As Byte, Optional ByVal ConDeshacer As Boolean)
With MapData(x, y)
    If .OBJInfo.objindex <> 0 Then
        If ConDeshacer Then _
            Call modEdicion.Deshacer_Add("Quitar objeto")
            
        If .Graphic(3).grhIndex = .ObjGrh.grhIndex Then .Graphic(3).grhIndex = 0
        
        .ObjGrh.grhIndex = 0
        .OBJInfo.objindex = 0
        .OBJInfo.Amount = 0
        
        MapInfo.Changed = 1
    End If
End With
End Sub

Public Sub InsertarTrigger(ByVal x As Byte, ByVal y As Byte, ByVal Trigger As Byte, Optional ByVal ConDeshacer As Boolean)
With MapData(x, y)
    If .Trigger <> Trigger Then
        If ConDeshacer Then _
            Call modEdicion.Deshacer_Add("Insertar Trigger " & Trigger)
            
        .Trigger = Trigger
        MapInfo.Changed = 1 'Set changed flag
    End If
End With
End Sub
