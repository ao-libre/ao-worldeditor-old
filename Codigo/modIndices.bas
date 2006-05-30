Attribute VB_Name = "modIndices"
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
' modIndices
'
' @remarks Funciones Especificas al Trabajo con Indices
' @author gshaxor@gmail.com
' @version 0.1.04
' @date 20060520

Option Explicit


''
' Carga los indices de Graficos
'

Public Sub CargarIndicesDeGraficos()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************

On Error GoTo ErrorHandler

    Dim Grh As Integer
    Dim Frame As Integer
    Dim TempInt As Integer
    
    If FileExist(DirIndex & "Graficos.ind", vbArchive) = False Then
        MsgBox "Falta el archivo 'graficos.ind' en " & DirIndex, vbCritical
        End
    End If
    
    'Resize arrays
    ReDim GrhData(1 To MaxGrhs) As GrhData
    
    'Open files
    Open DirIndex & "Graficos.ind" For Binary Access Read As #1
    Seek #1, 1
    
    Get #1, , MiCabecera
    Get #1, , TempInt
    Get #1, , TempInt
    Get #1, , TempInt
    Get #1, , TempInt
    Get #1, , TempInt
    
    'Fill Grh List
    
    'Get first Grh Number
    Get #1, , Grh
    
    Do Until Grh <= 0
            
        'Get number of frames
        Get #1, , GrhData(Grh).NumFrames
        If GrhData(Grh).NumFrames <= 0 Then GoTo ErrorHandler
        
        If GrhData(Grh).NumFrames > 1 Then
        
            'Read a animation GRH set
            For Frame = 1 To GrhData(Grh).NumFrames
            
                Get #1, , GrhData(Grh).Frames(Frame)
                If GrhData(Grh).Frames(Frame) <= 0 Or GrhData(Grh).Frames(Frame) > MaxGrhs Then GoTo ErrorHandler
            
            Next Frame
        
            Get #1, , GrhData(Grh).Speed
            If GrhData(Grh).Speed <= 0 Then GoTo ErrorHandler
            
            'Compute width and height
            GrhData(Grh).pixelHeight = GrhData(GrhData(Grh).Frames(1)).pixelHeight
            If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler
            
            GrhData(Grh).pixelWidth = GrhData(GrhData(Grh).Frames(1)).pixelWidth
            If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler
            
            GrhData(Grh).TileWidth = GrhData(GrhData(Grh).Frames(1)).TileWidth
            If GrhData(Grh).TileWidth <= 0 Then GoTo ErrorHandler
            
            GrhData(Grh).TileHeight = GrhData(GrhData(Grh).Frames(1)).TileHeight
            If GrhData(Grh).TileHeight <= 0 Then GoTo ErrorHandler
        
        Else
        
            'Read in normal GRH data
            Get #1, , GrhData(Grh).FileNum
            If GrhData(Grh).FileNum <= 0 Then GoTo ErrorHandler
            
            Get #1, , GrhData(Grh).sX
            If GrhData(Grh).sX < 0 Then GoTo ErrorHandler
            
            Get #1, , GrhData(Grh).sY
            If GrhData(Grh).sY < 0 Then GoTo ErrorHandler
                
            Get #1, , GrhData(Grh).pixelWidth
            If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler
            
            Get #1, , GrhData(Grh).pixelHeight
            If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler
            
            'Compute width and height
            GrhData(Grh).TileWidth = GrhData(Grh).pixelWidth / TilePixelHeight
            GrhData(Grh).TileHeight = GrhData(Grh).pixelHeight / TilePixelWidth
            
            GrhData(Grh).Frames(1) = Grh
                
        End If
    
        'Get Next Grh Number
        Get #1, , Grh
    
    Loop
    '************************************************
    
    Close #1
    
    Exit Sub

ErrorHandler:
Close #1
    MsgBox "Error al intentar cargar el Graficos " & Grh & " de graficos.ind en " & DirIndex & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub

''
' Carga los indices de Superficie
'

Public Sub CargarIndicesSuperficie()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************

On Error GoTo Fallo
    If FileExist(IniPath & "GrhIndex\indices.ini", vbArchive) = False Then
        MsgBox "Falta el archivo 'GrhIndex\indices.ini'", vbCritical
        End
    End If
    Dim Leer As New clsIniReader
    Dim i As Integer
    Leer.Initialize IniPath & "GrhIndex\indices.ini"
    MaxSup = Leer.GetValue("INIT", "Referencias")
    ReDim SupData(MaxSup) As SupData
    frmMain.lListado(0).Clear
    For i = 0 To MaxSup
        SupData(i).name = Leer.GetValue("REFERENCIA" & i, "Nombre")
        SupData(i).Grh = Val(Leer.GetValue("REFERENCIA" & i, "GrhIndice"))
        SupData(i).Width = Val(Leer.GetValue("REFERENCIA" & i, "Ancho"))
        SupData(i).Height = Val(Leer.GetValue("REFERENCIA" & i, "Alto"))
        SupData(i).Block = IIf(Val(Leer.GetValue("REFERENCIA" & i, "Bloquear")) = 1, True, False)
        SupData(i).Capa = Val(Leer.GetValue("REFERENCIA" & i, "Capa"))
        frmMain.lListado(0).AddItem SupData(i).name & " - #" & i
    Next
    DoEvents
    Exit Sub
Fallo:
    MsgBox "Error al intentar cargar el indice " & i & " de GrhIndex\indices.ini" & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly
End Sub

''
' Carga los indices de Objetos
'

Public Sub CargarIndicesOBJ()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

On Error GoTo Fallo
    If FileExist(DirDats & "\OBJ.dat", vbArchive) = False Then
        MsgBox "Falta el archivo 'OBJ.dat' en " & DirDats, vbCritical
        End
    End If
    Dim Obj As Integer
    Dim Leer As New clsIniReader
    Call Leer.Initialize(DirDats & "\OBJ.dat")
    frmMain.lListado(3).Clear
    NumOBJs = Val(Leer.GetValue("INIT", "NumOBJs"))
    ReDim ObjData(1 To NumOBJs) As ObjData
    For Obj = 1 To NumOBJs
        frmCargando.X.Caption = "Cargando Datos de Objetos..." & Obj & "/" & NumOBJs
        DoEvents
        ObjData(Obj).name = Leer.GetValue("OBJ" & Obj, "Name")
        ObjData(Obj).GrhIndex = Val(Leer.GetValue("OBJ" & Obj, "GrhIndex"))
        ObjData(Obj).ObjType = Val(Leer.GetValue("OBJ" & Obj, "ObjType"))
        ObjData(Obj).Ropaje = Val(Leer.GetValue("OBJ" & Obj, "NumRopaje"))
        ObjData(Obj).Info = Leer.GetValue("OBJ" & Obj, "Info")
        ObjData(Obj).WeaponAnim = Val(Leer.GetValue("OBJ" & Obj, "Anim"))
        ObjData(Obj).Texto = Leer.GetValue("OBJ" & Obj, "Texto")
        ObjData(Obj).GrhSecundario = Val(Leer.GetValue("OBJ" & Obj, "GrhSec"))
        frmMain.lListado(3).AddItem ObjData(Obj).name & " - #" & Obj
    Next Obj
    Exit Sub
Fallo:
MsgBox "Error al intentar cargar el Objteto " & Obj & " de OBJ.dat en " & DirDats & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub

''
' Carga los indices de Triggers
'

Public Sub CargarIndicesTriggers()
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************

On Error GoTo Fallo
    If FileExist(DirIndex & "Triggers.ini", vbArchive) = False Then
        MsgBox "Falta el archivo 'Triggers.ini' en " & DirIndex, vbCritical
        End
    End If
    Dim NumT As Integer
    Dim T As Integer
    Dim Leer As New clsIniReader
    Call Leer.Initialize(DirIndex & "Triggers.ini")
    frmMain.lListado(4).Clear
    NumT = Val(Leer.GetValue("INIT", "NumTriggers"))
    For T = 1 To NumT
         frmMain.lListado(4).AddItem Leer.GetValue("Trig" & T, "Name") & " - #" & (T - 1)
    Next T

Exit Sub
Fallo:
    MsgBox "Error al intentar cargar el Trigger " & T & " de Triggers.ini en " & DirIndex & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub

''
' Carga los indices de Cuerpos
'

Public Sub CargarIndicesDeCuerpos()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************

On Error GoTo Fallo
    If FileExist(DirIndex & "Personajes.ind", vbArchive) = False Then
        MsgBox "Falta el archivo 'Personajes.ind' en " & DirIndex, vbCritical
        End
    End If
    Dim N As Integer
    Dim i As Integer
    N = FreeFile
    Open DirIndex & "Personajes.ind" For Binary Access Read As #N
    'cabecera
    Get #N, , MiCabecera
    'num de cabezas
    Get #N, , NumBodies
    'Resize array
    ReDim BodyData(0 To NumBodies + 1) As tBodyData
    ReDim MisCuerpos(0 To NumBodies + 1) As tIndiceCuerpo
    For i = 1 To NumBodies
        Get #N, , MisCuerpos(i)
        InitGrh BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0
        InitGrh BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0
        InitGrh BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0
        InitGrh BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0
        BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
        BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
    Next i
    Close #N
Exit Sub
Fallo:
    MsgBox "Error al intentar cargar el Cuerpo " & i & " de Personajes.ind en " & DirIndex & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub

''
' Carga los indices de Cabezas
'

Public Sub CargarIndicesDeCabezas()
On Error GoTo Fallo
    If FileExist(DirIndex & "Cabezas.ind", vbArchive) = False Then
        MsgBox "Falta el archivo 'Cabezas.ind' en " & DirIndex, vbCritical
        End
    End If
    Dim N As Integer
    Dim i As Integer
    Dim MisCabezas() As tIndiceCabeza
    N = FreeFile
    Open DirIndex & "Cabezas.ind" For Binary Access Read As #N
    'cabecera
    Get #N, , MiCabecera
    'num de cabezas
    Get #N, , Numheads
    'Resize array
    ReDim HeadData(0 To Numheads + 1) As tHeadData
    ReDim MisCabezas(0 To Numheads + 1) As tIndiceCabeza
    For i = 1 To Numheads
        Get #N, , MisCabezas(i)
        InitGrh HeadData(i).Head(1), MisCabezas(i).Head(1), 0
        InitGrh HeadData(i).Head(2), MisCabezas(i).Head(2), 0
        InitGrh HeadData(i).Head(3), MisCabezas(i).Head(3), 0
        InitGrh HeadData(i).Head(4), MisCabezas(i).Head(4), 0
    Next i
    Close #N
Exit Sub
Fallo:
    MsgBox "Error al intentar cargar la Cabeza " & i & " de Cabezas.ind en " & DirIndex & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub

''
' Carga los indices de NPCs
'

Public Sub CargarIndicesNPC()
'*************************************************
'Author: ^[GS]^
'Last modified: 26/05/06
'*************************************************

On Error GoTo Fallo
    If FileExist(DirDats & "\NPCs.dat", vbArchive) = False Then
        MsgBox "Falta el archivo 'NPCs.dat' en " & DirDats, vbCritical
        End
    End If
    If FileExist(DirDats & "\NPCs-HOSTILES.dat", vbArchive) = False Then
        MsgBox "Falta el archivo 'NPCs-HOSTILES.dat' en " & DirDats, vbCritical
        End
    End If
    Dim Trabajando As String
    Dim NPC As Integer
    Dim Leer As New clsIniReader
    frmMain.lListado(1).Clear
    frmMain.lListado(2).Clear
    Call Leer.Initialize(DirDats & "\NPCs.dat")
    NumNPCs = Val(Leer.GetValue("INIT", "NumNPCs"))
    Call Leer.Initialize(DirDats & "\NPCs-HOSTILES.dat")
    NumNPCsHOST = Val(Leer.GetValue("INIT", "NumNPCs"))
    ReDim NpcData(1 To NumNPCsHOST + 499) As NpcData
    Trabajando = "Dats\NPCs.dat"
    Call Leer.Initialize(DirDats & "\NPCs.dat")
    For NPC = 1 To NumNPCs
        NpcData(NPC).name = Leer.GetValue("NPC" & NPC, "Name")
        NpcData(NPC).Body = Val(Leer.GetValue("NPC" & NPC, "Body"))
        NpcData(NPC).Head = Val(Leer.GetValue("NPC" & NPC, "Head"))
        NpcData(NPC).Heading = Val(Leer.GetValue("NPC" & NPC, "Heading"))
        If LenB(NpcData(NPC).name) <> 0 Then frmMain.lListado(1).AddItem NpcData(NPC).name & " - #" & NPC
    Next NPC
    Trabajando = "Dats\NPCs-HOSTILES.dat"
    Call Leer.Initialize(DirDats & "\NPCs-HOSTILES.dat")
    For NPC = 1 To NumNPCsHOST
        NpcData(NPC + 499).name = Leer.GetValue("NPC" & (NPC + 499), "Name")
        NpcData(NPC + 499).Body = Val(Leer.GetValue("NPC" & (NPC + 499), "Body"))
        NpcData(NPC + 499).Head = Val(Leer.GetValue("NPC" & (NPC + 499), "Head"))
        NpcData(NPC + 499).Heading = Val(Leer.GetValue("NPC" & (NPC + 499), "Heading"))
        If LenB(NpcData(NPC + 499).name) <> 0 Then frmMain.lListado(2).AddItem NpcData(NPC + 499).name & " - #" & (NPC + 499)
    Next NPC
    Exit Sub
Fallo:
    MsgBox "Error al intentar cargar el NPC " & NPC & " de " & Trabajando & " en " & DirDats & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub
