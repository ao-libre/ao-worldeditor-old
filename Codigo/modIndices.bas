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
' @version 0.1.05
' @date 20060530

Option Explicit

' Carga los indices de Superficie
'

Public Sub CargarIndicesSuperficie()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 29/05/06
    '*************************************************

    On Error GoTo Fallo

    If FileExist(inipath & "GrhIndex\indices.ini", vbArchive) = False Then
        Call MsgBox("Falta el archivo 'GrhIndex\indices.ini'", vbCritical)
        End
    End If

    Dim Leer As New clsIniManager
    Dim i    As Integer

    Call Leer.Initialize(inipath & "GrhIndex\indices.ini")
    
    MaxSup = Leer.GetValue("INIT", "Referencias")
    ReDim SupData(MaxSup) As SupData
    
    Call frmMain.lListado(0).Clear

    For i = 0 To MaxSup

        With SupData(i)
        
            .Name = Leer.GetValue("REFERENCIA" & i, "Nombre")
            .Grh = Val(Leer.GetValue("REFERENCIA" & i, "GrhIndice"))
            .Width = Val(Leer.GetValue("REFERENCIA" & i, "Ancho"))
            .Height = Val(Leer.GetValue("REFERENCIA" & i, "Alto"))
            .Block = IIf(Val(Leer.GetValue("REFERENCIA" & i, "Bloquear")) = 1, True, False)
            .Capa = Val(Leer.GetValue("REFERENCIA" & i, "Capa"))
            Call frmMain.lListado(0).AddItem(.Name & " - #" & i)
        
        End With
        
    Next
    
    DoEvents
    
    Exit Sub
Fallo:
    Call MsgBox("Error al intentar cargar el indice " & i & " de GrhIndex\indices.ini" & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly)

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
        Call MsgBox("Falta el archivo 'OBJ.dat' en " & DirDats, vbCritical)
        End

    End If

    Dim Obj  As Integer
    Dim Leer As New clsIniManager

    Call Leer.Initialize(DirDats & "\OBJ.dat")
    
    'Limpio la lista
    Call frmMain.lListado(3).Clear
    
    NumOBJs = Val(Leer.GetValue("INIT", "NumOBJs"))
    ReDim ObjData(1 To NumOBJs) As ObjData

    For Obj = 1 To NumOBJs
        frmCargando.X.Caption = "Cargando Datos de Objetos..." & Obj & "/" & NumOBJs
        DoEvents
        
        With ObjData(Obj)
        
            .Name = Leer.GetValue("OBJ" & Obj, "Name")
            .GrhIndex = Val(Leer.GetValue("OBJ" & Obj, "GrhIndex"))
            .ObjType = Val(Leer.GetValue("OBJ" & Obj, "ObjType"))
            .Ropaje = Val(Leer.GetValue("OBJ" & Obj, "NumRopaje"))
            .Info = Leer.GetValue("OBJ" & Obj, "Info")
            .WeaponAnim = Val(Leer.GetValue("OBJ" & Obj, "Anim"))
            .Texto = Leer.GetValue("OBJ" & Obj, "Texto")
            .GrhSecundario = Val(Leer.GetValue("OBJ" & Obj, "GrhSec"))
            Call frmMain.lListado(3).AddItem(.Name & " - #" & Obj)
        
        End With

    Next Obj

    Exit Sub
Fallo:
    Call MsgBox("Error al intentar cargar el Objteto " & Obj & " de OBJ.dat en " & DirDats & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly)

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
    Dim T    As Integer
    Dim Leer As New clsIniManager

    Call Leer.Initialize(DirIndex & "Triggers.ini")
    
    Call frmMain.lListado(4).Clear
    
    NumT = Val(Leer.GetValue("INIT", "NumTriggers"))

    For T = 1 To NumT
        Call frmMain.lListado(4).AddItem(Leer.GetValue("Trig" & T, "Name") & " - #" & (T - 1))
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

    If Not FileExist(DirIndex & "Personajes.ind", vbArchive) Then
        Call MsgBox("Falta el archivo 'Personajes.ind' en " & DirIndex, vbCritical)
        End

    End If
    
    Dim n As Integer
    Dim i As Integer
    
    n = FreeFile
    Open DirIndex & "Personajes.ind" For Binary Access Read As #n
        
        'cabecera
        Get #n, , MiCabecera
        'num de cabezas
        Get #n, , NumBodies
            
        'Resize array
        ReDim BodyData(1 To NumBodies) As tBodyData
        ReDim MisCuerpos(1 To NumBodies) As tIndiceCuerpo
            
        For i = 1 To NumBodies
            Get #n, , MisCuerpos(i)
                
            Call Grh_Initialize(BodyData(i).Walk(1), MisCuerpos(i).Body(1), , , 0)
            Call Grh_Initialize(BodyData(i).Walk(2), MisCuerpos(i).Body(2), , , 0)
            Call Grh_Initialize(BodyData(i).Walk(3), MisCuerpos(i).Body(3), , , 0)
            Call Grh_Initialize(BodyData(i).Walk(4), MisCuerpos(i).Body(4), , , 0)
                
            BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
        Next i

    Close #n
    
    Exit Sub
Fallo:
    Call MsgBox("Error al intentar cargar el Cuerpo " & i & " de Personajes.ind en " & DirIndex & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly)

End Sub

''
' Carga los indices de Cabezas
'

Public Sub CargarIndicesDeCabezas()

    On Error GoTo Fallo

    If Not FileExist(DirIndex & "Cabezas.ind", vbArchive) Then
        MsgBox "Falta el archivo 'Cabezas.ind' en " & DirIndex, vbCritical
        End

    End If
    
    Dim n            As Integer
    Dim i            As Long
    Dim MisCabezas() As tIndiceCabeza
    
    n = FreeFile()
    
    Open DirIndex & "Cabezas.ind" For Binary Access Read As #n
    
        'cabecera
        Get #n, , MiCabecera
        
        'num de cabezas
        Get #n, , Numheads
        
        'Resize array
        ReDim HeadData(0 To Numheads) As tHeadData
        ReDim MisCabezas(0 To Numheads) As tIndiceCabeza
            
        For i = 1 To Numheads
            Get #n, , MisCabezas(i)
                
            If MisCabezas(i).Head(1) Then
                Call Grh_Initialize(HeadData(i).Head(1), MisCabezas(i).Head(1), , , 0)
                Call Grh_Initialize(HeadData(i).Head(2), MisCabezas(i).Head(2), , , 0)
                Call Grh_Initialize(HeadData(i).Head(3), MisCabezas(i).Head(3), , , 0)
                Call Grh_Initialize(HeadData(i).Head(4), MisCabezas(i).Head(4), , , 0)
    
            End If
    
        Next i

    Close #n
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
    On Error Resume Next

    'On Error GoTo Fallo
    If FileExist(DirDats & "\NPCs.dat", vbArchive) = False Then
        MsgBox "Falta el archivo 'NPCs.dat' en " & DirDats, vbCritical
        End

    End If
    
    Dim Trabajando As String
    Dim NPC        As Integer
    Dim Leer       As New clsIniManager

    Call frmMain.lListado(1).Clear
    Call frmMain.lListado(2).Clear
    
    Call Leer.Initialize(DirDats & "\NPCs.dat")
    
    NumNPCs = Val(Leer.GetValue("INIT", "NumNPCs"))
    ReDim NpcData(1000) As NpcData
    
    Trabajando = "Dats\NPCs.dat"

    For NPC = 1 To NumNPCs
        
        With NpcData(NPC)
        
            .Name = Leer.GetValue("NPC" & NPC, "Name")
        
            .Body = Val(Leer.GetValue("NPC" & NPC, "Body"))
            .Head = Val(Leer.GetValue("NPC" & NPC, "Head"))
            .Heading = Val(Leer.GetValue("NPC" & NPC, "Heading"))

            If LenB(NpcData(NPC).Name) <> 0 Then
                Call frmMain.lListado(1).AddItem(.Name & " - #" & NPC)
            End If
        
        End With
        
    Next

    Exit Sub
Fallo:
    Call MsgBox("Error al intentar cargar el NPC " & NPC & " de " & Trabajando & " en " & DirDats & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly)

End Sub

