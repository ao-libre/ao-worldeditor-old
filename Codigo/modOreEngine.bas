Attribute VB_Name = "modOreEngine"

'Particle Groups
Public TotalStreams As Integer

Public StreamData() As Stream
 
'RGB Type
Public Type RGB

    R As Long
    G As Long
    B As Long

End Type
 
Public Type Stream

    Name As String
    NumOfParticles As Long
    NumGrhs As Long
    id As Long
    x1 As Long
    Y1 As Long
    x2 As Long
    Y2 As Long
    angle As Long
    vecx1 As Long
    vecx2 As Long
    vecy1 As Long
    vecy2 As Long
    life1 As Long
    life2 As Long
    friction As Long
    spin As Byte
    spin_speedL As Single
    spin_speedH As Single
    AlphaBlend As Byte
    gravity As Byte
    grav_strength As Long
    bounce_strength As Long
    XMove As Byte
    YMove As Byte
    move_x1 As Long
    move_x2 As Long
    move_y1 As Long
    move_y2 As Long
    Grh_list() As Long
    colortint(0 To 3) As RGB
   
    Speed As Single
    life_counter As Long

End Type
 
'index de la particula que debe ser = que le pusimos al server
Public Enum ParticulaMedit

    CHICO = 34
    MEDIANO = 35
    GRANDE = 37
    XGRANDE = 38
    XXGRANDE = 39

End Enum
 
'Old fashion BitBlt function
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

'Added by Juan Martín Sotuyo Dodero
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Sub CargarParticulas()

    Dim StreamFile As String
    Dim loopc      As Long
    Dim i          As Long
    Dim GrhListing As String
    Dim TempSet    As String
    Dim ColorSet   As Long
    
    Dim Leer       As clsIniManager
    Set Leer = New clsIniManager
    
    Call Leer.Initialize(DirIndex & "Particulas.ini")
    
    'resize StreamData array
    TotalStreams = Val(Leer.GetValue("INIT", "Total"))
    ReDim StreamData(1 To TotalStreams) As Stream
 
    'fill StreamData array with info from Particles.ini
    For loopc = 1 To TotalStreams

        With StreamData(loopc)
            .Name = Leer.GetValue(Val(loopc), "Name")
            .NumOfParticles = Leer.GetValue(Val(loopc), "NumOfParticles")
            .x1 = Leer.GetValue(Val(loopc), "X1")
            .Y1 = Leer.GetValue(Val(loopc), "Y1")
            .x2 = Leer.GetValue(Val(loopc), "X2")
            .Y2 = Leer.GetValue(Val(loopc), "Y2")
            .angle = Leer.GetValue(Val(loopc), "Angle")
            .vecx1 = Leer.GetValue(Val(loopc), "VecX1")
            .vecx2 = Leer.GetValue(Val(loopc), "VecX2")
            .vecy1 = Leer.GetValue(Val(loopc), "VecY1")
            .vecy2 = Leer.GetValue(Val(loopc), "VecY2")
            .life1 = Leer.GetValue(Val(loopc), "Life1")
            .life2 = Leer.GetValue(Val(loopc), "Life2")
            .friction = Leer.GetValue(Val(loopc), "Friction")
            .spin = Leer.GetValue(Val(loopc), "Spin")
            .spin_speedL = Leer.GetValue(Val(loopc), "Spin_SpeedL")
            .spin_speedH = Leer.GetValue(Val(loopc), "Spin_SpeedH")
            .AlphaBlend = Leer.GetValue(Val(loopc), "AlphaBlend")
            .gravity = Leer.GetValue(Val(loopc), "Gravity")
            .grav_strength = Leer.GetValue(Val(loopc), "Grav_Strength")
            .bounce_strength = Leer.GetValue(Val(loopc), "Bounce_Strength")
            .XMove = Leer.GetValue(Val(loopc), "XMove")
            .YMove = Leer.GetValue(Val(loopc), "YMove")
            .move_x1 = Leer.GetValue(Val(loopc), "move_x1")
            .move_x2 = Leer.GetValue(Val(loopc), "move_x2")
            .move_y1 = Leer.GetValue(Val(loopc), "move_y1")
            .move_y2 = Leer.GetValue(Val(loopc), "move_y2")
            .life_counter = Leer.GetValue(Val(loopc), "life_counter")
            .Speed = Val(Leer.GetValue(Val(loopc), "Speed"))
            .NumGrhs = Leer.GetValue(Val(loopc), "NumGrhs")
       
            ReDim .Grh_list(1 To .NumGrhs)
            GrhListing = Leer.GetValue(Val(loopc), "Grh_List")
       
            For i = 1 To .NumGrhs
                .Grh_list(i) = general_field_read(Str(i), GrhListing, 44)
            Next i

            .Grh_list(i - 1) = .Grh_list(i - 1)

            For ColorSet = 1 To 4
                TempSet = Leer.GetValue(Val(loopc), "ColorSet" & ColorSet)
                .colortint(ColorSet - 1).R = general_field_read(1, TempSet, 44)
                .colortint(ColorSet - 1).G = general_field_read(2, TempSet, 44)
                .colortint(ColorSet - 1).B = general_field_read(3, TempSet, 44)
            Next ColorSet
            
        End With
            
    Next loopc
    
    Set Leer = Nothing
 
End Sub
 
Public Function General_Particle_Create(ByVal ParticulaInd As Long, ByVal X As Integer, ByVal Y As Integer, Optional ByVal particle_life As Long = 0) As Long

    Dim Rgb_List(0 To 3) As Long
    
    With StreamData(ParticulaInd)
    
        Rgb_List(0) = RGB(.colortint(0).R, .colortint(0).G, .colortint(0).B)
        Rgb_List(1) = RGB(.colortint(1).R, .colortint(1).G, .colortint(1).B)
        Rgb_List(2) = RGB(.colortint(2).R, .colortint(2).G, .colortint(2).B)
        Rgb_List(3) = RGB(.colortint(3).R, .colortint(3).G, .colortint(3).B)
 
        General_Particle_Create = Particle_Group_Create(X, Y, .Grh_list, Rgb_List(), .NumOfParticles, ParticulaInd, _
           .AlphaBlend, IIf(particle_life = 0, .life_counter, particle_life), .Speed, , .x1, .Y1, .angle, _
           .vecx1, .vecx2, .vecy1, .vecy2, _
           .life1, .life2, .friction, .spin_speedL, _
           .gravity, .grav_strength, .bounce_strength, .x2, _
           .Y2, .XMove, .move_x1, .move_x2, .move_y1, _
           .move_y2, .YMove, .spin_speedH, .spin)
    
    End With
 
End Function


Private Function Particle_Group_Next_Open() As Long

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    '
    '*****************************************************************
    On Error GoTo ErrorHandler:

    Dim loopc As Long
   
    loopc = 1

    Do Until particle_group_list(loopc).Active = False

        If loopc = particle_group_last Then
            Particle_Group_Next_Open = particle_group_last + 1
            Exit Function

        End If

        loopc = loopc + 1
    Loop
   
    Particle_Group_Next_Open = loopc
    Exit Function
ErrorHandler:
    Particle_Group_Next_Open = 1

End Function

Public Function Particle_Group_Create(ByVal map_x As Integer, ByVal map_y As Integer, ByRef GrhIndex_list() As Long, ByRef Rgb_List() As Long, _
   Optional ByVal particle_count As Long = 20, Optional ByVal stream_type As Long = 1, _
   Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
   Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
   Optional ByVal x1 As Integer, Optional ByVal Y1 As Integer, Optional ByVal angle As Integer, _
   Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
   Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
   Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
   Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
   Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
   Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal Y2 As Integer, _
   Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
   Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
   Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean) As Long
   
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 12/15/2002
    'Returns the particle_group_index if successful, else 0
    '**************************************************************
    If (map_x <> -1) And (map_y <> -1) Then
        If Map_Particle_Group_Get(map_x, map_y) = 0 Then
            Particle_Group_Create = Particle_Group_Next_Open
            Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, GrhIndex_list(), Rgb_List(), alpha_blend, alive_counter, frame_speed, id, x1, Y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, Y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin
        Else
            Particle_Group_Create = Particle_Group_Next_Open
            Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, GrhIndex_list(), Rgb_List(), alpha_blend, alive_counter, frame_speed, id, x1, Y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, Y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin

        End If

    End If

End Function
 
Public Function Particle_Group_Remove(ByVal particle_group_index As Long) As Boolean

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 1/04/2003
    '
    '*****************************************************************
    'Make sure it's a legal index
    If Particle_Group_Check(particle_group_index) Then
        Particle_Group_Destroy particle_group_index
        Particle_Group_Remove = True

    End If

End Function
 
Public Function Particle_Group_Remove_All() As Boolean

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 1/04/2003
    '
    '*****************************************************************
    Dim index As Long
   
    For index = 1 To particle_group_last

        'Make sure it's a legal index
        If Particle_Group_Check(index) Then
            Particle_Group_Destroy index

        End If

    Next index
   
    Particle_Group_Remove_All = True

End Function
 
Public Function Particle_Group_Find(ByVal id As Long) As Long

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 1/04/2003
    'Find the index related to the handle
    '*****************************************************************
    On Error GoTo ErrorHandler:

    Dim loopc As Long
   
    loopc = 1

    Do Until particle_group_list(loopc).id = id

        If loopc = particle_group_last Then
            Particle_Group_Find = 0
            Exit Function

        End If

        loopc = loopc + 1
    Loop
   
    Particle_Group_Find = loopc
    Exit Function
ErrorHandler:
    Particle_Group_Find = 0

End Function
 
Private Sub Particle_Group_Make(ByVal particle_group_index As Long, ByVal map_x As Integer, ByVal map_y As Integer, _
   ByVal particle_count As Long, ByVal stream_type As Long, ByRef GrhIndex_list() As Long, ByRef Rgb_List() As Long, _
   Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
   Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
   Optional ByVal x1 As Integer, Optional ByVal Y1 As Integer, Optional ByVal angle As Integer, _
   Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
   Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
   Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
   Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
   Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
   Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal Y2 As Integer, _
   Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
   Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
   Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean)

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    'Makes a new particle effect
    '*****************************************************************
    'Update array size
    If particle_group_index > particle_group_last Then
        particle_group_last = particle_group_index
        ReDim Preserve particle_group_list(1 To particle_group_last)

    End If

    particle_group_count = particle_group_count + 1
   
    'Make active
    particle_group_list(particle_group_index).Active = True
   
    'Map pos
    If (map_x <> -1) And (map_y <> -1) Then
        particle_group_list(particle_group_index).map_x = map_x
        particle_group_list(particle_group_index).map_y = map_y

    End If
   
    'Grh list
    ReDim particle_group_list(particle_group_index).GrhIndex_list(1 To UBound(GrhIndex_list))
    particle_group_list(particle_group_index).GrhIndex_list() = GrhIndex_list()
    particle_group_list(particle_group_index).GrhIndex_count = UBound(GrhIndex_list)
   
    'Sets alive vars
    If alive_counter = -1 Then
        particle_group_list(particle_group_index).alive_counter = -1
        particle_group_list(particle_group_index).never_die = True
    Else
        particle_group_list(particle_group_index).alive_counter = alive_counter
        particle_group_list(particle_group_index).never_die = False

    End If
   
    'alpha blending
    particle_group_list(particle_group_index).alpha_blend = alpha_blend
   
    'stream type
    particle_group_list(particle_group_index).stream_type = stream_type
   
    'speed
    particle_group_list(particle_group_index).frame_speed = frame_speed
   
    particle_group_list(particle_group_index).x1 = x1
    particle_group_list(particle_group_index).Y1 = Y1
    particle_group_list(particle_group_index).x2 = x2
    particle_group_list(particle_group_index).Y2 = Y2
    particle_group_list(particle_group_index).angle = angle
    particle_group_list(particle_group_index).vecx1 = vecx1
    particle_group_list(particle_group_index).vecx2 = vecx2
    particle_group_list(particle_group_index).vecy1 = vecy1
    particle_group_list(particle_group_index).vecy2 = vecy2
    particle_group_list(particle_group_index).life1 = life1
    particle_group_list(particle_group_index).life2 = life2
    particle_group_list(particle_group_index).fric = fric
    particle_group_list(particle_group_index).spin = spin
    particle_group_list(particle_group_index).spin_speedL = spin_speedL
    particle_group_list(particle_group_index).spin_speedH = spin_speedH
    particle_group_list(particle_group_index).gravity = gravity
    particle_group_list(particle_group_index).grav_strength = grav_strength
    particle_group_list(particle_group_index).bounce_strength = bounce_strength
    particle_group_list(particle_group_index).XMove = XMove
    particle_group_list(particle_group_index).YMove = YMove
    particle_group_list(particle_group_index).move_x1 = move_x1
    particle_group_list(particle_group_index).move_x2 = move_x2
    particle_group_list(particle_group_index).move_y1 = move_y1
    particle_group_list(particle_group_index).move_y2 = move_y2
   
    particle_group_list(particle_group_index).Rgb_List(0) = Rgb_List(0)
    particle_group_list(particle_group_index).Rgb_List(1) = Rgb_List(1)
    particle_group_list(particle_group_index).Rgb_List(2) = Rgb_List(2)
    particle_group_list(particle_group_index).Rgb_List(3) = Rgb_List(3)
   
    'create particle stream
    particle_group_list(particle_group_index).particle_count = particle_count
    ReDim particle_group_list(particle_group_index).particle_stream(1 To particle_count)
   
    'plot particle group on map
    MapData(map_x, map_y).particle_group_index = particle_group_index

End Sub
 
Private Sub Particle_Render(ByRef temp_particle As Particle, ByVal screen_x As Integer, ByVal screen_Y As Integer, _
   ByVal GrhIndex As Long, ByRef Rgb_List() As Long, _
   Optional ByVal alpha_blend As Boolean, Optional ByVal no_move As Boolean, _
   Optional ByVal x1 As Integer, Optional ByVal Y1 As Integer, Optional ByVal angle As Integer, _
   Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
   Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
   Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
   Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
   Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
   Optional ByVal bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal Y2 As Integer, _
   Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
   Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
   Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
   Optional grh_resizex As Integer, Optional grh_resizey As Integer, _
   Optional ByVal Radio As Integer, Optional ByVal Count As Integer, Optional ByVal index As Integer)
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 4/24/2003
    '
    '**************************************************************

    If no_move = False Then
        If temp_particle.alive_counter = 0 Then
            Grh_Initialize temp_particle.Grh, GrhIndex, alpha_blend

            If Radio = 0 Then
                temp_particle.X = RandomNumber(x1, x2)
                temp_particle.Y = RandomNumber(Y1, Y2)
            Else
                temp_particle.X = (RandomNumber(x1, x2) + Radio) + Radio * Cos(PI * 2 * index / Count)
                temp_particle.Y = (RandomNumber(Y1, Y2) + Radio) + Radio * Sin(PI * 2 * index / Count)

            End If

            temp_particle.vector_x = RandomNumber(vecx1, vecx2)
            temp_particle.vector_y = RandomNumber(vecy1, vecy2)
            temp_particle.angle = angle
            temp_particle.alive_counter = RandomNumber(life1, life2)
            temp_particle.friction = fric
        Else

            'Continue old particle
            'Do gravity
            If gravity = True Then
                temp_particle.vector_y = temp_particle.vector_y + grav_strength

                If temp_particle.Y > 0 Then
                    'bounce
                    temp_particle.vector_y = bounce_strength

                End If

            End If

            'Do rotation
            If spin = True Then temp_particle.Grh.angle = temp_particle.Grh.angle + (RandomNumber(spin_speedL, spin_speedH) / 100)
            If temp_particle.angle >= 360 Then
                temp_particle.angle = 0

            End If
                                
            If XMove = True Then temp_particle.vector_x = RandomNumber(move_x1, move_x2)
            If YMove = True Then temp_particle.vector_y = RandomNumber(move_y1, move_y2)

        End If

        'Add in vector
        temp_particle.X = temp_particle.X + (temp_particle.vector_x \ temp_particle.friction)
        temp_particle.Y = temp_particle.Y + (temp_particle.vector_y \ temp_particle.friction)
    
        'decrement counter
        temp_particle.alive_counter = temp_particle.alive_counter - 1

    End If
 
    'Draw it
    If temp_particle.Grh.GrhIndex Then
        Grh_Render temp_particle.Grh, temp_particle.X + screen_x, temp_particle.Y + screen_Y, Rgb_List()

    End If

End Sub

Public Sub Particle_Group_Render(ByVal particle_group_index As Long, ByVal screen_x As Integer, ByVal screen_Y As Integer)

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 12/15/2002
    'Renders a particle stream at a paticular screen point
    '*****************************************************************
    On Error GoTo Err

    Dim loopc       As Long

    Dim temp_rgb(3) As Long

    Dim no_move     As Boolean

    'Set colors
    temp_rgb(0) = particle_group_list(particle_group_index).Rgb_List(0)
    temp_rgb(1) = particle_group_list(particle_group_index).Rgb_List(1)
    temp_rgb(2) = particle_group_list(particle_group_index).Rgb_List(2)
    temp_rgb(3) = particle_group_list(particle_group_index).Rgb_List(3)
       
    If particle_group_list(particle_group_index).alive_counter Then
    
        'See if it is time to move a particle
        particle_group_list(particle_group_index).frame_counter = particle_group_list(particle_group_index).frame_counter + timer_ticks_per_frame

        If particle_group_list(particle_group_index).frame_counter > particle_group_list(particle_group_index).frame_speed Then
            particle_group_list(particle_group_index).frame_counter = 0
            no_move = False
        Else
            no_move = True

        End If
    
        'If it's still alive render all the particles inside
        For loopc = 1 To particle_group_list(particle_group_index).particle_count
        
            'Render particle
            Particle_Render particle_group_list(particle_group_index).particle_stream(loopc), _
               screen_x, screen_Y, _
               particle_group_list(particle_group_index).GrhIndex_list(Round(RandomNumber(1, particle_group_list(particle_group_index).GrhIndex_count), 0)), _
               temp_rgb(), _
               particle_group_list(particle_group_index).alpha_blend, no_move, _
               particle_group_list(particle_group_index).x1, particle_group_list(particle_group_index).Y1, particle_group_list(particle_group_index).angle, _
               particle_group_list(particle_group_index).vecx1, particle_group_list(particle_group_index).vecx2, _
               particle_group_list(particle_group_index).vecy1, particle_group_list(particle_group_index).vecy2, _
               particle_group_list(particle_group_index).life1, particle_group_list(particle_group_index).life2, _
               particle_group_list(particle_group_index).fric, particle_group_list(particle_group_index).spin_speedL, _
               particle_group_list(particle_group_index).gravity, particle_group_list(particle_group_index).grav_strength, _
               particle_group_list(particle_group_index).bounce_strength, particle_group_list(particle_group_index).x2, _
               particle_group_list(particle_group_index).Y2, particle_group_list(particle_group_index).XMove, _
               particle_group_list(particle_group_index).move_x1, particle_group_list(particle_group_index).move_x2, _
               particle_group_list(particle_group_index).move_y1, particle_group_list(particle_group_index).move_y2, _
               particle_group_list(particle_group_index).YMove, particle_group_list(particle_group_index).spin_speedH, _
               particle_group_list(particle_group_index).spin, , , , _
               , particle_group_list(particle_group_index).particle_count, loopc
                            
        Next loopc
        
        If no_move = False Then

            'Update the group alive counter
            If particle_group_list(particle_group_index).never_die = False Then
                particle_group_list(particle_group_index).alive_counter = particle_group_list(particle_group_index).alive_counter - 1

            End If

        End If
    
    Else
        'If it's dead destroy it
        particle_group_list(particle_group_index).particle_count = particle_group_list(particle_group_index).particle_count - 1

        If particle_group_list(particle_group_index).particle_count <= 0 Then Particle_Group_Destroy particle_group_index

    End If
    
Err:
    temp_rgb(0) = 0
    temp_rgb(1) = 0
    temp_rgb(2) = 0
    temp_rgb(3) = 0

End Sub
 
Public Function Particle_Type_Get(ByVal particle_Index As Long) As Long

    '*****************************************************************
    'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
    'Last Modify Date: 8/27/2003
    'Returns the stream type of a particle stream
    '*****************************************************************
    If Particle_Group_Check(particle_Index) Then
        Particle_Type_Get = particle_group_list(particle_Index).stream_type

    End If

End Function
 
Private Function Particle_Group_Check(ByVal particle_group_index As Long) As Boolean

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 1/04/2003
    '
    '**************************************************************
    'check index
    If particle_group_index > 0 And particle_group_index <= particle_group_last Then
        If particle_group_list(particle_group_index).Active Then
            Particle_Group_Check = True

        End If

    End If

End Function
 
Public Function Particle_Group_Map_Pos_Set(ByVal particle_group_index As Long, ByVal map_x As Long, ByVal map_y As Long) As Boolean

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 5/27/2003
    'Returns true if successful, else false
    '**************************************************************
    'Make sure it's a legal index
    If Particle_Group_Check(particle_group_index) Then

        'Make sure it's a legal move
        If InMapBounds(map_x, map_y) Then
            'Move it
            particle_group_list(particle_group_index).map_x = map_x
            particle_group_list(particle_group_index).map_y = map_y
   
            Particle_Group_Map_Pos_Set = True

        End If

    End If

End Function
 
Private Sub Particle_Group_Destroy(ByVal particle_group_index As Long)

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    '
    '**************************************************************
    Dim temp As particle_group
    
    If particle_group_list(particle_group_index).map_x > 0 And particle_group_list(particle_group_index).map_y > 0 Then
        MapData(particle_group_list(particle_group_index).map_x, particle_group_list(particle_group_index).map_y).particle_group_index = 0

    End If
    
    particle_group_list(particle_group_index) = temp
            
    'Update array size
    If particle_group_index = particle_group_last Then

        Do Until particle_group_list(particle_group_last).Active
            particle_group_last = particle_group_last - 1

            If particle_group_last = 0 Then
                particle_group_count = 0
                Exit Sub

            End If

        Loop
        ReDim Preserve particle_group_list(1 To particle_group_last)

    End If

    particle_group_count = particle_group_count - 1

End Sub

Public Function Map_Particle_Group_Get(ByVal map_x As Integer, ByVal map_y As Integer) As Long
 
    If InMapBounds(map_x, map_y) Then
        Map_Particle_Group_Get = MapData(map_x, map_y).particle_group_index
    Else
        Map_Particle_Group_Get = 0

    End If

End Function


