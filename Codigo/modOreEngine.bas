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

    name As String
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
    Dim Leer       As clsIniReader

    Call Leer.Initialize(DirIndex & "Particulas.ini")
    
    'resize StreamData array
    TotalStreams = Val(Leer.GetValue("INIT", "Total"))
    ReDim StreamData(1 To TotalStreams) As Stream
 
    'fill StreamData array with info from Particles.ini
    For loopc = 1 To TotalStreams

        With StreamData(loopc)
            .name = Leer.GetValue(Val(loopc), "Name")
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

