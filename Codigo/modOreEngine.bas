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
Dim loopc As Long
Dim i As Long
Dim GrhListing As String
Dim TempSet As String
Dim ColorSet As Long
   
StreamFile = DirIndex & "Particulas.ini"
TotalStreams = Val(general_var_get(StreamFile, "INIT", "Total"))
 
'resize StreamData array
ReDim StreamData(1 To TotalStreams) As Stream
 
    'fill StreamData array with info from Particles.ini
    For loopc = 1 To TotalStreams
        StreamData(loopc).name = general_var_get(StreamFile, Val(loopc), "Name")
        StreamData(loopc).NumOfParticles = general_var_get(StreamFile, Val(loopc), "NumOfParticles")
        StreamData(loopc).x1 = general_var_get(StreamFile, Val(loopc), "X1")
        StreamData(loopc).Y1 = general_var_get(StreamFile, Val(loopc), "Y1")
        StreamData(loopc).x2 = general_var_get(StreamFile, Val(loopc), "X2")
        StreamData(loopc).Y2 = general_var_get(StreamFile, Val(loopc), "Y2")
        StreamData(loopc).angle = general_var_get(StreamFile, Val(loopc), "Angle")
        StreamData(loopc).vecx1 = general_var_get(StreamFile, Val(loopc), "VecX1")
        StreamData(loopc).vecx2 = general_var_get(StreamFile, Val(loopc), "VecX2")
        StreamData(loopc).vecy1 = general_var_get(StreamFile, Val(loopc), "VecY1")
        StreamData(loopc).vecy2 = general_var_get(StreamFile, Val(loopc), "VecY2")
        StreamData(loopc).life1 = general_var_get(StreamFile, Val(loopc), "Life1")
        StreamData(loopc).life2 = general_var_get(StreamFile, Val(loopc), "Life2")
        StreamData(loopc).friction = general_var_get(StreamFile, Val(loopc), "Friction")
        StreamData(loopc).spin = general_var_get(StreamFile, Val(loopc), "Spin")
        StreamData(loopc).spin_speedL = general_var_get(StreamFile, Val(loopc), "Spin_SpeedL")
        StreamData(loopc).spin_speedH = general_var_get(StreamFile, Val(loopc), "Spin_SpeedH")
        StreamData(loopc).AlphaBlend = general_var_get(StreamFile, Val(loopc), "AlphaBlend")
        StreamData(loopc).gravity = general_var_get(StreamFile, Val(loopc), "Gravity")
        StreamData(loopc).grav_strength = general_var_get(StreamFile, Val(loopc), "Grav_Strength")
        StreamData(loopc).bounce_strength = general_var_get(StreamFile, Val(loopc), "Bounce_Strength")
        StreamData(loopc).XMove = general_var_get(StreamFile, Val(loopc), "XMove")
        StreamData(loopc).YMove = general_var_get(StreamFile, Val(loopc), "YMove")
        StreamData(loopc).move_x1 = general_var_get(StreamFile, Val(loopc), "move_x1")
        StreamData(loopc).move_x2 = general_var_get(StreamFile, Val(loopc), "move_x2")
        StreamData(loopc).move_y1 = general_var_get(StreamFile, Val(loopc), "move_y1")
        StreamData(loopc).move_y2 = general_var_get(StreamFile, Val(loopc), "move_y2")
        StreamData(loopc).life_counter = general_var_get(StreamFile, Val(loopc), "life_counter")
        StreamData(loopc).Speed = Val(general_var_get(StreamFile, Val(loopc), "Speed"))
        StreamData(loopc).NumGrhs = general_var_get(StreamFile, Val(loopc), "NumGrhs")
       
        ReDim StreamData(loopc).Grh_list(1 To StreamData(loopc).NumGrhs)
        GrhListing = general_var_get(StreamFile, Val(loopc), "Grh_List")
       
        For i = 1 To StreamData(loopc).NumGrhs
            StreamData(loopc).Grh_list(i) = general_field_read(Str(i), GrhListing, 44)
        Next i
        StreamData(loopc).Grh_list(i - 1) = StreamData(loopc).Grh_list(i - 1)
        For ColorSet = 1 To 4
            TempSet = general_var_get(StreamFile, Val(loopc), "ColorSet" & ColorSet)
            StreamData(loopc).colortint(ColorSet - 1).R = general_field_read(1, TempSet, 44)
            StreamData(loopc).colortint(ColorSet - 1).G = general_field_read(2, TempSet, 44)
            StreamData(loopc).colortint(ColorSet - 1).B = general_field_read(3, TempSet, 44)
        Next ColorSet
    Next loopc
 
End Sub
 
Public Function General_Particle_Create(ByVal ParticulaInd As Long, ByVal X As Integer, ByVal Y As Integer, Optional ByVal particle_life As Long = 0) As Long

Dim Rgb_List(0 To 3) As Long
Rgb_List(0) = RGB(StreamData(ParticulaInd).colortint(0).R, StreamData(ParticulaInd).colortint(0).G, StreamData(ParticulaInd).colortint(0).B)
Rgb_List(1) = RGB(StreamData(ParticulaInd).colortint(1).R, StreamData(ParticulaInd).colortint(1).G, StreamData(ParticulaInd).colortint(1).B)
Rgb_List(2) = RGB(StreamData(ParticulaInd).colortint(2).R, StreamData(ParticulaInd).colortint(2).G, StreamData(ParticulaInd).colortint(2).B)
Rgb_List(3) = RGB(StreamData(ParticulaInd).colortint(3).R, StreamData(ParticulaInd).colortint(3).G, StreamData(ParticulaInd).colortint(3).B)
 
General_Particle_Create = Particle_Group_Create(X, Y, StreamData(ParticulaInd).Grh_list, Rgb_List(), StreamData(ParticulaInd).NumOfParticles, ParticulaInd, _
    StreamData(ParticulaInd).AlphaBlend, IIf(particle_life = 0, StreamData(ParticulaInd).life_counter, particle_life), StreamData(ParticulaInd).Speed, , StreamData(ParticulaInd).x1, StreamData(ParticulaInd).Y1, StreamData(ParticulaInd).angle, _
    StreamData(ParticulaInd).vecx1, StreamData(ParticulaInd).vecx2, StreamData(ParticulaInd).vecy1, StreamData(ParticulaInd).vecy2, _
    StreamData(ParticulaInd).life1, StreamData(ParticulaInd).life2, StreamData(ParticulaInd).friction, StreamData(ParticulaInd).spin_speedL, _
    StreamData(ParticulaInd).gravity, StreamData(ParticulaInd).grav_strength, StreamData(ParticulaInd).bounce_strength, StreamData(ParticulaInd).x2, _
    StreamData(ParticulaInd).Y2, StreamData(ParticulaInd).XMove, StreamData(ParticulaInd).move_x1, StreamData(ParticulaInd).move_x2, StreamData(ParticulaInd).move_y1, _
    StreamData(ParticulaInd).move_y2, StreamData(ParticulaInd).YMove, StreamData(ParticulaInd).spin_speedH, StreamData(ParticulaInd).spin)
 
End Function
 


