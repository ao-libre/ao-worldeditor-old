Attribute VB_Name = "modGrh"

Private Const LoopAdEternum As Integer = 999

'Holds data about where a bmp can be found,
'How big it is and animation info
Public Type Grh_Data

    Active As Boolean
    texture_index As Long
    Src_X As Integer
    Src_Y As Integer
    src_width As Integer
    src_height As Integer
    
    frame_count As Integer
    frame_list(1 To 25) As Long
    frame_speed As Single
    MiniMap_color As Long

End Type

'Points to a Grh_Data and keeps animation info
Public Type Grh

    grh_index As Integer
    alpha_blend As Boolean
    angle As Single
    frame_speed As Single
    frame_counter As Single
    Started As Boolean
    LoopTimes As Integer

End Type

'Grh Data Array
Public Grh_list()            As Grh_Data

Public grh_count             As Long

Dim AnimBaseSpeed            As Single

Public timer_ticks_per_frame As Single

Public base_tile_size        As Integer

Public Sub Grh_Initialize(ByRef Grh As Grh, ByVal grh_index As Long, Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single, Optional ByVal Started As Byte = 2, Optional ByVal LoopTimes As Integer = LoopAdEternum)

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    'Sets up a grh. MUST be done before rendering
    '*****************************************************************
    If grh_index <= 0 Then Exit Sub

    'Copy of parameters
    Grh.grh_index = grh_index
    Grh.alpha_blend = alpha_blend
    Grh.angle = angle
    Grh.LoopTimes = LoopTimes
    
    'Start it if it's a animated grh
    If Started = 2 Then
        If Grh_list(Grh.grh_index).frame_count > 1 Then
            Grh.Started = True
        Else
            Grh.Started = False

        End If

    Else
        Grh.Started = Started

    End If
    
    'Set frame counters
    Grh.frame_counter = 1
    
    Grh.frame_speed = Grh_list(Grh.grh_index).frame_speed

End Sub

Private Sub Grh_Load_All()

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 1/04/2003
    'Loads Grh.dat
    '**************************************************************
    'On Error GoTo ErrorHandler
    Dim Grh         As Long
    Dim Frame       As Long
    Dim FileVersion As Long
    Dim initpath    As String

    initpath = inipath & PATH_INIT
    
    'Open files
    Open DirIndex & "Graficos.ind" For Binary As #1
    Seek #1, 1
    
    Get #1, , FileVersion
    
    'Get number of grhs
    Get #1, , grh_count

    'Resize arrays
    ReDim Grh_list(1 To grh_count) As Grh_Data
    'Fill Grh List
    
    'Get first Grh Number
    Get #1, , Grh
    
    Do Until Grh <= 0
        
        Grh_list(Grh).Active = True
        
        'Get number of frames
        Get #1, , Grh_list(Grh).frame_count

        If Grh_list(Grh).frame_count <= 0 Then GoTo ErrorHandler
        
        If Grh_list(Grh).frame_count > 1 Then
        
            'Read a animation GRH set
            For Frame = 1 To Grh_list(Grh).frame_count
            
                Get #1, , Grh_list(Grh).frame_list(Frame)

                If Grh_list(Grh).frame_list(Frame) <= 0 Or Grh_list(Grh).frame_list(Frame) > grh_count Then GoTo ErrorHandler
            
            Next Frame
        
            Get #1, , Grh_list(Grh).frame_speed

            If Grh_list(Grh).frame_speed = 0 Then GoTo ErrorHandler
            
            'Compute width and height
            Grh_list(Grh).src_height = Grh_list(Grh_list(Grh).frame_list(1)).src_height

            If Grh_list(Grh).src_height <= 0 Then GoTo ErrorHandler
            
            Grh_list(Grh).src_width = Grh_list(Grh_list(Grh).frame_list(1)).src_width

            If Grh_list(Grh).src_width <= 0 Then GoTo ErrorHandler
        
        Else
        
            'Read in normal GRH data
            Get #1, , Grh_list(Grh).texture_index

            If Grh_list(Grh).texture_index <= 0 Then GoTo ErrorHandler
            
            Get #1, , Grh_list(Grh).Src_X

            If Grh_list(Grh).Src_X < 0 Then GoTo ErrorHandler
            
            Get #1, , Grh_list(Grh).Src_Y

            If Grh_list(Grh).Src_Y < 0 Then GoTo ErrorHandler
                
            Get #1, , Grh_list(Grh).src_width

            If Grh_list(Grh).src_width <= 0 Then GoTo ErrorHandler
            
            Get #1, , Grh_list(Grh).src_height

            If Grh_list(Grh).src_height <= 0 Then GoTo ErrorHandler
            
            Grh_list(Grh).frame_list(1) = Grh
                
        End If
    
        'Get Next Grh Number
        Get #1, , Grh
    
    Loop
    '************************************************
    
    Close #1

    Dim Count As Long
 
    Open DirIndex & "\minimap.dat" For Binary As #1
    Seek #1, 1

    For Count = 1 To 20459

        If Grh_list(Count).Active Then
            Get #1, , Grh_list(Count).MiniMap_color

        End If

    Next Count

    Close #1
    Exit Sub
ErrorHandler:
    Close #1
    MsgBox "Error while loading the grh.dat! Stopped at GRH number: " & Grh

End Sub

Public Sub Grh_Render(ByRef Grh As Grh, ByVal screen_x As Long, ByVal screen_Y As Long, ByRef Rgb_List() As Long, Optional ByVal center As Boolean)

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 2/28/2003
    '
    '**************************************************************
    Dim tile_width  As Single

    Dim tile_height As Single

    Dim grh_index   As Long
    
    If Grh.grh_index = 0 Then Exit Sub
    
    'Animation
    If Grh.Started Then
        Grh.frame_counter = Grh.frame_counter + (timer_ticks_per_frame * Grh.frame_speed / 1000)

        If Grh.frame_counter > Grh_list(Grh.grh_index).frame_count Then
            If Grh.LoopTimes < 2 Then
                Grh.frame_counter = 1
                Grh.Started = False
            Else
                Grh.frame_counter = 1

                If Grh.LoopTimes <> LoopAdEternum Then
                    Grh.LoopTimes = Grh.LoopTimes - 1

                End If

            End If

        End If

    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    If Grh.frame_counter <= 0 Then Grh.frame_counter = 1
    grh_index = Grh_list(Grh.grh_index).frame_list(Grh.frame_counter)
    
    If grh_index = 0 Then Exit Sub 'This is an error condition
    
    'Center Grh over X,Y pos
    If center Then
        tile_width = Grh_list(grh_index).src_width / base_tile_size
        tile_height = Grh_list(grh_index).src_height / base_tile_size

        If tile_width <> 1 Then
            screen_x = screen_x - Int(tile_width * base_tile_size / 2) + base_tile_size / 2

        End If

        If tile_height <> 1 Then
            screen_Y = screen_Y - Int(tile_height * base_tile_size) + base_tile_size

        End If

    End If
    
    'Draw it to device
    DXEngine_TextureRender Grh_list(grh_index).texture_index, screen_x, screen_Y, Grh_list(grh_index).src_width, Grh_list(grh_index).src_height, Rgb_List, Grh_list(grh_index).Src_X, Grh_list(grh_index).Src_Y, Grh_list(grh_index).src_width, Grh_list(grh_index).src_height, Grh.alpha_blend, Grh.angle

End Sub

Public Sub Grh_Render_To_Hdc(ByVal grh_index As Long, desthdc As Long, ByVal screen_x As Long, ByVal screen_Y As Long, Optional transparent As Boolean = False)

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 5/02/2003
    'This method is SLOW... Don't use in a loop if you care about
    'speed!
    '*************************************************************
    If Grh_Check(grh_index) = False Then
        Exit Sub

    End If

    Dim Src_X      As Long

    Dim Src_Y      As Long

    Dim src_width  As Long

    Dim src_height As Long

    Dim file_index As Long

    'If it's animated switch grh_index to first frame
    If Grh_list(grh_index).frame_count <> 1 Then
        grh_index = Grh_list(grh_index).frame_list(1)

    End If

    file_index = Grh_list(grh_index).texture_index
    Src_X = Grh_list(grh_index).Src_X
    Src_Y = Grh_list(grh_index).Src_Y
    src_width = Grh_list(grh_index).src_width
    src_height = Grh_list(grh_index).src_height

    Call DXEngine_TextureToHdcRender(file_index, desthdc, screen_x, screen_Y, Src_X, Src_Y, src_width, src_height, transparent)

End Sub

Public Function GUI_Grh_Render(ByVal grh_index As Long, X As Long, Y As Long, Optional ByVal angle As Single, Optional ByVal alpha_blend As Boolean, Optional ByVal Color As Long) As Boolean

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 5/15/2003
    '
    '**************************************************************
    Dim temp_grh    As Grh

    Dim rpg_list(3) As Long

    If Grh_Check(grh_index) = False Then
        Exit Function

    End If

    rpg_list(0) = Color
    rpg_list(1) = Color
    rpg_list(2) = Color
    rpg_list(3) = Color

    Grh_Initialize temp_grh, grh_index, alpha_blend, angle
    
    Grh_Render temp_grh, X, Y, rpg_list
    
    GUI_Grh_Render = True

End Function

Public Sub Animations_Initialize(ByVal AnimationSpeed As Single, ByVal tile_size As Integer)
    Grh_Load_All
    base_tile_size = tile_size
    AnimBaseSpeed = AnimationSpeed

End Sub

Public Sub AnimSpeedCalculate(ByVal timer_elapsed_time As Single)
    timer_ticks_per_frame = AnimBaseSpeed * timer_elapsed_time

End Sub

Public Function Grh_Check(ByVal grh_index As Long) As Boolean

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 1/04/2003
    '
    '**************************************************************
    'check grh_index
    If grh_index > 0 And grh_index <= grh_count Then
        If Grh_list(grh_index).Active Then
            Grh_Check = True

        End If

    End If

End Function

Public Function GetMMColor(ByVal GrhIndex As Long) As Long
    GetMMColor = Grh_list(GrhIndex).MiniMap_color

End Function

