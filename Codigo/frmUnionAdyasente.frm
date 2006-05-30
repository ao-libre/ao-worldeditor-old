VERSION 5.00
Begin VB.Form frmUnionAdyacente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Union con Mapas Adyasentes"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   Icon            =   "frmUnionAdyasente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   6165
   StartUpPosition =   2  'CenterScreen
   Begin WorldEditor.lvButtons_H cmdAplicar 
      Height          =   375
      Left            =   3240
      TabIndex        =   29
      Top             =   4080
      Width           =   1335
      _extentx        =   2355
      _extenty        =   661
      caption         =   "&Aplicar"
      capalign        =   2
      backstyle       =   2
      cgradient       =   0
      font            =   "frmUnionAdyasente.frx":628A
      mode            =   0
      value           =   0
      cback           =   -2147483633
   End
   Begin VB.TextBox PosLim 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   360
      TabIndex        =   26
      Text            =   "89"
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox PosLim 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   5640
      TabIndex        =   24
      Text            =   "12"
      Top             =   3120
      Width           =   375
   End
   Begin VB.TextBox PosLim 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   5520
      TabIndex        =   22
      Text            =   "11"
      Top             =   3480
      Width           =   375
   End
   Begin VB.TextBox PosLim 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   480
      TabIndex        =   20
      Text            =   "90"
      Top             =   360
      Width           =   375
   End
   Begin VB.CheckBox Aplicar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aplicar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   1200
      TabIndex        =   19
      Top             =   2160
      Width           =   975
   End
   Begin VB.CheckBox Aplicar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aplicar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   18
      Top             =   2880
      Width           =   975
   End
   Begin VB.CheckBox Aplicar 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aplicar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3960
      TabIndex        =   17
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox Mapa 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   1800
      TabIndex        =   16
      Text            =   "0"
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox Mapa 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   2520
      TabIndex        =   15
      Text            =   "0"
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox Mapa 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   4200
      TabIndex        =   14
      Text            =   "0"
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox PosLim 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   840
      TabIndex        =   13
      Text            =   "11"
      Top             =   3600
      Width           =   375
   End
   Begin VB.TextBox PosLim 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   4800
      TabIndex        =   12
      Text            =   "90"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox PosLim 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   5640
      TabIndex        =   11
      Text            =   "10"
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox PosLim 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   360
      TabIndex        =   10
      Text            =   "91"
      Top             =   3240
      Width           =   375
   End
   Begin VB.CheckBox Aplicar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aplicar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3360
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Mapa 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   2520
      TabIndex        =   0
      Text            =   "0"
      Top             =   960
      Width           =   735
   End
   Begin WorldEditor.lvButtons_H cmdCancelar 
      Height          =   375
      Left            =   4680
      TabIndex        =   30
      Top             =   4080
      Width           =   1335
      _extentx        =   2355
      _extenty        =   661
      caption         =   "&Cancelar"
      capalign        =   2
      backstyle       =   2
      cgradient       =   0
      font            =   "frmUnionAdyasente.frx":62B6
      mode            =   0
      value           =   0
      cback           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H cmdDefault 
      Height          =   375
      Left            =   120
      TabIndex        =   31
      Top             =   4080
      Width           =   1095
      _extentx        =   1931
      _extenty        =   661
      caption         =   "&Default"
      capalign        =   2
      backstyle       =   2
      cgradient       =   0
      font            =   "frmUnionAdyasente.frx":62E2
      mode            =   0
      value           =   0
      cback           =   -2147483633
   End
   Begin VB.Label Label13 
      Caption         =   "NOTA: Mapa 0, borra el translado de mapa."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1320
      TabIndex        =   28
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   840
      X2              =   840
      Y1              =   840
      Y2              =   3360
   End
   Begin VB.Label Label12 
      Caption         =   "X:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   720
      Width           =   255
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   5280
      X2              =   5280
      Y1              =   720
      Y2              =   3240
   End
   Begin VB.Label Label11 
      Caption         =   "X:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5400
      TabIndex        =   25
      Top             =   3120
      Width           =   255
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      X1              =   1080
      X2              =   5280
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label Label10 
      Caption         =   "Y:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5280
      TabIndex        =   23
      Top             =   3480
      Width           =   255
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      X1              =   840
      X2              =   5040
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label9 
      Caption         =   "Y:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   360
      Width           =   255
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00404040&
      X1              =   960
      X2              =   5160
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00404040&
      X1              =   5160
      X2              =   5160
      Y1              =   3480
      Y2              =   600
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00404040&
      X1              =   5160
      X2              =   960
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00404040&
      X1              =   960
      X2              =   960
      Y1              =   3480
      Y2              =   600
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00008000&
      X1              =   120
      X2              =   6000
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label Label8 
      Caption         =   "Y:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5400
      TabIndex        =   9
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "Y:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "X:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "X:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4560
      TabIndex        =   6
      Top             =   120
      Width           =   255
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      X1              =   720
      X2              =   4920
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      X1              =   1200
      X2              =   5400
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mapa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mapa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mapa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mapa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   960
      Width           =   495
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   5040
      X2              =   5040
      Y1              =   360
      Y2              =   3240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   1080
      X2              =   1080
      Y1              =   840
      Y2              =   3600
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   2895
      Left            =   960
      Top             =   600
      Width           =   4215
   End
   Begin VB.Menu mnuDefault 
      Caption         =   "mnuDefault"
      Visible         =   0   'False
      Begin VB.Menu mnuBasica 
         Caption         =   "11,10 y 90,91 - Basica"
      End
      Begin VB.Menu mnuUlla 
         Caption         =   "9,7 y 92,94 - Ullathorpe"
      End
   End
End
Attribute VB_Name = "frmUnionAdyacente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Option Explicit

Private Sub Aplicar_Click(index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Dim i As Byte
cmdAplicar.Enabled = False
For i = 0 To 3
    If Aplicar(i).value = 1 Then cmdAplicar.Enabled = True
Next
End Sub

Private Sub cmdAplicar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next
Dim Y As Integer
Dim X As Integer

If Not MapaCargado Then
    Exit Sub
End If

' ARRIBA
If Mapa(0).Text > -1 And Aplicar(0).value = 1 Then
    Y = PosLim(1).Text
    For X = (PosLim(3).Text + 1) To (PosLim(2).Text - 1)
        If MapData(X, Y).Blocked = 0 Then
            MapData(X, Y).TileExit.Map = Mapa(0).Text
            If Mapa(0).Text = 0 Then
                MapData(X, Y).TileExit.X = 0
                MapData(X, Y).TileExit.Y = 0
            Else
                MapData(X, Y).TileExit.X = X
                MapData(X, Y).TileExit.Y = PosLim(4).Text
            End If
        End If
    Next
End If

' DERECHA
If Mapa(1).Text > -1 And Aplicar(1).value = 1 Then
    X = PosLim(2).Text
    For Y = (PosLim(1).Text + 1) To (PosLim(0).Text - 1)
        If MapData(X, Y).Blocked = 0 Then
            MapData(X, Y).TileExit.Map = Mapa(1).Text
                If Mapa(1).Text = 0 Then
                    MapData(X, Y).TileExit.X = 0
                    MapData(X, Y).TileExit.Y = 0
                Else
                    MapData(X, Y).TileExit.X = PosLim(6).Text
                    MapData(X, Y).TileExit.Y = Y
                End If
        End If
    Next
End If

' ABAJO
If Mapa(2).Text > -1 And Aplicar(2).value = 1 Then
    Y = PosLim(0).Text
    For X = (PosLim(3).Text + 1) To (PosLim(2).Text - 1)
        If MapData(X, Y).Blocked = 0 Then
            MapData(X, Y).TileExit.Map = Mapa(2).Text
                If Mapa(2).Text = 0 Then
                    MapData(X, Y).TileExit.X = 0
                    MapData(X, Y).TileExit.Y = 0
                Else
                    MapData(X, Y).TileExit.X = X
                    MapData(X, Y).TileExit.Y = PosLim(5).Text
                End If
        End If
    Next
End If

' IZQUIERDA
If Mapa(3).Text > -1 And Aplicar(3).value = 1 Then
    X = PosLim(3).Text
    For Y = (PosLim(1).Text + 1) To (PosLim(0).Text - 1)
        If MapData(X, Y).Blocked = 0 Then
            MapData(X, Y).TileExit.Map = Mapa(3).Text
                If Mapa(3).Text = 0 Then
                    MapData(X, Y).TileExit.X = 0
                    MapData(X, Y).TileExit.Y = 0
                Else
                    MapData(X, Y).TileExit.X = PosLim(7).Text
                    MapData(X, Y).TileExit.Y = Y
                End If
        End If
    Next
End If

'Set changed flag
MapInfo.Changed = 1
DoEvents

Unload Me
End Sub

Private Sub cmdCancelar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Unload Me
End Sub

Private Sub cmdDefault_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Me.PopupMenu mnuDefault
End Sub

''
'   Lee los Translados existentes en lugares claves en el Mapa
'

Private Sub LeerMapaExit()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next

Dim X As Integer
Dim Y As Integer

' ARRIBA
Mapa(0).Text = 0
Y = PosLim(1).Text
For X = (PosLim(3).Text + 1) To (PosLim(2).Text - 1)
        If MapData(X, Y).TileExit.Map > 0 Then
            Mapa(0).Text = MapData(X, Y).TileExit.Map
            Exit For
        End If
Next
Aplicar(0).value = 0

' DERECHA
Mapa(1).Text = 0
X = PosLim(2).Text
For Y = (PosLim(1).Text + 1) To (PosLim(0).Text - 1)
        If MapData(X, Y).TileExit.Map > 0 Then
            Mapa(1).Text = MapData(X, Y).TileExit.Map
            Exit For
        End If
Next
Aplicar(1).value = 0

' ABAJO
Mapa(2).Text = 0
Y = PosLim(0).Text
For X = (PosLim(3).Text + 1) To (PosLim(2).Text - 1)
        If MapData(X, Y).TileExit.Map > 0 Then
            Mapa(2).Text = MapData(X, Y).TileExit.Map
            Exit For
        End If
Next
Aplicar(2).value = 0

' IZQUIERDA
Mapa(3).Text = 0
X = PosLim(3).Text
For Y = (PosLim(1).Text + 1) To (PosLim(0).Text - 1)
        If MapData(X, Y).TileExit.Map > 0 Then
            Mapa(3).Text = MapData(X, Y).TileExit.Map
            Exit For
        End If
Next
Aplicar(3).value = 0


End Sub

Private Sub Form_Load()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call mnuBasica_Click
End Sub

Private Sub Mapa_Change(index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Aplicar(index).value = 1
End Sub

Private Sub Mapa_KeyPress(index As Integer, KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then
    KeyAscii = 0
    Exit Sub
End If

End Sub

Private Sub Mapa_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 26/05/06
'*************************************************
If LenB(Mapa(index).Text) = 0 Then Mapa(index).Text = 0
If Mapa(index).Text > 255 Then Mapa(index).Text = 255
End Sub

Private Sub mnuBasica_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
PosLim(0).Text = 91
PosLim(1).Text = 10
PosLim(2).Text = 90
PosLim(3).Text = 11
PosLim(4).Text = 90
PosLim(5).Text = 11
PosLim(6).Text = 12
PosLim(7).Text = 89
Call LeerMapaExit
End Sub

Private Sub mnuUlla_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
PosLim(0).Text = 94
PosLim(1).Text = 7
PosLim(2).Text = 92
PosLim(3).Text = 9
PosLim(4).Text = 93
PosLim(5).Text = 8
PosLim(6).Text = 10
PosLim(7).Text = 91
Call LeerMapaExit
End Sub

Private Sub PosLim_KeyPress(index As Integer, KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then
    KeyAscii = 0
    Exit Sub
End If

End Sub

Private Sub PosLim_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 26/05/06
'*************************************************
On Error Resume Next
If LenB(PosLim(index).Text) = 0 Then PosLim(index).Text = 1
If PosLim(index).Text > 99 Then PosLim(index) = 99
If PosLim(index).Text < 1 Then PosLim(index) = 1

Dim Y As Integer
Dim X As Integer

' ARRIBA
Y = PosLim(1).Text
For X = (PosLim(3).Text + 1) To (PosLim(2).Text - 1)
        If MapData(X, Y).TileExit.Map > 0 Then
            Mapa(0).Text = MapData(X, Y).TileExit.Map
            Aplicar(0).value = 0
            Exit For
        End If
Next

' DERECHA
X = PosLim(2).Text
For Y = (PosLim(1).Text + 1) To (PosLim(0).Text - 1)
        If MapData(X, Y).TileExit.Map > 0 Then
            Mapa(1).Text = MapData(X, Y).TileExit.Map
            Aplicar(1).value = 0
            Exit For
        End If
Next

' ABAJO
Y = PosLim(0).Text
For X = (PosLim(3).Text + 1) To (PosLim(2).Text - 1)
        If MapData(X, Y).TileExit.Map > 0 Then
            Mapa(2).Text = MapData(X, Y).TileExit.Map
            Aplicar(2).value = 0
            Exit For
        End If
Next

' IZQUIERDA
X = PosLim(3).Text
For Y = (PosLim(1).Text + 1) To (PosLim(0).Text - 1)
        If MapData(X, Y).TileExit.Map > 0 Then
            Mapa(3).Text = MapData(X, Y).TileExit.Map
            Aplicar(3).value = 0
            Exit For
        End If
Next

End Sub
