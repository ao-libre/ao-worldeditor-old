Attribute VB_Name = "modMusic"
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
' modMusic
'
' @author Torres Patricio (Pato)
' @version 1.0.0
' @date 20110110

Option Explicit

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long

Private isLoad As Boolean

Public Sub PlayMusic(ByRef strPath As String)
'*************************************************
'Author: Torres Patricio (Pato)
'Last modified: 01/10/10
'*************************************************
On Error GoTo ErrHandler

    If isLoad Then Call StopMusic
    
    Call LoadMusic(strPath)
    
    Call mciSendString("play mymusic", 0&, 0, 0)
    
    Exit Sub
ErrHandler:
End Sub

Private Function LoadMusic(ByRef strPath As String) As Boolean
'*************************************************
'Author: Torres Patricio (Pato)
'Last modified: 01/10/10
'*************************************************
On Error GoTo ErrHandler

    Dim ShortPath As String
    Dim res As Long
    
    ShortPath = Space$(256)
    res = GetShortPathName(strPath, ShortPath, 256)
    
    ShortPath = Left$(ShortPath, res)
     
    Call mciSendString("Open " & ShortPath & " Alias mymusic", 0&, 0, 0)
    
    LoadMusic = True
    isLoad = True
    
    Exit Function
ErrHandler:
End Function

Public Sub StopMusic()
'*************************************************
'Author: Torres Patricio (Pato)
'Last modified: 01/10/10
'*************************************************
On Error GoTo ErrHandler

    If isLoad Then
        If IsPlaying Then Call mciSendString("stop mymusic", 0&, 0, 0)
        
        Call mciSendString("close mymusic", 0&, 0, 0)
        isLoad = False
    End If
    Exit Sub
ErrHandler:
End Sub

Private Function IsPlaying() As Boolean
'*************************************************
'Author: Torres Patricio (Pato)
'Last modified: 01/10/10
'*************************************************
On Error GoTo ErrHandler
    Dim str As String * 10
 
    Call mciSendString("status mymusic mode", str, Len(str), 0)
    IsPlaying = (InStr(1, str, "playing") > 0)
    
    Exit Function
ErrHandler:
End Function
