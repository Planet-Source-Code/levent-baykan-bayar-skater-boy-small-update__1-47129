Attribute VB_Name = "ModDX"
Option Explicit
'This is the main game module including everything we need

'API for user input,I didn't use dxinput for now
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'API for copying obj array ,see dx_killobject
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

'UDT for objects
Type gObject
OX As Integer 'current X pos in screen
OY As Integer
x As Integer 'original X pos
y As Integer
W As Integer
h As Integer
Pic As Integer 'pic surf code
Spawns As Integer 'spawns when killed,if not -1
Pinned As Boolean 'pinned to screen,does not effected by gravity
onAir As Boolean 'check if obj is falling
Solid As Boolean 'if solid then sprite can't walk in them
Breakable As Boolean 'if obj is breakble
End Type

'UDT for object.txt,objects.txt data files
Public Type objData
    oName As String
    oPic As Integer
    oSolid As Boolean
    oBreak As Boolean
    oPinned As Boolean
    oSpawns As Integer
    oWidth As Integer
    oHeight As Integer
End Type
Global ObjX() As objData


Dim SurfDescP As DDSURFACEDESC2 'Surface desc.
Dim SurfDescB As DDSURFACEDESC2, SurfDescBk As DDSURFACEDESC2
Dim Clipper As DirectDrawClipper
Dim Primary As DirectDrawSurface7
Public OFFBack As DirectDrawSurface7
Dim BackSurf(1 To 3) As DirectDrawSurface7
Dim PlayOnce As Boolean 'To detect if sound palyed before
Dim ObjSurf() As DirectDrawSurface7
Dim SurfDescO() As DDSURFACEDESC2
Dim Sobj() As String
Dim S2obj() As String

Global Objs() As gObject
Global Lives As Integer 'number of lives

Global gPak As New Pak 'PAK file object
Global gamePath As String 'game path for file access
Global Bset As Integer 'backg picture surf. index
Global fps As Long ':)
Global iX As Integer 'sprite pos
Global Running As Boolean 'check is running
Global FPSLimit As Integer 'limit fps
Global WorldName As String * 32 'name of current map
Global nextMap As String * 32 'next maps name, will be loaded at the end of current level
Global wWidth As Integer 'global width of world
Dim ErrSystem(14) As Boolean 'not works for now,but supposed to show in which sub error happened,for debugging purposes
Dim SCCKEY As DDCOLORKEY
Sub DX_Init()
Dim i As Integer
Dim S2obj() As String
On Error GoTo erS
Set DDraw = DX.DirectDrawCreate("") 'create ddraw etc
DDraw.SetCooperativeLevel Form1.hWnd, DDSCL_NORMAL
    Bset = 2
    
    SurfDescP.lFlags = DDSD_CAPS
    SurfDescP.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    If Dir(gamePath + "log.txt") <> "" Then Kill gamePath + "log.txt" 'create log file
    Set Primary = DDraw.CreateSurface(SurfDescP)
   
    
    
        DX_loadBack
        DX_loadObject
        DX_LoadObjData
        DX_LoadWorld "w1.wrd"

    OFFBack.SetFont Form1.Font
    
    Set Clipper = DDraw.CreateClipper(0)
    Clipper.SetHWnd Form1.hWnd
    Primary.SetClipper Clipper
Exit Sub
erS:
If ErrSystem(0) = False Then
DX_log "An error occured! " & Err.Number & ":" & Err.Description
ErrSystem(0) = True
End If
End Sub
Sub DX_Draw_Back()
    On Error GoTo erS
    '------------------
    'Simple just draws backsurf(bset ) to backbuffer
    'wwidth/1600 times
    'so if wwidth is 3200 it will be drawn to subsequent times
    '------------------
    Dim Destrect As RECT
    Dim srcRect(1) As RECT
    Dim a As Long
    Dim i As Integer
    'Form1.Caption = SurfDescO.lWidth & "  " & SurfDescO.lHeight
    BackSurf(Bset).GetSurfaceDesc SurfDescBk
    
    srcRect(0).Left = 0 ' iX
    srcRect(0).Top = 0
    srcRect(0).Right = 1600  'iX + 320 'SurfDescO.lWidth
    srcRect(0).Bottom = 192 'SurfDescO.lHeight

    
   
    For i = 0 To wWidth \ 1600 - 1
    
    Destrect.Left = i * 1600
    Destrect.Right = (i * 1600) + 1600 '320
    Destrect.Top = 0
    Destrect.Bottom = 192
    
    a = OFFBack.Blt(Destrect, BackSurf(Bset), srcRect(0), DDBLT_WAIT)
    'a = OFFBack.Blt(Destrect, BackSurf(Bset), srcRect(1), DDBLT_WAIT)
    
    Next
    
    srcRect(0).Left = 0 ' iX
    srcRect(0).Top = 0
    srcRect(0).Right = wWidth - (i * 1600) 'iX + 320 'SurfDescO.lWidth
    srcRect(0).Bottom = 192 'SurfDescO.lHeight
    
    Destrect.Left = (i * 1600)
    Destrect.Right = wWidth '320
    Destrect.Top = 0
    Destrect.Bottom = 192
    
    a = OFFBack.Blt(Destrect, BackSurf(Bset), srcRect(0), DDBLT_WAIT)

Exit Sub
erS:
If ErrSystem(1) = False Then
DX_log "An error occured! " & Err.Number & ":" & Err.Description
ErrSystem(1) = True
End If
End Sub
Public Sub DX_Draw_Objects()
    On Error GoTo erS
    '---------------
    'Draws all objects onto backbuffer from last to 0
    'to draw sprite above all other things
    '----------------
    
    Dim Destrect As RECT
    Dim srcRect As RECT
    Dim a As Long
    Dim i As Integer
    
    For i = UBound(Objs) To 0 Step -1
    
    'a little optimization
    If Objs(i).OX + Objs(i).W >= 0 And Objs(i).OX <= 320 Then
    
    srcRect.Left = 0
    srcRect.Top = 0
    srcRect.Bottom = SurfDescO(Objs(i).Pic).lHeight
    srcRect.Right = SurfDescO(Objs(i).Pic).lWidth
    
    Destrect.Left = Objs(i).x
    Destrect.Top = Objs(i).y
    Destrect.Right = Objs(i).x + Objs(i).W
    Destrect.Bottom = Objs(i).y + Objs(i).h
    
    
    a = OFFBack.Blt(Destrect, ObjSurf(Objs(i).Pic), srcRect, DDBLT_KEYSRC)
    End If
    
    Next
Exit Sub
erS:
If ErrSystem(0) = False Then
DX_log "An error occured! " & Err.Number & ":" & Err.Description
ErrSystem(0) = True
End If
End Sub
Public Sub DX_FLIP()
On Error GoTo erS
    '---------------
    'FLIPS backbuffer onto primary(visible) surface
    'Nothing special
    '---------------
    
    Dim Destrect As RECT
    Dim sRect As RECT
    Dim s2Rect As RECT
    Dim i As Integer
    
    sRect.Left = iX
    sRect.Right = iX + 320
    sRect.Top = 0
    sRect.Bottom = 192
    
    
    
    DX.GetWindowRect Form1.Picture1.hWnd, Destrect

    For i = 1 To Lives
    s2Rect.Left = iX + ((i - 1) * 16)
    s2Rect.Right = s2Rect.Left + 16
    s2Rect.Top = 0
    s2Rect.Bottom = 19
    OFFBack.Blt s2Rect, ObjSurf(4), tRect(0, 0, 16, 19), DDBLT_KEYSRC
    Next
    
    OFFBack.DrawText iX + 4, 20, WorldName, False
    
    Primary.Blt Destrect, OFFBack, sRect, DDBLT_WAIT
Exit Sub
erS:
If ErrSystem(0) = False Then
DX_log "An error occured! " & Err.Number & ":" & Err.Description
ErrSystem(0) = True
End If
End Sub
Public Sub Wait(howlong)
    '-----------------------
    'wait #ofmiliseconds
    'example wait 3 will wait 3 miliseconds
    '-----------------------
    Dim temptime As Variant
    
    
    temptime = Timer
    Do
        DoEvents
        'DX_SMDoevents
    Loop While Timer < temptime + howlong
End Sub
Private Sub DX_LoadObjData()
    '-------------------
    'loads text data files
    '-------------------
    
Dim i As Integer
Dim j As Integer

Dim SoList() As String
Dim OList() As String

Dim OSList() As String
Dim OSList2() As String
Dim tmS2 As String

Dim pn As String, vf As Boolean, fs As Long, fof As Long

fof = gPak.GetFile(gamePath + "objects.txt", pn, fs, vf)
        If fof = -1 Then
            tmS2 = gPak.LoadText("objects.txt", 0, fs)
        ElseIf fof = -2 Then
            Exit Sub
        Else
            tmS2 = gPak.LoadText(pn, fof, fs)
        End If
        
fof = gPak.GetFile(gamePath + "object.txt", pn, fs, vf)
        If fof = -1 Then
            OSList = Split(gPak.LoadText("object.txt", 0, fs), vbCrLf)
        ElseIf fof = -2 Then
            Exit Sub
        Else
            OSList = Split(gPak.LoadText(pn, fof, fs), vbCrLf)
        End If
        
        OList = Split(tmS2, vbCrLf)
        
        
        
        ReDim ObjX(UBound(OList))
        
            For i = 0 To UBound(OList)
                        SoList = Split(OList(i), ",")
                        'List1.AddItem SoList(0)
                        ObjX(i).oName = SoList(0)
                        ObjX(i).oPic = SoList(1)
                        ObjX(i).oSolid = SoList(2)
                        ObjX(i).oSpawns = SoList(3)
            Next i
            For i = 0 To UBound(OSList)
                        OSList2 = Split(OSList(i), ",")
                        For j = 0 To UBound(OList)
                        If ObjX(j).oPic = i Then
                            ObjX(j).oWidth = OSList2(1)
                            ObjX(j).oHeight = OSList2(2)
                        End If
                        Next j
            Next i
End Sub
Private Sub DX_loadBack()
'loads back surfaces: checks if its in pak file
'if its not then look at current folder structure
'if file exists in both pak and current folder ,
'folder will be prefered

Dim i As Integer
Dim tmDC As Long
Dim pn As String, vf As Boolean, fs As Long, fof As Long
    SurfDescBk.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
    SurfDescBk.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    SurfDescBk.lHeight = 192:    SurfDescBk.lWidth = 1600
    
    i = 1
    Do
        fof = gPak.GetFile(gamePath + "gfx\back\back" & i & ".bmp", pn, fs, vf)
        If fof = -2 Then Exit Do
        
        Set BackSurf(i) = DDraw.CreateSurface(SurfDescBk)
        tmDC = BackSurf(i).GetDC
        If fof = -1 Then
            gPak.DrawBitmap "gfx\back\back" & i & ".bmp", 0, tmDC
            DX_log "Loaded Bitmap [gfx\back\back" & i & ".bmp]" & " from " & gamePath & "."
        ElseIf fof = -2 Then
            DX_log "Can't Load Bitmap [gfx\back\back" & i & ".bmp] !"
            Exit Do
        Else
            gPak.DrawBitmap pn, fof, tmDC
            DX_log "Loaded Bitmap [gfx\back\back" & i & ".bmp]" & " from " & pn & " file."
        End If
        BackSurf(i).ReleaseDC tmDC
        i = i + 1
    Loop
    
  
End Sub
Private Sub DX_loadObject()
Dim i As Integer
Dim pn As String, vf As Boolean, fs As Long, fof As Long
Dim rectEmpty As RECT
Dim tmDC As Long
fof = gPak.GetFile(gamePath + "object.txt", pn, fs, vf)
        If fof = -1 Then
            Sobj = Split(gPak.LoadText("object.txt", 0, fs), vbCrLf)
            DX_log "Loaded Datafile [object.txt] from [" & gamePath & "]"
        ElseIf fof = -2 Then
            DX_log "Can't load object data !"
            Exit Sub
        Else
            Sobj = Split(gPak.LoadText(pn, fof, fs), vbCrLf)
            DX_log "Loaded Datafile [object.txt] from [" & pn & "]"
        End If

ReDim Preserve ObjSurf(UBound(Sobj))
ReDim Preserve SurfDescO(UBound(Sobj))
For i = 0 To UBound(Sobj)
        S2obj = Split(Sobj(i), ",")
        SurfDescO(i).lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        SurfDescO(i).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
        SurfDescO(i).lHeight = S2obj(2): SurfDescO(i).lWidth = S2obj(1)
        Set ObjSurf(i) = DDraw.CreateSurface(SurfDescO(i))
         
        ObjSurf(i).SetColorKey DDCKEY_SRCBLT, SCCKEY
        
        tmDC = ObjSurf(i).GetDC
        fof = gPak.GetFile(App.Path + "\gfx\obj\" & S2obj(0), pn, fs, vf)
        If fof = -1 Then
            gPak.DrawBitmap "\gfx\obj\" & S2obj(0), 0, tmDC
            DX_log "Loaded Bitmap [\gfx\obj\" & S2obj(0) & "] from " & gamePath & "."
        ElseIf fof = -2 Then
            DX_log "Can't Load Bitmap [\gfx\obj\" & S2obj(0) & "]!"
            Exit Sub
        Else
            gPak.DrawBitmap pn, fof, tmDC
            DX_log "Loaded Bitmap [\gfx\obj\" & S2obj(0) & "] from " & pn & " file."
        End If
        ObjSurf(i).ReleaseDC tmDC
        ObjSurf(i).Lock rectEmpty, SurfDescO(i), DDLOCK_WAIT, 0
        SCCKEY.high = ObjSurf(i).GetLockedPixel(0, 0)
        SCCKEY.low = SCCKEY.high
        ObjSurf(i).Unlock rectEmpty
        ObjSurf(i).SetColorKey DDCKEY_SRCBLT, SCCKEY
        
Next
End Sub
Public Sub DX_Gravity()
Dim i As Integer
Dim t As Integer
Dim v As Integer

    For i = 0 To UBound(Objs)
        'If you want to use realistic object fall
        '(i.e. increasing speed while falling down)
        'uncomment lines below
        'Realistic Gravity Calculation by NDB
        't = Sqr((192 - Objs(i).y) / 5)
        'v = 0.982 * t
        '-----------------------------------
       
        If Objs(i).y + Objs(i).h <= 160 And DX_Coll(i) = -1 And Objs(i).Pinned = False Then Objs(i).y = Objs(i).y + 1 ' make this '1' -> 't' if you want to enable realistic gravity
        
        If Objs(i).y + Objs(i).h >= 160 Then Objs(i).onAir = False
    Next
    DX_CollVert 0
End Sub
Public Sub DX_getInput()
Dim i As Integer
'Check user input
If GetAsyncKeyState(vbKeySpace) <> 0 And Objs(0).onAir = False Then
    Objs(0).onAir = True
    If GetAsyncKeyState(vbKeyLeft) <> 0 Then
        For i = 1 To 30
        Objs(0).y = Objs(0).y - 1
        iX = iX - 1
        Next
    ElseIf GetAsyncKeyState(vbKeyRight) <> 0 Then
        For i = 1 To 30
        Objs(0).y = Objs(0).y - 1
        iX = iX + 1
        Next
    Else
        'If DX_CollVert(0) = -1 Then
        For i = 1 To 30
        Objs(0).y = Objs(0).y - 1
        Next
        'End If
    End If
    
    PlayOnce = False
    'GoTo Ex:
End If

 

If GetAsyncKeyState(vbKeyRight) <> 0 Then
    
    If DX_SideColl(0) <> 1 And iX < wWidth - 320 Then iX = iX + 2
End If

If GetAsyncKeyState(vbKeyLeft) <> 0 Then
    If iX > 0 And DX_SideColl(0) <> 2 Then iX = iX - 2
    
End If

Ex:

Objs(0).x = 148 + iX
Objs(0).OX = 148

For i = 1 To UBound(Objs)
    Objs(i).OX = Objs(i).x - iX
Next

End Sub
Public Function DX_Coll(OMe As Integer) As Integer
Dim i As Integer
DX_Coll = -1
'Vertical collision detection for gravity
'
'
For i = 0 To UBound(Objs)
    If i <> OMe Then
        If (Objs(i).OX + Objs(i).W > Objs(OMe).OX) And (Objs(i).OX < (Objs(OMe).OX + Objs(OMe).W)) And _
         (Objs(i).y = Objs(OMe).y + Objs(OMe).h) Then
        DX_Coll = i
        
        If i = 0 And Objs(i).Breakable Then DX_PlaySound "ugh.wav", True
          
        
        Objs(OMe).onAir = False
        ' = i & " " & OMe
        Exit Function
        End If
        
    End If
Next

DX_Coll = -1
End Function
Public Function DX_CollVert(OMe As Integer) As Integer
Dim i As Integer
'vertical collision detection
DX_CollVert = -1


For i = 0 To UBound(Objs)
    If i <> OMe Then
        If (Objs(i).OX + Objs(i).W > Objs(0).OX) And (Objs(i).OX < (Objs(0).OX + Objs(0).W)) And _
         (Objs(0).y = Objs(i).y + Objs(i).h) Then
            
            DX_CollVert = i
            
            If Objs(i).Breakable Then
            DX_KillObject i
            DX_PlaySound "crack.wav", False
            End If
        
        
        
        ' = i & " " & OMe
        Exit Function
        End If
        
    End If
Next

DX_CollVert = -1
End Function
Private Function DX_SideColl(OInd As Integer) As Integer
Dim i As Integer
'left-right collision detection

For i = 0 To UBound(Objs)
    If i <> OInd Then
        If Round((Objs(i).y + Objs(i).h) / 10) = Round((Objs(OInd).y + Objs(OInd).h) / 10) Then
            
            If Objs(i).OX + Objs(i).W = Objs(OInd).OX And Objs(i).Solid = True Then
            DX_SideColl = 2
            Exit Function
            End If
            
            If Objs(0).OX + Objs(0).W = Objs(i).OX And Objs(i).Solid = True Then
            DX_SideColl = 1
            Exit Function
            End If
            
        End If
    
    End If
Next
DX_SideColl = 0
End Function
Sub DX_PlaySound(SndName As String, POnce As Boolean)

'Plays sound file in the pak



Dim pn As String, vf As Boolean, fs As Long, fof As Long
fof = gPak.GetFile(gamePath + "sfx\" + SndName, pn, fs, vf)
        If fof = -1 Then
            gPak.LoadWave "sfx\" + SndName, 0, jBuffer
        ElseIf fof = -2 Then
            Exit Sub
        Else
            gPak.LoadWave pn, fof, jBuffer
        End If
        jBuffer.Play DSBPLAY_DEFAULT
PlayOnce = POnce
End Sub
 Sub DX_KillObject(ObjIndex As Integer)
    '--------------
    'Kills a specified object redims current object array
    'decreasing its object count
    'spawns a object if it originally spawns something
    'such as a crate spawning a heart
    '--------------
    
    On Error GoTo erS
    Dim tm() As gObject
    Dim i As Integer
    
    
    If Objs(ObjIndex).Spawns = -1 Then
    ReDim tm(UBound(Objs))
    CopyMemory tm(0), Objs(0), 24 * (UBound(Objs) + 1)
    ReDim Objs(UBound(Objs) - 1)
    For i = 0 To UBound(Objs)
        If i < ObjIndex Then
            Objs(i) = tm(i)
        ElseIf i = ObjIndex Then
            'Objs(i - 1) = tm(i)
        ElseIf i > ObjIndex Then
            Objs(i) = tm(i + 1)
        End If
    Next i
    Else
    Objs(ObjIndex).Pic = ObjX(Objs(ObjIndex).Spawns).oPic
    'MsgBox Objs(ObjIndex).Spawns
    Objs(ObjIndex).y = Objs(ObjIndex).y - 40 'jump up the new object
    Objs(ObjIndex).Pinned = False 'make the new object fall
    Objs(ObjIndex).Breakable = False 'make it unbreakable
    End If
'    ReDim tm(UBound(Objs))
'    CopyMemory tm(0), Objs(0), 24 * (UBound(Objs) + 1)
'    ReDim Objs(UBound(Objs) - 1)
'
'    For i = 0 To UBound(Objs)
'        If i < ObjIndex Then
'            Objs(i) = tm(i)
'        ElseIf i = ObjIndex Then
'            'Objs(i - 1) = tm(i)
'        ElseIf i > ObjIndex Then
'            Objs(i) = tm(i + 1)
'        End If
'    Next i

    Exit Sub
erS:


    DX_log "Attempted to kill invalid object! No:" & ObjIndex & " but MaxObject Index:" & UBound(Objs)
    Erase tm
    Running = False


End Sub
Sub DX_LoadWorld(SFileName As String)
'Dim ONum As Integer
'
'Open gamepath + "\maps\" + SFileName For Binary As #1
'Get #1, , WorldName
'Get #1, , Bset
'Get #1, , wWidth
'Get #1, , ONum
'ReDim Objs(ONum)
'Get #1, , Objs
'
'Close #1

'--------------
'Loads a map
'--------------

Dim pn As String, vf As Boolean, fs As Long, fof As Long
fof = gPak.GetFile(gamePath + "maps\" + SFileName, pn, fs, vf)
        
        If fof = -1 Then
            gPak.LoadMap "maps\" + SFileName, 0
            DX_log "Loaded Map [" + SFileName + "] from " & gamePath
        ElseIf fof = -2 Then
            DX_log "Error! Map not found! [" + SFileName + "]"
            Exit Sub
        Else
            DX_log "Loaded Map [" + SFileName + "] from " & pn
            gPak.LoadMap pn, fof
        End If
    
    Set OFFBack = Nothing
    SurfDescB.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    SurfDescB.lHeight = 192
    SurfDescB.lWidth = wWidth
    
    Set OFFBack = DDraw.CreateSurface(SurfDescB)
    OFFBack.SetForeColor 16777215
    OFFBack.SetFontTransparency True
End Sub
Sub DX_SMDoevents()
'little doevents

DoEvents
'DX_Draw_Back
DX_Draw_Objects
DoEvents
End Sub

Public Sub DX_log(LogIt As String)
'logs everything
Dim fx As Integer

fx = FreeFile
Open gamePath + "log.txt" For Append As #fx
Print #fx, Time & " - " & LogIt
Close #fx
End Sub
Function tRect(l As Integer, t As Integer, r As Integer, b As Integer) As RECT
'useful sub,like in delphi
tRect.Left = l
tRect.Bottom = b
tRect.Right = r
tRect.Top = t
End Function
Sub DX_Terminate()
'kills everything
Dim i As Integer

For i = 0 To UBound(ObjSurf)
Set ObjSurf(i) = Nothing
Next

For i = 1 To 3
Set BackSurf(i) = Nothing
Next


Set DDraw = Nothing
Set gPak = Nothing
End Sub

