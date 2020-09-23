Attribute VB_Name = "ModDX"
Option Explicit
'This is a slightly modified versiopn of DX.Bas

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Type gObject
OX As Integer
OY As Integer
X As Integer
Y As Integer
W As Integer
h As Integer
Pic As Integer
Spawns As Integer
Pinned As Boolean
onAir As Boolean
Solid As Boolean
Breakable As Boolean

End Type

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

Dim SurfDescP As DDSURFACEDESC2
Public SurfDescB As DDSURFACEDESC2, SurfDescBk As DDSURFACEDESC2
Dim Clipper As DirectDrawClipper
Dim Primary As DirectDrawSurface7
Public OFFBack As DirectDrawSurface7
Dim BackSurf(1 To 3) As DirectDrawSurface7
Dim PlayOnce As Boolean
Dim ObjSurf() As DirectDrawSurface7
Dim SurfDescO() As DDSURFACEDESC2
Dim Sobj() As String
Dim S2obj() As String

Global Objs() As gObject
Global Lives As Integer

Global gPak As New Pak
Global gamePath As String
Global Bset As Integer
Global fps As Long
Global iX As Integer
Global Running As Boolean
Global FPSLimit As Integer
Global WorldName As String * 32
Global wWidth As Integer
Global nextMap As String * 32
Dim SCCKEY As DDCOLORKEY

'Some Basic INI Access
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long



Sub DX_Init()
Dim i As Integer
Dim S2obj() As String

Set DDraw = DX.DirectDrawCreate("")
DDraw.SetCooperativeLevel Form1.hWnd, DDSCL_NORMAL
    Bset = 2
    
    SurfDescP.lFlags = DDSD_CAPS
    SurfDescP.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    If Dir(gamePath + "log.txt") <> "" Then Kill gamePath + "log.txt"
    Set Primary = DDraw.CreateSurface(SurfDescP)
   
    
    
        DX_loadBack
        DX_loadObject
    
    ReDim Objs(0)

    Objs(0).X = 132
    Objs(0).Y = 128
    Objs(0).W = 32: Objs(0).h = 32
    Objs(0).Pic = 1
    Objs(0).Pinned = False
    Objs(0).onAir = False

'    Randomize Timer
'    For i = 1 To 128
'
'    Objs(i).x = ((100 + (Rnd * 1300)) \ 16) * 16
'    Objs(i).y = ((20 + (Rnd * 50)) \ 16) * 16
'    Objs(i).OX = Objs(i).x
'    'Objs(i).OX = Objs(i).X
'    Objs(i).W = 16
'    Objs(i).h = 16
'    Objs(i).Pic = 3
'    Objs(i).Pinned = False
'    Objs(i).onAir = False
'    Objs(i).Solid = True
'    Next
   
    SurfDescB.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    SurfDescB.lHeight = 192
    SurfDescB.lWidth = wWidth
    
    Set OFFBack = DDraw.CreateSurface(SurfDescB)
    OFFBack.SetFont Form1.Picture1.Font
    
    Set Clipper = DDraw.CreateClipper(0)
    Clipper.SetHWnd Form1.Picture1.hWnd
    Primary.SetClipper Clipper
    
End Sub
Sub DX_Draw_Back()
      
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

End Sub

Public Sub DX_Draw_Objects()
    Dim Destrect As RECT
    Dim srcRect As RECT
    Dim a As Long
    Dim i As Integer
    
    For i = UBound(Objs) To 0 Step -1
    If Objs(i).OX >= 0 And Objs(i).OX <= 320 Then
    srcRect.Left = 0
    srcRect.Top = 0
    srcRect.Bottom = SurfDescO(Objs(i).Pic).lHeight
    srcRect.Right = SurfDescO(Objs(i).Pic).lWidth
    
    Destrect.Left = Objs(i).X
    Destrect.Top = Objs(i).Y
    Destrect.Right = Objs(i).X + Objs(i).W
    Destrect.Bottom = Objs(i).Y + Objs(i).h
    
    
    a = OFFBack.Blt(Destrect, ObjSurf(Objs(i).Pic), srcRect, DDBLT_KEYSRC)
    End If
    
    Next
    
End Sub
Public Sub DX_FLIP()
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
    
    OFFBack.DrawText iX + 260, 0, WorldName, False
    
    
    
    
    Primary.Blt Destrect, OFFBack, sRect, DDBLT_WAIT
   
End Sub
Public Sub Wait(howlong)
    ' USAGE: wait #ofseconds; example wait 3 will wait 3 seconds
    Dim temptime As Variant
    
    
    temptime = Timer
    Do
        DoEvents
        'DX_SMDoevents
    Loop While Timer < temptime + howlong
End Sub
Private Sub DX_loadBack()
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
        Form1.List2.AddItem "back" & i & ".bmp"
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
            DX_log "Loaded object data from " & gamePath & "."
        ElseIf fof = -2 Then
            DX_log "Can't load object data !"
            Exit Sub
        Else
            Sobj = Split(gPak.LoadText(pn, fof, fs), vbCrLf)
            DX_log "Loaded object data from " & pn & " file."
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
        fof = gPak.GetFile(gamePath + "gfx\obj\" & S2obj(0), pn, fs, vf)
        
        
        If fof = -1 Then
            gPak.DrawBitmap "gfx\obj\" & S2obj(0), 0, tmDC
            DX_log "Loaded " & "\gfx\obj\" & S2obj(0) & " from " & gamePath & "."
        ElseIf fof = -2 Then
            DX_log "Can't Load " & "\gfx\obj\" & S2obj(0) & "!"
            Exit Sub
        Else
            gPak.DrawBitmap pn, fof, tmDC
            DX_log "Loaded " & "\gfx\obj\" & S2obj(0) & " from " & pn & " file."
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
        'Realistic Gravity Calculation by NDB
        t = Sqr((192 - Objs(i).Y) / 5)
        v = 0.982 * t
        '-----------------------------------
        'If i = 3 Then Form1.Label3.Caption = v & " " & t
        'Form1.Label5.Caption = ""
        'Form1.Label5.Caption = DX_Coll(i)
        If Objs(i).Y + Objs(i).h <= 160 And DX_Coll(i) = -1 And Objs(i).Pinned = False Then Objs(i).Y = Objs(i).Y + 1
        
        If Objs(i).Y + Objs(i).h >= 160 Then Objs(i).onAir = False
    Next
    
End Sub
Public Sub DX_getInput()
Dim i As Integer

If GetAsyncKeyState(vbKeySpace) <> 0 And Objs(0).onAir = False Then
    Objs(0).onAir = True
    If GetAsyncKeyState(vbKeyLeft) <> 0 Then
    For i = 1 To 30
    Objs(0).Y = Objs(0).Y - 1
    iX = iX - 1
    Next
    ElseIf GetAsyncKeyState(vbKeyRight) <> 0 Then
    For i = 1 To 30
    Objs(0).Y = Objs(0).Y - 1
    iX = iX + 1
    
    Next
    Else
    For i = 1 To 30
    Objs(0).Y = Objs(0).Y - 1
    Next
    
    End If
    PlayOnce = False
    'GoTo Ex:
End If



If GetAsyncKeyState(vbKeyRight) <> 0 Then
    
    If DX_SideColl(0) <> 1 And iX < wWidth - 320 Then iX = iX + 1
End If

If GetAsyncKeyState(vbKeyLeft) <> 0 Then
    If iX > 0 And DX_SideColl(0) <> 2 Then iX = iX - 1
    
End If

Ex:

Objs(0).X = 148 + iX
Objs(0).OX = 148

For i = 1 To UBound(Objs)
    Objs(i).OX = Objs(i).X - iX
Next
'Form1.Label3.Caption = Objs(1).OX & " " & Objs(1).X
End Sub
Public Function DX_Coll(OMe As Integer) As Integer
Dim i As Integer
DX_Coll = -1


For i = 0 To UBound(Objs)
    If i <> OMe Then
        If (Objs(i).OX + Objs(i).W > Objs(OMe).OX) And (Objs(i).OX < (Objs(OMe).OX + Objs(OMe).W)) And _
         (Objs(i).Y = Objs(OMe).Y + Objs(OMe).h) Then
        DX_Coll = i
        If i = 0 Then DX_PlaySound "ugh.wav", True
        Objs(OMe).onAir = False
        ' = i & " " & OMe
        Exit Function
        End If
        
    End If
Next


DX_Coll = -1
End Function
Private Function DX_SideColl(OInd As Integer) As Integer
Dim i As Integer

For i = 0 To UBound(Objs)
    If i <> OInd Then
        If Round((Objs(i).Y + Objs(i).h) / 10) = Round((Objs(OInd).Y + Objs(OInd).h) / 10) Then
            
            If Objs(i).OX + Objs(i).W = Objs(OInd).OX And Objs(i).Solid = True Then
            DX_SideColl = 2
            Exit Function
            End If
            
            If Objs(0).OX + Objs(0).W = Objs(i).OX And Objs(i).Solid = True Then
            DX_SideColl = 1
            Exit Function
            End If
            
            'Form1.Label3.Caption = "ok " & i
        End If
    
    End If
Next
DX_SideColl = 0
End Function
Sub DX_PlaySound(SndName As String, POnce As Boolean)

If PlayOnce Then Exit Sub

Dim pn As String, vf As Boolean, fs As Long, fof As Long
fof = gPak.GetFile(gamePath + "sfx\" + SndName, pn, fs, vf)
                If fof = -1 Then
            gPak.LoadWave "\sfx\" + SndName, 0, jBuffer
        ElseIf fof = -2 Then
            Exit Sub
        Else
            gPak.LoadWave pn, fof, jBuffer
        End If
        jBuffer.Play DSBPLAY_DEFAULT
PlayOnce = POnce
End Sub
Sub DX_LoadWorld(FName As String)
Dim pn As String, vf As Boolean, fs As Long, fof As Long
End Sub
Sub DX_SMDoevents()
DoEvents
'DX_Draw_Back
DX_Draw_Objects
DoEvents
End Sub
Sub DX_Lines()
Dim i As Integer
Dim j As Integer
For i = 0 To 20
        OFFBack.DrawLine iX + i * 16, 0, iX + (i * 16), 192
    Next i
    For j = 0 To 12
        OFFBack.DrawLine iX, j * 16, iX + 320, j * 16
    Next j

End Sub
Public Sub DX_log(LogIt As String)
Dim fx As Integer

fx = FreeFile
Open gamePath + "log.txt" For Append As #fx
Print #fx, Time & " - " & LogIt
Close #fx
End Sub
Function tRect(l As Integer, t As Integer, r As Integer, b As Integer) As RECT
tRect.Left = l
tRect.Bottom = b
tRect.Right = r
tRect.Top = t
End Function
Sub DX_Terminate()
Dim i As Integer

For i = 0 To UBound(ObjSurf)
Set ObjSurf(i) = Nothing
Next

For i = 1 To 3
Set BackSurf(i) = Nothing
Next


Set DDraw = Nothing
End Sub
Sub DX_KillObject(ObjIndex As Integer)
    On Error GoTo erS
    Dim tm() As gObject
    Dim i As Integer
    
    'If Objs(ObjIndex).Spawns = -1 Then
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
    'Else
    'Objs(ObjIndex).Pic = ObjX(Objs(ObjIndex).Spawns).oPic
    'Objs(ObjIndex).Y = Objs(ObjIndex).Y - 40
    'End If
    Exit Sub
erS:
    
    MsgBox "Attempted to kill invalid object! No:" & ObjIndex & " but MaxObject Index:" & 0, 16
    
End Sub
Public Function ReadINI(strsection As String, strkey As String, strfullpath As String) As String
   Dim strbuffer As String
   Let strbuffer$ = String$(750, Chr$(0&))
   Let ReadINI$ = Left$(strbuffer$, GetPrivateProfileString(strsection$, ByVal LCase$(strkey$), "", strbuffer, Len(strbuffer), strfullpath$))
End Function

Public Sub WriteINI(strsection As String, strkey As String, strkeyvalue As String, strfullpath As String)
    Call WritePrivateProfileString(strsection$, strkey$, strkeyvalue$, strfullpath$)
End Sub

