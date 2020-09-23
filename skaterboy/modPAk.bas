Attribute VB_Name = "modPAk"
Option Explicit

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal DX As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
  
Global Const GMEM_FIXED = &H0
Global Const SRCCOPY = &HCC0020
Global Const DIB_RGB_COLORS = 0


'BASS MEDIA SYSTEM---
Public SMHandle As Long    'Stream/Module Handle
Public DataPtr As Long     'Data Pointer
'--------------------

Public Type BITMAPINFOHEADER '40 bytes
    biSize            As Long
    biWidth           As Long
    biHeight          As Long
    biPlanes          As Integer
    biBitCount        As Integer
    biCompression     As Long
    biSizeImage       As Long
    biXPelsPerMeter   As Long
    biYPelsPerMeter   As Long
    biClrUsed         As Long
    biClrImportant    As Long
End Type

Public Type Color
    b As Byte
    G As Byte
    r As Byte
End Type

Public Type PALETTEENTRY
    peRed As Byte
    peGreen As Byte
    peBlue As Byte
    peReserved As Byte
End Type


Public Type BITMAPINFO
    BMInfoHeader As BITMAPINFOHEADER
    Colors(255) As PALETTEENTRY
End Type


Public Type BITMAP
    bmTYPE As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Public Type BITMAPFILEHEADER
    bfType            As String * 2
    bfSize            As Long
    bfReserved1       As Integer
    bfReserved2       As Integer
    bfOhFileBits         As Long
End Type

Public Type Chunk
    Chuky As String * 4
    DataLen As Long
End Type




Public Type WAVETYPE
    strHead As String * 12
    strFormatID As String * 4
    lngChunkSize As Long
    intFormat As Integer
    intChannels As Integer
    lngSamplesPerSec As Long
    lngAvgBytesPerSec As Long
    intBlockAlign As Integer
    intBitsPerSample As Integer
End Type
'DirectX Variables
Global DX As DirectX7
Global DXSound As DirectSound
Global DDraw As DirectDraw7

Global jBuffer As DirectSoundBuffer


Public Sub Initialize(frmInit As Form)

    'Initialize DirectSound
    Set DX = New DirectX7
    Set DXSound = DX.DirectSoundCreate("")
    
    'Set the DirectSound object's cooperative level (Priority gives us sole control)
    DXSound.SetCooperativeLevel frmInit.hWnd, DSSCL_PRIORITY
        
        
    'BASS MEDIA SYSTEM----------------------------
        
    'If BASS_GetStringVersion <> "1.7" Then
    '    MsgBox "BASS version 1.7 was not loaded", vbCritical, "Bass.Dll"
    '    End
    'End If
    
    'If (BASS_Init(-1, 44100, 0, frmInit.hWnd) = 0) Then
    '    MsgBox "Error: Couldn't Initialize Digital Output", vbCritical, "Digital output"
    '    End
    'Else
    '    BASS_Start
    'End If
    '---------------------------------------------
        
        
End Sub


Private Sub StopAll()
    'Call BASS_Stop
    'Call BASS_StreamFree(SMHandle)
    'Call BASS_MusicFree(SMHandle)
    '
    'Call BASS_Free
    'free data from allocate memory
    'Call GlobalFree(ByVal DataPtr)
    'free Class Module
 
End Sub



Public Sub Terminate()

    'Terminate all
    StopAll
    
    Set jBuffer = Nothing
    Set DXSound = Nothing
    Set DX = Nothing

End Sub


