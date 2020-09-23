VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Pak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class for pak file access
'Pak files are used in some games such as Quake 1-2 ,Half Life etc.
'A Pak file is simply an uncompressed zip file,which consists of game resources
'such as bitmaps,waves or any other specific game files
'this module first loads all pak files from a specified folder
'looks up in them and adds files names in an array
'then,later when GetFile is called
'it searches all file names  and return file offset and pak file name
'Later specific functions such as DrawBitmap uses this infos to load data from pak file

Private Type pakheader_t
    Magic       As String * 4   ' Name of the new WAD format
    diroffset   As Long         ' Position of WAD directory from start of file
    dirsize     As Long         ' Number of entries * 0x40 (64 char)
End Type

Private Type pakentry_t
    filename    As String * 56  ' Name of the file, Unix style, with extension,
    ' 50 chars, padded with '\0'.
    offset      As Long         ' Position of the entry in PACK file
    size   As Long              ' Size of the entry in PACK file
End Type

Private Type pakMain
    Header As pakheader_t
    Entry() As pakentry_t
End Type

Dim PAKS() As pakMain

Dim h As pakheader_t
Dim e() As pakentry_t
Dim E1 As pakentry_t




Public Function LoadPaks(InThisPath As String) As Integer
    
    Dim loopvar As Integer
    Dim C As Integer
    
    ReDim PAKS(C)
    
    Do
        If Dir(InThisPath + "pak" & C & ".pak") = "" Then Exit Do
        If C <> 0 Then ReDim Preserve PAKS(C)
        
        Open InThisPath + "pak" & C & ".pak" For Binary As #1
        Get #1, , PAKS(C).Header
        Seek #1, PAKS(C).Header.diroffset + 1
        Get #1, , E1
        ReDim PAKS(C).Entry(PAKS(C).Header.dirsize / 64 - 1)
        
        For loopvar = 0 To UBound(PAKS(C).Entry)
            'Form1.List1.AddItem E1.filename
            PAKS(C).Entry(loopvar).filename = Replace(E1.filename, Chr(0), " ")
            PAKS(C).Entry(loopvar).offset = E1.offset
            PAKS(C).Entry(loopvar).size = E1.size
            Get #1, , E1
            
        Next
        Close #1
        C = C + 1
        
    Loop
    If C = 0 Then DX_log "Can't find any pak files!"
    LoadPaks = C
   
End Function
Public Function GetFile(FName As String, pakName As String, FSize As Long, VFile As Boolean) As Long
    Dim i As Integer
    Dim j As Integer
    Dim iFind As Boolean
    
    FName = Replace(FName, gamePath, "") 'Pure filename
    FName = Replace(FName, "\", "/") '    Unix Style
    
    iFind = False
    
    For j = 0 To UBound(PAKS)
        For i = 0 To UBound(PAKS(j).Entry)
            'MsgBox LCase(FName) & vbCr & LCase(PAKS(j).Entry(i).filename), , Len(FName) & " " & Len(PAKS(j).Entry(i).filename)
            If LCase(FName) = Trim(LCase(PAKS(j).Entry(i).filename)) Then
                pakName = "pak" & j & ".pak"
                VFile = True
                FSize = PAKS(j).Entry(i).size
                GetFile = PAKS(j).Entry(i).offset
                iFind = True
                
            End If
        Next i
    Next j
    
    FName = Replace(FName, "/", "\") ' back to windows
    
    If Dir(gamePath + FName) <> "" Then
        pakName = ""
        VFile = False
        FSize = FileLen(gamePath + FName)
        GetFile = -1
        iFind = True
    End If
    
    If iFind = False Then GetFile = -2
    
End Function
Public Function DrawBitmap(pName As String, BOffset As Long, ToDC As Long) As Long
Dim fHead As BITMAPFILEHEADER
Dim iHead As BITMAPINFO
Dim intBMPFile As Integer
Dim BMPData() As Byte

Dim bmTYPE As BITMAPINFO

intBMPFile = FreeFile

Open gamePath + pName For Binary As #intBMPFile
If BOffset <> 0 Then Seek #intBMPFile, BOffset + 1
Get #intBMPFile, , fHead
Get #intBMPFile, , iHead.BMInfoHeader

    If fHead.bfType <> "BM" Then
        MsgBox "Invalid Bitmap Header! " & fHead.bfType, 16
        Close #intBMPFile
        Exit Function
    End If
    If iHead.BMInfoHeader.biCompression <> 0 Then
        MsgBox "Unsupported compression algorithm", 16
        Close #intBMPFile
        Exit Function
    End If
        
    If iHead.BMInfoHeader.biClrUsed <> 0 Then
            For i = 0 To iHead.BMInfoHeader.biClrUsed - 1
                Get #intBMPFile, , iHead.Colors(i).peBlue
                Get #intBMPFile, , iHead.Colors(i).peGreen
                Get #intBMPFile, , iHead.Colors(i).peRed
                Get #intBMPFile, , iHead.Colors(i).peReserved
            Next i
            
        ElseIf iHead.BMInfoHeader.biBitCount = 8 Then
            Get #intBMPFile, , iHead.Colors
        End If
        'Size the BMPData array
        If iHead.BMInfoHeader.biBitCount = 8 Then
            ReDim BMPData(FileSize(iHead.BMInfoHeader.biWidth, iHead.BMInfoHeader.biHeight))
        Else
            ReDim BMPData(iHead.BMInfoHeader.biSizeImage - 1)
        End If
        'Fill the BMPData array
        Get #intBMPFile, , BMPData
        'Ensure info is correct
        If iHead.BMInfoHeader.biBitCount = 8 Then
            fHead.bfOhFileBits = 1078
            iHead.BMInfoHeader.biSizeImage = FileSize(iHead.BMInfoHeader.biWidth, iHead.BMInfoHeader.biHeight)
            iHead.BMInfoHeader.biClrUsed = 0
            iHead.BMInfoHeader.biClrImportant = 0
            iHead.BMInfoHeader.biXPelsPerMeter = 0
            iHead.BMInfoHeader.biYPelsPerMeter = 0
        End If
        'Draw It
        StretchDIBits ToDC, 0, 0, iHead.BMInfoHeader.biWidth, iHead.BMInfoHeader.biHeight, 0, 0, iHead.BMInfoHeader.biWidth, iHead.BMInfoHeader.biHeight, BMPData(0), iHead, DIB_RGB_COLORS, SRCCOPY
        Erase BMPData
Close #intBMPFile
End Function
Private Function FileSize(lngWidth As Long, lngHeight As Long) As Long
    'Return the size of the image portion of the bitmap
    If lngWidth Mod 4 > 0 Then
        FileSize = ((lngWidth \ 4) + 1) * 4 * lngHeight - 1
    Else
        FileSize = lngWidth * lngHeight - 1
    End If
End Function
Public Sub LoadWave(pName As String, WOffSet As Long, ToBuffer As DirectSoundBuffer)
Dim gudtHeader As WAVETYPE
Dim glngChunkSize As Long
Dim gbytData() As Byte

Dim intWAVFile As Integer
Dim i As Long
Dim strTemp As String * 4
Dim blnFound As Boolean

Dim udtBufferDesc As DSBUFFERDESC
Dim udtFormat As WAVEFORMATEX
   
    'Open the wave
    intWAVFile = FreeFile()
    Open gamePath + pName For Binary Access Read As intWAVFile
        'Get the header info
        
        If WOffSet <> 0 Then
        Get intWAVFile, WOffSet + 1, gudtHeader
        Else
        Get intWAVFile, , gudtHeader
        WOffSet = 4
        End If
        'MsgBox gudtHeader.strFormatID
        'Find the "data" portion of the file
        For i = WOffSet To LOF(intWAVFile)
            Get intWAVFile, i, strTemp
            If strTemp = "data" Then
                
                blnFound = True
                Exit For
            End If
        Next i
        'Ensure this is a wave file
        If blnFound = False Then
            MsgBox "Invalid wave data.", vbCritical, "Invalid Wave"
            Close intWAVFile
            Exit Sub
        End If
        'Get the data information
        Get intWAVFile, , glngChunkSize
        ReDim gbytData(glngChunkSize)
        Get intWAVFile, , gbytData
    Close intWAVFile
    
      'Set the Wave Format
    With udtFormat
        .nFormatTag = gudtHeader.intFormat
        .nChannels = gudtHeader.intChannels
        .lSamplesPerSec = gudtHeader.lngSamplesPerSec
        .nBitsPerSample = gudtHeader.intBitsPerSample
        .nBlockAlign = gudtHeader.intBlockAlign
        .lAvgBytesPerSec = gudtHeader.lngAvgBytesPerSec
    End With
    
    'Create the buffer
    udtBufferDesc.lBufferBytes = glngChunkSize
    Set ToBuffer = DXSound.CreateSoundBuffer(udtBufferDesc, udtFormat)
    
    'Load the buffer with data
    ToBuffer.WriteBuffer 0, glngChunkSize, gbytData(0), DSBLOCK_ENTIREBUFFER
    
    Erase gbytData
    
End Sub
Public Function LoadText(pName As String, TOffset As Long, TSize As Long) As String
Dim intText As Integer
Dim tm As String
    intText = FreeFile
    
    Open gamePath + pName For Binary As #intText
    If TOffset <> 0 Then Seek #intText, TOffset + 1
    tm = String(TSize, Chr(0))
    Get #intText, , tm
    Close #intText
    
    LoadText = tm
    
End Function
Public Sub LoadMap(pName As String, MOffset As Long)
Dim intMap As Integer
Dim ONum As Integer
'Exit Sub
intMap = FreeFile
Open gamePath + pName For Binary As #intMap
If MOffset <> 0 Then Seek #intMap, MOffset + 1
Get intMap, , WorldName
Get intMap, , Bset
Get intMap, , wWidth
Get intMap, , ONum
ReDim Objs(ONum)
Get #intMap, , Objs
Get #intMap, , nextMap
Close #intMap
    

End Sub
Public Sub LoadMedia(pName As String, MOffset As Long, FSize As Long)
'    Dim DataLength As Long      'file length
'    Dim DataStore() As Byte     'data array
'    Dim intMedia As Integer
'
'
'    Call BASS_StreamFree(SMHandle)
'    Call BASS_MusicFree(SMHandle)
'
'    'Free Memory
'    Call GlobalFree(ByVal DataPtr)
'
'    intMedia = FreeFile
'    DataLength = FSize
'
'    If DataLength = 0 Then
'
'        Exit Sub
'    End If
'
'    'Allocate a Pointer to Memory
'    DataPtr = GlobalAlloc(GMEM_FIXED, DataLength)
'
'    Open gamePath + pName For Binary Lock Read Write As #intMedia
'    ReDim DataStore(DataLength) As Byte
'
'    'Insert All The File Data into a Byte Array
'    If MOffSet <> 0 Then Seek #1, MOffSet + 1
'    Get intMedia, , DataStore
'
'    'copy file into allocated memory location
'    Call CopyMemory(ByVal DataPtr, DataStore(0), DataLength)
'
'    Close #intMedia
'
'    'read data from memory location
'    SMHandle = BASS_StreamCreateFile(BASSTRUE, DataPtr, 0, DataLength, 0)
'
'    If SMHandle = 0 Then
'        SMHandle = BASS_MusicLoad(BASSTRUE, DataPtr, 0, DataLength, BASS_MUSIC_LOOP Or BASS_MUSIC_RAMP Or BASS_MUSIC_CALCLEN)
'        If SMHandle <> 0 Then
'            isMusic = True
'            MusicTimer = Timer
'        End If
'    End If
'
'    If SMHandle = 0 Then
'        'free stream and music (if any)
'        Call BASS_StreamFree(SMHandle)
'        Call BASS_MusicFree(SMHandle)
'
'        ' free memory
'        Call GlobalFree(ByVal DataPtr)
'
'        MsgBox "Error: Couldn't Get Data from Memory", vbCritical, "Memory Data error..."
'
'
'        Exit Sub
'    End If
'
'
'
'
'    Call BASS_MusicPlay(SMHandle)
'    Call BASS_StreamPlay(SMHandle, 0, BASS_SAMPLE_LOOP)
'

End Sub

