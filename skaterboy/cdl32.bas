Attribute VB_Name = "Modcdl32"
Public Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Function CD(OwnerHwnd As Long, Filter As String, Title As String, Optional defExt As String, Optional InitDir As String) As String
Dim ofn As OPENFILENAME
    ofn.lStructSize = Len(ofn)
    ofn.hWndOwner = OwnerHwnd
    ofn.hInstance = App.hInstance
    ofn.lpstrFilter = Filter
        ofn.lpstrFile = Space$(254)
        ofn.nMaxFile = 255
        ofn.lpstrDefExt = defExt
        ofn.lpstrFileTitle = Space$(254)
        ofn.nMaxFileTitle = 255
        'ofn.lpstrInitialDir = App.Path
        If InitDir <> "" Then ofn.lpstrInitialDir = InitDir
        ofn.lpstrTitle = Title
        ofn.flags = 0
        Dim a
        a = GetOpenFileName(ofn)

        If (a) Then
                If Mid(Trim(ofn.lpstrFile), Len(Trim(ofn.lpstrFile)), 1) = Chr(0) Then ofn.lpstrFile = Mid(Trim(ofn.lpstrFile), 1, Len(Trim(ofn.lpstrFile)) - 1)
                CD = ofn.lpstrFile
                
                Else
                Exit Function
        End If
End Function
Public Function FileNameCDL(WithPath As String)
Dim sWithoutPath As String
Dim iLen As Integer
Dim iWhere As Integer
sWithoutPath = WithPath
Do Until InStr(sWithoutPath, "\") = 0
iLen = Len(sWithoutPath)
iWhere = InStr(sWithoutPath, "\")
sWithoutPath = Right(sWithoutPath, iLen - iWhere)
Loop
FileNameCDL = sWithoutPath
End Function
