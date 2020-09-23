VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7920
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   162
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "editor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   301
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   528
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check2 
      Caption         =   "Breakable"
      Height          =   285
      Left            =   6840
      TabIndex        =   17
      Top             =   1170
      Width           =   1005
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Pinned"
      Height          =   240
      Left            =   6840
      TabIndex        =   16
      Top             =   855
      Width           =   1005
   End
   Begin VB.ListBox List2 
      Height          =   1035
      Left            =   3600
      TabIndex        =   9
      Top             =   3330
      Width           =   1320
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6840
      Top             =   1890
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   510
      Left            =   6840
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   8
      Top             =   270
      Width           =   510
   End
   Begin VB.TextBox Text1 
      Height          =   960
      Left            =   4410
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   4500
      Width           =   3435
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   285
      LargeChange     =   32
      Left            =   45
      Max             =   1600
      SmallChange     =   16
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3015
      Width           =   4875
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1995
      Left            =   4950
      TabIndex        =   4
      Top             =   2385
      Width           =   2895
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1125
         TabIndex        =   14
         Top             =   225
         Width           =   1635
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1125
         TabIndex        =   12
         Top             =   855
         Width           =   1635
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1125
         TabIndex        =   11
         Text            =   "1600"
         Top             =   540
         Width           =   1635
      End
      Begin VB.Label Label4 
         Caption         =   "World Name:"
         Height          =   240
         Left            =   90
         TabIndex        =   15
         Top             =   270
         Width           =   960
      End
      Begin VB.Label Label3 
         Caption         =   "Next Map:"
         Height          =   240
         Left            =   90
         TabIndex        =   13
         Top             =   900
         Width           =   870
      End
      Begin VB.Label Label2 
         Caption         =   "World Size:"
         Height          =   240
         Left            =   90
         TabIndex        =   10
         Top             =   585
         Width           =   870
      End
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   4950
      TabIndex        =   3
      Top             =   270
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save World"
      Height          =   465
      Left            =   90
      TabIndex        =   2
      Top             =   3915
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load World"
      Height          =   465
      Left            =   90
      TabIndex        =   1
      Top             =   3375
      Width           =   1185
   End
   Begin VB.PictureBox Picture1 
      Height          =   2940
      Left            =   45
      ScaleHeight     =   192
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   0
      Top             =   45
      Width           =   4860
   End
   Begin VB.Label Label1 
      Caption         =   "Object List:"
      Height          =   195
      Left            =   4950
      TabIndex        =   5
      Top             =   45
      Width           =   1905
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Dim SoList() As String
Dim OList() As String

Dim OSList() As String
Dim OSList2() As String


Dim pn As String, vf As Boolean, fs As Long, fof As Long

Dim CObj As Integer
Dim PrecMode As Boolean

Private Sub Form_Load()

Dim i As Integer
Dim j As Integer

gamePath = App.Path + "\"
WorldName = "World 1"
wWidth = 1600
Initialize Me
CObj = -1

gPak.LoadPaks gamePath
DX_Init
PrecMode = CBool(ReadINI("Editor", "PreciseMode", App.Path + "\editor.ini"))



fof = gPak.GetFile(gamePath + "objects.txt", pn, fs, vf)
        If fof = -1 Then
            Text1.Text = gPak.LoadText("objects.txt", 0, fs)
        ElseIf fof = -2 Then
            Exit Sub
        Else
            Text1.Text = gPak.LoadText(pn, fof, fs)
        End If
        
fof = gPak.GetFile(gamePath + "object.txt", pn, fs, vf)
        If fof = -1 Then
            OSList = Split(gPak.LoadText("object.txt", 0, fs), vbCrLf)
        ElseIf fof = -2 Then
            Exit Sub
        Else
            OSList = Split(gPak.LoadText(pn, fof, fs), vbCrLf)
        End If
        
        OList = Split(Text1.Text, vbCrLf)
        
        
        
        ReDim ObjX(UBound(OList))
        
            For i = 0 To UBound(OList)
                        SoList = Split(OList(i), ",")
                        List1.AddItem SoList(0)
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

Private Sub Form_Unload(Cancel As Integer)
DX_Terminate
Terminate
End
End Sub



Private Sub Check1_Click()
If CObj <> -1 Then ObjX(CObj).oPinned = Not ObjX(CObj).oPinned
End Sub

Private Sub Check2_Click()
If CObj <> -1 Then ObjX(CObj).oBreak = Not ObjX(CObj).oBreak
End Sub

Private Sub Command1_Click()
Dim fn As String
fn = CD(hWnd, "*.wrd" + Chr(0) + "*.wrd", "Load World", "*.wrd")
If fn = "" Then Exit Sub
LoadWorld fn
End Sub

Private Sub Command2_Click()
Dim fn As String
fn = CD(hWnd, "*.wrd" + Chr(0) + "*.wrd", "Save World", "*.wrd")
If fn = "" Then Exit Sub
SaveWorld fn
End Sub

Private Sub HScroll1_Change()
iX = HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
iX = HScroll1.Value
End Sub

Private Sub List1_Click()
Dim i As Integer
    For i = 0 To UBound(OList)
        If ObjX(i).oName = List1.Text Then
        Picture2.Cls
        fof = gPak.GetFile(gamePath + "gfx\obj\" + Mid(OSList(ObjX(i).oPic), 1, InStr(1, OSList(ObjX(i).oPic), ",") - 1), pn, fs, vf)
        'MsgBox gamePath + "gfx\" + Mid(OSList(Objs(i).oPic), 1, InStr(1, OSList(Objs(i).oPic), ",") - 1)
        If fof = -1 Then
             gPak.DrawBitmap "gfx\obj\" + Mid(OSList(ObjX(i).oPic), 1, InStr(1, OSList(ObjX(i).oPic), ",") - 1), 0, Picture2.hdc
        ElseIf fof = -2 Then
            
            Exit Sub
        Else
            gPak.DrawBitmap pn, fof, Picture2.hdc
        End If
        CObj = i
        Picture2.Refresh
        Exit For
        End If
    Next
End Sub
Sub SaveWorld(SFileName As String)
Dim ONum As Integer

If Dir(SFileName) <> "" Then Kill SFileName
Open SFileName For Binary As #1
Put #1, , WorldName
Put #1, , Bset
Put #1, , wWidth
ONum = UBound(Objs)
Put #1, , ONum
Put #1, , Objs
nextMap = Replace(nextMap, " ", Chr(0))
Put #1, , nextMap

Close #1
End Sub
Sub LoadWorld(SFileName As String)
Dim ONum As Integer


Open SFileName For Binary As #1
Get #1, , WorldName
Get #1, , Bset
Get #1, , wWidth
Get #1, , ONum
ReDim Objs(ONum)
Get #1, , Objs
Get #1, , nextMap
Close #1
Text2.Text = wWidth
HScroll1.Max = wWidth
Text3.Text = nextMap
Text4.Text = WorldName
    
    Set OFFBack = Nothing
    SurfDescB.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    SurfDescB.lHeight = 192
    SurfDescB.lWidth = wWidth
    
    Set OFFBack = DDraw.CreateSurface(SurfDescB)
    
End Sub

Private Sub List2_Click()
    Bset = List2.ListIndex + 1
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim f As Integer
If CObj = -1 Then Exit Sub
    If Button = 1 Then
    
    ReDim Preserve Objs(UBound(Objs) + 1)
    If PrecMode Then
    Objs(UBound(Objs)).X = iX + (X \ 16) * 16
    Objs(UBound(Objs)).Y = (Y \ 16) * 16
    Else
    Objs(UBound(Objs)).X = iX + X
    Objs(UBound(Objs)).Y = Y '(Y \ 16) * 16
    End If
    'Objs(UBound(Objs)).Y = Y '(Y \ 16) * 16
    
    Objs(UBound(Objs)).Pic = ObjX(CObj).oPic
    Objs(UBound(Objs)).Solid = ObjX(CObj).oSolid
    Objs(UBound(Objs)).OX = Objs(UBound(Objs)).X
    Objs(UBound(Objs)).W = ObjX(CObj).oWidth
    Objs(UBound(Objs)).h = ObjX(CObj).oHeight
    Objs(UBound(Objs)).onAir = False
    Objs(UBound(Objs)).Breakable = ObjX(CObj).oBreak
    Objs(UBound(Objs)).Pinned = ObjX(CObj).oPinned
    Objs(UBound(Objs)).Spawns = ObjX(CObj).oSpawns
    
    ElseIf Button = 2 Then
        For f = 0 To UBound(Objs)
            If Objs(f).X < X And Objs(f).X + Objs(f).W > X And _
               Objs(f).Y < Y And Objs(f).Y + Objs(f).h > Y Then
               DX_KillObject f
               'Objs(f).Pic = 0
               Exit For
            End If
        Next
    
    
    
    End If
    
    
    
End Sub

Private Sub Text2_Change()
If Not IsNumeric(Text2) Then MsgBox "Please enter number!", 32
wWidth = Val(Text2)
End Sub

Private Sub Text3_Change()
nextMap = Text3.Text
End Sub

Private Sub Text4_Change()
WorldName = Text4.Text
End Sub

Private Sub Timer1_Timer()
    DX_Draw_Back
    DX_Lines
    DX_Draw_Objects
   
    DX_FLIP
    DX_Gravity
    DX_getInput
    Caption = UBound(Objs)
    'If iX = 1280 Then iX = 0
    'If cx = 0 Then DX_getInput
    'cx = cx + 1
    'If cx = 1 Then cx = 0
    DoEvents
End Sub
