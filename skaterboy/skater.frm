VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Skater Boy!"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4890
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   162
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "skater.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   220
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   326
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4185
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "100"
      ToolTipText     =   "Increase this value up to 1000 if you have performence problems"
      Top             =   2970
      Width           =   690
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4140
      Top             =   1755
   End
   Begin VB.PictureBox Picture1 
      Height          =   2940
      Left            =   0
      ScaleHeight     =   192
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   321
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4875
   End
   Begin VB.Label Label3 
      Caption         =   "Limit FPS to:"
      Height          =   240
      Left            =   3285
      TabIndex        =   2
      Top             =   3015
      Width           =   915
   End
   Begin VB.Label Label1 
      Height          =   240
      Left            =   0
      TabIndex        =   1
      Top             =   2970
      Width           =   825
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cx As Byte


Private Sub Form_Load()

Running = True 'set true so game runs
gamePath = App.Path + "\" 'game path for gfx,sfx etc
FPSLimit = 10 'fps limiter for faster computers
Lives = 5 'number of lives
Show 'show form before loop
Initialize Me 'Initialize DX7 sound system [modPak.bas]


gPak.LoadPaks gamePath 'load all files names' in pak files at gamepath
DX_Init 'Init DX7 surfaces,Load them ,load data, files map etc.

'ok run !!

Do While Running
    
    DX_Draw_Back 'draw background
    DX_Draw_Objects 'draw all objects including sprite
    
    DX_FLIP 'flip all back surf to screen
    DX_Gravity 'check gravity
    DX_getInput 'receive user input
    
    DoEvents
    fps = fps + 1 'increase fps look timer1_timer
    
    Wait FPSLimit / 1000 'simple fps limiter
    
    If Not Running Then End
Loop
'out of loop
DX_Terminate 'kill all surfaces,free memory
Terminate 'unload dxsound system
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
Running = False
End Sub


Private Sub Text1_Change()
If Not IsNumeric(Text1.Text) Then Text1.Text = 0
If Val(Text1.Text) > 1000 Then Text1.Text = 1000
If Val(Text1.Text) = 0 Then Text1.Text = 1
FPSLimit = 1000 / Val(Text1.Text)
End Sub

Private Sub Timer1_Timer()
Label1.Caption = "FPS: " & fps 'Show FPS         '& " - iX: " & iX & " wW: [" & wWidth - 320 & "]"

If iX >= wWidth - 320 And nextMap <> String(32, Chr(0)) Then 'load new map
        iX = 0 'set sprite pos
        
        DX_SMDoevents 'draw backg and wait
        OFFBack.DrawText 120, 100, "Congratulations!", False
        OFFBack.DrawText 100, 110, "You have finished " & WorldName, False
        DX_FLIP
        Wait 2
       
        'MsgBox Len(Mid(nextMap, 1, InStr(1, nextMap, Chr(0))))
        DX_LoadWorld Mid(nextMap, 1, InStr(1, nextMap, Chr(0)) - 1)
End If
    
fps = 0 'set it to 0 so we can refresh fps info every second.
End Sub
