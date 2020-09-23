VERSION 5.00
Begin VB.Form frmSimon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simon"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6180
   Icon            =   "frmSimon.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "frmSimon.frx":08DA
   ScaleHeight     =   334
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   412
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrAttract 
      Enabled         =   0   'False
      Interval        =   333
      Left            =   240
      Top             =   4560
   End
   Begin VB.PictureBox picButton 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1395
      Index           =   3
      Left            =   3165
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   164
      TabIndex        =   3
      Top             =   390
      Width           =   2460
   End
   Begin VB.PictureBox picButton 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2070
      Index           =   2
      Left            =   3210
      ScaleHeight     =   2070
      ScaleWidth      =   2430
      TabIndex        =   2
      Top             =   1965
      Width           =   2430
   End
   Begin VB.PictureBox picButton 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2070
      Index           =   1
      Left            =   510
      ScaleHeight     =   2070
      ScaleWidth      =   2430
      TabIndex        =   1
      Top             =   1965
      Width           =   2430
   End
   Begin VB.PictureBox picButton 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1380
      Index           =   0
      Left            =   570
      ScaleHeight     =   92
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   162
      TabIndex        =   0
      Top             =   390
      Width           =   2430
   End
   Begin VB.Label lblHigh 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5160
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblScore 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.Image imgRgn 
      Height          =   1395
      Index           =   3
      Left            =   3000
      Picture         =   "frmSimon.frx":655B4
      Top             =   5280
      Visible         =   0   'False
      Width           =   2460
   End
   Begin VB.Image imgRgn 
      Height          =   2070
      Index           =   2
      Left            =   3000
      Picture         =   "frmSimon.frx":6576B
      Top             =   6720
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.Image imgRgn 
      Height          =   2070
      Index           =   1
      Left            =   360
      Picture         =   "frmSimon.frx":659B7
      Top             =   6720
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.Image imgRgn 
      Height          =   1380
      Index           =   0
      Left            =   360
      Picture         =   "frmSimon.frx":65C06
      Top             =   5280
      Visible         =   0   'False
      Width           =   2430
   End
End
Attribute VB_Name = "frmSimon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'' === SIMON ===
'' Paul Bahlawan June 2, 2006
''
'' mGradient.bas by Carles P.V. (see module)
'' mRegionShape2.bas by LaVolpe (see module)
'' Base graphic borrowed from http://www.neave.com/games/simon/
''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Option Explicit

'MIDI stuff
Private Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
Private Declare Function midiOutOpen Lib "winmm.dll" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Private Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal dwMsg As Long) As Long
Private hmidi As Long

'Graphical stuff
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateRectRgnIndirect Lib "gdi32.dll" (ByRef lpRect As RECT) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32.dll" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'Game stuff
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private gameMode As Integer ' 0=game over,  1=show sequence,  2=player input,  3=busy
Private gameSpeed As Long
Private sequence As String
Private cntInput As Integer
Private tone As Long


Private Sub Form_Load()
Dim rc As Long
Dim curDevice As Long
Dim x As Integer
    'set up buttons
    Me.Show
    DoEvents
    For x = 0 To 3
        setShape x
        lightOFF x
    Next x
    DoEvents
    
    'open midi device
    midiOutClose hmidi
    rc = midiOutOpen(hmidi, curDevice, 0, 0, 0)
    If rc <> 0 Then
        MsgBox "Error " & rc & " - Could not open midi device.", , "Simon"
    End If
    
    Randomize
    tone = 35
    tmrAttract.Enabled = True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    'pop up settings
    If gameMode = 0 And Button = 2 Then
        frmSettings.Show vbModal
        tone = Val(frmSettings.txtTone.Text)
    End If
End Sub

Private Sub Form_Paint()
Dim x As Integer
'Here, I'm attempting to prevent my buttons from turning black...
    For x = 0 To 3
        lightOFF x
    Next x
End Sub

Private Sub picButton_MouseDown(index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim i As Long
    If gameMode = 2 Then
        If index = Val(Mid$(sequence, cntInput + 1, 1)) Then
            'Correct!
            lightON index
            playNote index
            cntInput = cntInput + 1
        Else
            'Wrong! Game over.
            gameMode = 3
            setInstrument 55
            playNote 0
            For i = 0 To 3 'flash the correct colour
                lightON Val(Mid$(sequence, cntInput + 1, 1))
                Sleep 200
                lightOFF Val(Mid$(sequence, cntInput + 1, 1))
                Sleep 150
                DoEvents
            Next i
            If Val(lblScore.Caption) > Val(lblHigh.Caption) Then 'high score?
                lblHigh.Caption = lblScore.Caption
            End If
            Sleep 500
            stopNote 0
            gameMode = 0
            tmrAttract.Enabled = True
        End If
    End If
End Sub

Private Sub picButton_MouseUp(index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    If gameMode = 0 Then
        StartNewGame
    End If
    
    If gameMode = 2 Then
        lightOFF index
        stopNote index
        If cntInput = Len(sequence) Then 'player completed this round...
            lblScore.Caption = cntInput
            lblScore.Refresh
            ShowSeq 'start next round
        End If
    End If
End Sub

Private Sub Form_Terminate()
    midiOutClose hmidi
End Sub

Private Sub playNote(ByVal mNote As Long)
Dim midimsg As Long
    midimsg = ((45 + mNote * 5) * &H100) + &H7F009F
    midiOutShortMsg hmidi, midimsg
End Sub

Private Sub stopNote(ByVal mNote As Long)
Dim midimsg As Long
    midimsg = ((45 + mNote * 5) * &H100) + &H8F
    midiOutShortMsg hmidi, midimsg
End Sub

Private Sub Form_Unload(Cancel As Integer)
    midiOutClose hmidi
    Unload frmSettings
End Sub

Private Sub setInstrument(instrm As Long)
Dim midimsg As Long
    midimsg = (instrm * 256) + &HCF
    midiOutShortMsg hmidi, midimsg
End Sub

Private Sub setShape(index As Integer)
Dim winRgn As Long  ' combined region
Dim testRgn As Long ' base shaped region from sample images
Dim pRect As RECT

    ' create the base shaped region
    testRgn = CreateShapedRegion2(imgRgn(index).Picture.Handle)
    
    ' create the new window region & return the winRgn pointer
    winRgn = CreateRectRgnIndirect(pRect)
    CombineRgn winRgn, testRgn, winRgn, 5
    
    DeleteObject testRgn    ' the original shaped region is no longer needed
    
    ' update the changes as needed
    SetWindowRgn picButton(index).hwnd, winRgn, True
    picButton(index).Refresh
End Sub

Private Sub lightON(index As Integer)
    Select Case index
        Case 0
            Call mGradient.PaintGradientC(picButton(index).hdc, 0, 0, picButton(index).Width, picButton(index).Height, RGB(255, 200, 200), RGB(255, 20, 20))
        Case 1
            Call mGradient.PaintGradientC(picButton(index).hdc, 0, 0, picButton(index).Width, picButton(index).Height, RGB(200, 200, 255), RGB(20, 20, 255))
        Case 2
            Call mGradient.PaintGradientC(picButton(index).hdc, 0, 0, picButton(index).Width, picButton(index).Height, RGB(255, 255, 200), RGB(255, 255, 20))
        Case 3
            Call mGradient.PaintGradientC(picButton(index).hdc, 0, 0, picButton(index).Width, picButton(index).Height, RGB(200, 255, 200), RGB(20, 255, 20))
    End Select
End Sub

Private Sub lightOFF(index As Integer)
    Select Case index
        Case 0
            Call mGradient.PaintGradientL(picButton(index).hdc, 0, 0, picButton(index).Width, picButton(index).Height, RGB(130, 0, 0), RGB(220, 80, 80), 125)
        Case 1
            Call mGradient.PaintGradientL(picButton(index).hdc, 0, 0, picButton(index).Width, picButton(index).Height, RGB(0, 0, 130), RGB(80, 80, 220), 240)
        Case 2
            Call mGradient.PaintGradientL(picButton(index).hdc, 0, 0, picButton(index).Width, picButton(index).Height, RGB(130, 130, 0), RGB(220, 220, 80), 310)
        Case 3
            Call mGradient.PaintGradientL(picButton(index).hdc, 0, 0, picButton(index).Width, picButton(index).Height, RGB(0, 130, 0), RGB(80, 220, 80), 50)
    End Select
End Sub

Private Sub StartNewGame()
Dim x As Integer
    gameMode = 3
    tmrAttract.Enabled = False
    For x = 0 To 3
        lightOFF x
    Next x
    lblScore.Caption = "0"
    DoEvents
    sequence = ""
    gameSpeed = 700
    ShowSeq
End Sub

Private Sub ShowSeq()
Dim x As Long
Dim tmp As Integer
    gameMode = 1
    sequence = sequence & Trim$(Str$(Int(Rnd * 4))) 'add a new key to the sequence
    Sleep 300 'dramatic pause
    setInstrument tone
    For x = 1 To Len(sequence)  'show the whole sequence
        tmp = Val(Mid$(sequence, x, 1))
        Sleep gameSpeed / 3
        lightON tmp
        playNote tmp
        Sleep gameSpeed
        lightOFF tmp
        stopNote tmp
        DoEvents
    Next x
    gameSpeed = gameSpeed - 40 'speed up for the next round
    If gameSpeed < 100 Then gameSpeed = 100
    cntInput = 0
    gameMode = 2
End Sub

'Attract mode
Private Sub tmrAttract_Timer()
Static prev As Integer
    lightOFF prev
    prev = (prev + 1) Mod 4 'Int(Rnd * 4)
    lightON prev
End Sub
