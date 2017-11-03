VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form Form15 
   BackColor       =   &H00400000&
   BorderStyle     =   0  'None
   Caption         =   "Combate"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form30"
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1111
      Left            =   11520
      Top             =   600
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00004080&
      Default         =   -1  'True
      Height          =   1935
      Left            =   8640
      Picture         =   "Form15.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7080
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   11520
      Top             =   120
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Height          =   135
      Left            =   120
      TabIndex        =   4
      Top             =   9000
      Width           =   75
   End
   Begin VB.Image Image12 
      Height          =   735
      Left            =   5520
      Top             =   5160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image Image9 
      Height          =   615
      Index           =   3
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   735
   End
   Begin VB.Image Image9 
      Height          =   615
      Index           =   2
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   735
   End
   Begin VB.Image Image9 
      Height          =   615
      Index           =   1
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   735
   End
   Begin VB.Image Image8 
      Height          =   615
      Index           =   3
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   735
   End
   Begin VB.Image Image8 
      Height          =   615
      Index           =   2
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   735
   End
   Begin VB.Image Image8 
      Height          =   615
      Index           =   1
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   735
   End
   Begin VB.Image Image5 
      Height          =   975
      Left            =   8640
      Top             =   5040
      Width           =   735
   End
   Begin VB.Image Image4 
      Height          =   975
      Left            =   2640
      Top             =   5040
      Width           =   735
   End
   Begin VB.Image Image9 
      Height          =   615
      Index           =   0
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   735
   End
   Begin VB.Image Image8 
      Height          =   615
      Index           =   0
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   10560
      TabIndex        =   14
      Top             =   6240
      Width           =   615
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   10560
      TabIndex        =   13
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   10560
      TabIndex        =   12
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   10560
      TabIndex        =   11
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image Image7 
      Height          =   615
      Index           =   3
      Left            =   10560
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   615
   End
   Begin VB.Image Image7 
      Height          =   615
      Index           =   2
      Left            =   10560
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   615
   End
   Begin VB.Image Image7 
      Height          =   615
      Index           =   1
      Left            =   10560
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   615
   End
   Begin VB.Image Image7 
      Height          =   615
      Index           =   0
      Left            =   10560
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   615
   End
   Begin VB.Image Image6 
      Height          =   615
      Index           =   0
      Left            =   840
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   840
      TabIndex        =   10
      Top             =   6240
      Width           =   615
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   9
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   8
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   7
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image Image6 
      Height          =   615
      Index           =   3
      Left            =   840
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   615
   End
   Begin VB.Image Image6 
      Height          =   615
      Index           =   2
      Left            =   840
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   615
   End
   Begin VB.Image Image6 
      Height          =   615
      Index           =   1
      Left            =   840
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   615
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1950
      Left            =   0
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   1950
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  'Fixed Single
      Height          =   990
      Index           =   6
      Left            =   7680
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   975
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  'Fixed Single
      Height          =   990
      Index           =   5
      Left            =   6720
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   975
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  'Fixed Single
      Height          =   990
      Index           =   4
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   975
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  'Fixed Single
      Height          =   990
      Index           =   3
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   975
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  'Fixed Single
      Height          =   990
      Index           =   2
      Left            =   3840
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   975
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  'Fixed Single
      Height          =   990
      Index           =   1
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   975
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  'Fixed Single
      Height          =   990
      Index           =   0
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   975
   End
   Begin VB.Image Image11 
      Height          =   615
      Left            =   6000
      Stretch         =   -1  'True
      Tag             =   "Urso"
      Top             =   3720
      Width           =   855
   End
   Begin VB.Image Image10 
      Height          =   615
      Left            =   5040
      Stretch         =   -1  'True
      Tag             =   "Lobo"
      Top             =   3720
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   3720
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12000
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   30
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   30
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   0   'False
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   0   'False
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   0   'False
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   0   'False
      EnablePositionControls=   0   'False
      EnableFullScreenControls=   0   'False
      EnableTracker   =   0   'False
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   0
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -190
      WindowlessVideo =   0   'False
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mana"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10800
      TabIndex        =   6
      Top             =   8400
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Vida"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   2
      Top             =   8400
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   10800
      TabIndex        =   1
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   9600
      TabIndex        =   0
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   14.25
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   1920
      TabIndex        =   5
      Top             =   7080
      Width           =   6735
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim play, cont, cant, rand, dom, maxv1, maxv2, maxm1, maxm2 As Integer
Dim turno, car, p, dou, cri As Integer
Dim v1(0 To 6), v2(0 To 6), p1(0 To 6), p2(0 To 6) As String
Dim nome1, nome2, tmp, mags, img, cem1, cem2, cen1, cen2 As String
Dim santa1 As Boolean
Dim santa2 As Boolean
Dim h1 As Boolean
Dim h2 As Boolean
Dim ok As Boolean
Dim c1(0 To 3), c2(0 To 3), a1(0 To 3), a2(0 To 3), ix(0 To 3) As String

Sub toreg()
    Get #play, 1, reg
    Image2.Picture = LoadPicture(reg.foto)
    If play = 1 Then
        Label1.Caption = maxv1
        Label2.Caption = maxm1
    Else
        Label1.Caption = maxv2
        Label2.Caption = maxm2
    End If
End Sub

Sub criatura(cn, cc, cp As String)
Dim lol As Integer
    lol = 0
    If play = 1 Then
        While Image6(lol).Tag <> "" And lol < 3
            lol = lol + 1
        Wend
        If Image6(lol).Tag = "" Then
            Image6(lol).Tag = cn
            Image6(lol).Picture = LoadPicture(cp)
            Label7(lol).Caption = cc
            c1(lol) = ""
            a1(lol) = cc
        Else
            Image6(0).Tag = cn
            Image6(0).Picture = LoadPicture(cp)
            Label7(0).Caption = cc
            c1(0) = ""
            a1(0) = cc
        End If
    Else
        While Image7(lol).Tag <> "" And lol < 3
            lol = lol + 1
        Wend
        If Image7(lol).Tag = "" Then
            Image7(lol).Tag = cn
            Image7(lol).Picture = LoadPicture(cp)
            Label8(lol).Caption = cc
            c2(lol) = ""
            a2(lol) = cc
        Else
            Image7(0).Tag = cn
            Image7(0).Picture = LoadPicture(cp)
            Label8(0).Caption = cc
            c2(0) = ""
            a2(0) = cc
        End If
    End If
End Sub

Sub detecta(d As Integer)
Dim q As Integer
Dim st As String
    q = 0
    dou = 0
    p = 0
    For i = 0 To 6
        Image3(i).BorderStyle = 1
    Next
    mags = ""
    ok = False
    Select Case Image3(d).Tag
        Case "Vitalidade do Ar (0/+1) <-> Mana - 8"
            If Val(Label2.Caption) >= 8 Then
                Image3(d).BorderStyle = 0
                mags = "Vitalidade do Ar (0/+1) <-> Mana - 8"
            End If
        Case "Vapores Venenosos (Envenenar) <-> Mana - 10"
            If Val(Label2.Caption) >= 10 Then
                Image3(d).BorderStyle = 0
                mags = "Vapores Venenosos (Envenenar) <-> Mana - 10"
            End If
        Case "Chamar Ventos (X-Ataque-X) <-> Mana - 10"
            If Val(Label2.Caption) >= 10 Then
                If play = 1 Then
                    maxm1 = maxm1 - 10
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 10
                    Label2.Caption = maxm2
                    Image5.Left = 8000
                End If
                img = path & "Combates\Chamar Ventos\" & play & "\" & reg.sexo & "\"
                p = 7
                Timer1.Enabled = True
                For i = 0 To 3
                    Image6(i).Left = 840
                    Label7(i).Left = 840
                    Image7(i).Left = 10560
                    Label8(i).Left = 10560
                Next
                organiza d
            End If
        Case "Corpo de Ar (0/+2) <-> Mana - 15"
            If Val(Label2.Caption) >= 15 Then
                Image3(d).BorderStyle = 0
                mags = "Corpo de Ar (0/+2) <-> Mana - 15"
            End If
        Case "Elementar do Ar (7/7) <-> Mana - 20"
            If Val(Label2.Caption) >= 20 Then
                If play = 1 Then
                    maxm1 = maxm1 - 20
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 20
                    Label2.Caption = maxm2
                    Image5.Left = 8000
                End If
                ok = True
                img = path & "Combates\Summon\" & play & "\" & reg.sexo & "\"
                p = 6
                Timer1.Enabled = True
                organiza d
                criatura "Elementar de Ar", "7/7", path & "Criaturas\Elementar de Ar" & play & ".bmp"
            End If
        Case "Força da Terra (+1/0) <-> Mana - 8"
            If Val(Label2.Caption) >= 8 Then
                Image3(d).BorderStyle = 0
                mags = "Força da Terra (+1/0) <-> Mana - 8"
            End If
        Case "Lançar Pedra (-1 Vida) <-> Mana - 5"
            If Val(Label2.Caption) >= 5 Then
                Image3(d).BorderStyle = 0
                mags = "Lançar Pedra (-1 Vida) <-> Mana - 5"
            End If
        Case "Barreira de Pedra (0/8) <-> Mana - 12"
            If Val(Label2.Caption) >= 12 Then
                If play = 1 Then
                    maxm1 = maxm1 - 12
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 12
                    Label2.Caption = maxm2
                    Image5.Left = 8000
                End If
                ok = True
                img = path & "Combates\Summon\" & play & "\" & reg.sexo & "\"
                p = 6
                Timer1.Enabled = True
                organiza d
                criatura "Barreira de Pedra", "0/8", path & "Criaturas\Barreira de Pedra" & play & ".bmp"
            End If
        Case "Corpo de Pedra (+2/0) <-> Mana - 15"
            If Val(Label2.Caption) >= 15 Then
                Image3(d).BorderStyle = 0
                mags = "Corpo de Pedra (+2/0) <-> Mana - 15"
            End If
        Case "Elementar de Pedra (7/7) <-> Mana - 20"
            If Val(Label2.Caption) >= 20 Then
                If play = 1 Then
                    maxm1 = maxm1 - 20
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 20
                    Label2.Caption = maxm2
                    Image5.Left = 8000
                End If
                ok = True
                img = path & "Combates\Summon\" & play & "\" & reg.sexo & "\"
                p = 6
                Timer1.Enabled = True
                organiza d
                criatura "Elementar de Pedra", "7/7", path & "Criaturas\Elementar de Pedra" & play & ".bmp"
            End If
        Case "Agilidade de Fogo (1 Turno Inbloqueável) <-> Mana - 13"
            If Val(Label2.Caption) >= 13 Then
                Image3(d).BorderStyle = 0
                mags = "Agilidade de Fogo (1 Turno Inbloqueável) <-> Mana - 13"
            End If
        Case "Barreira de Fogo (0/3) <-> Mana - 5"
            If Val(Label2.Caption) >= 5 Then
                If play = 1 Then
                    maxm1 = maxm1 - 5
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 5
                    Label2.Caption = maxm2
                    Image5.Left = 8000
                End If
                ok = True
                img = path & "Combates\Summon\" & play & "\" & reg.sexo & "\"
                p = 6
                Timer1.Enabled = True
                organiza d
                criatura "Barreira de Fogo", "0/3", path & "Criaturas\Barreira de Fogo" & play & ".bmp"
            End If
        Case "Bola de Fogo (-2 Vida) <-> Mana - 10"
            If Val(Label2.Caption) >= 10 Then
                Image3(d).BorderStyle = 0
                mags = "Bola de Fogo (-2 Vida) <-> Mana - 10"
            End If
        Case "Corpo de Fogo (Inbloqueável) <-> Mana - 25"
            If Val(Label2.Caption) >= 25 Then
                Image3(d).BorderStyle = 0
                mags = "Corpo de Fogo (Inbloqueável) <-> Mana - 25"
            End If
        Case "Elementar de Fogo (7/7) <-> Mana - 20"
            If Val(Label2.Caption) >= 20 Then
                If play = 1 Then
                    maxm1 = maxm1 - 20
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 20
                    Label2.Caption = maxm2
                    Image5.Left = 8000
                End If
                ok = True
                img = path & "Combates\Summon\" & play & "\" & reg.sexo & "\"
                p = 6
                Timer1.Enabled = True
                organiza d
                criatura "Elementar de Fogo", "7/7", path & "Criaturas\Elementar de Fogo" & play & ".bmp"
            End If
        Case "Pureza de Água (1 Turno Abilidade de Voar) <-> Mana - 11"
            If Val(Label2.Caption) >= 11 Then
                Image3(d).BorderStyle = 0
                mags = "Pureza de Água (1 Turno Abilidade de Voar) <-> Mana - 11"
            End If
        Case "Chamar Névoa (X-Ataque-X) <-> Mana - 10"
            If Val(Label2.Caption) >= 10 Then
                If play = 1 Then
                    maxm1 = maxm1 - 10
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 10
                    Label2.Caption = maxm2
                    Image5.Left = 8000
                End If
                img = path & "Combates\Chamar Nevoa\" & play & "\" & reg.sexo & "\"
                p = 7
                Timer1.Enabled = True
                For i = 0 To 3
                    Image6(i).Left = 840
                    Label7(i).Left = 840
                    Image7(Index).Left = 10560
                    Label8(Index).Left = 10560
                Next
                organiza d
            End If
        Case "Tempestade de Gelo (-2 Vida) <-> Mana - 10"
            If Val(Label2.Caption) >= 10 Then
                Image3(d).BorderStyle = 0
                mags = "Tempestade de Gelo (-2 Vida) <-> Mana - 10"
            End If
        Case "Corpo de Água (Abilidade de Voar) <-> Mana - 20"
            If Val(Label2.Caption) >= 20 Then
                Image3(d).BorderStyle = 0
                mags = "Corpo de Água (Abilidade de Voar) <-> Mana - 20"
            End If
        Case "Elementar de Água (7/7) <-> Mana - 20"
            If Val(Label2.Caption) >= 20 Then
                If play = 1 Then
                    maxm1 = maxm1 - 20
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 20
                    Label2.Caption = maxm2
                    Image5.Left = 8000
                End If
                ok = True
                img = path & "Combates\Summon\" & play & "\" & reg.sexo & "\"
                p = 6
                Timer1.Enabled = True
                organiza d
                criatura "Elementar de Água", "7/7", path & "Criaturas\Elementar de Agua" & play & ".bmp"
            End If
        Case "Escudo Protector (X-Magia Inimiga de Ataque-X) <-> Mana - 10"
            If Val(Label2.Caption) >= 10 Then
                Image3(d).BorderStyle = 0
                mags = "Escudo Protector (X-Magia Inimiga de Ataque-X) <-> Mana - 10"
            End If
        Case "Raio Eléctrico (-1 Vida) <-> Mana - 5"
            If Val(Label2.Caption) >= 5 Then
                Image3(d).BorderStyle = 0
                mags = "Raio Eléctrico (-1 Vida) <-> Mana - 5"
            End If
        Case "Barreira de Força (0/5) <-> Mana - 8"
            If Val(Label2.Caption) >= 8 Then
                If play = 1 Then
                    maxm1 = maxm1 - 8
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 8
                    Label2.Caption = maxm2
                    Image5.Left = 8000
                End If
                ok = True
                img = path & "Combates\Summon\" & play & "\" & reg.sexo & "\"
                p = 6
                Timer1.Enabled = True
                organiza d
                criatura "Barreira de Força", "0/5", path & "Criaturas\Barreira de Forca" & play & ".bmp"
            End If
        Case "Trovão (-2 Vida) <-> Mana - 10"
            If Val(Label2.Caption) >= 10 Then
                Image3(d).BorderStyle = 0
                mags = "Trovão (-2 Vida) <-> Mana - 10"
            End If
        Case "Desintegrar (-1 Criatura) <-> Mana - 20"
            If Val(Label2.Caption) >= 20 Then
                Image3(d).BorderStyle = 0
                mags = "Desintegrar (-1 Criatura) <-> Mana - 20"
            End If
        Case "Encantar Criatura (1 Turno +1 Criatura) <-> Mana - 5 -> Lobo (3/3) ou 10 -> Urso (5/5)"
            Image3(d).BorderStyle = 0
            mags = "Encantar Criatura (1 Turno +1 Criatura) <-> Mana - 5 -> Lobo (3/3) ou 10 -> Urso (5/5)"
        Case "Floresta (X-Ac. +Ataque/+Vida-X) <-> Mana - 5"
            If Val(Label2.Caption) >= 5 Then
                If play = 1 Then
                    maxm1 = maxm1 - 5
                    Label2.Caption = maxm1
                    q = 0
                    While (Image8(q).Tag <> "Floresta") And (q < 3)
                        q = q + 1
                    Wend
                    If (Image8(q).Tag = "Floresta") Then
                        flor1 = flor1 + 1
                    Else
                        q = 0
                        While (Image8(q).Tag <> "") And (q < 3)
                            q = q + 1
                        Wend
                        If (Image8(q).Tag = "") Then
                            Image8(q).Tag = "Floresta"
                            Image8(q).Picture = LoadPicture(path & "Combates\Floresta.bmp")
                            flor1 = 0
                        End If
                    End If
                Else
                    maxm2 = maxm2 - 5
                    Label2.Caption = maxm2
                    Image5.Left = 8000
                    q = 0
                    While (Image9(q).Tag <> "Floresta") And (q < 3)
                        q = q + 1
                    Wend
                    If (Image9(q).Tag = "Floresta") Then
                        flor2 = flor2 + 1
                    Else
                        q = 0
                        While (Image9(q).Tag <> "") And (q < 3)
                            q = q + 1
                        Wend
                        If (Image9(q).Tag = "") Then
                            Image9(q).Tag = "Floresta"
                            Image9(q).Picture = LoadPicture(path & "Combates\Floresta.bmp")
                            flor2 = 0
                        End If
                    End If
                End If
                img = path & "Combates\Floresta\" & play & "\" & reg.sexo & "\"
                p = 5
                Timer1.Enabled = True
                organiza d
            End If
        Case "Controlar Criatura (+1 Criatura) <-> Mana - 10 -> Lobo (3/3) ou 15 -> Urso (5/5)"
            Image3(d).BorderStyle = 0
            mags = "Controlar Criatura (+1 Criatura) <-> Mana - 10 -> Lobo (3/3) ou 15 -> Urso (5/5)"
        Case "Invocar Criatura (Criatura 4/4) <-> Mana - 12"
            If Val(Label2.Caption) >= 12 Then
                If play = 1 Then
                    maxm1 = maxm1 - 12
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 12
                    Label2.Caption = maxm2
                    Image5.Left = 8000
                End If
                ok = True
                img = path & "Combates\Summon\" & play & "\" & reg.sexo & "\"
                p = 6
                Timer1.Enabled = True
                organiza d
                If d Mod 2 = 0 Then
                    criatura "Urso", "4/4", path & "Urso" & play & ".bmp"
                Else
                    criatura "Lobo", "4/4", path & "Lobo" & play & ".bmp"
                End If
            End If
        Case "Regeneração (+1 Vida - Jogador/Vida Toda - Criatura) <-> Mana - 10"
            If Val(Label2.Caption) >= 10 Then
                Image3(d).BorderStyle = 0
                mags = "Regeneração (+1 Vida - Jogador/Vida Toda - Criatura) <-> Mana - 10"
            End If
        Case "Resistir Magia (X-Magia Inimiga-X) <-> Mana - 10"
            If Val(Label2.Caption) >= 10 Then
                Image3(d).BorderStyle = 0
                mags = "Resistir Magia (X-Magia Inimiga-X) <-> Mana - 10"
            End If
        Case "Dispersar Magias (X-Magias-X) <-> Mana - 15"
            If Val(Label2.Caption) >= 15 Then
                If play = 1 Then
                    maxm1 = maxm1 - 15
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 15
                    Label2.Caption = maxm2
                End If
                img = path & "Combates\Dispersar Magias\" & play & "\" & reg.sexo & "\"
                p = 3
                Timer1.Enabled = True
                For i = 0 To 3
                    Image8(i).Tag = ""
                    Image9(i).Tag = ""
                    Image8(i).Picture = LoadPicture()
                    Image9(i).Picture = LoadPicture()
                Next
            End If
        Case "Enfraquecer (-1/0) <-> Mana - 10"
            If Val(Label2.Caption) >= 10 Then
                Image3(d).BorderStyle = 0
                mags = "Enfraquecer (-1/0) <-> Mana - 10"
            End If
        Case "Escudo Reflector (<=Reflectir Magias Inimigas<=) <-> Mana - 20"
            If Val(Label2.Caption) >= 20 Then
                img = path & "Combates\Escudo Reflector\" & play & "\" & reg.sexo & "\"
                p = 5
                Timer1.Enabled = True
                If play = 1 Then
                    maxm1 = maxm1 - 20
                    Label2.Caption = maxm1
                    For i = 0 To 3
                        Image8(i).Tag = Image9(i).Tag
                        Image8(i).Picture = LoadPicture()
                        Image9(i).Tag = ""
                        Image9(i).Picture = LoadPicture()
                    Next
                Else
                    maxm2 = maxm2 - 20
                    Label2.Caption = maxm2
                    Image5.Left = 8000
                    For i = 0 To 3
                        Image9(i).Tag = Image8(i).Tag
                        Image9(i).Picture = LoadPicture()
                        Image8(i).Tag = ""
                        Image8(i).Picture = LoadPicture()
                    Next
                End If
            End If
        Case "Polymorph (Transforma Criatura 0/0) <-> Mana - 20"
            If Val(Label2.Caption) >= 20 Then
                Image3(d).BorderStyle = 0
                mags = "Polymorph (Transforma Criatura 0/0) <-> Mana - 20"
            End If
        Case "Prejudicar (-1 Vida) <-> Mana - 5"
            If Val(Label2.Caption) >= 5 Then
                Image3(d).BorderStyle = 0
                mags = "Prejudicar (-1 Vida) <-> Mana - 5"
            End If
        Case "Conjurar Espírito (+1 Criatura do Cemitério - Perda de -2/-2) <-> Mana - 10"
            If Val(Label2.Caption) >= 10 Then
                q = 0
                ok = True
                img = path & "Combates\Summon\" & play & "\" & reg.sexo & "\"
                p = 6
                Timer1.Enabled = True
                organiza d
                If play = 1 Then
                    maxm1 = maxm1 - 10
                    Label2.Caption = maxm1
                    cen1 = (Val(Mid(cen1, 1, 1)) - 2) & "/" & (Val(Mid(cen1, 3, 1)) - 2)
                    If Val(Mid(cen1, 1, 1)) <= 0 Then
                        cen1 = 1 & "/" & (Val(Mid(cen1, 3, 1)))
                    End If
                    If Val(Mid(cen1, 3, 1)) <= 0 Then
                        cen1 = (Val(Mid(cen1, 1, 1))) & "/" & 1
                    End If
                    While Image6(q).Tag <> "" And q < 3
                        q = q + 1
                    Wend
                    If Image6(q).Tag = "" Then
                        Image6(q).Tag = "Espírito"
                        Image6(q).Picture = LoadPicture(cem1)
                        Label7(q).Caption = cen1
                        c1(q) = ""
                        a1(q) = cen1
                    Else
                        Image6(0).Tag = "Espírito"
                        Image6(0).Picture = LoadPicture(cem1)
                        Label7(0).Caption = cen1
                        c1(0) = ""
                        a1(0) = cen1
                    End If
                Else
                    maxm2 = maxm2 - 10
                    Label2.Caption = maxm2
                    Image5.Left = 8000
                    cen2 = (Val(Mid(cen2, 1, 1)) - 2) & "/" & (Val(Mid(cen2, 3, 1)) - 2)
                    If Val(Mid(cen2, 1, 1)) <= 0 Then
                        cen2 = 1 & "/" & (Val(Mid(cen2, 3, 1)))
                    End If
                    If Val(Mid(cen2, 3, 1)) <= 0 Then
                        cen2 = (Val(Mid(cen2, 1, 1))) & "/" & 1
                    End If
                    While Image7(q).Tag <> "" And q < 3
                        q = q + 1
                    Wend
                    If Image7(q).Tag = "" Then
                        Image7(q).Tag = "Espírito"
                        Image7(q).Picture = LoadPicture(cem2)
                        Label8(q).Caption = cen2
                        c2(q) = ""
                        a2(q) = cen2
                    Else
                        Image7(0).Tag = "Espírito"
                        Image7(0).Picture = LoadPicture(cem2)
                        Label8(0).Caption = cen2
                        c2(0) = ""
                        a2(0) = cen2
                    End If
                End If
            End If
        Case "Invocar Zombie (+1 Criatura do Cemitério) <-> Mana - Custo do Original de Invocação -1"
            If Val(Label2.Caption) >= 10 Then
                q = 0
                ok = True
                img = path & "Combates\Summon\" & play & "\" & reg.sexo & "\"
                p = 6
                Timer1.Enabled = True
                organiza d
                If play = 1 Then
                    maxm1 = maxm1 - 10
                    Label2.Caption = maxm1
                    cen1 = (Val(Mid(cen1, 1, 1))) & "/" & (Val(Mid(cen1, 3, 1)))
                    If Val(Mid(cen1, 1, 1)) <= 0 Then
                        cen1 = 1 & "/" & (Val(Mid(cen1, 3, 1)))
                    End If
                    If Val(Mid(cen1, 3, 1)) <= 0 Then
                        cen1 = (Val(Mid(cen1, 1, 1))) & "/" & 1
                    End If
                    While Image6(q).Tag <> "" And q < 3
                        q = q + 1
                    Wend
                    If Image6(q).Tag = "" Then
                        Image6(q).Tag = "Espírito"
                        Image6(q).Picture = LoadPicture(cem1)
                        Label7(q).Caption = cen1
                        c1(q) = ""
                        a1(q) = cen1
                    Else
                        Image6(0).Tag = "Espírito"
                        Image6(0).Picture = LoadPicture(cem1)
                        Label7(0).Caption = cen1
                        c1(0) = ""
                        a1(0) = cen1
                    End If
                Else
                    maxm2 = maxm2 - 10
                    Label2.Caption = maxm2
                    Image5.Left = 8000
                    cen2 = (Val(Mid(cen2, 1, 1))) & "/" & (Val(Mid(cen2, 3, 1)))
                    If Val(Mid(cen2, 1, 1)) <= 0 Then
                        cen2 = 1 & "/" & (Val(Mid(cen2, 3, 1)))
                    End If
                    If Val(Mid(cen2, 3, 1)) <= 0 Then
                        cen2 = (Val(Mid(cen2, 1, 1))) & "/" & 1
                    End If
                    While Image7(q).Tag <> "" And q < 3
                        q = q + 1
                    Wend
                    If Image7(q).Tag = "" Then
                        Image7(q).Tag = "Espírito"
                        Image7(q).Picture = LoadPicture(cem2)
                        Label8(q).Caption = cen2
                        c2(q) = ""
                        a2(q) = cen2
                    Else
                        Image7(0).Tag = "Espírito"
                        Image7(0).Picture = LoadPicture(cem2)
                        Label8(0).Caption = cen2
                        c2(0) = ""
                        a2(0) = cen2
                    End If
                End If
            End If
        Case "Criar Zombie (5/3) <-> Mana - 13"
            If Val(Label2.Caption) >= 13 Then
                If play = 1 Then
                    maxm1 = maxm1 - 13
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 13
                    Label2.Caption = maxm2
                    Image5.Left = 8000
                End If
                ok = True
                img = path & "Combates\Summon\" & play & "\" & reg.sexo & "\"
                p = 6
                Timer1.Enabled = True
                organiza d
                criatura "Zombie", "5/3", path & "Criaturas\Zombie" & play & ".bmp"
            End If
        Case "Sugar Vida (+1 Vida Jogador, -1 Vida Oponente) <-> Mana - 12"
            If Val(Label2.Caption) >= 12 Then
                organiza d
                If play = 1 Then
                    maxm1 = maxm1 - 12
                    Label2.Caption = maxm1
                    q = 0
                    While (Image8(q).Tag <> "") And (q < 3)
                        q = q + 1
                    Wend
                    If (Image8(q).Tag = "") Then
                        Image8(q).Tag = "Sugar Vida"
                        Image8(q).Picture = LoadPicture(path & "Combates\Sugar Vida.bmp")
                    End If
                Else
                    maxm2 = maxm2 - 12
                    Label2.Caption = maxm2
                    Image5.Left = 8000
                    q = 0
                    While (Image9(q).Tag <> "") And (q < 3)
                        q = q + 1
                    Wend
                    If (Image9(q).Tag = "") Then
                        Image9(q).Tag = "Sugar Vida"
                        Image9(q).Picture = LoadPicture(path & "Combates\Sugar Vida.bmp")
                    End If
                End If
            End If
        Case "Curativo (+1 Vida) <-> Mana - 5"
            If Val(Label2.Caption) >= 5 Then
                Image3(d).BorderStyle = 0
                mags = "Curativo (+1 Vida) <-> Mana - 5"
            End If
        Case "Desentoxicar (X-Veneno-X) <-> Mana - 10"
            If Val(Label2.Caption) >= 10 Then
                Image3(d).BorderStyle = 0
                mags = "Desentoxicar (X-Veneno-X) <-> Mana - 10"
            End If
        Case "Curar (+2 Vida) <-> Mana - 10"
            If Val(Label2.Caption) >= 10 Then
                Image3(d).BorderStyle = 0
                mags = "Curar (+2 Vida) <-> Mana - 10"
            End If
        Case "Santuário (X-Zombies-X) <-> Mana - 20"
            If Val(Label2.Caption) >= 20 Then
                img = path & "Combates\Resistir Magia ou Santuario\" & play & "\" & reg.sexo & "\"
                p = 5
                Timer1.Enabled = True
                organiza d
                If play = 1 Then
                    maxm1 = maxm1 - 20
                    Label2.Caption = maxm1
                    q = 0
                    While (Image8(q).Tag <> "") And (q < 3)
                        q = q + 1
                    Wend
                    If (Image8(q).Tag = "") Then
                        Image8(q).Tag = "Santuário"
                        Image8(q).Picture = LoadPicture(path & "Combates\Santuario.bmp")
                        santa1 = True
                    End If
                Else
                    maxm2 = maxm2 - 20
                    Label2.Caption = maxm2
                    Image5.Left = 8000
                    q = 0
                    While (Image9(q).Tag <> "") And (q < 3)
                        q = q + 1
                    Wend
                    If (Image9(q).Tag = "") Then
                        Image9(q).Tag = "Santuário"
                        Image9(q).Picture = LoadPicture(path & "Combates\Santuario.bmp")
                        santa2 = True
                    End If
                End If
            End If
        Case "Ressuscitar (+1 Criatura do Cemitério) <-> Mana - Custo Igual ao Original de Invocação"
            If Val(Label2.Caption) >= 10 Then
                q = 0
                ok = True
                img = path & "Combates\Summon\" & play & "\" & reg.sexo & "\"
                p = 6
                Timer1.Enabled = True
                organiza d
                If play = 1 Then
                    maxm1 = maxm1 - 10
                    Label2.Caption = maxm1
                    cen1 = (Val(Mid(cen1, 1, 1))) & "/" & (Val(Mid(cen1, 3, 1)))
                    If Val(Mid(cen1, 1, 1)) <= 0 Then
                        cen1 = 1 & "/" & (Val(Mid(cen1, 3, 1)))
                    End If
                    If Val(Mid(cen1, 3, 1)) <= 0 Then
                        cen1 = (Val(Mid(cen1, 1, 1))) & "/" & 1
                    End If
                    While Image6(q).Tag <> "" And q < 3
                        q = q + 1
                    Wend
                    If Image6(q).Tag = "" Then
                        Image6(q).Tag = "Espírito"
                        Image6(q).Picture = LoadPicture(cem1)
                        Label7(q).Caption = cen1
                        c1(q) = ""
                        a1(q) = cen1
                    Else
                        Image6(0).Tag = "Espírito"
                        Image6(0).Picture = LoadPicture(cem1)
                        Label7(0).Caption = cen1
                        c1(0) = ""
                        a1(0) = cen1
                    End If
                Else
                    maxm2 = maxm2 - 10
                    Label2.Caption = maxm2
                    Image5.Left = 8000
                    cen2 = (Val(Mid(cen2, 1, 1))) & "/" & (Val(Mid(cen2, 3, 1)))
                    If Val(Mid(cen2, 1, 1)) <= 0 Then
                        cen2 = 1 & "/" & (Val(Mid(cen2, 3, 1)))
                    End If
                    If Val(Mid(cen2, 3, 1)) <= 0 Then
                        cen2 = (Val(Mid(cen2, 1, 1))) & "/" & 1
                    End If
                    While Image7(q).Tag <> "" And q < 3
                        q = q + 1
                    Wend
                    If Image7(q).Tag = "" Then
                        Image7(q).Tag = "Espírito"
                        Image7(q).Picture = LoadPicture(cem2)
                        Label8(q).Caption = cen2
                        c2(q) = ""
                        a2(q) = cen2
                    Else
                        Image7(0).Tag = "Espírito"
                        Image7(0).Picture = LoadPicture(cem2)
                        Label8(0).Caption = cen2
                        c2(0) = ""
                        a2(0) = cen2
                    End If
                End If
            End If
        Case "Troll (1/1) <-> Mana - 5"
            If Val(Label2.Caption) >= 5 Then
                If play = 1 Then
                    maxm1 = maxm1 - 5
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 5
                    Label2.Caption = maxm2
                    Image5.Left = 8000
                End If
                ok = True
                img = path & "Combates\Summon\" & play & "\" & reg.sexo & "\"
                p = 6
                Timer1.Enabled = True
                organiza d
                criatura "Troll", "1/1", path & "Criaturas\Troll" & play & ".bmp"
            End If
        Case "Orco (2/3) <-> Mana - 10"
            If Val(Label2.Caption) >= 10 Then
                If play = 1 Then
                    maxm1 = maxm1 - 10
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 10
                    Label2.Caption = maxm2
                    Image5.Left = 8000
                End If
                ok = True
                img = path & "Combates\Summon\" & play & "\" & reg.sexo & "\"
                p = 6
                Timer1.Enabled = True
                organiza d
                criatura "Orco", "2/3", path & "Criaturas\Orco" & play & ".bmp"
            End If
        Case "Ogre (4/5) <-> Mana - 15"
            If Val(Label2.Caption) >= 15 Then
                If play = 1 Then
                    maxm1 = maxm1 - 15
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 15
                    Label2.Caption = maxm2
                    Image5.Left = 8000
                End If
                ok = True
                img = path & "Combates\Summon\" & play & "\" & reg.sexo & "\"
                p = 6
                Timer1.Enabled = True
                organiza d
                criatura "Ogre", "4/5", path & "Criaturas\Ogre" & play & ".bmp"
            End If
        Case "Almas Perdidas (Condenado 6/7) <-> Mana - 20"
            If Val(Label2.Caption) >= 20 Then
                If play = 1 Then
                    maxm1 = maxm1 - 20
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 20
                    Label2.Caption = maxm2
                    Image5.Left = 8000
                End If
                ok = True
                img = path & "Combates\Summon\" & play & "\" & reg.sexo & "\"
                p = 6
                Timer1.Enabled = True
                organiza d
                criatura "Condenado", "6/7", path & "Criaturas\Condenado" & play & ".bmp"
            End If
        Case "Portas do Inferno (Demónio 8/8) <-> Mana - 25"
            If Val(Label2.Caption) >= 25 Then
                If play = 1 Then
                    maxm1 = maxm1 - 25
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 25
                    Label2.Caption = maxm2
                    Image5.Left = 8000
                End If
                ok = True
                img = path & "Combates\Summon\" & play & "\" & reg.sexo & "\"
                p = 6
                Timer1.Enabled = True
                organiza d
                criatura "Demónio", "8/8", path & "Criaturas\Demonio" & play & ".bmp"
            End If
        Case "Poção de Vida"
            If Val(Label1.Caption) <= reg.vida - 2 Then
                If play = 1 Then
                    maxv1 = maxv1 + 2
                    Label1.Caption = maxv1
                Else
                    maxv2 = maxv2 + 2
                    Label1.Caption = maxv2
                End If
                reg.pvida = reg.pvida - 1
                Put #play, 1, reg
                organiza d
            End If
        Case "Poção de Mana"
            If Val(Label2.Caption) <= reg.mana - 10 Then
                If play = 1 Then
                    maxm1 = maxm1 + 10
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 + 10
                    Label2.Caption = maxm2
                End If
                reg.pmana = reg.pmana - 1
                Put #play, 1, reg
                organiza d
            End If
    End Select
End Sub

Sub organiza(s As Integer)
    If s < 6 Then
        For i = s To 5
            If play = 1 Then
                v1(i) = v1(i + 1)
                p1(i) = p1(i + 1)
                Image3(i).Picture = LoadPicture(p1(i))
            Else
                v2(i) = v2(i + 1)
                p2(i) = p2(i + 1)
                Image3(i).Picture = LoadPicture(p2(i))
            End If
            Image3(i).Tag = Image3(i + 1).Tag
        Next
    End If
    Image3(6).Tag = ""
    Image3(6).Picture = LoadPicture()
    If play = 1 Then
        v1(6) = ""
        p1(6) = ""
    Else
        v2(6) = ""
        p2(6) = ""
    End If
End Sub

Sub gerorg()
    If modo = 0 Then
        car = 0
        Do Until (Image3(car).Tag = "") Or (car = 6)
            car = car + 1
        Loop
    End If
End Sub

Private Sub Command1_Click()
    For i = 0 To 6
        Image3(i).BorderStyle = 1
        Image3(i).Enabled = True
    Next
    mags = ""
    If turno = 1 Then
        If play = 1 Then
            play = 2
            For i = 0 To 6
                Image3(i).Tag = v2(i)
                Image3(i).Picture = LoadPicture(p2(i))
            Next
        Else
            play = 1
            For i = 0 To 6
                Image3(i).Tag = v1(i)
                Image3(i).Picture = LoadPicture(p1(i))
            Next
        End If
        For i = 0 To 6
            Image3(i).Enabled = False
        Next
        For i = 0 To 6
            Select Case Image3(i).Tag
                Case "Chamar Ventos (X-Ataque-X) <-> Mana - 10"
                    Image3(i).Enabled = True
                Case "Chamar Névoa (X-Ataque-X) <-> Mana - 10"
                    Image3(i).Enabled = True
                Case "Escudo Protector (X-Magia Inimiga de Ataque-X) <-> Mana - 10"
                    Image3(i).Enabled = True
                Case "Resistir Magia (X-Magia Inimiga-X) <-> Mana - 10"
                    Image3(i).Enabled = True
                Case "Dispersar Magias (X-Magias-X) <-> Mana - 15"
                    Image3(i).Enabled = True
                Case "Escudo Reflector (<=Reflectir Magias Inimigas<=) <-> Mana - 20"
                    Image3(i).Enabled = True
                Case "Santuário (X-Zombies-X) <-> Mana - 20"
                    Image3(i).Enabled = True
                Case "Poção de Vida"
                    Image3(i).Enabled = True
                Case "Poção de Mana"
                    Image3(i).Enabled = True
            End Select
        Next
        toreg
        turno = 2
    Else
        turno = 1
        gerorg
        abload
        Get #play, 1, reg
        Label2.Caption = reg.mana
        If play = 1 Then
            maxm1 = reg.mana
            For i = 0 To 3
                If Image7(i).Left = 9480 Then
                    maxv1 = maxv1 - Val(Mid(Label8(i).Caption, 1, 1))
                    Image7(i).Left = 10560
                End If
            Next
            For i = 0 To 3
                Select Case Image9(i).Tag
                    Index = Val(Mid(ix(q), 7, 1))
                    Case "Prejudicar"
                        ok = True
                        img = path & "Combates\Prejudicar\" & play & "\" & reg.sexo & "\"
                        p = 6
                        Timer1.Enabled = True
                        If Mid(ix(i), 1, 6) = "image6" Then
                            st = Val(Mid(Label7(Index).Caption, 3, 1)) - 1
                            Label7(Index).Caption = Mid(Label7(Index).Caption, 1, 2) & st
                            If st <= 0 Then
                                Image6(Index).Picture = LoadPicture()
                                Image6(Index).Tag = ""
                                Label7(Index).Caption = ""
                                a1(Index) = ""
                                c1(Index) = ""
                            End If
                        End If
                        If Mid(ix(i), 1, 6) = "image7" Then
                            st = Val(Mid(Label8(Index).Caption, 3, 1)) - 1
                            Label8(Index).Caption = Mid(Label8(Index).Caption, 1, 2) & st
                            If st <= 0 Then
                                Image7(Index).Picture = LoadPicture()
                                Image7(Index).Tag = ""
                                Label8(Index).Caption = ""
                                a1(Index) = ""
                                c1(Index) = ""
                            End If
                        End If
                        If Mid(ix(i), 1, 6) = "image4" Then
                            maxv1 = maxv1 - 1
                            Label1.Caption = maxv1
                        End If
                        If Mid(ix(i), 1, 6) = "image5" Then
                            maxv2 = maxv2 - 1
                            Label1.Caption = maxv2
                        End If
                    Case "Desintegrar"
                        img = path & "Combates\Desintegrar ou Polymorph\" & play & "\" & reg.sexo & "\"
                        p = 7
                        Timer1.Enabled = True
                        If Mid(ix(i), 1, 6) = "image6" Then
                            Image6(Index).Picture = LoadPicture()
                            Image6(Index).Tag = ""
                            Label7(Index).Caption = ""
                            c1(Index) = ""
                            a1(Index) = ""
                        End If
                        If Mid(ix(i), 1, 6) = "image7" Then
                            Image7(Index).Picture = LoadPicture()
                            Image7(Index).Tag = ""
                            Label8(Index).Caption = ""
                            c2(Index) = ""
                            a2(Index) = ""
                        End If
                    Case "Trovão"
                        img = path & "Combates\Trovao\" & play & "\" & reg.sexo & "\"
                        p = 7
                        Timer1.Enabled = True
                        If Mid(ix(i), 1, 6) = "image6" Then
                            st = Val(Mid(Label7(Index).Caption, 3, 1)) - 2
                            Label7(Index).Caption = Mid(Label7(Index).Caption, 1, 2) & st
                            If st <= 0 Then
                                Image6(Index).Picture = LoadPicture()
                                Image6(Index).Tag = ""
                                Label7(Index).Caption = ""
                                c1(Index) = ""
                                a1(Index) = ""
                            End If
                        End If
                        If Mid(ix(i), 1, 6) = "image7" Then
                            st = Val(Mid(Label8(Index).Caption, 3, 1)) - 2
                            Label8(Index).Caption = Mid(Label8(Index).Caption, 1, 2) & st
                            If st <= 0 Then
                                Image7(Index).Picture = LoadPicture()
                                Image7(Index).Tag = ""
                                Label8(Index).Caption = ""
                                c2(Index) = ""
                                a2(Index) = ""
                            End If
                        End If
                        If Mid(ix(i), 1, 6) = "image4" Then
                            maxv1 = maxv1 - 2
                            Label1.Caption = maxv1
                        End If
                        If Mid(ix(i), 1, 6) = "image5" Then
                            maxv2 = maxv2 - 2
                            Label1.Caption = maxv2
                        End If
                    Case "Raio Electrico"
                        img = path & "Combates\Raio Eletrico\" & play & "\" & reg.sexo & "\"
                        p = 5
                        Timer1.Enabled = True
                        If Mid(ix(i), 1, 6) = "image6" Then
                            st = Val(Mid(Label7(Index).Caption, 3, 1)) - 1
                            Label7(Index).Caption = Mid(Label7(Index).Caption, 1, 2) & st
                            If st <= 0 Then
                                Image6(Index).Picture = LoadPicture()
                                Image6(Index).Tag = ""
                                Label7(Index).Caption = ""
                                c1(Index) = ""
                                a1(Index) = ""
                            End If
                        End If
                        If Mid(ix(i), 1, 6) = "image7" Then
                            st = Val(Mid(Label8(Index).Caption, 3, 1)) - 1
                            Label8(Index).Caption = Mid(Label8(Index).Caption, 1, 2) & st
                            If st <= 0 Then
                                Image7(Index).Picture = LoadPicture()
                                Image7(Index).Tag = ""
                                Label8(Index).Caption = ""
                                c2(Index) = ""
                                a2(Index) = ""
                            End If
                        End If
                        If Mid(ix(i), 1, 6) = "image4" Then
                            maxv1 = maxv1 - 1
                            Label1.Caption = maxv1
                        End If
                        If Mid(ix(i), 1, 6) = "image5" Then
                            maxv2 = maxv2 - 1
                            Label1.Caption = maxv2
                        End If
                    Case "Tempestade de Gelo"
                        ok = True
                        img = path & "Combates\Tempestade de Gelo\" & play & "\" & reg.sexo & "\"
                        p = 9
                        Timer1.Enabled = True
                        If Mid(ix(i), 1, 6) = "image6" Then
                            st = Val(Mid(Label7(Index).Caption, 3, 1)) - 2
                            Label7(Index).Caption = Mid(Label7(Index).Caption, 1, 2) & st
                            If st <= 0 Then
                                Image6(Index).Picture = LoadPicture()
                                Image6(Index).Tag = ""
                                Label7(Index).Caption = ""
                                c1(Index) = ""
                                a1(Index) = ""
                            End If
                        End If
                        If Mid(ix(i), 1, 6) = "image7" Then
                            st = Val(Mid(Label8(Index).Caption, 3, 1)) - 2
                            Label8(Index).Caption = Mid(Label8(Index).Caption, 1, 2) & st
                            If st <= 0 Then
                                Image7(Index).Picture = LoadPicture()
                                Image7(Index).Tag = ""
                                Label8(Index).Caption = ""
                                c2(Index) = ""
                                a2(Index) = ""
                            End If
                        End If
                        If Mid(ix(i), 1, 6) = "image4" Then
                            maxv1 = maxv1 - 2
                            Label1.Caption = maxv1
                        End If
                        If Mid(ix(i), 1, 6) = "image5" Then
                            maxv2 = maxv2 - 2
                            Label1.Caption = maxv2
                        End If
                    Case "Bola de Fogo"
                        ok = True
                        img = path & "Combates\Bola de Fogo\" & play & "\" & reg.sexo & "\"
                        p = 8
                        Timer1.Enabled = True
                        If Mid(ix(i), 1, 6) = "image6" Then
                            st = Val(Mid(Label7(Index).Caption, 3, 1)) - 2
                            Label7(Index).Caption = Mid(Label7(Index).Caption, 1, 2) & st
                            If st <= 0 Then
                                Image6(Index).Picture = LoadPicture()
                                Image6(Index).Tag = ""
                                Label7(Index).Caption = ""
                                c1(Index) = ""
                                a1(Index) = ""
                            End If
                        End If
                        If Mid(ix(i), 1, 6) = "image7" Then
                            st = Val(Mid(Label8(Index).Caption, 3, 1)) - 2
                            Label8(Index).Caption = Mid(Label8(Index).Caption, 1, 2) & st
                            If st <= 0 Then
                                Image7(Index).Picture = LoadPicture()
                                Image7(Index).Tag = ""
                                Label8(Index).Caption = ""
                                c2(Index) = ""
                                a2(Index) = ""
                            End If
                        End If
                        If Mid(ix(i), 1, 6) = "image4" Then
                            maxv1 = maxv1 - 2
                            Label1.Caption = maxv1
                        End If
                        If Mid(ix(i), 1, 6) = "image5" Then
                            maxv2 = maxv2 - 2
                            Label1.Caption = maxv2
                        End If
                    Case "Sugar Vida"
                        img = path & "Combates\Sugar Vida\" & play & "\" & reg.sexo & "\"
                        p = 10
                        Timer1.Enabled = True
                        organiza cri
                        If play = 1 Then
                            maxm1 = maxm1 - 12
                            maxv1 = maxv1 + 1
                            maxv2 = maxv2 - 1
                            Label2.Caption = maxm1
                            Label1.Caption = maxv1
                        Else
                            maxm2 = maxm2 - 12
                            maxv2 = maxv2 + 1
                            maxv1 = maxv1 - 1
                            Label2.Caption = maxm2
                            Label1.Caption = maxv1
                        End If
                    Case "Lançar Pedra"
                        ok = True
                        img = path & "Combates\Lancar Pedra\" & play & "\" & reg.sexo & "\"
                        p = 8
                        Timer1.Enabled = True
                        If Mid(ix(i), 1, 6) = "image6" Then
                            st = Val(Mid(Label7(Index).Caption, 3, 1)) - 1
                            Label7(Index).Caption = Mid(Label7(Index).Caption, 1, 2) & st
                            If st <= 0 Then
                                Image6(Index).Picture = LoadPicture()
                                Image6(Index).Tag = ""
                                Label7(Index).Caption = ""
                                c1(Index) = ""
                                a1(Index) = ""
                            End If
                        End If
                        If Mid(ix(i), 1, 6) = "image7" Then
                            st = Val(Mid(Label8(Index).Caption, 3, 1)) - 1
                            Label8(Index).Caption = Mid(Label8(Index).Caption, 1, 2) & st
                            If st <= 0 Then
                                Image7(Index).Picture = LoadPicture()
                                Image7(Index).Tag = ""
                                Label8(Index).Caption = ""
                                c2(Index) = ""
                                a2(Index) = ""
                            End If
                        End If
                        If Mid(ix(i), 1, 6) = "image4" Then
                            maxv1 = maxv1 - 1
                            Label1.Caption = maxv1
                        End If
                        If Mid(ix(i), 1, 6) = "image5" Then
                            maxv2 = maxv2 - 1
                            Label1.Caption = maxv2
                        End If
                End Select
                ix(q) = ""
                Image9(q).Picture = LoadPicture()
                Image9(q).Tag = ""
            Next
        Else
            maxm2 = reg.mana
            For i = 0 To 3
                If Image6(i).Left = 1680 Then
                    maxv2 = maxv2 - Val(Mid(Label7(i).Caption, 1, 1))
                    Image6(i).Left = 840
                End If
            Next
            For i = 0 To 3
                Select Case Image8(i).Tag
                    Index = Val(Mid(ix(q), 7, 1))
                    Case "Prejudicar"
                        ok = True
                        img = path & "Combates\Prejudicar\" & play & "\" & reg.sexo & "\"
                        p = 6
                        Timer1.Enabled = True
                        If Mid(ix(i), 1, 6) = "image6" Then
                            st = Val(Mid(Label7(Index).Caption, 3, 1)) - 1
                            Label7(Index).Caption = Mid(Label7(Index).Caption, 1, 2) & st
                            If st <= 0 Then
                                Image6(Index).Picture = LoadPicture()
                                Image6(Index).Tag = ""
                                Label7(Index).Caption = ""
                                a1(Index) = ""
                                c1(Index) = ""
                            End If
                        End If
                        If Mid(ix(i), 1, 6) = "image7" Then
                            st = Val(Mid(Label8(Index).Caption, 3, 1)) - 1
                            Label8(Index).Caption = Mid(Label8(Index).Caption, 1, 2) & st
                            If st <= 0 Then
                                Image7(Index).Picture = LoadPicture()
                                Image7(Index).Tag = ""
                                Label8(Index).Caption = ""
                                a1(Index) = ""
                                c1(Index) = ""
                            End If
                        End If
                        If Mid(ix(i), 1, 6) = "image4" Then
                            maxv1 = maxv1 - 1
                            Label1.Caption = maxv1
                        End If
                        If Mid(ix(i), 1, 6) = "image5" Then
                            maxv2 = maxv2 - 1
                            Label1.Caption = maxv2
                        End If
                    Case "Desintegrar"
                        img = path & "Combates\Desintegrar ou Polymorph\" & play & "\" & reg.sexo & "\"
                        p = 7
                        Timer1.Enabled = True
                        If Mid(ix(i), 1, 6) = "image6" Then
                            Image6(Index).Picture = LoadPicture()
                            Image6(Index).Tag = ""
                            Label7(Index).Caption = ""
                            c1(Index) = ""
                            a1(Index) = ""
                        End If
                        If Mid(ix(i), 1, 6) = "image7" Then
                            Image7(Index).Picture = LoadPicture()
                            Image7(Index).Tag = ""
                            Label8(Index).Caption = ""
                            c2(Index) = ""
                            a2(Index) = ""
                        End If
                    Case "Trovão"
                        img = path & "Combates\Trovao\" & play & "\" & reg.sexo & "\"
                        p = 7
                        Timer1.Enabled = True
                        If Mid(ix(i), 1, 6) = "image6" Then
                            st = Val(Mid(Label7(Index).Caption, 3, 1)) - 2
                            Label7(Index).Caption = Mid(Label7(Index).Caption, 1, 2) & st
                            If st <= 0 Then
                                Image6(Index).Picture = LoadPicture()
                                Image6(Index).Tag = ""
                                Label7(Index).Caption = ""
                                c1(Index) = ""
                                a1(Index) = ""
                            End If
                        End If
                        If Mid(ix(i), 1, 6) = "image7" Then
                            st = Val(Mid(Label8(Index).Caption, 3, 1)) - 2
                            Label8(Index).Caption = Mid(Label8(Index).Caption, 1, 2) & st
                            If st <= 0 Then
                                Image7(Index).Picture = LoadPicture()
                                Image7(Index).Tag = ""
                                Label8(Index).Caption = ""
                                c2(Index) = ""
                                a2(Index) = ""
                            End If
                        End If
                        If Mid(ix(i), 1, 6) = "image4" Then
                            maxv1 = maxv1 - 2
                            Label1.Caption = maxv1
                        End If
                        If Mid(ix(i), 1, 6) = "image5" Then
                            maxv2 = maxv2 - 2
                            Label1.Caption = maxv2
                        End If
                    Case "Raio Electrico"
                        img = path & "Combates\Raio Eletrico\" & play & "\" & reg.sexo & "\"
                        p = 5
                        Timer1.Enabled = True
                        If Mid(ix(i), 1, 6) = "image6" Then
                            st = Val(Mid(Label7(Index).Caption, 3, 1)) - 1
                            Label7(Index).Caption = Mid(Label7(Index).Caption, 1, 2) & st
                            If st <= 0 Then
                                Image6(Index).Picture = LoadPicture()
                                Image6(Index).Tag = ""
                                Label7(Index).Caption = ""
                                c1(Index) = ""
                                a1(Index) = ""
                            End If
                        End If
                        If Mid(ix(i), 1, 6) = "image7" Then
                            st = Val(Mid(Label8(Index).Caption, 3, 1)) - 1
                            Label8(Index).Caption = Mid(Label8(Index).Caption, 1, 2) & st
                            If st <= 0 Then
                                Image7(Index).Picture = LoadPicture()
                                Image7(Index).Tag = ""
                                Label8(Index).Caption = ""
                                c2(Index) = ""
                                a2(Index) = ""
                            End If
                        End If
                        If Mid(ix(i), 1, 6) = "image4" Then
                            maxv1 = maxv1 - 1
                            Label1.Caption = maxv1
                        End If
                        If Mid(ix(i), 1, 6) = "image5" Then
                            maxv2 = maxv2 - 1
                            Label1.Caption = maxv2
                        End If
                    Case "Tempestade de Gelo"
                        ok = True
                        img = path & "Combates\Tempestade de Gelo\" & play & "\" & reg.sexo & "\"
                        p = 9
                        Timer1.Enabled = True
                        If Mid(ix(i), 1, 6) = "image6" Then
                            st = Val(Mid(Label7(Index).Caption, 3, 1)) - 2
                            Label7(Index).Caption = Mid(Label7(Index).Caption, 1, 2) & st
                            If st <= 0 Then
                                Image6(Index).Picture = LoadPicture()
                                Image6(Index).Tag = ""
                                Label7(Index).Caption = ""
                                c1(Index) = ""
                                a1(Index) = ""
                            End If
                        End If
                        If Mid(ix(i), 1, 6) = "image7" Then
                            st = Val(Mid(Label8(Index).Caption, 3, 1)) - 2
                            Label8(Index).Caption = Mid(Label8(Index).Caption, 1, 2) & st
                            If st <= 0 Then
                                Image7(Index).Picture = LoadPicture()
                                Image7(Index).Tag = ""
                                Label8(Index).Caption = ""
                                c2(Index) = ""
                                a2(Index) = ""
                            End If
                        End If
                        If Mid(ix(i), 1, 6) = "image4" Then
                            maxv1 = maxv1 - 2
                            Label1.Caption = maxv1
                        End If
                        If Mid(ix(i), 1, 6) = "image5" Then
                            maxv2 = maxv2 - 2
                            Label1.Caption = maxv2
                        End If
                    Case "Bola de Fogo"
                        ok = True
                        img = path & "Combates\Bola de Fogo\" & play & "\" & reg.sexo & "\"
                        p = 8
                        Timer1.Enabled = True
                        If Mid(ix(i), 1, 6) = "image6" Then
                            st = Val(Mid(Label7(Index).Caption, 3, 1)) - 2
                            Label7(Index).Caption = Mid(Label7(Index).Caption, 1, 2) & st
                            If st <= 0 Then
                                Image6(Index).Picture = LoadPicture()
                                Image6(Index).Tag = ""
                                Label7(Index).Caption = ""
                                c1(Index) = ""
                                a1(Index) = ""
                            End If
                        End If
                        If Mid(ix(i), 1, 6) = "image7" Then
                            st = Val(Mid(Label8(Index).Caption, 3, 1)) - 2
                            Label8(Index).Caption = Mid(Label8(Index).Caption, 1, 2) & st
                            If st <= 0 Then
                                Image7(Index).Picture = LoadPicture()
                                Image7(Index).Tag = ""
                                Label8(Index).Caption = ""
                                c2(Index) = ""
                                a2(Index) = ""
                            End If
                        End If
                        If Mid(ix(i), 1, 6) = "image4" Then
                            maxv1 = maxv1 - 2
                            Label1.Caption = maxv1
                        End If
                        If Mid(ix(i), 1, 6) = "image5" Then
                            maxv2 = maxv2 - 2
                            Label1.Caption = maxv2
                        End If
                    Case "Sugar Vida"
                        img = path & "Combates\Sugar Vida\" & play & "\" & reg.sexo & "\"
                        p = 10
                        Timer1.Enabled = True
                        organiza cri
                        If play = 1 Then
                            maxm1 = maxm1 - 12
                            maxv1 = maxv1 + 1
                            maxv2 = maxv2 - 1
                            Label2.Caption = maxm1
                            Label1.Caption = maxv1
                        Else
                            maxm2 = maxm2 - 12
                            maxv2 = maxv2 + 1
                            maxv1 = maxv1 - 1
                            Label2.Caption = maxm2
                            Label1.Caption = maxv1
                        End If
                    Case "Lançar Pedra"
                        ok = True
                        img = path & "Combates\Lancar Pedra\" & play & "\" & reg.sexo & "\"
                        p = 8
                        Timer1.Enabled = True
                        If Mid(ix(i), 1, 6) = "image6" Then
                            st = Val(Mid(Label7(Index).Caption, 3, 1)) - 1
                            Label7(Index).Caption = Mid(Label7(Index).Caption, 1, 2) & st
                            If st <= 0 Then
                                Image6(Index).Picture = LoadPicture()
                                Image6(Index).Tag = ""
                                Label7(Index).Caption = ""
                                c1(Index) = ""
                                a1(Index) = ""
                            End If
                        End If
                        If Mid(ix(i), 1, 6) = "image7" Then
                            st = Val(Mid(Label8(Index).Caption, 3, 1)) - 1
                            Label8(Index).Caption = Mid(Label8(Index).Caption, 1, 2) & st
                            If st <= 0 Then
                                Image7(Index).Picture = LoadPicture()
                                Image7(Index).Tag = ""
                                Label8(Index).Caption = ""
                                c2(Index) = ""
                                a2(Index) = ""
                            End If
                        End If
                        If Mid(ix(i), 1, 6) = "image4" Then
                            maxv1 = maxv1 - 1
                            Label1.Caption = maxv1
                        End If
                        If Mid(ix(i), 1, 6) = "image5" Then
                            maxv2 = maxv2 - 1
                            Label1.Caption = maxv2
                        End If
                End Select
                ix(q) = ""
                Image8(q).Picture = LoadPicture()
                Image8(q).Tag = ""
            Next
        End If
    End If
End Sub

Sub abload()
Dim check1, check2, l As Integer
    l = 0
    While (Image3(l).Tag <> "Poção de Vida") And (l < 6)
        l = l + 1
    Wend
    If Image3(l).Tag = "Poção de Vida" Then
        check1 = 1
    Else
        check1 = 0
    End If
    l = 0
    While (Image3(l).Tag <> "Poção de Mana") And (l < 6)
        l = l + 1
    Wend
    If Image3(l).Tag = "Poção de Mana" Then
        check2 = 1
    Else
        check2 = 0
    End If
    toreg
    cont = 0
    Randomize
    Do
        rand = Int(Rnd * 12) + 1
        Select Case rand
            Case 1
                If reg.ar = True Then
                    dom = Int(Rnd * 5) + 1
                    Select Case dom
                        Case 1
                            Image3(cont + car).Tag = "Vitalidade do Ar (0/+1) <-> Mana - 8"
                        Case 2
                            Image3(cont + car).Tag = "Vapores Venenosos (Envenenar) <-> Mana - 10"
                        Case 3
                            Image3(cont + car).Tag = "Chamar Ventos (X-Ataque-X) <-> Mana - 10"
                        Case 4
                            Image3(cont + car).Tag = "Corpo de Ar (0/+2) <-> Mana - 15"
                        Case 5
                            Image3(cont + car).Tag = "Elementar do Ar (7/7) <-> Mana - 20"
                    End Select
                    If play = 1 Then
                        p1(cont + car) = path & "Magias\ar.bmp"
                        v1(cont + car) = Image3(cont + car).Tag
                    Else
                        p2(cont + car) = path & "Magias\ar.bmp"
                        v2(cont + car) = Image3(cont + car).Tag
                    End If
                    Image3(cont + car).Picture = LoadPicture(path & "Magias\ar.bmp")
                    cont = cont + 1
                End If
            Case 2
                If reg.terra = True Then
                    dom = Int(Rnd * 5) + 1
                    Select Case dom
                        Case 1
                            Image3(cont + car).Tag = "Força da Terra (+1/0) <-> Mana - 8"
                        Case 2
                            Image3(cont + car).Tag = "Lançar Pedra (-1 Vida) <-> Mana - 5"
                        Case 3
                            Image3(cont + car).Tag = "Barreira de Pedra (0/8) <-> Mana - 12"
                        Case 4
                            Image3(cont + car).Tag = "Corpo de Pedra (+2/0) <-> Mana - 15"
                        Case 5
                            Image3(cont + car).Tag = "Elementar de Pedra (7/7) <-> Mana - 20"
                    End Select
                    If play = 1 Then
                        p1(cont + car) = path & "Magias\terra.bmp"
                        v1(cont + car) = Image3(cont + car).Tag
                    Else
                        p2(cont + car) = path & "Magias\terra.bmp"
                        v2(cont + car) = Image3(cont + car).Tag
                    End If
                    Image3(cont + car).Picture = LoadPicture(path & "Magias\terra.bmp")
                    cont = cont + 1
                End If
            Case 3
                If reg.fogo = True Then
                    dom = Int(Rnd * 5) + 1
                    Select Case dom
                        Case 1
                            Image3(cont + car).Tag = "Agilidade de Fogo (1 Turno Inbloqueável) <-> Mana - 13"
                        Case 2
                            Image3(cont + car).Tag = "Barreira de Fogo (0/3) <-> Mana - 5"
                        Case 3
                            Image3(cont + car).Tag = "Bola de Fogo (-2 Vida) <-> Mana - 10"
                        Case 4
                            Image3(cont + car).Tag = "Corpo de Fogo (Inbloqueável) <-> Mana - 25"
                        Case 5
                            Image3(cont + car).Tag = "Elementar de Fogo (7/7) <-> Mana - 20"
                    End Select
                    If play = 1 Then
                        p1(cont + car) = path & "Magias\fogo.bmp"
                        v1(cont + car) = Image3(cont + car).Tag
                    Else
                        p2(cont + car) = path & "Magias\fogo.bmp"
                        v2(cont + car) = Image3(cont + car).Tag
                    End If
                    Image3(cont + car).Picture = LoadPicture(path & "Magias\fogo.bmp")
                    cont = cont + 1
                End If
            Case 4
                If reg.agua = True Then
                    dom = Int(Rnd * 5) + 1
                    Select Case dom
                        Case 1
                            Image3(cont + car).Tag = "Pureza de Água (1 Turno Abilidade de Voar) <-> Mana - 11"
                        Case 2
                            Image3(cont + car).Tag = "Chamar Névoa (X-Ataque-X) <-> Mana - 10"
                        Case 3
                            Image3(cont + car).Tag = "Tempestade de Gelo (-2 Vida) <-> Mana - 10"
                        Case 4
                            Image3(cont + car).Tag = "Corpo de Água (Abilidade de Voar) <-> Mana - 20"
                        Case 5
                            Image3(cont + car).Tag = "Elementar de Água (7/7) <-> Mana - 20"
                    End Select
                    If play = 1 Then
                        p1(cont + car) = path & "Magias\agua.bmp"
                        v1(cont + car) = Image3(cont + car).Tag
                    Else
                        p2(cont + car) = path & "Magias\agua.bmp"
                        v2(cont + car) = Image3(cont + car).Tag
                    End If
                    Image3(cont + car).Picture = LoadPicture(path & "Magias\agua.bmp")
                    cont = cont + 1
                End If
            Case 5
                If reg.forca = True Then
                    dom = Int(Rnd * 5) + 1
                    Select Case dom
                        Case 1
                            Image3(cont + car).Tag = "Escudo Protector (X-Magia Inimiga de Ataque-X) <-> Mana - 10"
                        Case 2
                            Image3(cont + car).Tag = "Raio Eléctrico (-1 Vida) <-> Mana - 5"
                        Case 3
                            Image3(cont + car).Tag = "Barreira de Força (0/5) <-> Mana - 8"
                        Case 4
                            Image3(cont + car).Tag = "Trovão (-2 Vida) <-> Mana - 10"
                        Case 5
                            Image3(cont + car).Tag = "Desintegrar (-1 Criatura) <-> Mana - 20"
                    End Select
                    If play = 1 Then
                        p1(cont + car) = path & "Magias\forca.bmp"
                        v1(cont + car) = Image3(cont + car).Tag
                    Else
                        p2(cont + car) = path & "Magias\forca.bmp"
                        v2(cont + car) = Image3(cont + car).Tag
                    End If
                    Image3(cont + car).Picture = LoadPicture(path & "Magias\forca.bmp")
                    cont = cont + 1
                End If
            Case 6
                If reg.natureza = True Then
                    dom = Int(Rnd * 5) + 1
                    Select Case dom
                        Case 1
                            Image3(cont + car).Tag = "Encantar Criatura (1 Turno +1 Criatura) <-> Mana - 5 -> Lobo (3/3) ou 10 -> Urso (5/5)"
                        Case 2
                            Image3(cont + car).Tag = "Floresta (X-Ac. +Ataque/+Vida-X) <-> Mana - 5"
                        Case 3
                            Image3(cont + car).Tag = "Controlar Criatura (+1 Criatura) <-> Mana - 10 -> Lobo (3/3) ou 15 -> Urso (5/5)"
                        Case 4
                            Image3(cont + car).Tag = "Invocar Criatura (Criatura 4/4) <-> Mana - 12"
                        Case 5
                            Image3(cont + car).Tag = "Regeneração (+1 Vida - Jogador/Vida Toda - Criatura) <-> Mana - 10"
                    End Select
                    If play = 1 Then
                        p1(cont + car) = path & "Magias\natureza.bmp"
                        v1(cont + car) = Image3(cont + car).Tag
                    Else
                        p2(cont + car) = path & "Magias\natureza.bmp"
                        v2(cont + car) = Image3(cont + car).Tag
                    End If
                    Image3(cont + car).Picture = LoadPicture(path & "Magias\natureza.bmp")
                    cont = cont + 1
                End If
            Case 7
                If reg.meta = True Then
                    dom = Int(Rnd * 5) + 1
                    Select Case dom
                        Case 1
                            Image3(cont + car).Tag = "Resistir Magia (X-Magia Inimiga-X) <-> Mana - 10"
                        Case 2
                            Image3(cont + car).Tag = "Dispersar Magias (X-Magias-X) <-> Mana - 15"
                        Case 3
                            Image3(cont + car).Tag = "Enfraquecer (-1/0) <-> Mana - 10"
                        Case 4
                            Image3(cont + car).Tag = "Escudo Reflector (<=Reflectir Magias Inimigas<=) <-> Mana - 20"
                        Case 5
                            Image3(cont + car).Tag = "Polymorph (Transforma Criatura 0/0) <-> Mana - 20"
                    End Select
                    If play = 1 Then
                        p1(cont + car) = path & "Magias\meta.bmp"
                        v1(cont + car) = Image3(cont + car).Tag
                    Else
                        p2(cont + car) = path & "Magias\meta.bmp"
                        v2(cont + car) = Image3(cont + car).Tag
                    End If
                    Image3(cont + car).Picture = LoadPicture(path & "Magias\meta.bmp")
                    cont = cont + 1
                End If
            Case 8
                If reg.nnegra = True Then
                    dom = Int(Rnd * 5) + 1
                    Select Case dom
                        Case 1
                            Image3(cont + car).Tag = "Prejudicar (-1 Vida) <-> Mana - 5"
                        Case 2
                            Image3(cont + car).Tag = "Conjurar Espírito (+1 Criatura do Cemitério - Perda de -2/-2) <-> Mana - 10"
                        Case 3
                            Image3(cont + car).Tag = "Invocar Zombie (+1 Criatura do Cemitério) <-> Mana - Custo do Original de Invocação -1"
                        Case 4
                            Image3(cont + car).Tag = "Criar Zombie (5/3) <-> Mana - 13"
                        Case 5
                            Image3(cont + car).Tag = "Sugar Vida (+1 Vida Jogador, -1 Vida Oponente) <-> Mana - 12"
                    End Select
                    If play = 1 Then
                        p1(cont + car) = path & "Magias\nnegra.bmp"
                        v1(cont + car) = Image3(cont + car).Tag
                    Else
                        p2(cont + car) = path & "Magias\nnegra.bmp"
                        v2(cont + car) = Image3(cont + car).Tag
                    End If
                    Image3(cont + car).Picture = LoadPicture(path & "Magias\nnegra.bmp")
                    cont = cont + 1
                End If
            Case 9
                If reg.nbranca = True Then
                    dom = Int(Rnd * 5) + 1
                    Select Case dom
                        Case 1
                            Image3(cont + car).Tag = "Curativo (+1 Vida) <-> Mana - 5"
                        Case 2
                            Image3(cont + car).Tag = "Desentoxicar (X-Veneno-X) <-> Mana - 10"
                        Case 3
                            Image3(cont + car).Tag = "Curar (+2 Vida) <-> Mana - 10"
                        Case 4
                            Image3(cont + car).Tag = "Santuário (X-Zombies-X) <-> Mana - 20"
                        Case 5
                            Image3(cont + car).Tag = "Ressuscitar (+1 Criatura do Cemitério) <-> Mana - Custo Igual ao Original de Invocação"
                    End Select
                    If play = 1 Then
                        p1(cont + car) = path & "Magias\nbranca.bmp"
                        v1(cont + car) = Image3(cont + car).Tag
                    Else
                        p2(cont + car) = path & "Magias\nbranca.bmp"
                        v2(cont + car) = Image3(cont + car).Tag
                    End If
                    Image3(cont + car).Picture = LoadPicture(path & "Magias\nbranca.bmp")
                    cont = cont + 1
                End If
            Case 10
                If reg.invoca = True Then
                    dom = Int(Rnd * 5) + 1
                    Select Case dom
                        Case 1
                            Image3(cont + car).Tag = "Troll (1/1) <-> Mana - 5"
                        Case 2
                            Image3(cont + car).Tag = "Orco (2/3) <-> Mana - 10"
                        Case 3
                            Image3(cont + car).Tag = "Ogre (4/5) <-> Mana - 15"
                        Case 4
                            Image3(cont + car).Tag = "Almas Perdidas (Condenado 6/7) <-> Mana - 20"
                        Case 5
                            Image3(cont + car).Tag = "Portas do Inferno (Demónio 8/8) <-> Mana - 25"
                    End Select
                    If play = 1 Then
                        p1(cont + car) = path & "Magias\invoca.bmp"
                        v1(cont + car) = Image3(cont + car).Tag
                    Else
                        p2(cont + car) = path & "Magias\invoca.bmp"
                        v2(cont + car) = Image3(cont + car).Tag
                    End If
                    Image3(cont + car).Picture = LoadPicture(path & "Magias\invoca.bmp")
                    cont = cont + 1
                End If
            Case 11
                If reg.pvida > 0 And check1 = 0 Then
                    Image3(cont).Tag = "Poção de Vida"
                    If play = 1 Then
                        p1(cont + car) = path & "Magias\pocao.bmp"
                        v1(cont + car) = Image3(cont + car).Tag
                    Else
                        p2(cont + car) = path & "Magias\pocao.bmp"
                        v2(cont + car) = Image3(cont + car).Tag
                    End If
                    Image3(cont + car).Picture = LoadPicture(path & "Magias\pocao.bmp")
                    cont = cont + 1
                    check1 = 1
                End If
            Case 12
                If reg.pmana > 0 And check2 = 0 Then
                    Image3(cont + car).Tag = "Poção de Mana"
                    If play = 1 Then
                        p1(cont + car) = path & "Magias\pocao.bmp"
                        v1(cont + car) = Image3(cont + car).Tag
                    Else
                        p2(cont + car) = path & "Magias\pocao.bmp"
                        v2(cont + car) = Image3(cont + car).Tag
                    End If
                    Image3(cont + car).Picture = LoadPicture(path & "Magias\pocao.bmp")
                    cont = cont + 1
                    check2 = 1
                End If
        End Select
    Loop Until cont = cant
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label5.Caption = "Acabar Turno"
End Sub

Sub final()
Dim gold, xp As Integer
    If maxv1 <= 0 Then
        Get #1, 1, reg
        gold = Int(reg.ouro * 0.1)
        xp = Int(reg.exp / 10)
        Get #2, 1, reg
        reg.ouro = reg.ouro + gold
        reg.exp = reg.exp + xp
        reg.pc = reg.pc + 1
        Put #2, 1, reg
        Close #1
        Close #2
        Form2.Visible = True
        Unload Form15
    End If
    If maxv2 <= 0 Then
        Get #2, 1, reg
        gold = Int(reg.ouro * 0.1)
        xp = Int(reg.exp / 10)
        Get #1, 1, reg
        reg.ouro = reg.ouro + gold
        reg.exp = reg.exp + xp
        reg.pc = reg.pc + 1
        Put #1, 1, reg
        Close #1
        Close #2
        Form2.Visible = True
        Unload Form15
    End If
End Sub

Private Sub Command2_Click()
    Close #1
    Close #2
    Form2.Visible = True
    Unload Form15
End Sub

Private Sub Form_Load()
    cant = 7
    santa1 = True
    santa2 = True
    ok = False
    h1 = False
    h2 = False
    cem1 = ""
    cem2 = ""
    cen1 = ""
    cen2 = ""
    mags = ""
    Image10.Picture = LoadPicture(path & "Criaturas\Lobo2.bmp")
    Image11.Picture = LoadPicture(path & "Criaturas\Urso1.bmp")
    car = 0
    turno = 1
    Open pathfx1 For Random As #1 Len = Len(reg)
    Open pathfx2 For Random As #2 Len = Len(reg)
    Get #1, 1, reg
    If reg.sexo = "m" Then
        Image4.Picture = LoadPicture(path & "Combates\M1.bmp")
    Else
        Image4.Picture = LoadPicture(path & "Combates\F1.bmp")
    End If
    maxv1 = reg.vida
    maxm1 = reg.mana
    nome1 = RTrim(reg.nome)
    Get #2, 1, reg
    If reg.sexo = "m" Then
        Image5.Picture = LoadPicture(path & "Combates\M2.bmp")
    Else
        Image5.Picture = LoadPicture(path & "Combates\F2.bmp")
    End If
    maxv2 = reg.vida
    maxm2 = reg.mana
    nome2 = RTrim(reg.nome)
    play = 2
    abload
    play = 1
    abload
    If modo = 0 Then
        cant = 1
        car = 6
    Else
        cant = 7
        car = 0
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label5.Caption = ""
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label5.Caption = "Arena"
End Sub

Private Sub Image10_Click()
    dou = 0
    p = 0
    ok = False
    Select Case mags
        Case "Encantar Criatura (1 Turno +1 Criatura) <-> Mana - 5 -> Lobo (3/3) ou 10 -> Urso (5/5)"
            If Val(Label2.Caption) >= 5 Then
                If play = 1 Then
                    maxm1 = maxm1 - 5
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 5
                    Label2.Caption = maxm2
                    Image5.Left = 8000
                End If
                ok = True
                img = path & "Combates\Summon\" & play & "\" & reg.sexo & "\"
                p = 6
                Timer1.Enabled = True
                organiza cri
                criatura "Lobo (1T)", "3/3", path & "Criaturas\Lobo" & play & ".bmp"
            End If
        Case "Controlar Criatura (+1 Criatura) <-> Mana - 10 -> Lobo (3/3) ou 15 -> Urso (5/5)"
            If Val(Label2.Caption) >= 10 Then
                If play = 1 Then
                    maxm1 = maxm1 - 10
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 10
                    Label2.Caption = maxm2
                    Image5.Left = 8000
                End If
                ok = True
                img = path & "Combates\Summon\" & play & "\" & reg.sexo & "\"
                p = 6
                Timer1.Enabled = True
                organiza cri
                criatura "Lobo", "3/3", path & "Criaturas\Lobo" & play & ".bmp"
            End If
    End Select
    mags = ""
End Sub

Private Sub Image10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label5.Caption = Image10.Tag
End Sub

Private Sub Image11_Click()
    dou = 0
    p = 0
    ok = False
    Select Case mags
        Case "Encantar Criatura (1 Turno +1 Criatura) <-> Mana - 5 -> Lobo (3/3) ou 10 -> Urso (5/5)"
            If Val(Label2.Caption) >= 10 Then
                If play = 1 Then
                    maxm1 = maxm1 - 10
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 10
                    Label2.Caption = maxm2
                    Image5.Left = 8000
                End If
                ok = True
                img = path & "Combates\Summon\" & play & "\" & reg.sexo & "\"
                p = 6
                Timer1.Enabled = True
                organiza cri
                criatura "Urso (1T)", "5/5", path & "Criaturas\Urso" & play & ".bmp"
            End If
        Case "Controlar Criatura (+1 Criatura) <-> Mana - 10 -> Lobo (3/3) ou 15 -> Urso (5/5)"
            If Val(Label2.Caption) >= 15 Then
                If play = 1 Then
                    maxm1 = maxm1 - 15
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 15
                    Label2.Caption = maxm2
                    Image5.Left = 8000
                End If
                ok = True
                img = path & "Combates\Summon\" & play & "\" & reg.sexo & "\"
                p = 6
                Timer1.Enabled = True
                organiza cri
                criatura "Urso", "5/5", path & "Criaturas\Urso" & play & ".bmp"
            End If
    End Select
    mags = ""
End Sub

Private Sub Image11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label5.Caption = Image11.Tag
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label5.Caption = RTrim(reg.nome)
End Sub

Private Sub Image3_Click(Index As Integer)
    detecta Index
    cri = Index
End Sub

Private Sub Image3_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label5.Caption = Image3(Index).Tag
End Sub

Private Sub Image4_Click()
    ok = False
    p = 0
    dou = 0
    Select Case mags
        Case "Vapores Venenosos (Envenenar) <-> Mana - 10"
            img = path & "Combates\Vapores Venenosos\" & play & "\" & reg.sexo & "\"
            p = 7
            Timer1.Enabled = True
            If play = 1 Then
                maxm1 = maxm1 - 10
                Label2.Caption = maxm1
            Else
                maxm2 = maxm2 - 10
                Label2.Caption = maxm2
            End If
            h1 = True
            organiza cri
        Case "Lançar Pedra (-1 Vida) <-> Mana - 5"
            If play = 1 Then
                maxm1 = maxm1 - 5
                Label2.Caption = maxm1
                q = 0
                While (Image8(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image8(q).Tag = "") Then
                    Image8(q).Tag = "Lançar Pedra"
                    Image8(q).Picture = LoadPicture(path & "Combates\Lancar Pedra.bmp")
                    ix(q) = "image4"
                End If
            Else
                maxm2 = maxm2 - 5
                Label2.Caption = maxm2
                q = 0
                While (Image9(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image9(q).Tag = "") Then
                    Image9(q).Tag = "Lançar Pedra"
                    Image9(q).Picture = LoadPicture(path & "Combates\Lancar Pedra.bmp")
                    ix(q) = "image4"
                End If
            End If
            organiza cri
        Case "Bola de Fogo (-2 Vida) <-> Mana - 10"
            If play = 1 Then
                maxm1 = maxm1 - 5
                Label2.Caption = maxm1
                q = 0
                While (Image8(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image8(q).Tag = "") Then
                    Image8(q).Tag = "Bola de Fogo"
                    Image8(q).Picture = LoadPicture(path & "Combates\Bola de Fogo.bmp")
                    ix(q) = "image4"
                End If
            Else
                maxm2 = maxm2 - 5
                Label2.Caption = maxm2
                q = 0
                While (Image9(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image9(q).Tag = "") Then
                    Image9(q).Tag = "Bola de Fogo"
                    Image9(q).Picture = LoadPicture(path & "Combates\Bola de Fogo.bmp")
                    ix(q) = "image4"
                End If
            End If
            organiza cri
        Case "Tempestade de Gelo (-2 Vida) <-> Mana - 10"
            If play = 1 Then
                maxm1 = maxm1 - 5
                Label2.Caption = maxm1
                q = 0
                While (Image8(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image8(q).Tag = "") Then
                    Image8(q).Tag = "Tempestade de Gelo"
                    Image8(q).Picture = LoadPicture(path & "Combates\Tempestade de Gelo.bmp")
                    ix(q) = "image4"
                End If
            Else
                maxm2 = maxm2 - 5
                Label2.Caption = maxm2
                q = 0
                While (Image9(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image9(q).Tag = "") Then
                    Image9(q).Tag = "Tempestade de Gelo"
                    Image9(q).Picture = LoadPicture(path & "Combates\Tempestade de Gelo.bmp")
                    ix(q) = "image4"
                End If
            End If
            organiza cri
        Case "Raio Eléctrico (-1 Vida) <-> Mana - 5"
            If play = 1 Then
                maxm1 = maxm1 - 5
                Label2.Caption = maxm1
                q = 0
                While (Image8(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image8(q).Tag = "") Then
                    Image8(q).Tag = "Raio Electrico"
                    Image8(q).Picture = LoadPicture(path & "Combates\Raio Electrico.bmp")
                    ix(q) = "image4"
                End If
            Else
                maxm2 = maxm2 - 5
                Label2.Caption = maxm2
                q = 0
                While (Image9(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image9(q).Tag = "") Then
                    Image9(q).Tag = "Raio Electrico"
                    Image9(q).Picture = LoadPicture(path & "Combates\Raio Electrico.bmp")
                    ix(q) = "image4"
                End If
            End If
            organiza cri
        Case "Trovão (-2 Vida) <-> Mana - 10"
            If play = 1 Then
                maxm1 = maxm1 - 5
                Label2.Caption = maxm1
                q = 0
                While (Image8(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image8(q).Tag = "") Then
                    Image8(q).Tag = "Trovão"
                    Image8(q).Picture = LoadPicture(path & "Combates\Trovao.bmp")
                    ix(q) = "image4"
                End If
            Else
                maxm2 = maxm2 - 5
                Label2.Caption = maxm2
                q = 0
                While (Image9(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image9(q).Tag = "") Then
                    Image9(q).Tag = "Trovão"
                    Image9(q).Picture = LoadPicture(path & "Combates\Trovão.bmp")
                    ix(q) = "image4"
                End If
            End If
            organiza cri
        Case "Regeneração (+1 Vida - Jogador/Vida Toda - Criatura) <-> Mana - 10"
            If play = 1 Then
                maxm1 = maxm1 - 5
                Label2.Caption = maxm1
            Else
                maxm2 = maxm2 - 5
                Label2.Caption = maxm2
            End If
            maxv1 = maxv1 + 1
            Label1.Caption = maxv1
            organiza cri
        Case "Prejudicar (-1 Vida) <-> Mana - 5"
            If play = 1 Then
                maxm1 = maxm1 - 5
                Label2.Caption = maxm1
                q = 0
                While (Image8(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image8(q).Tag = "") Then
                    Image8(q).Tag = "Prejudicar"
                    Image8(q).Picture = LoadPicture(path & "Combates\Prejudicar.bmp")
                    ix(q) = "image4"
                End If
            Else
                maxm2 = maxm2 - 5
                Label2.Caption = maxm2
                q = 0
                While (Image9(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image9(q).Tag = "") Then
                    Image9(q).Tag = "Prejudicar"
                    Image9(q).Picture = LoadPicture(path & "Combates\Prejudicar.bmp")
                    ix(q) = "image4"
                End If
            End If
            organiza cri
        Case "Curativo (+1 Vida) <-> Mana - 5"
            Get #play, 1, reg
            If reg.vida > maxv1 Then
                If play = 1 Then
                    maxm1 = maxm1 - 5
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 5
                    Label2.Caption = maxm2
                End If
                maxv1 = maxv1 + 1
                Label1.Caption = maxv1
                organiza cri
            End If
        Case "Desentoxicar (X-Veneno-X) <-> Mana - 10"
            If play = 1 Then
                maxm1 = maxm1 - 10
                Label2.Caption = maxm1
            Else
                maxm2 = maxm2 - 10
                Label2.Caption = maxm2
            End If
            h1 = False
            organiza cri
        Case "Curar (+2 Vida) <-> Mana - 10"
            Get #play, 1, reg
            If reg.vida > maxv1 Then
                If play = 1 Then
                    maxm1 = maxm1 - 10
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 10
                    Label2.Caption = maxm2
                End If
                maxv1 = maxv1 + 1
                Label1.Caption = maxv1
                organiza cri
            End If
    End Select
    mags = ""
    For i = 0 To 6
        Image3(i).BorderStyle = 1
    Next
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label5.Caption = nome1
End Sub

Private Sub Image5_Click()
    ok = False
    p = 0
    dou = 0
    Select Case mags
        Case "Vapores Venenosos (Envenenar) <-> Mana - 10"
            img = path & "Combates\Vapores Venenosos\" & play & "\" & reg.sexo & "\"
            p = 7
            Timer1.Enabled = True
            If play = 1 Then
                maxm1 = maxm1 - 10
                Label2.Caption = maxm1
            Else
                maxm2 = maxm2 - 10
                Label2.Caption = maxm2
            End If
            h2 = True
            organiza cri
        Case "Lançar Pedra (-1 Vida) <-> Mana - 5"
            If play = 1 Then
                maxm1 = maxm1 - 5
                Label2.Caption = maxm1
                q = 0
                While (Image8(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image8(q).Tag = "") Then
                    Image8(q).Tag = "Lançar Pedra"
                    Image8(q).Picture = LoadPicture(path & "Combates\Lancar Pedra.bmp")
                    ix(q) = "image5"
                End If
            Else
                maxm2 = maxm2 - 5
                Label2.Caption = maxm2
                q = 0
                While (Image9(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image9(q).Tag = "") Then
                    Image9(q).Tag = "Lançar Pedra"
                    Image9(q).Picture = LoadPicture(path & "Combates\Lancar Pedra.bmp")
                    ix(q) = "image5"
                End If
            End If
            organiza cri
        Case "Bola de Fogo (-2 Vida) <-> Mana - 10"
            If play = 1 Then
                maxm1 = maxm1 - 5
                Label2.Caption = maxm1
                q = 0
                While (Image8(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image8(q).Tag = "") Then
                    Image8(q).Tag = "Bola de Fogo"
                    Image8(q).Picture = LoadPicture(path & "Combates\Bola de Fogo.bmp")
                    ix(q) = "image5"
                End If
            Else
                maxm2 = maxm2 - 5
                Label2.Caption = maxm2
                q = 0
                While (Image9(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image9(q).Tag = "") Then
                    Image9(q).Tag = "Bola de Fogo"
                    Image9(q).Picture = LoadPicture(path & "Combates\Bola de Fogo.bmp")
                    ix(q) = "image5"
                End If
            End If
            organiza cri
        Case "Tempestade de Gelo (-2 Vida) <-> Mana - 10"
            If play = 1 Then
                maxm1 = maxm1 - 5
                Label2.Caption = maxm1
                q = 0
                While (Image8(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image8(q).Tag = "") Then
                    Image8(q).Tag = "Tempestade de Gelo"
                    Image8(q).Picture = LoadPicture(path & "Combates\Tempestade de Gelo.bmp")
                    ix(q) = "image5"
                End If
            Else
                maxm2 = maxm2 - 5
                Label2.Caption = maxm2
                q = 0
                While (Image9(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image9(q).Tag = "") Then
                    Image9(q).Tag = "Tempestade de Gelo"
                    Image9(q).Picture = LoadPicture(path & "Combates\Tempestade de Gelo.bmp")
                    ix(q) = "image5"
                End If
            End If
            organiza cri
        Case "Raio Eléctrico (-1 Vida) <-> Mana - 5"
            If play = 1 Then
                maxm1 = maxm1 - 5
                Label2.Caption = maxm1
                q = 0
                While (Image8(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image8(q).Tag = "") Then
                    Image8(q).Tag = "Raio Electrico"
                    Image8(q).Picture = LoadPicture(path & "Combates\Raio Electrico.bmp")
                    ix(q) = "image5"
                End If
            Else
                maxm2 = maxm2 - 5
                Label2.Caption = maxm2
                q = 0
                While (Image9(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image9(q).Tag = "") Then
                    Image9(q).Tag = "Raio Electrico"
                    Image9(q).Picture = LoadPicture(path & "Combates\Raio Electrico.bmp")
                    ix(q) = "image5"
                End If
            End If
            organiza cri
        Case "Trovão (-2 Vida) <-> Mana - 10"
            If play = 1 Then
                maxm1 = maxm1 - 5
                Label2.Caption = maxm1
                q = 0
                While (Image8(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image8(q).Tag = "") Then
                    Image8(q).Tag = "Trovão"
                    Image8(q).Picture = LoadPicture(path & "Combates\Trovao.bmp")
                    ix(q) = "image5"
                End If
            Else
                maxm2 = maxm2 - 5
                Label2.Caption = maxm2
                q = 0
                While (Image9(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image9(q).Tag = "") Then
                    Image9(q).Tag = "Trovão"
                    Image9(q).Picture = LoadPicture(path & "Combates\Trovão.bmp")
                    ix(q) = "image5"
                End If
            End If
            organiza cri
        Case "Regeneração (+1 Vida - Jogador/Vida Toda - Criatura) <-> Mana - 10"
            If play = 1 Then
                maxm1 = maxm1 - 5
                Label2.Caption = maxm1
            Else
                maxm2 = maxm2 - 5
                Label2.Caption = maxm2
            End If
            maxv2 = maxv2 + 1
            Label1.Caption = maxv2
            organiza cri
        Case "Prejudicar (-1 Vida) <-> Mana - 5"
            If play = 1 Then
                maxm1 = maxm1 - 5
                Label2.Caption = maxm1
                q = 0
                While (Image8(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image8(q).Tag = "") Then
                    Image8(q).Tag = "Prejudicar"
                    Image8(q).Picture = LoadPicture(path & "Combates\Prejudicar.bmp")
                    ix(q) = "image5"
                End If
            Else
                maxm2 = maxm2 - 5
                Label2.Caption = maxm2
                q = 0
                While (Image9(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image9(q).Tag = "") Then
                    Image9(q).Tag = "Prejudicar"
                    Image9(q).Picture = LoadPicture(path & "Combates\Prejudicar.bmp")
                    ix(q) = "image5"
                End If
            End If
            organiza cri
        Case "Curativo (+1 Vida) <-> Mana - 5"
            Get #play, 1, reg
            If reg.vida > maxv2 Then
                If play = 1 Then
                    maxm1 = maxm1 - 5
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 5
                    Label2.Caption = maxm2
                End If
                maxv2 = maxv2 + 1
                Label1.Caption = maxv2
                organiza cri
            End If
        Case "Desentoxicar (X-Veneno-X) <-> Mana - 10"
            If play = 1 Then
                maxm1 = maxm1 - 10
                Label2.Caption = maxm1
            Else
                maxm2 = maxm2 - 10
                Label2.Caption = maxm2
            End If
            h2 = False
            organiza cri
        Case "Curar (+2 Vida) <-> Mana - 10"
            Get #play, 1, reg
            If reg.vida > maxv2 Then
                If play = 1 Then
                    maxm1 = maxm1 - 10
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 10
                    Label2.Caption = maxm2
                End If
                maxv2 = maxv2 + 2
                Label1.Caption = maxv2
                organiza cri
            End If
    End Select
    mags = ""
    For i = 0 To 6
        Image3(i).BorderStyle = 1
    Next
End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label5.Caption = nome2
End Sub

Private Sub Image6_Click(Index As Integer)
Dim st As Integer
Dim q As Integer
    dou = 0
    p = 0
    q = 0
    ok = False
    Select Case mags
        Case "Vitalidade do Ar (0/+1) <-> Mana - 8"
            If Val(Mid(Label7(Index).Caption, 3, 1)) < 9 Then
                ok = True
                img = path & "Combates\Vitalidade do Ar\" & play & "\" & reg.sexo & "\"
                p = 8
                Timer1.Enabled = True
                st = Val(Mid(Label7(Index).Caption, 3, 1)) + 1
                Label7(Index).Caption = Mid(Label7(Index).Caption, 1, 2) & st
                If play = 1 Then
                    maxm1 = maxm1 - 8
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 8
                    Label2.Caption = maxm2
                End If
                organiza cri
            End If
        Case "Vapores Venenosos (Envenenar) <-> Mana - 10"
            If c1(Index) <> "Veneno" Then
                img = path & "Combates\Vapores Venenosos\" & play & "\" & reg.sexo & "\"
                p = 7
                Timer1.Enabled = True
                If play = 1 Then
                    maxm1 = maxm1 - 10
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 10
                    Label2.Caption = maxm2
                End If
                c1(Index) = "Veneno"
                organiza cri
            End If
        Case "Corpo de Ar (0/+2) <-> Mana - 15"
            If Val(Mid(Label7(Index).Caption, 3, 1)) < 8 Then
                ok = True
                img = path & "Combates\Corpo de Ar\" & play & "\" & reg.sexo & "\"
                p = 8
                Timer1.Enabled = True
                st = Val(Mid(Label7(Index).Caption, 3, 1)) + 2
                Label7(Index).Caption = Mid(Label7(Index).Caption, 1, 2) & st
                If play = 1 Then
                    maxm1 = maxm1 - 15
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 15
                    Label2.Caption = maxm2
                End If
                organiza cri
            End If
        Case "Força da Terra (+1/0) <-> Mana - 8"
            If Val(Mid(Label7(Index).Caption, 1, 1)) > 0 And Val(Mid(Label7(Index).Caption, 1, 1)) < 9 Then
                ok = True
                img = path & "Combates\Forca da Terra\" & play & "\" & reg.sexo & "\"
                p = 8
                Timer1.Enabled = True
                st = Val(Mid(Label7(Index).Caption, 1, 1)) + 1
                Label7(Index).Caption = st & Mid(Label7(Index).Caption, 2, 2)
                If play = 1 Then
                    maxm1 = maxm1 - 8
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 8
                    Label2.Caption = maxm2
                End If
                organiza cri
            End If
        Case "Lançar Pedra (-1 Vida) <-> Mana - 5"
            If play = 1 Then
                maxm1 = maxm1 - 5
                Label2.Caption = maxm1
                q = 0
                While (Image8(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image8(q).Tag = "") Then
                    Image8(q).Tag = "Lançar Pedra"
                    Image8(q).Picture = LoadPicture(path & "Combates\Lancar Pedra.bmp")
                    ix(q) = "image6" & Index
                End If
            Else
                maxm2 = maxm2 - 5
                Label2.Caption = maxm2
                q = 0
                While (Image9(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image9(q).Tag = "") Then
                    Image9(q).Tag = "Lançar Pedra"
                    Image9(q).Picture = LoadPicture(path & "Combates\Lancar Pedra.bmp")
                    ix(q) = "image6" & Index
                End If
            End If
            organiza cri
        Case "Corpo de Pedra (+2/0) <-> Mana - 15"
            If Val(Mid(Label7(Index).Caption, 1, 1)) > 0 And Val(Mid(Label7(Index).Caption, 1, 1)) < 8 Then
                ok = True
                img = path & "Combates\Corpo de Pedra\" & play & "\" & reg.sexo & "\"
                p = 8
                Timer1.Enabled = True
                st = Val(Mid(Label7(Index).Caption, 1, 1)) + 2
                Label7(Index).Caption = st & Mid(Label7(Index).Caption, 2, 2)
                If play = 1 Then
                    maxm1 = maxm1 - 15
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 15
                    Label2.Caption = maxm2
                End If
                organiza cri
            End If
        Case "Agilidade de Fogo (1 Turno Inbloqueável) <-> Mana - 13"
            If c1(Index) <> "Bloque1" Then
                ok = True
                img = path & "Combates\Agilidade do Fogo\" & play & "\" & reg.sexo & "\"
                p = 8
                Timer1.Enabled = True
                If play = 1 Then
                    maxm1 = maxm1 - 13
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 13
                    Label2.Caption = maxm2
                End If
                c1(Index) = "Bloque1"
                organiza cri
            End If
        Case "Bola de Fogo (-2 Vida) <-> Mana - 10"
            If play = 1 Then
                maxm1 = maxm1 - 10
                Label2.Caption = maxm1
                q = 0
                While (Image8(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image8(q).Tag = "") Then
                    Image8(q).Tag = "Bola de Fogo"
                    Image8(q).Picture = LoadPicture(path & "Combates\Bola de Fogo.bmp")
                    ix(q) = "image6" & Index
                End If
            Else
                maxm2 = maxm2 - 10
                Label2.Caption = maxm2
                q = 0
                While (Image9(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image9(q).Tag = "") Then
                    Image9(q).Tag = "Bola de Fogo"
                    Image9(q).Picture = LoadPicture(path & "Combates\Bola de Fogo.bmp")
                    ix(q) = "image6" & Index
                End If
            End If
            organiza cri
        Case "Corpo de Fogo (Inbloqueável) <-> Mana - 25"
            If c1(Index) <> "Bloques" Then
                ok = True
                img = path & "Combates\Corpo de Fogo\" & play & "\" & reg.sexo & "\"
                p = 8
                Timer1.Enabled = True
                If play = 1 Then
                    maxm1 = maxm1 - 25
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 25
                    Label2.Caption = maxm2
                End If
                c1(Index) = "Bloques"
                organiza cri
            End If
        Case "Pureza de Água (1 Turno Abilidade de Voar) <-> Mana - 11"
            If c1(Index) <> "Voar1" Then
                ok = True
                img = path & "Combates\Pureza da Agua\" & play & "\" & reg.sexo & "\"
                p = 8
                Timer1.Enabled = True
                If play = 1 Then
                    maxm1 = maxm1 - 11
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 11
                    Label2.Caption = maxm2
                End If
                c1(Index) = "Voar1"
                organiza cri
            End If
        Case "Tempestade de Gelo (-2 Vida) <-> Mana - 10"
            If play = 1 Then
                maxm1 = maxm1 - 10
                Label2.Caption = maxm1
                q = 0
                While (Image8(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image8(q).Tag = "") Then
                    Image8(q).Tag = "Tempestade de Gelo"
                    Image8(q).Picture = LoadPicture(path & "Combates\Tempestade de Gelo.bmp")
                    ix(q) = "image6" & Index
                End If
            Else
                maxm2 = maxm2 - 10
                Label2.Caption = maxm2
                q = 0
                While (Image9(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image9(q).Tag = "") Then
                    Image9(q).Tag = "Tempestade de Gelo"
                    Image9(q).Picture = LoadPicture(path & "Combates\Tempestade de Gelo.bmp")
                    ix(q) = "image6" & Index
                End If
            End If
            organiza cri
        Case "Corpo de Água (Abilidade de Voar) <-> Mana - 20"
            If c1(Index) <> "Voars" Then
                ok = True
                img = path & "Combates\Corpo de Agua\" & play & "\" & reg.sexo & "\"
                p = 8
                Timer1.Enabled = True
                If play = 1 Then
                    maxm1 = maxm1 - 20
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 20
                    Label2.Caption = maxm2
                End If
                c1(Index) = "Voars"
                organiza cri
            End If
        Case "Raio Eléctrico (-1 Vida) <-> Mana - 5"
            If play = 1 Then
                maxm1 = maxm1 - 5
                Label2.Caption = maxm1
                q = 0
                While (Image8(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image8(q).Tag = "") Then
                    Image8(q).Tag = "Raio Electrico"
                    Image8(q).Picture = LoadPicture(path & "Combates\Raio Electrico.bmp")
                    ix(q) = "image6" & Index
                End If
            Else
                maxm2 = maxm2 - 5
                Label2.Caption = maxm2
                q = 0
                While (Image9(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image9(q).Tag = "") Then
                    Image9(q).Tag = "Raio Electrico"
                    Image9(q).Picture = LoadPicture(path & "Combates\Raio Electrico.bmp")
                    ix(q) = "image6" & Index
                End If
            End If
            organiza cri
        Case "Trovão (-2 Vida) <-> Mana - 10"
            If play = 1 Then
                maxm1 = maxm1 - 10
                Label2.Caption = maxm1
                q = 0
                While (Image8(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image8(q).Tag = "") Then
                    Image8(q).Tag = "Trovão"
                    Image8(q).Picture = LoadPicture(path & "Combates\Trovao.bmp")
                    ix(q) = "image6" & Index
                End If
            Else
                maxm2 = maxm2 - 10
                Label2.Caption = maxm2
                q = 0
                While (Image9(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image9(q).Tag = "") Then
                    Image9(q).Tag = "Trovão"
                    Image9(q).Picture = LoadPicture(path & "Combates\Trovao.bmp")
                    ix(q) = "image6" & Index
                End If
            End If
            organiza cri
        Case "Desintegrar (-1 Criatura) <-> Mana - 20"
            If play = 1 Then
                maxm1 = maxm1 - 20
                Label2.Caption = maxm1
                q = 0
                While (Image8(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image8(q).Tag = "") Then
                    Image8(q).Tag = "Desintegrar"
                    Image8(q).Picture = LoadPicture(path & "Combates\Desintegrar ou Polymorph.bmp")
                    ix(q) = "image6" & Index
                End If
            Else
                maxm2 = maxm2 - 20
                Label2.Caption = maxm2
                q = 0
                While (Image9(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image9(q).Tag = "") Then
                    Image9(q).Tag = "Desintegrar"
                    Image9(q).Picture = LoadPicture(path & "Combates\Desintegrar ou Polymorph.bmp")
                    ix(q) = "image6" & Index
                End If
            End If
            organiza cri
        Case "Regeneração (+1 Vida - Jogador/Vida Toda - Criatura) <-> Mana - 10"
            If Val(Mid(a1(Index), 3, 1)) > Val(Mid(Label7(Index).Caption, 3, 1)) Then
                img = path & "Combates\Regeneracao ou Curar ou Curativo ou Desintoxicar\" & play & "\" & reg.sexo & "\"
                p = 5
                Timer1.Enabled = True
                If play = 1 Then
                    maxm1 = maxm1 - 10
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 10
                    Label2.Caption = maxm2
                End If
                st = Val(Mid(a1(Index), 3, 1))
                Label7(Index).Caption = Mid(Label7(Index).Caption, 1, 2) & st
                organiza cri
            End If
        Case "Enfraquecer (-1/0) <-> Mana - 10"
            If Val(Mid(Label7(Index).Caption, 1, 1)) > 0 Then
                ok = True
                img = path & "Combates\Enfraquecer\" & play & "\" & reg.sexo & "\"
                p = 8
                Timer1.Enabled = True
                st = Val(Mid(Label7(Index).Caption, 1, 1)) - 1
                Label7(Index).Caption = st & Mid(Label7(Index).Caption, 2, 2)
                If play = 1 Then
                    maxm1 = maxm1 - 10
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 10
                    Label2.Caption = maxm2
                End If
                organiza cri
            End If
        Case "Polymorph (Transforma Criatura 0/0) <-> Mana - 20"
            If play = 1 Then
                maxm1 = maxm1 - 20
                Label2.Caption = maxm1
            Else
                maxm2 = maxm2 - 20
                Label2.Caption = maxm2
            End If
            img = path & "Combates\Desintegrar ou Polymorph\" & play & "\" & reg.sexo & "\"
            p = 7
            Timer1.Enabled = True
            Image6(Index).Picture = LoadPicture(path & "Criaturas\Ovelha1.bmp")
            Image6(Index).Tag = "Ovelha"
            Label7(Index).Caption = "0/0"
            organiza cri
        Case "Prejudicar (-1 Vida) <-> Mana - 5"
            If play = 1 Then
                maxm1 = maxm1 - 5
                Label2.Caption = maxm1
                q = 0
                While (Image8(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image8(q).Tag = "") Then
                    Image8(q).Tag = "Prejudicar"
                    Image8(q).Picture = LoadPicture(path & "Combates\Prejudicar.bmp")
                    ix(q) = "image6" & Index
                End If
            Else
                maxm2 = maxm2 - 5
                Label2.Caption = maxm2
                q = 0
                While (Image9(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image9(q).Tag = "") Then
                    Image9(q).Tag = "Prejudicar"
                    Image9(q).Picture = LoadPicture(path & "Combates\Prejudicar.bmp")
                    ix(q) = "image6" & Index
                End If
            End If
            organiza cri
        Case "Curativo (+1 Vida) <-> Mana - 5"
            If play = 1 Then
                maxm1 = maxm1 - 5
                Label2.Caption = maxm1
            Else
                maxm2 = maxm2 - 5
                Label2.Caption = maxm2
            End If
            img = path & "Combates\Regeneracao ou Curar ou Curativo ou Desintoxicar\" & play & "\" & reg.sexo & "\"
            p = 5
            Timer1.Enabled = True
            If Val(Mid(Label7(Index).Caption, 3, 1)) < Val(Mid(a1(Index), 3, 1)) Then
                st = Val(Mid(Label7(Index).Caption, 3, 1)) + 1
                Label7(Index).Caption = Mid(Label7(Index).Caption, 1, 2) & st
            End If
            organiza cri
        Case "Desentoxicar (X-Veneno-X) <-> Mana - 10"
            If play = 1 Then
                maxm1 = maxm1 - 10
                Label2.Caption = maxm1
            Else
                maxm2 = maxm2 - 10
                Label2.Caption = maxm2
            End If
            If c1(Index) = "Veneno" Then
                c1(Index) = ""
            End If
            img = path & "Combates\Regeneracao ou Curar ou Curativo ou Desintoxicar\" & play & "\" & reg.sexo & "\"
            p = 5
            Timer1.Enabled = True
            organiza cri
        Case "Curar (+2 Vida) <-> Mana - 10"
            img = path & "Combates\Regeneracao ou Curar ou Curativo ou Desintoxicar\" & play & "\" & reg.sexo & "\"
            p = 5
            Timer1.Enabled = True
            If play = 1 Then
                maxm1 = maxm1 - 10
                Label2.Caption = maxm1
            Else
                maxm2 = maxm2 - 10
                Label2.Caption = maxm2
            End If
            If Val(Mid(Label7(Index).Caption, 3, 1)) < Val(Mid(a1(Index), 3, 1)) Then
                st = Val(Mid(Label7(Index).Caption, 3, 1)) + 2
                If Val(Mid(Label7(Index).Caption, 3, 1)) > Val(Mid(a1(Index), 3, 1)) Then
                    st = Val(Mid(Label7(Index).Caption, 3, 1)) - 1
                End If
                Label7(Index).Caption = Mid(Label7(Index).Caption, 1, 2) & st
            End If
            organiza cri
    End Select
    mags = ""
    For i = 0 To 6
        Image3(i).BorderStyle = 1
    Next
End Sub

Private Sub Image6_DblClick(Index As Integer)
    If play = 1 Then
        If Image6(Index).Tag <> "" Then
            If Image6(Index).Left = 840 Then
                Image6(Index).Left = 1680
                Label7(Index).Left = 1680
            Else
                Image6(Index).Left = 840
                Label7(Index).Left = 840
            End If
        End If
    End If
End Sub

Private Sub Image6_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label5.Caption = Image6(Index).Tag
End Sub

Private Sub Image7_Click(Index As Integer)
Dim st As Integer
Dim q As Integer
    dou = 0
    p = 0
    q = 0
    ok = False
    Select Case mags
        Case "Vitalidade do Ar (0/+1) <-> Mana - 8"
            If Val(Mid(Label8(Index).Caption, 3, 1)) < 9 Then
                ok = True
                img = path & "Combates\Vitalidade do Ar\" & play & "\" & reg.sexo & "\"
                p = 8
                Timer1.Enabled = True
                st = Val(Mid(Label8(Index).Caption, 3, 1)) + 1
                Label8(Index).Caption = Mid(Label8(Index).Caption, 1, 2) & st
                If play = 1 Then
                    maxm1 = maxm1 - 8
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 8
                    Label2.Caption = maxm2
                End If
                organiza cri
            End If
        Case "Vapores Venenosos (Envenenar) <-> Mana - 10"
            If c2(Index) <> "Veneno" Then
                img = path & "Combates\Vapores Venenosos\" & play & "\" & reg.sexo & "\"
                p = 7
                Timer1.Enabled = True
                If play = 1 Then
                    maxm1 = maxm1 - 10
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 10
                    Label2.Caption = maxm2
                End If
                c2(Index) = "Veneno"
                organiza cri
            End If
        Case "Corpo de Ar (0/+2) <-> Mana - 15"
            If Val(Mid(Label8(Index).Caption, 3, 1)) < 8 Then
                ok = True
                img = path & "Combates\Corpo de Ar\" & play & "\" & reg.sexo & "\"
                p = 8
                Timer1.Enabled = True
                st = Val(Mid(Label8(Index).Caption, 3, 1)) + 2
                Label8(Index).Caption = Mid(Label8(Index).Caption, 1, 2) & st
                If play = 1 Then
                    maxm1 = maxm1 - 15
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 15
                    Label2.Caption = maxm2
                End If
                organiza cri
            End If
        Case "Força da Terra (+1/0) <-> Mana - 8"
            If Val(Mid(Label8(Index).Caption, 1, 1)) > 0 And Val(Mid(Label8(Index).Caption, 1, 1)) < 9 Then
                ok = True
                img = path & "Combates\Forca da Terra\" & play & "\" & reg.sexo & "\"
                p = 8
                Timer1.Enabled = True
                st = Val(Mid(Label8(Index).Caption, 1, 1)) + 1
                Label8(Index).Caption = st & Mid(Label8(Index).Caption, 2, 2)
                If play = 1 Then
                    maxm1 = maxm1 - 8
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 8
                    Label2.Caption = maxm2
                End If
                organiza cri
            End If
        Case "Lançar Pedra (-1 Vida) <-> Mana - 5"
            If play = 1 Then
                maxm1 = maxm1 - 5
                Label2.Caption = maxm1
                q = 0
                While (Image8(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image8(q).Tag = "") Then
                    Image8(q).Tag = "Lançar Pedra"
                    Image8(q).Picture = LoadPicture(path & "Combates\Lancar Pedra.bmp")
                    ix(q) = "image7" & Index
                End If
            Else
                maxm2 = maxm2 - 5
                Label2.Caption = maxm2
                q = 0
                While (Image9(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image9(q).Tag = "") Then
                    Image9(q).Tag = "Lançar Pedra"
                    Image9(q).Picture = LoadPicture(path & "Combates\Lancar Pedra.bmp")
                    ix(q) = "image7" & Index
                End If
            End If
            organiza cri
        Case "Corpo de Pedra (+2/0) <-> Mana - 15"
            If Val(Mid(Label8(Index).Caption, 1, 1)) > 0 And Val(Mid(Label8(Index).Caption, 1, 1)) < 8 Then
                ok = True
                img = path & "Combates\Corpo de Pedra\" & play & "\" & reg.sexo & "\"
                p = 8
                Timer1.Enabled = True
                st = Val(Mid(Label8(Index).Caption, 1, 1)) + 2
                Label8(Index).Caption = st & Mid(Label8(Index).Caption, 2, 2)
                If play = 1 Then
                    maxm1 = maxm1 - 15
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 15
                    Label2.Caption = maxm2
                End If
                organiza cri
            End If
        Case "Agilidade de Fogo (1 Turno Inbloqueável) <-> Mana - 13"
            If c2(Index) <> "Bloque1" Then
                ok = True
                img = path & "Combates\Agilidade do Fogo\" & play & "\" & reg.sexo & "\"
                p = 8
                Timer1.Enabled = True
                If play = 1 Then
                    maxm1 = maxm1 - 13
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 13
                    Label2.Caption = maxm2
                End If
                c2(Index) = "Bloque1"
                organiza cri
            End If
        Case "Bola de Fogo (-2 Vida) <-> Mana - 10"
            If play = 1 Then
                maxm1 = maxm1 - 10
                Label2.Caption = maxm1
                q = 0
                While (Image8(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image8(q).Tag = "") Then
                    Image8(q).Tag = "Bola de Fogo"
                    Image8(q).Picture = LoadPicture(path & "Combates\Bola de Fogo.bmp")
                    ix(q) = "image7" & Index
                End If
            Else
                maxm2 = maxm2 - 10
                Label2.Caption = maxm2
                q = 0
                While (Image9(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image9(q).Tag = "") Then
                    Image9(q).Tag = "Bola de Fogo"
                    Image9(q).Picture = LoadPicture(path & "Combates\Bola de Fogo.bmp")
                    ix(q) = "image7" & Index
                End If
            End If
            organiza cri
        Case "Corpo de Fogo (Inbloqueável) <-> Mana - 25"
            If c2(Index) <> "Bloques" Then
                ok = True
                img = path & "Combates\Corpo de Fogo\" & play & "\" & reg.sexo & "\"
                p = 8
                Timer1.Enabled = True
                If play = 1 Then
                    maxm1 = maxm1 - 25
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 25
                    Label2.Caption = maxm2
                End If
                c2(Index) = "Bloques"
                organiza cri
            End If
        Case "Pureza de Água (1 Turno Abilidade de Voar) <-> Mana - 11"
            If c2(Index) <> "Voar1" Then
                ok = True
                img = path & "Combates\Pureza da Agua\" & play & "\" & reg.sexo & "\"
                p = 8
                Timer1.Enabled = True
                If play = 1 Then
                    maxm1 = maxm1 - 11
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 11
                    Label2.Caption = maxm2
                End If
                c2(Index) = "Voar1"
                organiza cri
            End If
        Case "Tempestade de Gelo (-2 Vida) <-> Mana - 10"
            If play = 1 Then
                maxm1 = maxm1 - 10
                Label2.Caption = maxm1
                q = 0
                While (Image8(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image8(q).Tag = "") Then
                    Image8(q).Tag = "Tempestade de Gelo"
                    Image8(q).Picture = LoadPicture(path & "Combates\Tempestade de Gelo.bmp")
                    ix(q) = "image7" & Index
                End If
            Else
                maxm2 = maxm2 - 10
                Label2.Caption = maxm2
                q = 0
                While (Image9(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image9(q).Tag = "") Then
                    Image9(q).Tag = "Tempestade de Gelo"
                    Image9(q).Picture = LoadPicture(path & "Combates\Tempestade de Gelo.bmp")
                    ix(q) = "image7" & Index
                End If
            End If
            organiza cri
        Case "Corpo de Água (Abilidade de Voar) <-> Mana - 20"
            If c2(Index) <> "Voars" Then
                ok = True
                img = path & "Combates\Corpo de Agua\" & play & "\" & reg.sexo & "\"
                p = 8
                Timer1.Enabled = True
                If play = 1 Then
                    maxm1 = maxm1 - 20
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 20
                    Label2.Caption = maxm2
                End If
                c2(Index) = "Voars"
                organiza cri
            End If
        Case "Raio Eléctrico (-1 Vida) <-> Mana - 5"
            If play = 1 Then
                maxm1 = maxm1 - 5
                Label2.Caption = maxm1
                q = 0
                While (Image8(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image8(q).Tag = "") Then
                    Image8(q).Tag = "Raio Electrico"
                    Image8(q).Picture = LoadPicture(path & "Combates\Raio Electrico.bmp")
                    ix(q) = "image7" & Index
                End If
            Else
                maxm2 = maxm2 - 5
                Label2.Caption = maxm2
                q = 0
                While (Image9(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image9(q).Tag = "") Then
                    Image9(q).Tag = "Raio Electrico"
                    Image9(q).Picture = LoadPicture(path & "Combates\Raio Electrico.bmp")
                    ix(q) = "image7" & Index
                End If
            End If
            organiza cri
        Case "Trovão (-2 Vida) <-> Mana - 10"
            If play = 1 Then
                maxm1 = maxm1 - 10
                Label2.Caption = maxm1
                q = 0
                While (Image8(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image8(q).Tag = "") Then
                    Image8(q).Tag = "Trovão"
                    Image8(q).Picture = LoadPicture(path & "Combates\Trovao.bmp")
                    ix(q) = "image7" & Index
                End If
            Else
                maxm2 = maxm2 - 10
                Label2.Caption = maxm2
                q = 0
                While (Image9(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image9(q).Tag = "") Then
                    Image9(q).Tag = "Trovão"
                    Image9(q).Picture = LoadPicture(path & "Combates\Trovao.bmp")
                    ix(q) = "image7" & Index
                End If
            End If
            organiza cri
        Case "Desintegrar (-1 Criatura) <-> Mana - 20"
            If play = 1 Then
                maxm1 = maxm1 - 20
                Label2.Caption = maxm1
                q = 0
                While (Image8(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image8(q).Tag = "") Then
                    Image8(q).Tag = "Desintegrar"
                    Image8(q).Picture = LoadPicture(path & "Combates\Desintegrar ou Polymorph.bmp")
                    ix(q) = "image7" & Index
                End If
            Else
                maxm2 = maxm2 - 20
                Label2.Caption = maxm2
                q = 0
                While (Image9(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image9(q).Tag = "") Then
                    Image9(q).Tag = "Desintegrar"
                    Image9(q).Picture = LoadPicture(path & "Combates\Desintegrar ou Polymorph.bmp")
                    ix(q) = "image7" & Index
                End If
            End If
            organiza cri
        Case "Regeneração (+1 Vida - Jogador/Vida Toda - Criatura) <-> Mana - 10"
            If Val(Mid(a2(Index), 3, 1)) > Val(Mid(Label8(Index).Caption, 3, 1)) Then
                img = path & "Combates\Regeneracao ou Curar ou Curativo ou Desintoxicar\" & play & "\" & reg.sexo & "\"
                p = 5
                Timer1.Enabled = True
                If play = 1 Then
                    maxm1 = maxm1 - 10
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 10
                    Label2.Caption = maxm2
                End If
                st = Val(Mid(a2(Index), 3, 1))
                Label8(Index).Caption = Mid(Label8(Index).Caption, 1, 2) & st
                organiza cri
            End If
        Case "Enfraquecer (-1/0) <-> Mana - 10"
            If Val(Mid(Label8(Index).Caption, 1, 1)) > 0 Then
                ok = True
                img = path & "Combates\Enfraquecer\" & play & "\" & reg.sexo & "\"
                p = 8
                Timer1.Enabled = True
                st = Val(Mid(Label8(Index).Caption, 1, 1)) - 1
                Label8(Index).Caption = st & Mid(Label8(Index).Caption, 2, 2)
                If play = 1 Then
                    maxm1 = maxm1 - 10
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 10
                    Label2.Caption = maxm2
                End If
                organiza cri
            End If
        Case "Polymorph (Transforma Criatura 0/0) <-> Mana - 20"
            If play = 1 Then
                maxm1 = maxm1 - 20
                Label2.Caption = maxm1
            Else
                maxm2 = maxm2 - 20
                Label2.Caption = maxm2
            End If
            img = path & "Combates\Desintegrar ou Polymorph\" & play & "\" & reg.sexo & "\"
            p = 7
            Timer1.Enabled = True
            Image7(Index).Picture = LoadPicture(path & "Criaturas\Ovelha2.bmp")
            Image7(Index).Tag = "Ovelha"
            Label8(Index).Caption = "0/0"
            organiza cri
        Case "Prejudicar (-1 Vida) <-> Mana - 5"
            If play = 1 Then
                maxm1 = maxm1 - 5
                Label2.Caption = maxm1
                q = 0
                While (Image8(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image8(q).Tag = "") Then
                    Image8(q).Tag = "Prejudicar"
                    Image8(q).Picture = LoadPicture(path & "Combates\Prejudicar.bmp")
                    ix(q) = "image7" & Index
                End If
            Else
                maxm2 = maxm2 - 5
                Label2.Caption = maxm2
                q = 0
                While (Image9(q).Tag <> "") And (q < 3)
                    q = q + 1
                Wend
                If (Image9(q).Tag = "") Then
                    Image9(q).Tag = "Prejudicar"
                    Image9(q).Picture = LoadPicture(path & "Combates\Prejudicar.bmp")
                    ix(q) = "image7" & Index
                End If
            End If
            organiza cri
        Case "Curativo (+1 Vida) <-> Mana - 5"
            If play = 1 Then
                maxm1 = maxm1 - 5
                Label2.Caption = maxm1
            Else
                maxm2 = maxm2 - 5
                Label2.Caption = maxm2
            End If
            img = path & "Combates\Regeneracao ou Curar ou Curativo ou Desintoxicar\" & play & "\" & reg.sexo & "\"
            p = 5
            Timer1.Enabled = True
            If Val(Mid(Label8(Index).Caption, 3, 1)) < Val(Mid(a2(Index), 3, 1)) Then
                st = Val(Mid(Label8(Index).Caption, 3, 1)) + 1
                Label8(Index).Caption = Mid(Label8(Index).Caption, 1, 2) & st
            End If
            organiza cri
        Case "Desentoxicar (X-Veneno-X) <-> Mana - 10"
            If play = 1 Then
                maxm1 = maxm1 - 10
                Label2.Caption = maxm1
            Else
                maxm2 = maxm2 - 10
                Label2.Caption = maxm2
            End If
            If c2(Index) = "Veneno" Then
                c2(Index) = ""
            End If
            img = path & "Combates\Regeneracao ou Curar ou Curativo ou Desintoxicar\" & play & "\" & reg.sexo & "\"
            p = 5
            Timer1.Enabled = True
            organiza cri
        Case "Curar (+2 Vida) <-> Mana - 10"
            img = path & "Combates\Regeneracao ou Curar ou Curativo ou Desintoxicar\" & play & "\" & reg.sexo & "\"
            p = 5
            Timer1.Enabled = True
            If play = 1 Then
                maxm1 = maxm1 - 10
                Label2.Caption = maxm1
            Else
                maxm2 = maxm2 - 10
                Label2.Caption = maxm2
            End If
            If Val(Mid(Label8(Index).Caption, 3, 1)) < Val(Mid(a2(Index), 3, 1)) Then
                st = Val(Mid(Label8(Index).Caption, 3, 1)) + 2
                If Val(Mid(Label8(Index).Caption, 3, 1)) > Val(Mid(a2(Index), 3, 1)) Then
                    st = Val(Mid(Label8(Index).Caption, 3, 1)) - 1
                End If
                Label8(Index).Caption = Mid(Label8(Index).Caption, 1, 2) & st
            End If
            organiza cri
    End Select
    mags = ""
    For i = 0 To 6
        Image3(i).BorderStyle = 1
    Next
End Sub

Private Sub Image7_DblClick(Index As Integer)
    If play = 2 Then
        If Image7(Index).Tag <> "" Then
            If Image7(Index).Left = 10560 Then
                Image7(Index).Left = 9480
                Label8(Index).Left = 9480
            Else
                Image7(Index).Left = 10560
                Label8(Index).Left = 10560
            End If
        End If
    End If
End Sub

Private Sub Image7_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label5.Caption = Image7(Index).Tag
End Sub

Private Sub Image8_Click(Index As Integer)
    Select Case mags
        Case "Escudo Protector (X-Magia Inimiga de Ataque-X) <-> Mana - 10"
            If Image8(Index).Tag = "Prejudicar" Or Image8(Index).Tag = "Desintegrar" Or Image8(Index).Tag = "Trovão" Or Image8(Index).Tag = "Raio Electrico" Or Image8(Index).Tag = "Tempestade de Gelo" Or Image8(Index).Tag = "Bola de Fogo" Or Image8(Index).Tag = "Sugar Vida" Or Image8(Index).Tag = "Lançar Pedra" Then
                img = path & "Combates\Escudo Protector\" & play & "\" & reg.sexo & "\"
                p = 5
                Timer1.Enabled = True
                If play = 1 Then
                    maxm1 = maxm1 - 10
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 10
                    Label2.Caption = maxm2
                End If
                Image8(Index).Picture = LoadPicture()
                Image8(Index).Tag = ""
                organiza cri
            End If
        Case "Resistir Magia (X-Magia Inimiga-X) <-> Mana - 10"
            img = path & "Combates\Resistir Magia ou Santuario\" & play & "\" & reg.sexo & "\"
            p = 5
            Timer1.Enabled = True
            If play = 1 Then
                maxm1 = maxm1 - 10
                Label2.Caption = maxm1
            Else
                maxm2 = maxm2 - 10
                Label2.Caption = maxm2
            End If
            Image8(Index).Picture = LoadPicture()
            Image8(Index).Tag = ""
            organiza cri
    End Select
    mags = ""
    For i = 0 To 6
        Image3(i).BorderStyle = 1
    Next
End Sub

Private Sub Image8_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label5.Caption = Image8(Index).Tag
End Sub

Private Sub Image9_Click(Index As Integer)
    Select Case mags
        Case "Escudo Protector (X-Magia Inimiga de Ataque-X) <-> Mana - 10"
            If Image9(Index).Tag = "Prejudicar" Or Image9(Index).Tag = "Desintegrar" Or Image9(Index).Tag = "Trovão" Or Image9(Index).Tag = "Raio Electrico" Or Image9(Index).Tag = "Tempestade de Gelo" Or Image9(Index).Tag = "Bola de Fogo" Or Image9(Index).Tag = "Sugar Vida" Or Image9(Index).Tag = "Lançar Pedra" Then
                img = path & "Combates\Escudo Protector\" & play & "\" & reg.sexo & "\"
                p = 5
                Timer1.Enabled = True
                If play = 1 Then
                    maxm1 = maxm1 - 10
                    Label2.Caption = maxm1
                Else
                    maxm2 = maxm2 - 10
                    Label2.Caption = maxm2
                End If
                Image9(Index).Picture = LoadPicture()
                Image9(Index).Tag = ""
                organiza cri
            End If
        Case "Resistir Magia (X-Magia Inimiga-X) <-> Mana - 10"
            img = path & "Combates\Resistir Magia ou Santuario\" & play & "\" & reg.sexo & "\"
            p = 5
            Timer1.Enabled = True
            If play = 1 Then
                maxm1 = maxm1 - 10
                Label2.Caption = maxm1
            Else
                maxm2 = maxm2 - 10
                Label2.Caption = maxm2
            End If
            Image9(Index).Picture = LoadPicture()
            Image9(Index).Tag = ""
            organiza cri
    End Select
    mags = ""
    For i = 0 To 6
        Image3(i).BorderStyle = 1
    Next
End Sub

Private Sub Image9_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label5.Caption = Image9(Index).Tag
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label5.Caption = "Vida"
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label5.Caption = "Mana"
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label5.Caption = "Vida"
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label5.Caption = ""
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label5.Caption = "Mana"
End Sub

Private Sub Timer1_Timer()
    If dou = 0 Then
        Command1.Enabled = False
        For i = 0 To 6
            Image3(i).Enabled = False
        Next
    End If
    dou = dou + 1
    If dou = p Then
        If play = 1 Then
            Image4.Picture = LoadPicture(path & "Combates\" & reg.sexo & play & ".bmp")
        Else
            Image5.Picture = LoadPicture(path & "Combates\" & reg.sexo & play & ".bmp")
        End If
        Image12.Visible = False
        Image12.Picture = LoadPicture()
        Command1.Enabled = True
        For i = 0 To 6
            Image3(i).Enabled = True
        Next
        If play = 2 Then
            Image5.Left = 8640
        End If
        Timer1.Enabled = False
    Else
        If dou = p - 1 And ok = True Then
            If play = 1 Then
                Image4.Picture = LoadPicture(path & "Combates\" & reg.sexo & play & ".bmp")
            Else
                Image5.Picture = LoadPicture(path & "Combates\" & reg.sexo & play & ".bmp")
            End If
            Image12.Visible = True
            Image12.Picture = LoadPicture(img & dou & ".bmp")
        Else
            If play = 1 Then
                Image4.Picture = LoadPicture(img & dou & ".bmp")
            Else
                Image5.Picture = LoadPicture(img & dou & ".bmp")
            End If
        End If
    End If
End Sub

Private Sub Timer2_Timer()
    final
End Sub
