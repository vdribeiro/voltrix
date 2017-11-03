VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form16 
   BorderStyle     =   0  'None
   Caption         =   "Config. VS"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form31"
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11160
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Todos os ficheiros de música|*.mp3;*.wav;*.wma;*.mid|MP3|*.mp3|WAV|*.wav|WMA|*.wma|MIDI|*.mid"
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Musica"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   8520
      TabIndex        =   13
      Top             =   2760
      Width           =   3255
      Begin VB.CommandButton Command2 
         Caption         =   "Musica [...]"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   2775
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Distribuição das Cartas"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   8520
      TabIndex        =   3
      Top             =   600
      Width           =   3255
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "Tradicional (+1)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   0
         Left            =   480
         TabIndex        =   5
         Top             =   600
         Width           =   2535
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "Aleatóriamente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   1
         Left            =   480
         TabIndex        =   4
         Top             =   1200
         Width           =   2295
      End
   End
   Begin VB.CommandButton Command1 
      Default         =   -1  'True
      Height          =   975
      Left            =   5160
      Picture         =   "Form16.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7800
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Arenas"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4095
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   7815
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "7 Carvalhos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "Cidade Perdida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Index           =   1
         Left            =   4080
         TabIndex        =   11
         Top             =   600
         Width           =   2895
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "Deserto Vermelho"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   1560
         Width           =   3375
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "Floresta Negra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Index           =   3
         Left            =   4080
         TabIndex        =   9
         Top             =   1560
         Width           =   2775
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "Masmorra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Index           =   4
         Left            =   240
         TabIndex        =   8
         Top             =   2520
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "Aldeia Abandonada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Index           =   5
         Left            =   4080
         TabIndex        =   7
         Top             =   2520
         Width           =   3495
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "Forte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Index           =   6
         Left            =   2760
         TabIndex        =   6
         Top             =   3360
         Width           =   1215
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "VS"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   4920
      TabIndex        =   1
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1950
      Left            =   7080
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1950
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1950
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1950
   End
   Begin VB.Image Image3 
      Height          =   9015
      Left            =   0
      Picture         =   "Form16.frx":49A6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s As Integer
Private Sub Command1_Click()
    Load Form15
    s = -1
    Do
        s = s + 1
    Loop Until (Option1(s).Value = True)
    Form15.Image1.Picture = LoadPicture(path & "Arenas\" & s + 1 & ".bmp")
    If Dir(Form15.MediaPlayer1.FileName) <> "" And Form15.MediaPlayer1.FileName <> "" Then
        Form15.MediaPlayer1.play
    End If
    Form15.Visible = True
    Form16.Visible = False
End Sub

Private Sub Command2_Click()
    CommonDialog1.InitDir = path
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
        Form15.MediaPlayer1.FileName = CommonDialog1.FileName
    End If
End Sub

Private Sub Form_Load()
    Option1(0).Value = True
    Option2(0).Value = True
    modo = 0
End Sub

Private Sub Option2_Click(Index As Integer)
    modo = Index
End Sub
