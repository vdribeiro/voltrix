VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form10 
   BorderStyle     =   0  'None
   Caption         =   "Personagem"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form25"
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton Command5 
      Caption         =   "Anterior"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   18
      Top             =   8040
      Width           =   2655
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "*.bmp"
      Filter          =   "Todos os ficheiros de imagem|*.bmp;*.jpg;*.gif|Bitmap|*.bmp|JPEG|*.jpg|GIF|*.gif"
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080FF80&
      Height          =   615
      Left            =   5880
      Picture         =   "Form10.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FF80&
      Height          =   615
      Left            =   2160
      Picture         =   "Form10.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Seguinte"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      TabIndex        =   7
      Top             =   8040
      Width           =   2655
   End
   Begin VB.Frame Frame2 
      Caption         =   "Sexo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   6240
      TabIndex        =   4
      Top             =   5040
      Width           =   2775
      Begin VB.OptionButton Option2 
         Caption         =   "Feminino"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   1440
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Masculino"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   840
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Raça"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   1560
      TabIndex        =   3
      Top             =   4440
      Width           =   3975
      Begin VB.OptionButton Option3 
         Caption         =   "Meio-Ogre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   1080
         TabIndex        =   17
         Top             =   3000
         Width           =   1575
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Meio-Orco"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   1080
         TabIndex        =   16
         Top             =   2640
         Width           =   1575
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Halfling"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   1080
         TabIndex        =   15
         Top             =   2280
         Width           =   1575
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Gnomo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   1080
         TabIndex        =   14
         Top             =   1920
         Width           =   1575
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Meio-Elfo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   13
         Top             =   1560
         Width           =   1575
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Elfo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   12
         Top             =   1200
         Width           =   1575
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Anão"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   11
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Humano"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   8
         Top             =   480
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2760
      MaxLength       =   25
      TabIndex        =   0
      Top             =   3480
      Width           =   5535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Importar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1560
      TabIndex        =   2
      Top             =   3480
      Width           =   1065
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1950
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   1950
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim raca As String
Dim sexo As String * 1
Dim n As Integer
Sub cdn()
    CommonDialog1.FileName = path & "Faces\" & raca & sexo & n & ".bmp"
End Sub

Private Sub Command1_Click()
    CommonDialog1.InitDir = path
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
        Image1.Picture = LoadPicture(CommonDialog1.FileName)
        reg.foto = CommonDialog1.FileName
    End If
End Sub

Private Sub Command2_Click()
    If Text1.Text <> "" Then
        reg.nome = Text1.Text
        reg.raca = raca
        reg.sexo = sexo
        reg.foto = CommonDialog1.FileName
        Form11.Visible = True
        Form10.Visible = False
    Else
        MsgBox "Deve preencher o campo -Nome-", vbOKOnly, "Personagem"
    End If
End Sub

Private Sub Command5_Click()
    Unload Form10
    Unload Form11
    Unload Form12
    Form13.Visible = True
End Sub


Private Sub Form_Load()
    raca = "h"
    sexo = "m"
    n = 1
    Image1.Picture = LoadPicture(path & "Faces\" & raca & sexo & n & ".bmp")
    cdn
End Sub

Private Sub Option1_click()
    sexo = "m"
    n = 1
    For i = 0 To 7
        Option3(i).Enabled = True
    Next
    Image1.Picture = LoadPicture(path & "Faces\" & raca & sexo & n & ".bmp")
    cdn
End Sub

Private Sub Option2_Click()
    sexo = "f"
    n = 1
    Option3(1).Enabled = False
    Option3(4).Enabled = False
    Option3(5).Enabled = False
    Option3(7).Enabled = False
    Image1.Picture = LoadPicture(path & "Faces\" & raca & sexo & n & ".bmp")
    cdn
End Sub

Private Sub Command3_Click()
    If n > 1 Then
        n = n - 1
    Else
        n = 4
    End If
    cdn
    Image1.Picture = LoadPicture(path & "Faces\" & raca & sexo & n & ".bmp")
End Sub

Private Sub Command4_Click()
    If n < 4 Then
        n = n + 1
    Else
        n = 1
    End If
    cdn
    Image1.Picture = LoadPicture(path & "Faces\" & raca & sexo & n & ".bmp")
End Sub

Private Sub Option3_Click(Index As Integer)
    n = 1
    Select Case Index
        Case 0
            raca = "h"
            Option2.Enabled = True
        Case 1
            raca = "d"
            Option2.Enabled = False
        Case 2
            raca = "e"
            Option2.Enabled = True
        Case 3
            raca = "he"
            Option2.Enabled = True
        Case 4
            raca = "g"
            Option2.Enabled = False
        Case 5
            raca = "ha"
            Option2.Enabled = False
        Case 6
            raca = "ho"
            Option2.Enabled = True
        Case 7
            raca = "hog"
            Option2.Enabled = False
    End Select
    cdn
    Image1.Picture = LoadPicture(path & "Faces\" & raca & sexo & "1.bmp")
End Sub
