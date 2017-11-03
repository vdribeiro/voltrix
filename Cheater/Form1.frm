VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      Cancel          =   -1  'True
      Caption         =   "INSERIR"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   60
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   19
      Top             =   7800
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SAIR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   60
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7080
      TabIndex        =   18
      Top             =   7800
      Width           =   4815
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "bmp|*.bmp|jpg|*.jpg|gif|*.gif"
   End
   Begin VB.CommandButton Command2 
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
      Left            =   600
      TabIndex        =   16
      Top             =   4320
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox Text2 
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
      Left            =   1680
      MaxLength       =   25
      TabIndex        =   15
      Top             =   3600
      Visible         =   0   'False
      Width           =   5535
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
      Left            =   7440
      TabIndex        =   6
      Top             =   3480
      Visible         =   0   'False
      Width           =   3975
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
         TabIndex        =   14
         Top             =   480
         Value           =   -1  'True
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
         TabIndex        =   13
         Top             =   840
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
         TabIndex        =   11
         Top             =   1560
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
         TabIndex        =   10
         Top             =   1920
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
         TabIndex        =   9
         Top             =   2280
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
         TabIndex        =   8
         Top             =   2640
         Width           =   1575
      End
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
         TabIndex        =   7
         Top             =   3000
         Width           =   1575
      End
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
      Left            =   3960
      TabIndex        =   3
      Top             =   4680
      Visible         =   0   'False
      Width           =   2775
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
         TabIndex        =   4
         Top             =   1440
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   2
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      IMEMode         =   3  'DISABLE
      Left            =   4680
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1950
      Left            =   960
      Stretch         =   -1  'True
      Top             =   5160
      Visible         =   0   'False
      Width           =   1950
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
      Left            =   480
      TabIndex        =   17
      Top             =   3600
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Mortal Kombat 4 Incomplete"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      TabIndex        =   1
      Top             =   840
      Width           =   3345
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim path, x As String
Dim n As Integer

Private Sub Command1_Click()
    If Text1.Text = "vidrovip" Then
        Label2.Visible = True
        Text2.Visible = True
        Command2.Visible = True
        Image1.Visible = True
        Frame1.Visible = True
        Frame2.Visible = True
        Command4.Visible = True
        Text1.PasswordChar = ""
        Text1.Locked = True
    Else
        MsgBox "Password Incorrecta", vbOKOnly, "CHEATER"
        Text1.Text = ""
    End If
End Sub

Private Sub Command6_Click()

End Sub

Sub gravar()
    Open x For Output As #1 Len = Len(reg)
    Close #1
    Open x For Random As #1 Len = Len(reg)
    Put #1, 1, reg
    Close #1
    MsgBox "Ficheiro gravado como " & x & ". Terminando aplicação...", vbOKOnly, "CHEATER"
    End
End Sub

Private Sub Command2_Click()
    CommonDialog1.InitDir = path
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
        Image1.Picture = LoadPicture(CommonDialog1.FileName)
        reg.foto = CommonDialog1.FileName
        If Text2.Text <> "" Then
            Command4.Enabled = True
        Else
            Command4.Enabled = False
        End If
    End If
End Sub

Private Sub Command3_Click()
    End
End Sub

Private Sub Command4_Click()
    reg.nome = Text2.Text
    reg.ouro = 1000
    reg.vida = 50
    reg.mana = 50
    reg.pvida = 10
    reg.pmana = 10
    reg.pc = 0
    reg.exp = 15000
    reg.ar = True
    reg.terra = True
    reg.fogo = True
    reg.agua = True
    reg.forca = True
    reg.natureza = True
    reg.meta = True
    reg.nnegra = True
    reg.nbranca = True
    reg.invoca = True
    x = InputBox("Insira o nome do ficheiro", "CHEATER", RTrim(reg.nome))
    If x <> "" Then
        x = path & "Personagens\" & x & ".dat"
        If Dir(x) <> "" Then
            bt = MsgBox("Ficheiro existente. Deseja substituir?", vbYesNo, "CHEATER")
            If bt = 6 Then
                gravar
            End If
        Else
            gravar
        End If
    End If
End Sub

Private Sub Form_Load()
    path = "c:\voltrix\"
    reg.raca = "h"
    reg.foto = ""
    reg.sexo = "m"
End Sub

Private Sub Option1_click()
    reg.sexo = "m"
    n = 1
    For i = 0 To 7
        Option3(i).Enabled = True
    Next
End Sub

Private Sub Option2_Click()
    reg.sexo = "f"
    n = 1
    Option3(1).Enabled = False
    Option3(4).Enabled = False
    Option3(5).Enabled = False
    Option3(7).Enabled = False
End Sub

Private Sub Option3_Click(Index As Integer)
    n = 1
    Select Case Index
        Case 0
            reg.raca = "h"
            Option2.Enabled = True
        Case 1
            reg.raca = "d"
            Option2.Enabled = False
        Case 2
            reg.raca = "e"
            Option2.Enabled = True
        Case 3
            reg.raca = "he"
            Option2.Enabled = True
        Case 4
            reg.raca = "g"
            Option2.Enabled = False
        Case 5
            reg.raca = "ha"
            Option2.Enabled = False
        Case 6
            reg.raca = "ho"
            Option2.Enabled = True
        Case 7
            reg.raca = "hog"
            Option2.Enabled = False
    End Select
End Sub

Private Sub Text2_Change()
    If Text2.Text <> "" And RTrim(reg.foto) <> "" Then
        Command4.Enabled = True
    Else
        Command4.Enabled = False
    End If
End Sub
