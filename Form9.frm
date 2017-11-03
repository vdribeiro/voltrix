VERSION 5.00
Begin VB.Form Form9 
   BorderStyle     =   0  'None
   Caption         =   "Menu2"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form24"
   Picture         =   "Form9.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Retroceder"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   525
      Left            =   5160
      TabIndex        =   3
      Top             =   5040
      Width           =   2115
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Editar Personagem"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   525
      Left            =   4320
      TabIndex        =   2
      Top             =   3960
      Width           =   3645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1P VS 2P"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   525
      Left            =   5280
      TabIndex        =   1
      Top             =   3000
      Width           =   1860
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1P VS COM"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   525
      Left            =   5040
      TabIndex        =   0
      Top             =   2040
      Width           =   2430
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim calc, yup As Integer

Private Sub Label1_Click()
    calc = 0
    pathfx1 = ""
    pathfx1 = InputBox("Insira o nome da personagem", "1P VS COM")
    If pathfx1 <> "" Then
        If Dir(path & "Personagens\" & pathfx1 & ".dat") <> "" Then
            pathfx1 = path & "Personagens\" & pathfx1 & ".dat"
            pathfx2 = path & "Personagens\cpu.dat"
            Open pathfx1 For Random As #1 Len = Len(reg)
            Get #1, 1, reg
            Close #1
            Form16.Image1.Picture = LoadPicture(reg.foto)
            Form16.Image2.Picture = LoadPicture(path & "Faces\cpu.jpg")
            If reg.ar = True Then
                calc = calc + 1
            End If
            If reg.terra = True Then
                calc = calc + 1
            End If
            If reg.fogo = True Then
                calc = calc + 1
            End If
            If reg.agua = True Then
                calc = calc + 1
            End If
            If reg.forca = True Then
                calc = calc + 1
            End If
            If reg.natureza = True Then
                calc = calc + 1
            End If
            If reg.meta = True Then
                calc = calc + 1
            End If
            If reg.nnegra = True Then
                calc = calc + 1
            End If
            If reg.nbranca = True Then
                calc = calc + 1
            End If
            If reg.invoca = True Then
                calc = calc + 1
            End If
            reg.ar = False
            reg.terra = False
            reg.fogo = False
            reg.agua = False
            reg.forca = False
            reg.natureza = False
            reg.meta = False
            reg.nnegra = False
            reg.nbranca = False
            reg.invoca = False
            Randomize
            Do
                yup = Int(Rnd * 10) + 1
                Select Case yup
                    Case 1
                        If reg.ar = False Then
                            calc = calc - 1
                            reg.ar = True
                        End If
                    Case 2
                        If reg.terra = False Then
                            calc = calc - 1
                            reg.terra = True
                        End If
                    Case 3
                        If reg.fogo = False Then
                            calc = calc - 1
                            reg.fogo = True
                        End If
                    Case 4
                        If reg.agua = False Then
                            calc = calc - 1
                            reg.agua = True
                        End If
                    Case 5
                        If reg.forca = False Then
                            calc = calc - 1
                            reg.forca = True
                        End If
                    Case 6
                        If reg.natureza = False Then
                            calc = calc - 1
                            reg.natureza = True
                        End If
                    Case 7
                        If reg.meta = False Then
                            calc = calc - 1
                            reg.meta = True
                        End If
                    Case 8
                        If reg.nnegra = False Then
                            calc = calc - 1
                            reg.nnegra = True
                        End If
                    Case 9
                        If reg.nbranca = False Then
                            calc = calc - 1
                            reg.nbranca = True
                        End If
                    Case 10
                        If reg.invoca = False Then
                            calc = calc - 1
                            reg.invoca = True
                        End If
                End Select
            Loop Until calc = 0
            reg.nome = "CPU"
            reg.raca = "h"
            reg.sexo = "m"
            reg.foto = path & "Faces\cpu.jpg"
            Open pathfx2 For Random As #1 Len = Len(reg)
            Put #1, 1, reg
            Close #1
            Form16.Visible = True
            Form9.Visible = False
        Else
            MsgBox "Não encontrado", vbOKOnly, "1P VS COM"
        End If
    End If
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.ForeColor = &HFF&
End Sub

Private Sub Label2_Click()
    pathfx1 = ""
    pathfx2 = ""
    pathfx1 = InputBox("Insira o nome do 1º personagem", "1P VS 2P")
    pathfx2 = InputBox("Insira o nome do 2º personagem", "1P VS 2P")
    If pathfx1 <> "" And pathfx2 <> "" Then
        If Dir(path & "Personagens\" & pathfx1 & ".dat") <> "" And Dir(path & "Personagens\" & pathfx2 & ".dat") <> "" Then
            pathfx1 = path & "Personagens\" & pathfx1 & ".dat"
            pathfx2 = path & "Personagens\" & pathfx2 & ".dat"
            Open pathfx1 For Random As #1 Len = Len(reg)
            Open pathfx2 For Random As #2 Len = Len(reg)
            Get #1, 1, reg
            Form16.Image1.Picture = LoadPicture(reg.foto)
            Get #2, 1, reg
            Form16.Image2.Picture = LoadPicture(reg.foto)
            Close #1
            Close #2
            Form16.Visible = True
            Form9.Visible = False
        Else
            MsgBox "Não encontrado", vbOKOnly, "1P VS 2P"
        End If
    End If
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label2.ForeColor = &HFF&
End Sub

Private Sub Label3_Click()
    Form13.Visible = True
    Form9.Visible = False
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label3.ForeColor = &HFF&
End Sub

Private Sub Label4_Click()
    Form2.Visible = True
    Form9.Visible = False
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label4.ForeColor = &HFF&
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.ForeColor = &HC0&
    Label2.ForeColor = &HC0&
    Label3.ForeColor = &HC0&
    Label4.ForeColor = &HC0&
End Sub
