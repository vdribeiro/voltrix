VERSION 5.00
Begin VB.Form Form11 
   BackColor       =   &H80000002&
   BorderStyle     =   0  'None
   Caption         =   "Magias"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form26"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Seguinte"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3360
      TabIndex        =   62
      Top             =   7920
      Width           =   5295
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   1695
      Left            =   3360
      TabIndex        =   61
      Top             =   6000
      Width           =   5295
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Curativo"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   8880
      TabIndex        =   56
      Top             =   6960
      Width           =   2895
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Desintoxicar"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   2
      Left            =   8880
      TabIndex        =   55
      Top             =   7320
      Width           =   2895
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Curar"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   3
      Left            =   8880
      TabIndex        =   54
      Top             =   7680
      Width           =   2895
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Santuário"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   4
      Left            =   8880
      TabIndex        =   53
      Top             =   8040
      Width           =   2895
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ressuscitar"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   5
      Left            =   8880
      TabIndex        =   51
      Top             =   8400
      Width           =   2895
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NECROMANCIA BRANCA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   975
      Index           =   0
      Left            =   8880
      TabIndex        =   50
      Top             =   6000
      Width           =   2895
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Prejudicar"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   60
      Top             =   6960
      Width           =   2895
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Conjurar Espírito"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   59
      Top             =   7320
      Width           =   2895
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Invocar Zombie"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   58
      Top             =   7680
      Width           =   2895
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Criar Zombie"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   57
      Top             =   8040
      Width           =   2895
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sugar Vida"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   5
      Left            =   240
      TabIndex        =   52
      Top             =   8400
      Width           =   2895
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NECROMANCIA NEGRA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   975
      Index           =   0
      Left            =   240
      TabIndex        =   49
      Top             =   6000
      Width           =   2895
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "INVOCAÇÃO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   495
      Index           =   0
      Left            =   8880
      TabIndex        =   48
      Top             =   3480
      Width           =   2895
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Troll"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   8880
      TabIndex        =   47
      Top             =   3960
      Width           =   2895
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Orco"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   2
      Left            =   8880
      TabIndex        =   46
      Top             =   4320
      Width           =   2895
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ogre"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   3
      Left            =   8880
      TabIndex        =   45
      Top             =   4680
      Width           =   2895
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Almas Perdidas"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   4
      Left            =   8880
      TabIndex        =   44
      Top             =   5040
      Width           =   2895
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Portas do Inferno"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   5
      Left            =   8880
      TabIndex        =   43
      Top             =   5400
      Width           =   2895
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "METAMORFOSE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   495
      Index           =   0
      Left            =   6000
      TabIndex        =   42
      Top             =   3480
      Width           =   2895
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NATUREZA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   495
      Index           =   0
      Left            =   3120
      TabIndex        =   41
      Top             =   3480
      Width           =   2895
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FORÇA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   40
      Top             =   3480
      Width           =   2895
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Polymorph"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   5
      Left            =   6000
      TabIndex        =   39
      Top             =   5400
      Width           =   2895
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Regeneração"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   5
      Left            =   3120
      TabIndex        =   38
      Top             =   5400
      Width           =   2895
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Desintegrar"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   5
      Left            =   240
      TabIndex        =   37
      Top             =   5400
      Width           =   2895
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Escudo Reflector"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   4
      Left            =   6000
      TabIndex        =   36
      Top             =   5040
      Width           =   2895
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Enfraquecedor"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   3
      Left            =   6000
      TabIndex        =   35
      Top             =   4680
      Width           =   2895
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dispersar Magias"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   2
      Left            =   6000
      TabIndex        =   34
      Top             =   4320
      Width           =   2895
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Resistir Magia"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   6000
      TabIndex        =   33
      Top             =   3960
      Width           =   2895
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Invocar Criatura"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   4
      Left            =   3120
      TabIndex        =   32
      Top             =   5040
      Width           =   2895
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Controlar Criatura"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   3
      Left            =   3120
      TabIndex        =   31
      Top             =   4680
      Width           =   2895
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Floresta"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   2
      Left            =   3120
      TabIndex        =   30
      Top             =   4320
      Width           =   2895
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Encantar Criatura"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   3120
      TabIndex        =   29
      Top             =   3960
      Width           =   2895
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Trovão"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   28
      Top             =   5040
      Width           =   2895
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Barreira de Força"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   27
      Top             =   4680
      Width           =   2895
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Raio Eléctrico"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   26
      Top             =   4320
      Width           =   2895
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Escudo Protector"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   25
      Top             =   3960
      Width           =   2895
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ÁGUA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   495
      Index           =   0
      Left            =   8880
      TabIndex        =   24
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FOGO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   495
      Index           =   0
      Left            =   6000
      TabIndex        =   23
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TERRA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   495
      Index           =   0
      Left            =   3120
      TabIndex        =   22
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Elementar de Água"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   5
      Left            =   8880
      TabIndex        =   21
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Elementar de Fogo"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   5
      Left            =   6000
      TabIndex        =   20
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Elementar de Pedra"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   5
      Left            =   3120
      TabIndex        =   19
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Elementar de Ar"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   5
      Left            =   240
      TabIndex        =   18
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Corpo de Água"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   4
      Left            =   8880
      TabIndex        =   17
      Top             =   2520
      Width           =   2895
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tempestade de Gelo"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   3
      Left            =   8880
      TabIndex        =   16
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Chamar Nevoa"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   2
      Left            =   8880
      TabIndex        =   15
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pureza de Água"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   8880
      TabIndex        =   14
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Corpo de Fogo"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   4
      Left            =   6000
      TabIndex        =   13
      Top             =   2520
      Width           =   2895
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bola de Fogo"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   3
      Left            =   6000
      TabIndex        =   12
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Barreira de Fogo"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   2
      Left            =   6000
      TabIndex        =   11
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Agilidade de Fogo"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   6000
      TabIndex        =   10
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Corpo de Pedra"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   4
      Left            =   3120
      TabIndex        =   9
      Top             =   2520
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Barreira de Pedra"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   3
      Left            =   3120
      TabIndex        =   8
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lançar Pedra"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   2
      Left            =   3120
      TabIndex        =   7
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Força da Terra"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   3120
      TabIndex        =   6
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Corpo de Ar"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   5
      Top             =   2520
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Chamar Ventos"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Vapores Venenosos"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Vitalidade do Ar"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "AR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Escolha da escola de magia"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    If reg.ar = False And reg.terra = False And reg.fogo = False And reg.agua = False And reg.forca = False And reg.natureza = False And reg.meta = False And reg.invoca = False And reg.nnegra = False And reg.nbranca = False Then
        MsgBox "Escolha uma Escola de Magia", vbOKOnly, "Escolas de Magia"
    Else
        gold = 500
        Form12.Visible = True
        Form11.Visible = False
    End If
End Sub
Sub borda()
    For k = 0 To 5
        Label2(k).BorderStyle = 1
    Next
    For k = 0 To 5
        Label3(k).BorderStyle = 1
    Next
    For k = 0 To 5
        Label4(k).BorderStyle = 1
    Next
    For k = 0 To 5
        Label5(k).BorderStyle = 1
    Next
    For k = 0 To 5
        Label6(k).BorderStyle = 1
    Next
    For k = 0 To 5
        Label7(k).BorderStyle = 1
    Next
    For k = 0 To 5
        Label8(k).BorderStyle = 1
    Next
    For k = 0 To 5
        Label9(k).BorderStyle = 1
    Next
    For k = 0 To 5
        Label10(k).BorderStyle = 1
    Next
    For k = 0 To 5
        Label11(k).BorderStyle = 1
    Next
End Sub
Sub mags()
    reg.ar = False
    reg.terra = False
    reg.fogo = False
    reg.agua = False
    reg.forca = False
    reg.natureza = False
    reg.meta = False
    reg.nbranca = False
    reg.nnegra = False
    reg.invoca = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label12.Caption = ""
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label12.Caption = ""
End Sub

Private Sub Label2_Click(Index As Integer)
    borda
    mags
    For k = 0 To 5
        Label2(k).BorderStyle = 0
    Next
    reg.ar = True
End Sub
Private Sub Label3_Click(Index As Integer)
    borda
    mags
    For k = 0 To 5
        Label3(k).BorderStyle = 0
    Next
    reg.terra = True
End Sub
Private Sub Label4_Click(Index As Integer)
    borda
    mags
    For k = 0 To 5
        Label4(k).BorderStyle = 0
    Next
    reg.fogo = True
End Sub
Private Sub Label5_Click(Index As Integer)
    borda
    mags
    For k = 0 To 5
        Label5(k).BorderStyle = 0
    Next
    reg.agua = True
End Sub
Private Sub Label6_Click(Index As Integer)
    borda
    mags
    For k = 0 To 5
        Label6(k).BorderStyle = 0
    Next
    reg.forca = True
End Sub
Private Sub Label7_Click(Index As Integer)
    borda
    mags
    For k = 0 To 5
        Label7(k).BorderStyle = 0
    Next
    reg.natureza = True
End Sub
Private Sub Label8_Click(Index As Integer)
    borda
    mags
    For k = 0 To 5
        Label8(k).BorderStyle = 0
    Next
    reg.meta = True
End Sub
Private Sub Label9_Click(Index As Integer)
    borda
    mags
    For k = 0 To 5
        Label9(k).BorderStyle = 0
    Next
    reg.invoca = True
End Sub
Private Sub Label10_Click(Index As Integer)
    borda
    mags
    For k = 0 To 5
        Label10(k).BorderStyle = 0
    Next
    reg.nnegra = True
End Sub
Private Sub Label11_Click(Index As Integer)
    borda
    mags
    For k = 0 To 5
        Label11(k).BorderStyle = 0
    Next
    reg.nbranca = True
End Sub

Private Sub Label2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0
            Label12.Caption = "O Colégio do Ar contém os feitiços que manipulam o primeiro Elemento da Terra, o ar e o vento."
        Case 1
            Label12.Caption = "A criatura alvo recebe 0/+1."
        Case 2
            Label12.Caption = "Envenena o jogador ou a criatura alvo (-1 de vida no inicio de cada turno)."
        Case 3
            Label12.Caption = "Impede as criaturas de atacarem no turno presente (pode e deve ser utilizada nos turnos de defesa)."
        Case 4
            Label12.Caption = "A criatura alvo recebe 0/+2."
        Case 5
            Label12.Caption = "Invoca um Elementar de Ar (7/7)."
    End Select
End Sub

Private Sub Label3_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0
            Label12.Caption = "O Colégio da Terra contém os feitiços que manipulam o segundo Elemento da Terra, a terra e a pedra."
        Case 1
            Label12.Caption = "A criatura alvo recebe +1/0."
        Case 2
            Label12.Caption = "Causa 1 ponto de dano ao jogador ou criatura alvo."
        Case 3
            Label12.Caption = "Invoca uma Barreira de Pedra (0/8)."
        Case 4
            Label12.Caption = "A criatura alvo recebe +2/0."
        Case 5
            Label12.Caption = "Invoca um Elementar de Terra (7/7)."
    End Select
End Sub

Private Sub Label4_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0
            Label12.Caption = "O Colégio do Fogo contém os feitiços que manipulam o terceiro Elemento da Terra, o fogo e o calor."
        Case 1
            Label12.Caption = "A criatura alvo fica inbloqueável durante um turno."
        Case 2
            Label12.Caption = "Invoca uma Barreira de Fogo (0/3)."
        Case 3
            Label12.Caption = "Causa 2 pontos de dano ao jogador ou criatura alvo."
        Case 4
            Label12.Caption = "A criatura alvo fica inbloqueável."
        Case 5
            Label12.Caption = "Invoca um Elementar de Fogo (7/7)."
    End Select
End Sub

Private Sub Label5_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0
            Label12.Caption = "O Colégio da Água contém os feitiços que manipulam o quarto Elemento da Terra, a água e o gelo."
        Case 1
            Label12.Caption = "A criatura alvo ganha a abilidade de voar durante um turno."
        Case 2
            Label12.Caption = "Impede as criaturas de atacarem no turno presente (pode e deve ser utilizada nos turnos de defesa)."
        Case 3
            Label12.Caption = "Causa 2 pontos de dano ao jogador ou criatura alvo."
        Case 4
            Label12.Caption = "A criatura alvo ganha a abilidade de voar."
        Case 5
            Label12.Caption = "Invoca um Elementar de Água (7/7)."
    End Select
End Sub

Private Sub Label6_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0
            Label12.Caption = "O Colégio da Força contém os feitiços que manipulam e direccionam a energia pura."
        Case 1
            Label12.Caption = "Protege o jogador de uma das magias do inimigo."
        Case 2
            Label12.Caption = "Causa 1 ponto de dano ao jogador ou criatura alvo."
        Case 3
            Label12.Caption = "Invoca uma Barreira de Força (0/5)."
        Case 4
            Label12.Caption = "Causa 2 pontos de dano ao jogador ou criatura alvo."
        Case 5
            Label12.Caption = "Desintegra a criatura alvo."
    End Select
End Sub

Private Sub Label7_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0
            Label12.Caption = "O Colégio da Natureza contém os feitiços que controlam as plantas, os animais e as forças naturais."
        Case 1
            Label12.Caption = "Controla a criatura alvo da natureza por um turno."
        Case 2
            Label12.Caption = "Impede as criaturas de poder X/X de atacarem (X = nº de vezes que o feitiço é lançado -> acumulativo)"
        Case 3
            Label12.Caption = "Controla a criatura alvo da natureza."
        Case 4
            Label12.Caption = "Invoca a criatura lobo ou urso."
        Case 5
            Label12.Caption = "Regenera a criatura alvo ou dá +1 de vida ao jogador."
    End Select
End Sub

Private Sub Label8_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0
            Label12.Caption = "O Colégio da Metamorfose contém os feitiços que mudam a substância do alvo e afectam outras magias."
        Case 1
            Label12.Caption = "Destrói uma das magias criadas pelo inimigo."
        Case 2
            Label12.Caption = "Destrói todas as magias em jogo."
        Case 3
            Label12.Caption = "A criatura alvo perde -1/0."
        Case 4
            Label12.Caption = "Os feitiços viram-se contra o feiticeiro."
        Case 5
            Label12.Caption = "Transforma a criatura alvo numa ovelha (0/0)."
    End Select
End Sub

Private Sub Label9_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0
            Label12.Caption = "O Colégio da Invocação contém os feitiços que tratam da invocação de criaturas."
        Case 1
            Label12.Caption = "Invoca um Troll (1/1)."
        Case 2
            Label12.Caption = "Invoca um Orco (2/3)."
        Case 3
            Label12.Caption = "Invoca um Ogre (4/5)."
        Case 4
            Label12.Caption = "Invoca um Condenado (6/7)."
        Case 5
            Label12.Caption = "Invoca um Demónio (8/8)."
    End Select
End Sub

Private Sub Label10_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0
            Label12.Caption = "O Colégio da Necromancia Negra contém feitiços que afectam negativamente a vida de uma criatura."
        Case 1
            Label12.Caption = "Causa 1 ponto de dano ao jogador ou criatura alvo."
        Case 2
            Label12.Caption = "Contola o espírito da criatura alvo do cemitério (X-2 -> X = Poder de ataque e vida da critura controlada)."
        Case 3
            Label12.Caption = "Contola a criatura alvo do cemitério."
        Case 4
            Label12.Caption = "Cria um Zombie (5/3)."
        Case 5
            Label12.Caption = "O jogador fica com +1 de vida e o inimigo com -1."
    End Select
End Sub

Private Sub Label11_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0
            Label12.Caption = "O Colégio da Necromancia Branca contém feitiços que afectam positivamente a vida de uma criatura."
        Case 1
            Label12.Caption = "O jogador ou a criatura alvo são curados com +1 de vida (não pode ultrapassar a vida inicial)."
        Case 2
            Label12.Caption = "O jogador ou a criatura alvo deixa de estar envenenada."
        Case 3
            Label12.Caption = "O jogador ou a criatura alvo são curados com +2 de vida (não pode ultrapassar a vida inicial)."
        Case 4
            Label12.Caption = "Impede os Zombies de atacar."
        Case 5
            Label12.Caption = "Devolve uma criatura do cemitério para controle do jogador."
    End Select
End Sub

