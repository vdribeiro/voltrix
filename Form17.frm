VERSION 5.00
Begin VB.Form Form17 
   BorderStyle     =   0  'None
   Caption         =   "Info"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form32"
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton Command3 
      Height          =   735
      Left            =   240
      Picture         =   "Form17.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Height          =   735
      Left            =   10440
      Picture         =   "Form17.frx":12A6
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Height          =   855
      Left            =   5640
      Picture         =   "Form17.frx":255C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8040
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   120
      TabIndex        =   15
      Top             =   4560
      Width           =   11775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   120
      TabIndex        =   14
      Top             =   4200
      Width           =   11775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   2040
      Width           =   11775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   11775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   11775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Width           =   11775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   9
      Top             =   3480
      Width           =   11775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   8
      Top             =   3840
      Width           =   11775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   5940
      TabIndex        =   7
      Top             =   120
      Width           =   150
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   120
      TabIndex        =   6
      Top             =   4920
      Width           =   11775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   120
      TabIndex        =   5
      Top             =   5280
      Width           =   11775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   120
      TabIndex        =   4
      Top             =   5640
      Width           =   11775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   120
      TabIndex        =   3
      Top             =   6000
      Width           =   11775
   End
End
Attribute VB_Name = "Form17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form6.Visible = True
    Form17.Visible = False
End Sub

Private Sub Command2_Click()
    Select Case Label2.Caption
        Case "Elfos"
            Form17.Label1(0).Caption = "Detestam orcos e ogres e � minima opurtunidade que tiverem, matam-os. Os elfos conseguem falar com os animais pois"
            Form17.Label1(1).Caption = "conhecem todas as suas linguas. Vivem cerca de 1000 anos para cima e de acordo com as suas idades tratam-se por"
            Form17.Label1(2).Caption = "irm�o/irm� se forem da mesma terra e tiverem sensivelmente a mesma idade; primo/prima se for um elfo de outra"
            Form17.Label1(3).Caption = "terra; pai/m�e se forem s�culos mais velhos; tio/tia a partir dos 1000 anos; elfos verdadeiramente mais velhos s�o"
            Form17.Label1(4).Caption = "tratados por av�/av�. A raridade dos 2000 anos j� foi alcan�ada 2 ou 3 vezes, em todo o caso estes s�o tratados por"
            Form17.Label1(5).Caption = "Anci�o/Anci�."
            Form17.Label1(6).Caption = ""
            Form17.Label1(7).Caption = ""
            Form17.Label1(8).Caption = ""
            Form17.Label1(9).Caption = ""
            Form17.Label1(10).Caption = ""
            Form17.Label1(11).Caption = ""
        Case "Meio-Orcos"
            Form17.Label1(0).Caption = "Os meio-orcos s�o, sem margem de d�vida, muito mais inteligentes que os seus pais orcos. Cont�do, tais seres raramente"
            Form17.Label1(1).Caption = "encontram trabalhos decentes. Quase sempre devido � sua apar�ncia s�o postos a trabalhar como lixeiros, limpa-chamin�s,"
            Form17.Label1(2).Caption = "limpa-retretes, etc... (trabalhos deste g�nero). Tal desconsidera��o a esta ra�a deixa-os profundamente revoltados e, por"
            Form17.Label1(3).Caption = "vezes, leva-os a obterem respeito pela for�a (com este acto s�o normalmente abatidos!). Estes pobres diabos, pelas"
            Form17.Label1(4).Caption = "dificuldades que lhes s�o postas, normalmente optam por serem bandidos. De uma maneira ou doutra, eles vivem cerca de"
            Form17.Label1(5).Caption = "60 anos, que para eles � bastante bom. Como se pode ver, os meio-orcos vivem balan�ando numa l�mina e quando caem �"
            Form17.Label1(6).Caption = "porque s�o empurrados."
            Form17.Label1(7).Caption = ""
            Form17.Label1(8).Caption = ""
            Form17.Label1(9).Caption = ""
            Form17.Label1(10).Caption = ""
            Form17.Label1(11).Caption = ""
        Case "Orcos"
            Form17.Label1(0).Caption = "Mesmo tendo poucos anos de vida, os orcos possuem uma vitalidade assutadora. S�o fortes e se magoados, a ferida �"
            Form17.Label1(1).Caption = "cicatrizada imediatamente. Se um dos membros for separado do resto do corpo, este continua a mexer-se durante um"
            Form17.Label1(2).Caption = "tempo, ali�s, se um dos membros sair, pode ser recolocado se feito por um cirurgi�o competente. Tamb�m s�o resistentes"
            Form17.Label1(3).Caption = "e practicamente imunes a qualquer tipo de doen�as. Aguentam temperaturas extremas, o que faz deles, resumindo e"
            Form17.Label1(4).Caption = "concluindo, criaturas com uma constiti��o corporal vigorosa. Os orcos organizam-se em tribos de 30 a 40 membros onde"
            Form17.Label1(5).Caption = "existe um lider, normalmente um orco macho. Contudo, por vezes v�em-se ogres a liderar tribos de orcos pois quem"
            Form17.Label1(6).Caption = "normalmente assume o titulo de chefe � o ser mais forte independentemente da ra�a. Os orcos s�o seres n�madas que"
            Form17.Label1(7).Caption = "usam pouco do c�rebro que t�m, a motiva��o que eles t�m � ou o medo ou o �dio. O medo - traduz-se no respeito e"
            Form17.Label1(8).Caption = "venera��o. O �dio - motivo de assass�nio ou outro tipo de viol�ncia a practicamente qualquer ser vivo. Os orcos mostram"
            Form17.Label1(9).Caption = "curiosamente um gosto pela tecnologia, especialmente pelas armas de fogo (pouco prov�vel algu�m lhe vender alguma)!!!"
            Form17.Label1(10).Caption = ""
            Form17.Label1(11).Caption = ""
    End Select
    Command3.Visible = True
    Command2.Visible = False
End Sub

Private Sub Command3_Click()
    Select Case Label2.Caption
        Case "Elfos"
            Form17.Label1(0).Caption = "Os elfos s�o criaturas espantosas. Medem cerca de 1,50 metros, ambos os membros (bra�os e pernas) s�o mais longos que"
            Form17.Label1(1).Caption = "os dos humanos, os dedos s�o longos e finos, t�m orelhas grandes e pontiagudas, boca e nariz pequenos. T�m uma pele"
            Form17.Label1(2).Caption = "branca como a neve, cabelo claro que pode aparecer de v�rias cores - as mais apreciadas pelos elfos s�o o vermelho e o"
            Form17.Label1(3).Caption = "preto. A maior peculariedade dos elfos s�o os olhos que podem ser azuis, prateados, verdes, cinzentos ou amarelos. �s"
            Form17.Label1(4).Caption = "vezes a cor dos olhos varia de acordo com o humor da criatura. Mas a cor n�o � a maior qualidade dos olhos desta ra�a"
            Form17.Label1(5).Caption = "fant�stica: possuem tamb�m uma vis�o exepcional que lhes permite ver no escuro e a longas distancias que fazem deles"
            Form17.Label1(6).Caption = "excelentes arqueiros e ca�adores. Os elfos s�o calmos e acolhedores, gentis e gostam de qualquer tipo de piadas ou"
            Form17.Label1(7).Caption = "partidas. Adoram festas e festivais e passam a sua maior parte do tempo a festejar. S�o generosos mas por um lado"
            Form17.Label1(8).Caption = "arrogantes e acham-se superiores �s outras ra�as com exep��o aos halflings quem eles amam extraordin�riamente."
            Form17.Label1(9).Caption = ""
            Form17.Label1(10).Caption = ""
            Form17.Label1(11).Caption = ""
        Case "Meio-Orcos"
            Form17.Label1(0).Caption = "Os meio-orcos s�o uma ra�a hibrida - o resultado da jun��o entre humanos e orcos. Normalmente tal jun��o � feita pela"
            Form17.Label1(1).Caption = "viola��o de uma mulher humana por um orco, tendo-se conhecido, no entanto, casos raros de humanos e orcos se"
            Form17.Label1(2).Caption = "reproduzirem de livre vontade. De uma maneira ou de outra, o resultado de tal jun��o � quase sempre visto com horror e"
            Form17.Label1(3).Caption = "�dio. Os meio-orcos variam largamente de apar�ncia. Eles t�m aproximadamente a mesma altura dos humanos - 1,75m."
            Form17.Label1(4).Caption = "No geral, tendem a parecer orcos robustos contudo livres das irregularidades da coluna que amaldi�oam os pais orcos. Os"
            Form17.Label1(5).Caption = "dentes e a pelugem s�o ligeiramente mais pequenos em compara��o aos orcos de puro sangue. Tais caracter�sticas s�o a"
            Form17.Label1(6).Caption = "regra geral, mas como a exep��o confirma regra, existem meio-orcos que conseguem por vezes passar despercebidos"
            Form17.Label1(7).Caption = "e serem confundidos com humanos. Mas, como sempre, n�o h� pormenores que escapem, tais como o nariz empinado, os"
            Form17.Label1(8).Caption = "pelos grossos e normalmente abundantes e as unhas e os caninos largos e afiados."
            Form17.Label1(9).Caption = ""
            Form17.Label1(10).Caption = ""
            Form17.Label1(11).Caption = ""
        Case "Orcos"
            Form17.Label1(0).Caption = "Bizarro � a palavra que define esta ra�a. Os orcos t�m tendedencias canibais - normalmente comem os mortos em vez de"
            Form17.Label1(1).Caption = "os enterrarem. Medem 1,65 metros no m�ximo, contudo � um pouco dificil de descobrir a altura exacta. A coluna �"
            Form17.Label1(2).Caption = "normalmente curvada e as costas com deforma��es como a de um corcunda, com ombros altos e cabe�a projectada para a"
            Form17.Label1(3).Caption = "frente. Tamb�m t�m orelhas pontiagudas - mas n�o como as dos elfos - um pouco mais deformadas. Boca grande para que"
            Form17.Label1(4).Caption = "caibam os seus dentes grandes e agu�ados. Se um orco perder um dente, este volta a crescer - � uma propriedade que eles"
            Form17.Label1(5).Caption = "partilham com os tubar�es. Vivem cerca de 40 anos, o que � curto, mas em compensa��o reproduzem-se rapidamente."
            Form17.Label1(6).Caption = "Uma f�mea orco pode ter filhos de 4 em 4 m�ses e os seus filhos crescem mais r�pido que os humanos - s�o capazes de"
            Form17.Label1(7).Caption = "andar ou at� mesmo correr aos 6 meses. Normalmente os orcos nascem aos pares ou at� aos triplos. G�meos e trig�meos"
            Form17.Label1(8).Caption = "s�o uma ocorr�ncia comum. Mas como as f�meas s� t�m 2 bra�os e 2 mamas, normalmente o terceiro filho � comido pelo"
            Form17.Label1(9).Caption = "pai como um modo de celebrar a sua fertilidade."
            Form17.Label1(10).Caption = ""
            Form17.Label1(11).Caption = ""
    End Select
    Command2.Visible = True
    Command3.Visible = False
End Sub
