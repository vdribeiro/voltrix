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
            Form17.Label1(0).Caption = "Detestam orcos e ogres e à minima opurtunidade que tiverem, matam-os. Os elfos conseguem falar com os animais pois"
            Form17.Label1(1).Caption = "conhecem todas as suas linguas. Vivem cerca de 1000 anos para cima e de acordo com as suas idades tratam-se por"
            Form17.Label1(2).Caption = "irmão/irmã se forem da mesma terra e tiverem sensivelmente a mesma idade; primo/prima se for um elfo de outra"
            Form17.Label1(3).Caption = "terra; pai/mãe se forem séculos mais velhos; tio/tia a partir dos 1000 anos; elfos verdadeiramente mais velhos são"
            Form17.Label1(4).Caption = "tratados por avô/avó. A raridade dos 2000 anos já foi alcançada 2 ou 3 vezes, em todo o caso estes são tratados por"
            Form17.Label1(5).Caption = "Ancião/Anciã."
            Form17.Label1(6).Caption = ""
            Form17.Label1(7).Caption = ""
            Form17.Label1(8).Caption = ""
            Form17.Label1(9).Caption = ""
            Form17.Label1(10).Caption = ""
            Form17.Label1(11).Caption = ""
        Case "Meio-Orcos"
            Form17.Label1(0).Caption = "Os meio-orcos são, sem margem de dúvida, muito mais inteligentes que os seus pais orcos. Contúdo, tais seres raramente"
            Form17.Label1(1).Caption = "encontram trabalhos decentes. Quase sempre devido à sua aparência são postos a trabalhar como lixeiros, limpa-chaminés,"
            Form17.Label1(2).Caption = "limpa-retretes, etc... (trabalhos deste género). Tal desconsideração a esta raça deixa-os profundamente revoltados e, por"
            Form17.Label1(3).Caption = "vezes, leva-os a obterem respeito pela força (com este acto são normalmente abatidos!). Estes pobres diabos, pelas"
            Form17.Label1(4).Caption = "dificuldades que lhes são postas, normalmente optam por serem bandidos. De uma maneira ou doutra, eles vivem cerca de"
            Form17.Label1(5).Caption = "60 anos, que para eles é bastante bom. Como se pode ver, os meio-orcos vivem balançando numa lâmina e quando caem é"
            Form17.Label1(6).Caption = "porque são empurrados."
            Form17.Label1(7).Caption = ""
            Form17.Label1(8).Caption = ""
            Form17.Label1(9).Caption = ""
            Form17.Label1(10).Caption = ""
            Form17.Label1(11).Caption = ""
        Case "Orcos"
            Form17.Label1(0).Caption = "Mesmo tendo poucos anos de vida, os orcos possuem uma vitalidade assutadora. São fortes e se magoados, a ferida é"
            Form17.Label1(1).Caption = "cicatrizada imediatamente. Se um dos membros for separado do resto do corpo, este continua a mexer-se durante um"
            Form17.Label1(2).Caption = "tempo, aliás, se um dos membros sair, pode ser recolocado se feito por um cirurgião competente. Também são resistentes"
            Form17.Label1(3).Caption = "e practicamente imunes a qualquer tipo de doenças. Aguentam temperaturas extremas, o que faz deles, resumindo e"
            Form17.Label1(4).Caption = "concluindo, criaturas com uma constitição corporal vigorosa. Os orcos organizam-se em tribos de 30 a 40 membros onde"
            Form17.Label1(5).Caption = "existe um lider, normalmente um orco macho. Contudo, por vezes vêem-se ogres a liderar tribos de orcos pois quem"
            Form17.Label1(6).Caption = "normalmente assume o titulo de chefe é o ser mais forte independentemente da raça. Os orcos são seres nómadas que"
            Form17.Label1(7).Caption = "usam pouco do cérebro que têm, a motivação que eles têm é ou o medo ou o ódio. O medo - traduz-se no respeito e"
            Form17.Label1(8).Caption = "veneração. O ódio - motivo de assassínio ou outro tipo de violência a practicamente qualquer ser vivo. Os orcos mostram"
            Form17.Label1(9).Caption = "curiosamente um gosto pela tecnologia, especialmente pelas armas de fogo (pouco provável alguém lhe vender alguma)!!!"
            Form17.Label1(10).Caption = ""
            Form17.Label1(11).Caption = ""
    End Select
    Command3.Visible = True
    Command2.Visible = False
End Sub

Private Sub Command3_Click()
    Select Case Label2.Caption
        Case "Elfos"
            Form17.Label1(0).Caption = "Os elfos são criaturas espantosas. Medem cerca de 1,50 metros, ambos os membros (braços e pernas) são mais longos que"
            Form17.Label1(1).Caption = "os dos humanos, os dedos são longos e finos, têm orelhas grandes e pontiagudas, boca e nariz pequenos. Têm uma pele"
            Form17.Label1(2).Caption = "branca como a neve, cabelo claro que pode aparecer de várias cores - as mais apreciadas pelos elfos são o vermelho e o"
            Form17.Label1(3).Caption = "preto. A maior peculariedade dos elfos são os olhos que podem ser azuis, prateados, verdes, cinzentos ou amarelos. Às"
            Form17.Label1(4).Caption = "vezes a cor dos olhos varia de acordo com o humor da criatura. Mas a cor não é a maior qualidade dos olhos desta raça"
            Form17.Label1(5).Caption = "fantástica: possuem também uma visão exepcional que lhes permite ver no escuro e a longas distancias que fazem deles"
            Form17.Label1(6).Caption = "excelentes arqueiros e caçadores. Os elfos são calmos e acolhedores, gentis e gostam de qualquer tipo de piadas ou"
            Form17.Label1(7).Caption = "partidas. Adoram festas e festivais e passam a sua maior parte do tempo a festejar. São generosos mas por um lado"
            Form17.Label1(8).Caption = "arrogantes e acham-se superiores às outras raças com exepção aos halflings quem eles amam extraordináriamente."
            Form17.Label1(9).Caption = ""
            Form17.Label1(10).Caption = ""
            Form17.Label1(11).Caption = ""
        Case "Meio-Orcos"
            Form17.Label1(0).Caption = "Os meio-orcos são uma raça hibrida - o resultado da junção entre humanos e orcos. Normalmente tal junção é feita pela"
            Form17.Label1(1).Caption = "violação de uma mulher humana por um orco, tendo-se conhecido, no entanto, casos raros de humanos e orcos se"
            Form17.Label1(2).Caption = "reproduzirem de livre vontade. De uma maneira ou de outra, o resultado de tal junção é quase sempre visto com horror e"
            Form17.Label1(3).Caption = "ódio. Os meio-orcos variam largamente de aparência. Eles têm aproximadamente a mesma altura dos humanos - 1,75m."
            Form17.Label1(4).Caption = "No geral, tendem a parecer orcos robustos contudo livres das irregularidades da coluna que amaldiçoam os pais orcos. Os"
            Form17.Label1(5).Caption = "dentes e a pelugem são ligeiramente mais pequenos em comparação aos orcos de puro sangue. Tais características são a"
            Form17.Label1(6).Caption = "regra geral, mas como a exepção confirma regra, existem meio-orcos que conseguem por vezes passar despercebidos"
            Form17.Label1(7).Caption = "e serem confundidos com humanos. Mas, como sempre, não há pormenores que escapem, tais como o nariz empinado, os"
            Form17.Label1(8).Caption = "pelos grossos e normalmente abundantes e as unhas e os caninos largos e afiados."
            Form17.Label1(9).Caption = ""
            Form17.Label1(10).Caption = ""
            Form17.Label1(11).Caption = ""
        Case "Orcos"
            Form17.Label1(0).Caption = "Bizarro é a palavra que define esta raça. Os orcos têm tendedencias canibais - normalmente comem os mortos em vez de"
            Form17.Label1(1).Caption = "os enterrarem. Medem 1,65 metros no máximo, contudo é um pouco dificil de descobrir a altura exacta. A coluna é"
            Form17.Label1(2).Caption = "normalmente curvada e as costas com deformações como a de um corcunda, com ombros altos e cabeça projectada para a"
            Form17.Label1(3).Caption = "frente. Também têm orelhas pontiagudas - mas não como as dos elfos - um pouco mais deformadas. Boca grande para que"
            Form17.Label1(4).Caption = "caibam os seus dentes grandes e aguçados. Se um orco perder um dente, este volta a crescer - é uma propriedade que eles"
            Form17.Label1(5).Caption = "partilham com os tubarões. Vivem cerca de 40 anos, o que é curto, mas em compensação reproduzem-se rapidamente."
            Form17.Label1(6).Caption = "Uma fêmea orco pode ter filhos de 4 em 4 mêses e os seus filhos crescem mais rápido que os humanos - são capazes de"
            Form17.Label1(7).Caption = "andar ou até mesmo correr aos 6 meses. Normalmente os orcos nascem aos pares ou até aos triplos. Gémeos e trigémeos"
            Form17.Label1(8).Caption = "são uma ocorrência comum. Mas como as fêmeas só têm 2 braços e 2 mamas, normalmente o terceiro filho é comido pelo"
            Form17.Label1(9).Caption = "pai como um modo de celebrar a sua fertilidade."
            Form17.Label1(10).Caption = ""
            Form17.Label1(11).Caption = ""
    End Select
    Command2.Visible = True
    Command3.Visible = False
End Sub
