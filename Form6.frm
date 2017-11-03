VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Raças"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton Command13 
      Caption         =   "Vários"
      Height          =   495
      Left            =   8760
      TabIndex        =   15
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Troll"
      Height          =   495
      Left            =   8760
      TabIndex        =   14
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Ogres"
      Height          =   495
      Left            =   8760
      TabIndex        =   13
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Orcos"
      Height          =   495
      Left            =   8760
      TabIndex        =   12
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Meio-Ogres"
      Height          =   495
      Left            =   2280
      TabIndex        =   11
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Meio-Orcos"
      Height          =   495
      Left            =   1080
      TabIndex        =   10
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Halflings"
      Height          =   495
      Left            =   2280
      TabIndex        =   9
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Gnomos"
      Height          =   495
      Left            =   1080
      TabIndex        =   8
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Anões"
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Meio-Elfos"
      Height          =   495
      Left            =   1080
      TabIndex        =   6
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Elfos"
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Humanos"
      Height          =   495
      Left            =   1080
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Height          =   735
      Left            =   5520
      Picture         =   "Form6.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8040
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OUTRAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Left            =   8400
      TabIndex        =   3
      Top             =   2520
      Width           =   2025
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JOGADOR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Left            =   1080
      TabIndex        =   2
      Top             =   2520
      Width           =   2385
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RAÇAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Left            =   5175
      TabIndex        =   1
      Top             =   360
      Width           =   1665
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   0
      Picture         =   "Form6.frx":2A7E
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Form4.Visible = True
    Form6.Visible = False
End Sub

Private Sub Command10_Click()
    Form17.Command2.Visible = True
    Form17.Command3.Visible = False
    Form17.BackColor = &HFF&
    Form17.Label2.Caption = "Orcos"
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
    Form17.Visible = True
    Form6.Visible = False
End Sub

Private Sub Command11_Click()
    Form17.Command2.Visible = False
    Form17.Command3.Visible = False
    Form17.BackColor = &H80&
    Form17.Label2.Caption = "Ogres"
    Form17.Label1(0).Caption = "Uma raça em extinção, os ogres são gigantes que medem cerca de 3,80 metros. A razão da extinção e de não se saber a"
    Form17.Label1(1).Caption = "esperança média de vida desta raça deve-se à brutalidade com que são tratados. Quando um ogre é visto é imediatamente"
    Form17.Label1(2).Caption = "morto, independentemente da raça que o veja. As tendências canibais desta raça também ajudam, comem normalmente os"
    Form17.Label1(3).Caption = "mortos. Se há estimativa que se possa fazer, eu diria então uns duzentos e tais anos. Os ogres são criaturas solitárias."
    Form17.Label1(4).Caption = "Preferem viver nas montanhas o que normalmente provoca conflitos com os anões que também optam por viver no"
    Form17.Label1(5).Caption = "mesmo sítio. Os ogres são exactamente o oposto dos anões, ou seja, são muito pouco inteligentes e desorganizados, o que"
    Form17.Label1(6).Caption = "se torna uma desvantagem quando há luta de territórios entre estas duas raças. Apesar da sua tremenda força, é raro um"
    Form17.Label1(7).Caption = "ogre ganhar a um anão. A raça que mais detestam são os elfos. Quanto ao método de caça destes seres - é fácil e prático -"
    Form17.Label1(8).Caption = "- escolhem a vítima, escondem-se e quando chegar o momento certo dão-lhe um murro na cabeça fracturando-lhe o crânio."
    Form17.Label1(9).Caption = ""
    Form17.Label1(10).Caption = ""
    Form17.Label1(11).Caption = ""
    Form17.Visible = True
    Form6.Visible = False
End Sub

Private Sub Command12_Click()
    Form17.Command2.Visible = False
    Form17.Command3.Visible = False
    Form17.BackColor = &H8080&
    Form17.Label2.Caption = "Troll"
    Form17.Label1(0).Caption = "Existem duas palavras que definem esta raça: brutos e estupidos. Os trolls são criaturas com mais de 2 metros de altura e"
    Form17.Label1(1).Caption = "que vivem cerca de 30 anos. São uma raça um tanto estranha. São preguiçosos, gordos e feios. O seu grande peso fá-los"
    Form17.Label1(2).Caption = "lentos, no entanto, esta característica não os impede de serem criaturas incrivelmente fortes. Os trolls gostam sempre de"
    Form17.Label1(3).Caption = "aparecer no momento errado. Contúdo, são fáceis de detectar: fazem muito barulho e cheiram a podridão. Estas criaturas,"
    Form17.Label1(4).Caption = "depois de mortas continuam a ser letais; exalam vapores venenosos quando abatidos."
    Form17.Label1(5).Caption = ""
    Form17.Label1(6).Caption = ""
    Form17.Label1(7).Caption = ""
    Form17.Label1(8).Caption = ""
    Form17.Label1(9).Caption = ""
    Form17.Label1(10).Caption = ""
    Form17.Label1(11).Caption = ""
    Form17.Visible = True
    Form6.Visible = False
End Sub

Private Sub Command13_Click()
    Form17.Command2.Visible = False
    Form17.Command3.Visible = False
    Form17.BackColor = &HFFFFFF
    Form17.Label2.Caption = "Vários"
    Form17.Label1(0).Caption = "No percurso do jogo poderás confrontar-te com outras criaturas tais como lobos, ursos, etc... que poderão servir como"
    Form17.Label1(1).Caption = "objectos de invocação/controle. Conforme a arena podem aparecer estas criaturas em quantidade aleatória."
    Form17.Label1(2).Caption = ""
    Form17.Label1(3).Caption = ""
    Form17.Label1(4).Caption = ""
    Form17.Label1(5).Caption = ""
    Form17.Label1(6).Caption = ""
    Form17.Label1(7).Caption = ""
    Form17.Label1(8).Caption = ""
    Form17.Label1(9).Caption = ""
    Form17.Label1(10).Caption = ""
    Form17.Label1(11).Caption = ""
    Form17.Visible = True
    Form6.Visible = False
End Sub

Private Sub Command2_Click()
    Form17.Command2.Visible = False
    Form17.Command3.Visible = False
    Form17.BackColor = &HFF8080
    Form17.Label2.Caption = "Humanos"
    Form17.Label1(0).Caption = "Os humanos são a raça mais populosa e com maior existência neste mundo. Podem cruzar-se com elfos, orcos e ogres."
    Form17.Label1(1).Caption = "São uma raça inteligente e curiosa e o seu carácter é variado e diverso. Tendem a formar grupos familiares onde mostram"
    Form17.Label1(2).Caption = "grande afecção aos parceiros e crianças. Por norma medem cerca de 1,75 metros e vivem cerca de 80 anos, mas,"
    Form17.Label1(3).Caption = "com sorte, conseguem chegar aos 100 anos. São encontrados a viver em tudo quanto é sitio e ocupam qualquer tipo de"
    Form17.Label1(4).Caption = "profissão independentemente do esforço que fazem para a conseguir. A raça que mais detestam são os ogres porque"
    Form17.Label1(5).Caption = "estes acham a sua carne deliciosa. Também mostram repugnância pelos orcos."
    Form17.Label1(6).Caption = ""
    Form17.Label1(7).Caption = ""
    Form17.Label1(8).Caption = ""
    Form17.Label1(9).Caption = ""
    Form17.Label1(10).Caption = ""
    Form17.Label1(11).Caption = ""
    Form17.Visible = True
    Form6.Visible = False
End Sub

Private Sub Command3_Click()
    Form17.Command2.Visible = True
    Form17.Command3.Visible = False
    Form17.BackColor = &HC000&
    Form17.Label2.Caption = "Elfos"
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
    Form17.Visible = True
    Form6.Visible = False
End Sub

Private Sub Command4_Click()
    Form17.Command2.Visible = False
    Form17.Command3.Visible = False
    Form17.BackColor = &HFFFF00
    Form17.Label2.Caption = "Meio-Elfos"
    Form17.Label1(0).Caption = "São uma raça hibrida, o produto da junção entre humanos e elfos. Tendem a ter uma estatura humana de cerca de 1,75m"
    Form17.Label1(1).Caption = "de altura, mas podem por vezes parecer totalmente elfos ou totalmente humanos, no entanto há características que não"
    Form17.Label1(2).Caption = "falham a esta raça como as orelhas pontiagudas que os elfos possuem e a cor de cabelo dos humanos e os complexos"
    Form17.Label1(3).Caption = "contra os ogres e orcos que ambas raças possuem. Esta raça não é totalmente discriminada entre elfos ou humanos, mas"
    Form17.Label1(4).Caption = "também não são totalmente aceites. Entre elfos são tratados como primos/primas mesmo se forem da mesma terra,"
    Form17.Label1(5).Caption = "o que mostra um pouco de rejeição. Entre humanos não têm problemas em serem aceites. Vivem para cima dos 400 anos."
    Form17.Label1(6).Caption = "Ao contrário dos pais elfos, não são arrogantes nem se mostram superiores às outras raças."
    Form17.Label1(7).Caption = ""
    Form17.Label1(8).Caption = ""
    Form17.Label1(9).Caption = ""
    Form17.Label1(10).Caption = ""
    Form17.Label1(11).Caption = ""
    Form17.Visible = True
    Form6.Visible = False
End Sub

Private Sub Command5_Click()
    Form17.Command2.Visible = False
    Form17.Command3.Visible = False
    Form17.BackColor = &H808000
    Form17.Label2.Caption = "Anões"
    Form17.Label1(0).Caption = "Os anões, apesar da sua estatura, são incrivelmente fortes - muito mais que os humanos e também vivem mais tempo do que"
    Form17.Label1(1).Caption = "estes ultimos. Com 1 metro de altura estas criaturas podem viver até aos 600 anos. Prendados com uma constituição"
    Form17.Label1(2).Caption = "corporal vigorosa, raramente ficam doentes e conseguem suportar temperaturas extremas. São altamente inteligentes e"
    Form17.Label1(3).Caption = "civilizados. Seguem as suas tradições à risca e se alguém as desobedecer ou criticar ficam profundamente ofendidos. Uma"
    Form17.Label1(4).Caption = "das tradições é a barba - é impossível ver um anão que não tenha barba. Mostram grande aptitude matemática e normal-"
    Form17.Label1(5).Caption = "mente seguem as profissões que requerem tal disciplina - arquitectura, engenharia, etc... No entanto, os anões são mais"
    Form17.Label1(6).Caption = "conhecidos pela beleza e precisão das suas esculturas. Esta raça organiza-se em clãs que tem as suas leis e, se quebradas,"
    Form17.Label1(7).Caption = "o castigo é, provavelmente, a morte. Vivem nas montanhas o que provoca conflitos com os ogres, que decidem viver no"
    Form17.Label1(8).Caption = "mesmo sitio. Curiosamente, nunca foram vistas mulheres desta raça. A explicação é que esta longevidade dos anões tem"
    Form17.Label1(9).Caption = "um preço a pagar - as mulheres demoram 10 anos a ter um filho e aquando estas situações ficam num periodo muito frágil"
    Form17.Label1(10).Caption = "ficando assim protegidas a todo o custo aos olhares estranhos."
    Form17.Label1(11).Caption = ""
    Form17.Visible = True
    Form6.Visible = False
End Sub

Private Sub Command6_Click()
    Form17.Command2.Visible = False
    Form17.Command3.Visible = False
    Form17.BackColor = &HFFFF&
    Form17.Label2.Caption = "Gnomos"
    Form17.Label1(0).Caption = "Esta raça de 1 metro de altura e de esperança média de vida de 500 anos não possui nenhum tipo de força extraordinária,"
    Form17.Label1(1).Caption = "ao contrário dos anões. Também são raramente vistos com barba. Mas em contrapartida possuem um senso avançado do"
    Form17.Label1(2).Caption = "cheiro. O faro dos gnomos é preciso e exepcional. Fazem exelentes perfumes. O engraçado dos gnomos é que o faro"
    Form17.Label1(3).Caption = "deles corresponde ao tamanho do nariz que é grande e carnudo. No entanto, a caracteristica mais marcante desta raça é a"
    Form17.Label1(4).Caption = "personalidade. Os gnomos gostam de viajar e aprender novas linguas e experiências. São uma raça ambiciosa, mas esta"
    Form17.Label1(5).Caption = "ambição pode ser justificada pelo desejo de proteger a família. Note-se que os gnomos não são muito capazes de se"
    Form17.Label1(6).Caption = "defenderem o que os motiva a ganhar dinheiro a todo o custo para poderem comprar proteção. Isto não significa que eles"
    Form17.Label1(7).Caption = "sejam cobardes, antes pelo contrário: é sabido que os gnomos fazem vinganças terríveis e sangrentas aos que magoam ou"
    Form17.Label1(8).Caption = "ameaçam a sua familia. Eles vivem durante muito tempo e lembram-se da ofensa com todo o promenor e são capazes de"
    Form17.Label1(9).Caption = "perseguir quem lhes faz mal a vida inteira. Quanto às mulheres dos gnomos: são raramente vistas pois ficam normalmente"
    Form17.Label1(10).Caption = "em casa sob a proteção de um guarda-costas."
    Form17.Label1(11).Caption = ""
    Form17.Visible = True
    Form6.Visible = False
End Sub

Private Sub Command7_Click()
    Form17.Command2.Visible = False
    Form17.Command3.Visible = False
    Form17.BackColor = &H80FF&
    Form17.Label2.Caption = "Halfling"
    Form17.Label1(0).Caption = "Os halflings ou hobbits são a raça das pessoas mais pequenas que existem. Medem no máximo 80 centímetros e vivem"
    Form17.Label1(1).Caption = "cerca de 400 anos. São uma raça vitima da gula, o que os faz gordos e, por vezes mais pesados que os anões que são mais"
    Form17.Label1(2).Caption = "altos. Mas em contrapartida são exelentes cozinheiros e produzem os melhores vinhos. Quanto à aparência: não têm barba,"
    Form17.Label1(3).Caption = "mas possuem grande pelugem no resto do corpo, especialmente nas pernas. A característica mais caricata desta raça são os"
    Form17.Label1(4).Caption = "pés que são grandes, peludos e têm uma pele dura e fléxivel. Com esta característica, os halflings não usam qualquer tipo"
    Form17.Label1(5).Caption = "de calçado, já que os acham desconfortáveis. Os halflings são uma raça gentil e acolhedora que costuma viver em áreas"
    Form17.Label1(6).Caption = "rurais. São gente simples que não desejam mais que uma vida calma. Adoram o conforto e a paz. Trabalham geralmente"
    Form17.Label1(7).Caption = "como agricultores e constroiem a sua própria casa. As mulheres desta raça, ao contrário das dos anões e gnomos, gozam"
    Form17.Label1(8).Caption = "grande liberdade e são atrevidas e faladoras, só não gostam é de viajar muito."
    Form17.Label1(9).Caption = ""
    Form17.Label1(10).Caption = ""
    Form17.Label1(11).Caption = ""
    Form17.Visible = True
    Form6.Visible = False
End Sub

Private Sub Command8_Click()
    Form17.Command2.Visible = True
    Form17.Command3.Visible = False
    Form17.BackColor = &H8080FF
    Form17.Label2.Caption = "Meio-Orcos"
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
    Form17.Visible = True
    Form6.Visible = False
End Sub

Private Sub Command9_Click()
    Form17.Command2.Visible = False
    Form17.Command3.Visible = False
    Form17.BackColor = &HC0&
    Form17.Label2.Caption = "Meio-Ogres"
    Form17.Label1(0).Caption = "Os meio-ogres são uma raça hibrida curiosa. São fruto da junção entre humanos e ogres espantosamente. Não se pode"
    Form17.Label1(1).Caption = "dizer ao certo como isto é possivel, o mais provável é serem concebidos por violação. O único cenário que se pode "
    Form17.Label1(2).Caption = "imaginar é o de um ogre macho violar uma humana, embora seja muito pouco provável a mulher sobreviver a tal ataque"
    Form17.Label1(3).Caption = "sexual. É também dificil de imaginar como um dos pais teria paciencia e/ou vontade de criar esse filho. É que mesmo"
    Form17.Label1(4).Caption = "depois de um ogre satisfazer o sua fome sexual, ele é capaz de também querer satisfazer as suas outras fomes. Por isso era"
    Form17.Label1(5).Caption = "de esperar que depois de violada, a mulher seria comida. Aliás, depois da mulher ser molestada por um ogre até era capaz"
    Form17.Label1(6).Caption = "da morte ser um alívio para ela. De uma maneira ou doutra, tal acto deve ser relacionado com a ingenuidade dos ogres de"
    Form17.Label1(7).Caption = "procriação, visto que eles estavam em extinção. Quanto às características dos meio-ogres, eles medem cerca de 3 metros e"
    Form17.Label1(8).Caption = "são fortes e resistentes como os pais ogres. Vivem cerca de 90 anos. São mais inteligentes e menos agressivosque os pais"
    Form17.Label1(9).Caption = "ogres. Incrivelmente e ao contrário dos ogres amam todas as pessoas mais baixas que eles, especialmente crianças, no"
    Form17.Label1(10).Caption = "entanto com a exepção dos elfos. Curiosamente, só são vistos meio-ogres machos. Tal razão pode dever-se ao facto de que"
    Form17.Label1(11).Caption = "talvez os ogres usem as filhas para propósitos de reprodução. Mais sinistro do que isto não pode haver."
    Form17.Visible = True
    Form6.Visible = False
End Sub
