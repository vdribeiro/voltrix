VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Ra�as"
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
      Caption         =   "V�rios"
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
      Caption         =   "An�es"
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
      Caption         =   "RA�AS"
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
    Form17.Visible = True
    Form6.Visible = False
End Sub

Private Sub Command11_Click()
    Form17.Command2.Visible = False
    Form17.Command3.Visible = False
    Form17.BackColor = &H80&
    Form17.Label2.Caption = "Ogres"
    Form17.Label1(0).Caption = "Uma ra�a em extin��o, os ogres s�o gigantes que medem cerca de 3,80 metros. A raz�o da extin��o e de n�o se saber a"
    Form17.Label1(1).Caption = "esperan�a m�dia de vida desta ra�a deve-se � brutalidade com que s�o tratados. Quando um ogre � visto � imediatamente"
    Form17.Label1(2).Caption = "morto, independentemente da ra�a que o veja. As tend�ncias canibais desta ra�a tamb�m ajudam, comem normalmente os"
    Form17.Label1(3).Caption = "mortos. Se h� estimativa que se possa fazer, eu diria ent�o uns duzentos e tais anos. Os ogres s�o criaturas solit�rias."
    Form17.Label1(4).Caption = "Preferem viver nas montanhas o que normalmente provoca conflitos com os an�es que tamb�m optam por viver no"
    Form17.Label1(5).Caption = "mesmo s�tio. Os ogres s�o exactamente o oposto dos an�es, ou seja, s�o muito pouco inteligentes e desorganizados, o que"
    Form17.Label1(6).Caption = "se torna uma desvantagem quando h� luta de territ�rios entre estas duas ra�as. Apesar da sua tremenda for�a, � raro um"
    Form17.Label1(7).Caption = "ogre ganhar a um an�o. A ra�a que mais detestam s�o os elfos. Quanto ao m�todo de ca�a destes seres - � f�cil e pr�tico -"
    Form17.Label1(8).Caption = "- escolhem a v�tima, escondem-se e quando chegar o momento certo d�o-lhe um murro na cabe�a fracturando-lhe o cr�nio."
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
    Form17.Label1(0).Caption = "Existem duas palavras que definem esta ra�a: brutos e estupidos. Os trolls s�o criaturas com mais de 2 metros de altura e"
    Form17.Label1(1).Caption = "que vivem cerca de 30 anos. S�o uma ra�a um tanto estranha. S�o pregui�osos, gordos e feios. O seu grande peso f�-los"
    Form17.Label1(2).Caption = "lentos, no entanto, esta caracter�stica n�o os impede de serem criaturas incrivelmente fortes. Os trolls gostam sempre de"
    Form17.Label1(3).Caption = "aparecer no momento errado. Cont�do, s�o f�ceis de detectar: fazem muito barulho e cheiram a podrid�o. Estas criaturas,"
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
    Form17.Label2.Caption = "V�rios"
    Form17.Label1(0).Caption = "No percurso do jogo poder�s confrontar-te com outras criaturas tais como lobos, ursos, etc... que poder�o servir como"
    Form17.Label1(1).Caption = "objectos de invoca��o/controle. Conforme a arena podem aparecer estas criaturas em quantidade aleat�ria."
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
    Form17.Label1(0).Caption = "Os humanos s�o a ra�a mais populosa e com maior exist�ncia neste mundo. Podem cruzar-se com elfos, orcos e ogres."
    Form17.Label1(1).Caption = "S�o uma ra�a inteligente e curiosa e o seu car�cter � variado e diverso. Tendem a formar grupos familiares onde mostram"
    Form17.Label1(2).Caption = "grande afec��o aos parceiros e crian�as. Por norma medem cerca de 1,75 metros e vivem cerca de 80 anos, mas,"
    Form17.Label1(3).Caption = "com sorte, conseguem chegar aos 100 anos. S�o encontrados a viver em tudo quanto � sitio e ocupam qualquer tipo de"
    Form17.Label1(4).Caption = "profiss�o independentemente do esfor�o que fazem para a conseguir. A ra�a que mais detestam s�o os ogres porque"
    Form17.Label1(5).Caption = "estes acham a sua carne deliciosa. Tamb�m mostram repugn�ncia pelos orcos."
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
    Form17.Visible = True
    Form6.Visible = False
End Sub

Private Sub Command4_Click()
    Form17.Command2.Visible = False
    Form17.Command3.Visible = False
    Form17.BackColor = &HFFFF00
    Form17.Label2.Caption = "Meio-Elfos"
    Form17.Label1(0).Caption = "S�o uma ra�a hibrida, o produto da jun��o entre humanos e elfos. Tendem a ter uma estatura humana de cerca de 1,75m"
    Form17.Label1(1).Caption = "de altura, mas podem por vezes parecer totalmente elfos ou totalmente humanos, no entanto h� caracter�sticas que n�o"
    Form17.Label1(2).Caption = "falham a esta ra�a como as orelhas pontiagudas que os elfos possuem e a cor de cabelo dos humanos e os complexos"
    Form17.Label1(3).Caption = "contra os ogres e orcos que ambas ra�as possuem. Esta ra�a n�o � totalmente discriminada entre elfos ou humanos, mas"
    Form17.Label1(4).Caption = "tamb�m n�o s�o totalmente aceites. Entre elfos s�o tratados como primos/primas mesmo se forem da mesma terra,"
    Form17.Label1(5).Caption = "o que mostra um pouco de rejei��o. Entre humanos n�o t�m problemas em serem aceites. Vivem para cima dos 400 anos."
    Form17.Label1(6).Caption = "Ao contr�rio dos pais elfos, n�o s�o arrogantes nem se mostram superiores �s outras ra�as."
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
    Form17.Label2.Caption = "An�es"
    Form17.Label1(0).Caption = "Os an�es, apesar da sua estatura, s�o incrivelmente fortes - muito mais que os humanos e tamb�m vivem mais tempo do que"
    Form17.Label1(1).Caption = "estes ultimos. Com 1 metro de altura estas criaturas podem viver at� aos 600 anos. Prendados com uma constitui��o"
    Form17.Label1(2).Caption = "corporal vigorosa, raramente ficam doentes e conseguem suportar temperaturas extremas. S�o altamente inteligentes e"
    Form17.Label1(3).Caption = "civilizados. Seguem as suas tradi��es � risca e se algu�m as desobedecer ou criticar ficam profundamente ofendidos. Uma"
    Form17.Label1(4).Caption = "das tradi��es � a barba - � imposs�vel ver um an�o que n�o tenha barba. Mostram grande aptitude matem�tica e normal-"
    Form17.Label1(5).Caption = "mente seguem as profiss�es que requerem tal disciplina - arquitectura, engenharia, etc... No entanto, os an�es s�o mais"
    Form17.Label1(6).Caption = "conhecidos pela beleza e precis�o das suas esculturas. Esta ra�a organiza-se em cl�s que tem as suas leis e, se quebradas,"
    Form17.Label1(7).Caption = "o castigo �, provavelmente, a morte. Vivem nas montanhas o que provoca conflitos com os ogres, que decidem viver no"
    Form17.Label1(8).Caption = "mesmo sitio. Curiosamente, nunca foram vistas mulheres desta ra�a. A explica��o � que esta longevidade dos an�es tem"
    Form17.Label1(9).Caption = "um pre�o a pagar - as mulheres demoram 10 anos a ter um filho e aquando estas situa��es ficam num periodo muito fr�gil"
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
    Form17.Label1(0).Caption = "Esta ra�a de 1 metro de altura e de esperan�a m�dia de vida de 500 anos n�o possui nenhum tipo de for�a extraordin�ria,"
    Form17.Label1(1).Caption = "ao contr�rio dos an�es. Tamb�m s�o raramente vistos com barba. Mas em contrapartida possuem um senso avan�ado do"
    Form17.Label1(2).Caption = "cheiro. O faro dos gnomos � preciso e exepcional. Fazem exelentes perfumes. O engra�ado dos gnomos � que o faro"
    Form17.Label1(3).Caption = "deles corresponde ao tamanho do nariz que � grande e carnudo. No entanto, a caracteristica mais marcante desta ra�a � a"
    Form17.Label1(4).Caption = "personalidade. Os gnomos gostam de viajar e aprender novas linguas e experi�ncias. S�o uma ra�a ambiciosa, mas esta"
    Form17.Label1(5).Caption = "ambi��o pode ser justificada pelo desejo de proteger a fam�lia. Note-se que os gnomos n�o s�o muito capazes de se"
    Form17.Label1(6).Caption = "defenderem o que os motiva a ganhar dinheiro a todo o custo para poderem comprar prote��o. Isto n�o significa que eles"
    Form17.Label1(7).Caption = "sejam cobardes, antes pelo contr�rio: � sabido que os gnomos fazem vingan�as terr�veis e sangrentas aos que magoam ou"
    Form17.Label1(8).Caption = "amea�am a sua familia. Eles vivem durante muito tempo e lembram-se da ofensa com todo o promenor e s�o capazes de"
    Form17.Label1(9).Caption = "perseguir quem lhes faz mal a vida inteira. Quanto �s mulheres dos gnomos: s�o raramente vistas pois ficam normalmente"
    Form17.Label1(10).Caption = "em casa sob a prote��o de um guarda-costas."
    Form17.Label1(11).Caption = ""
    Form17.Visible = True
    Form6.Visible = False
End Sub

Private Sub Command7_Click()
    Form17.Command2.Visible = False
    Form17.Command3.Visible = False
    Form17.BackColor = &H80FF&
    Form17.Label2.Caption = "Halfling"
    Form17.Label1(0).Caption = "Os halflings ou hobbits s�o a ra�a das pessoas mais pequenas que existem. Medem no m�ximo 80 cent�metros e vivem"
    Form17.Label1(1).Caption = "cerca de 400 anos. S�o uma ra�a vitima da gula, o que os faz gordos e, por vezes mais pesados que os an�es que s�o mais"
    Form17.Label1(2).Caption = "altos. Mas em contrapartida s�o exelentes cozinheiros e produzem os melhores vinhos. Quanto � apar�ncia: n�o t�m barba,"
    Form17.Label1(3).Caption = "mas possuem grande pelugem no resto do corpo, especialmente nas pernas. A caracter�stica mais caricata desta ra�a s�o os"
    Form17.Label1(4).Caption = "p�s que s�o grandes, peludos e t�m uma pele dura e fl�xivel. Com esta caracter�stica, os halflings n�o usam qualquer tipo"
    Form17.Label1(5).Caption = "de cal�ado, j� que os acham desconfort�veis. Os halflings s�o uma ra�a gentil e acolhedora que costuma viver em �reas"
    Form17.Label1(6).Caption = "rurais. S�o gente simples que n�o desejam mais que uma vida calma. Adoram o conforto e a paz. Trabalham geralmente"
    Form17.Label1(7).Caption = "como agricultores e constroiem a sua pr�pria casa. As mulheres desta ra�a, ao contr�rio das dos an�es e gnomos, gozam"
    Form17.Label1(8).Caption = "grande liberdade e s�o atrevidas e faladoras, s� n�o gostam � de viajar muito."
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
    Form17.Visible = True
    Form6.Visible = False
End Sub

Private Sub Command9_Click()
    Form17.Command2.Visible = False
    Form17.Command3.Visible = False
    Form17.BackColor = &HC0&
    Form17.Label2.Caption = "Meio-Ogres"
    Form17.Label1(0).Caption = "Os meio-ogres s�o uma ra�a hibrida curiosa. S�o fruto da jun��o entre humanos e ogres espantosamente. N�o se pode"
    Form17.Label1(1).Caption = "dizer ao certo como isto � possivel, o mais prov�vel � serem concebidos por viola��o. O �nico cen�rio que se pode "
    Form17.Label1(2).Caption = "imaginar � o de um ogre macho violar uma humana, embora seja muito pouco prov�vel a mulher sobreviver a tal ataque"
    Form17.Label1(3).Caption = "sexual. � tamb�m dificil de imaginar como um dos pais teria paciencia e/ou vontade de criar esse filho. � que mesmo"
    Form17.Label1(4).Caption = "depois de um ogre satisfazer o sua fome sexual, ele � capaz de tamb�m querer satisfazer as suas outras fomes. Por isso era"
    Form17.Label1(5).Caption = "de esperar que depois de violada, a mulher seria comida. Ali�s, depois da mulher ser molestada por um ogre at� era capaz"
    Form17.Label1(6).Caption = "da morte ser um al�vio para ela. De uma maneira ou doutra, tal acto deve ser relacionado com a ingenuidade dos ogres de"
    Form17.Label1(7).Caption = "procria��o, visto que eles estavam em extin��o. Quanto �s caracter�sticas dos meio-ogres, eles medem cerca de 3 metros e"
    Form17.Label1(8).Caption = "s�o fortes e resistentes como os pais ogres. Vivem cerca de 90 anos. S�o mais inteligentes e menos agressivosque os pais"
    Form17.Label1(9).Caption = "ogres. Incrivelmente e ao contr�rio dos ogres amam todas as pessoas mais baixas que eles, especialmente crian�as, no"
    Form17.Label1(10).Caption = "entanto com a exep��o dos elfos. Curiosamente, s� s�o vistos meio-ogres machos. Tal raz�o pode dever-se ao facto de que"
    Form17.Label1(11).Caption = "talvez os ogres usem as filhas para prop�sitos de reprodu��o. Mais sinistro do que isto n�o pode haver."
    Form17.Visible = True
    Form6.Visible = False
End Sub
