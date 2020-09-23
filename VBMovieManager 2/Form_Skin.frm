VERSION 5.00
Begin VB.Form Form_Skin 
   BackColor       =   &H00D9D9D9&
   BorderStyle     =   0  'None
   ClientHeight    =   6696
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12432
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.4
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   558
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1036
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Opcao_Over 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2640
      Picture         =   "Form_Skin.frx":0000
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5400
      Width           =   180
   End
   Begin VB.PictureBox Opcao_Normal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2400
      Picture         =   "Form_Skin.frx":01F2
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5400
      Width           =   180
   End
   Begin VB.Image Image_Slider 
      Height          =   216
      Left            =   3000
      Picture         =   "Form_Skin.frx":03E4
      Top             =   4080
      Width           =   132
   End
   Begin VB.Image Botao_Normal_2 
      Height          =   396
      Left            =   8760
      Picture         =   "Form_Skin.frx":06AE
      Top             =   5160
      Width           =   2172
   End
   Begin VB.Image Botao_Play 
      Height          =   264
      Left            =   7680
      Picture         =   "Form_Skin.frx":0E1F
      Top             =   5760
      Width           =   276
   End
   Begin VB.Image Botao_Relatorio 
      Height          =   264
      Left            =   7200
      Picture         =   "Form_Skin.frx":1491
      Top             =   5760
      Width           =   276
   End
   Begin VB.Image Botao_Eliminar 
      Height          =   264
      Left            =   6720
      Picture         =   "Form_Skin.frx":1B03
      Top             =   5760
      Width           =   276
   End
   Begin VB.Image Botao_Editar 
      Height          =   264
      Left            =   6240
      Picture         =   "Form_Skin.frx":2175
      Top             =   5760
      Width           =   276
   End
   Begin VB.Image Botao_Novo 
      Height          =   264
      Left            =   5760
      Picture         =   "Form_Skin.frx":27E7
      Top             =   5760
      Width           =   276
   End
   Begin VB.Image Fundo_Frame_Estrelas 
      Enabled         =   0   'False
      Height          =   435
      Left            =   7680
      Picture         =   "Form_Skin.frx":2E59
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   1380
   End
   Begin VB.Image Caixa_de_Pesquisa 
      Height          =   360
      Left            =   7560
      Picture         =   "Form_Skin.frx":3169
      Top             =   3120
      Width           =   1884
   End
   Begin VB.Image Image_Gratis 
      Height          =   1224
      Left            =   7200
      Picture         =   "Form_Skin.frx":38F0
      Top             =   360
      Width           =   1536
   End
   Begin VB.Image Botao_Instalador 
      Height          =   372
      Left            =   3240
      Picture         =   "Form_Skin.frx":D232
      Top             =   5520
      Width           =   972
   End
   Begin VB.Image Sombra_Letra_Activa 
      Height          =   288
      Left            =   1560
      Picture         =   "Form_Skin.frx":F000
      Top             =   240
      Width           =   300
   End
   Begin VB.Image Sombra_Letra_Over 
      Height          =   288
      Left            =   1080
      Picture         =   "Form_Skin.frx":F3B5
      Top             =   240
      Width           =   300
   End
   Begin VB.Image Sombra_Letra_Normal 
      Height          =   288
      Left            =   600
      Picture         =   "Form_Skin.frx":F6EB
      Top             =   240
      Width           =   300
   End
   Begin VB.Image On_Top_Over 
      Height          =   240
      Left            =   4680
      Picture         =   "Form_Skin.frx":F9A6
      Top             =   960
      Visible         =   0   'False
      Width           =   324
   End
   Begin VB.Image On_Top_Normal 
      Height          =   240
      Left            =   4200
      Picture         =   "Form_Skin.frx":10078
      Top             =   960
      Visible         =   0   'False
      Width           =   324
   End
   Begin VB.Image Estrela_Over_2 
      Height          =   192
      Left            =   1200
      Picture         =   "Form_Skin.frx":1074A
      Top             =   5040
      Width           =   192
   End
   Begin VB.Image Estrela_Normal_2 
      Height          =   192
      Left            =   960
      Picture         =   "Form_Skin.frx":10A8C
      Top             =   5040
      Width           =   192
   End
   Begin VB.Image Image_Sem_Capa 
      Height          =   1536
      Left            =   840
      Picture         =   "Form_Skin.frx":10DCE
      Top             =   1680
      Width           =   1536
   End
   Begin VB.Image Luz_Verde 
      Height          =   192
      Left            =   2400
      Picture         =   "Form_Skin.frx":1CE10
      Top             =   840
      Width           =   192
   End
   Begin VB.Image Luz_Vermelha 
      Height          =   192
      Left            =   2760
      Picture         =   "Form_Skin.frx":1D152
      Top             =   840
      Width           =   192
   End
   Begin VB.Image Icon_Update 
      Enabled         =   0   'False
      Height          =   384
      Left            =   4920
      Picture         =   "Form_Skin.frx":1D494
      Top             =   4920
      Width           =   384
   End
   Begin VB.Image Estrela_Normal 
      Height          =   192
      Left            =   960
      Picture         =   "Form_Skin.frx":1E0D6
      Top             =   4320
      Width           =   192
   End
   Begin VB.Image Estrela_Over 
      Height          =   192
      Left            =   1200
      Picture         =   "Form_Skin.frx":1E418
      Top             =   4320
      Width           =   192
   End
   Begin VB.Image Botao_Pausa_Normal 
      Height          =   192
      Left            =   2040
      Picture         =   "Form_Skin.frx":1E75A
      ToolTipText     =   "Pausa"
      Top             =   4440
      Visible         =   0   'False
      Width           =   168
   End
   Begin VB.Image Botao_Play_Normal 
      Height          =   192
      Left            =   2040
      Picture         =   "Form_Skin.frx":1EA5C
      ToolTipText     =   "Reproduzir"
      Top             =   4080
      Width           =   144
   End
   Begin VB.Image Botao_Seguinte_Normal 
      Height          =   168
      Left            =   2400
      Picture         =   "Form_Skin.frx":1ECDE
      ToolTipText     =   "Faixa seguinte"
      Top             =   4440
      Width           =   132
   End
   Begin VB.Image Botao_Anterior_Normal 
      Height          =   168
      Left            =   2400
      Picture         =   "Form_Skin.frx":1EF18
      ToolTipText     =   "Faixa anteiro"
      Top             =   4080
      Width           =   132
   End
   Begin VB.Image Botao_Over 
      Height          =   396
      Left            =   3600
      Picture         =   "Form_Skin.frx":1F152
      Top             =   2400
      Width           =   732
   End
   Begin VB.Image Botao_Normal 
      Height          =   396
      Left            =   3600
      Picture         =   "Form_Skin.frx":2094C
      Top             =   1800
      Width           =   732
   End
   Begin VB.Image Check_Over 
      Height          =   156
      Left            =   2640
      Picture         =   "Form_Skin.frx":22146
      Top             =   5160
      Width           =   156
   End
   Begin VB.Image Check_Normal 
      Height          =   156
      Left            =   2400
      Picture         =   "Form_Skin.frx":22390
      Top             =   5160
      Width           =   156
   End
   Begin VB.Image Icon_Info 
      Enabled         =   0   'False
      Height          =   384
      Left            =   4920
      Picture         =   "Form_Skin.frx":225DA
      Top             =   1920
      Width           =   384
   End
   Begin VB.Image Icon_Quest 
      Enabled         =   0   'False
      Height          =   384
      Left            =   4920
      Picture         =   "Form_Skin.frx":2321C
      Top             =   2520
      Width           =   384
   End
   Begin VB.Image Icon_Error 
      Enabled         =   0   'False
      Height          =   384
      Left            =   4920
      Picture         =   "Form_Skin.frx":23E5E
      Top             =   3120
      Width           =   384
   End
   Begin VB.Image Icon_Cut 
      Enabled         =   0   'False
      Height          =   384
      Left            =   4920
      Picture         =   "Form_Skin.frx":24AA0
      Top             =   3720
      Width           =   384
   End
   Begin VB.Image Icon_Trash 
      Enabled         =   0   'False
      Height          =   384
      Left            =   4920
      Picture         =   "Form_Skin.frx":256E2
      Top             =   4320
      Width           =   384
   End
   Begin VB.Image Som_Normal 
      Height          =   216
      Left            =   5640
      Picture         =   "Form_Skin.frx":26324
      Top             =   2880
      Width           =   264
   End
   Begin VB.Image Som_Over 
      Height          =   216
      Left            =   5640
      Picture         =   "Form_Skin.frx":2682E
      Top             =   3480
      Width           =   264
   End
End
Attribute VB_Name = "Form_Skin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
