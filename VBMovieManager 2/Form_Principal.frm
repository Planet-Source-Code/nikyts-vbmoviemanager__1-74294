VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form_Principal 
   Appearance      =   0  'Flat
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   0  'None
   Caption         =   "VBMovieManager"
   ClientHeight    =   8625
   ClientLeft      =   0
   ClientTop       =   45
   ClientWidth     =   14385
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "Form_Principal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   575
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   959
   Begin VB.PictureBox Barra_Lateral 
      Appearance      =   0  'Flat
      BackColor       =   &H00F8F8F8&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   0
      ScaleHeight     =   129
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   220
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1920
      Width           =   3300
      Begin MSFlexGridLib.MSFlexGrid Lista_Categorias 
         Height          =   1455
         Left            =   345
         TabIndex        =   2
         Top             =   720
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   2566
         _Version        =   393216
         Rows            =   20
         BackColor       =   16777215
         ForeColor       =   0
         BackColorFixed  =   15790578
         ForeColorFixed  =   0
         BackColorSel    =   14200408
         ForeColorSel    =   16777215
         BackColorBkg    =   16777215
         GridColor       =   16777215
         GridColorFixed  =   16777215
         Redraw          =   -1  'True
         FocusRect       =   0
         GridLines       =   0
         GridLinesFixed  =   0
         ScrollBars      =   2
         SelectionMode   =   1
         BorderStyle     =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   24
         X2              =   200
         Y1              =   32
         Y2              =   32
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Categorias"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Left            =   360
         TabIndex        =   10
         Top             =   240
         Width           =   1050
      End
      Begin VB.Image Fundo_Down_Barra_Lateral 
         Enabled         =   0   'False
         Height          =   135
         Left            =   0
         Picture         =   "Form_Principal.frx":57E2
         Top             =   600
         Width           =   3300
      End
      Begin VB.Image Fundo_Centro_Barra_Lateral 
         Enabled         =   0   'False
         Height          =   645
         Left            =   0
         Picture         =   "Form_Principal.frx":5B6F
         Top             =   120
         Width           =   3300
      End
      Begin VB.Image Fundo_Top_Barra_Lateral 
         Enabled         =   0   'False
         Height          =   135
         Left            =   0
         Picture         =   "Form_Principal.frx":5FA0
         Top             =   0
         Width           =   3300
      End
   End
   Begin VB.PictureBox Pic_Gadgets 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1224
      Left            =   0
      MouseIcon       =   "Form_Principal.frx":6490
      MousePointer    =   99  'Custom
      Picture         =   "Form_Principal.frx":BC72
      ScaleHeight     =   82
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   102
      TabIndex        =   80
      Top             =   360
      Width           =   1536
   End
   Begin VB.PictureBox Barra_Menu 
      Appearance      =   0  'Flat
      BackColor       =   &H00252525&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   552
      Left            =   0
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   801
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   360
      Width           =   12012
      Begin VB.Label Label_Gadgets 
         AutoSize        =   -1  'True
         BackColor       =   &H002E2E2E&
         BackStyle       =   0  'Transparent
         Caption         =   "Centro de programas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   5040
         TabIndex        =   85
         Top             =   240
         Width           =   1845
      End
      Begin VB.Label Label_Preferencias 
         AutoSize        =   -1  'True
         BackColor       =   &H002E2E2E&
         BackStyle       =   0  'Transparent
         Caption         =   "Preferências"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   192
         Left            =   7320
         TabIndex        =   83
         Top             =   240
         Width           =   1068
      End
      Begin VB.Label Label_Sobre 
         AutoSize        =   -1  'True
         BackColor       =   &H002E2E2E&
         BackStyle       =   0  'Transparent
         Caption         =   "Sobre"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   192
         Left            =   10080
         TabIndex        =   84
         Top             =   240
         Width           =   516
      End
      Begin VB.Label Label_Suporte 
         AutoSize        =   -1  'True
         BackColor       =   &H002E2E2E&
         BackStyle       =   0  'Transparent
         Caption         =   "Suporte"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   192
         Left            =   8880
         TabIndex        =   82
         Top             =   240
         Width           =   672
      End
      Begin VB.Image Fundo_Barra_Menu 
         Enabled         =   0   'False
         Height          =   690
         Left            =   0
         Picture         =   "Form_Principal.frx":155B4
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.PictureBox Barra_Botoes 
      Appearance      =   0  'Flat
      BackColor       =   &H00252525&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   576
      Left            =   120
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   833
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   8040
      Width           =   12492
      Begin VB.PictureBox Botao_Novo 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   0
         Picture         =   "Form_Principal.frx":159B9
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   108
         Width           =   345
      End
      Begin VB.PictureBox Botao_Editar 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1920
         Picture         =   "Form_Principal.frx":1602B
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   108
         Width           =   345
      End
      Begin VB.PictureBox Botao_Eliminar 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3600
         Picture         =   "Form_Principal.frx":1669D
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   108
         Width           =   345
      End
      Begin VB.PictureBox Botao_Relatorio 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   5520
         Picture         =   "Form_Principal.frx":16D0F
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   108
         Width           =   345
      End
      Begin VB.PictureBox Botao_Play 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   7320
         Picture         =   "Form_Principal.frx":17381
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   108
         Width           =   345
      End
      Begin VB.Image Botao_Redimensionar 
         Height          =   165
         Left            =   10800
         Picture         =   "Form_Principal.frx":179F3
         Top             =   240
         Width           =   135
      End
      Begin VB.Image Image13 
         Height          =   165
         Left            =   12600
         Picture         =   "Form_Principal.frx":17B69
         Top             =   885
         Width           =   135
      End
      Begin VB.Label Label_Novo 
         AutoSize        =   -1  'True
         BackColor       =   &H002E2E2E&
         BackStyle       =   0  'Transparent
         Caption         =   "Adicionar"
         ForeColor       =   &H00FFFFFF&
         Height          =   192
         Left            =   480
         TabIndex        =   79
         Top             =   180
         Width           =   792
      End
      Begin VB.Label Label_Editar 
         AutoSize        =   -1  'True
         BackColor       =   &H002E2E2E&
         BackStyle       =   0  'Transparent
         Caption         =   "Editar"
         ForeColor       =   &H00FFFFFF&
         Height          =   192
         Left            =   2400
         TabIndex        =   78
         Top             =   180
         Width           =   492
      End
      Begin VB.Label Label_Eliminar 
         AutoSize        =   -1  'True
         BackColor       =   &H002E2E2E&
         BackStyle       =   0  'Transparent
         Caption         =   "Eliminar"
         ForeColor       =   &H00FFFFFF&
         Height          =   192
         Left            =   4080
         TabIndex        =   77
         Top             =   180
         Width           =   696
      End
      Begin VB.Label Label_Play 
         AutoSize        =   -1  'True
         BackColor       =   &H002E2E2E&
         BackStyle       =   0  'Transparent
         Caption         =   "Visualizar"
         ForeColor       =   &H00FFFFFF&
         Height          =   192
         Left            =   7800
         TabIndex        =   76
         Top             =   180
         Width           =   828
      End
      Begin VB.Label Label_Relatorio 
         AutoSize        =   -1  'True
         BackColor       =   &H002E2E2E&
         BackStyle       =   0  'Transparent
         Caption         =   "Relatório"
         ForeColor       =   &H00FFFFFF&
         Height          =   192
         Left            =   6000
         TabIndex        =   75
         Top             =   180
         Width           =   768
      End
      Begin VB.Image Fundo_Barra_Botoes 
         Enabled         =   0   'False
         Height          =   615
         Left            =   -120
         Picture         =   "Form_Principal.frx":17CDF
         Top             =   0
         Width           =   600
      End
   End
   Begin VB.PictureBox Frame_Capa 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3132
      Left            =   120
      ScaleHeight     =   209
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   220
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4320
      Width           =   3300
      Begin VB.Image Image_Capa 
         Enabled         =   0   'False
         Height          =   1920
         Left            =   420
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1920
      End
      Begin VB.Image Image_Sem_Capa 
         Enabled         =   0   'False
         Height          =   1920
         Left            =   420
         Picture         =   "Form_Principal.frx":17FF6
         Top             =   480
         Width           =   1920
      End
      Begin VB.Image Fundo_Centro_Barra_Lateral_2 
         Enabled         =   0   'False
         Height          =   645
         Left            =   0
         Picture         =   "Form_Principal.frx":24038
         Top             =   240
         Width           =   3300
      End
      Begin VB.Image Fundo_Top_Barra_Lateral_2 
         Enabled         =   0   'False
         Height          =   135
         Left            =   0
         Picture         =   "Form_Principal.frx":24469
         Top             =   0
         Width           =   3300
      End
      Begin VB.Image Fundo_Down_Barra_Lateral_2 
         Enabled         =   0   'False
         Height          =   135
         Left            =   0
         Picture         =   "Form_Principal.frx":24959
         Top             =   960
         Width           =   3300
      End
   End
   Begin VB.PictureBox frame_Centro 
      Appearance      =   0  'Flat
      BackColor       =   &H00F8F8F8&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   3720
      ScaleHeight     =   297
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   537
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2040
      Width           =   8055
      Begin MSFlexGridLib.MSFlexGrid Lista_Filmes 
         Height          =   1335
         Left            =   2160
         TabIndex        =   0
         Top             =   240
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   2355
         _Version        =   393216
         Rows            =   1
         Cols            =   9
         BackColor       =   16777215
         ForeColor       =   0
         BackColorFixed  =   15790578
         ForeColorFixed  =   0
         BackColorSel    =   14200408
         ForeColorSel    =   16777215
         BackColorBkg    =   16777215
         GridColor       =   14737632
         GridColorFixed  =   14737632
         Redraw          =   -1  'True
         FocusRect       =   0
         GridLinesFixed  =   1
         SelectionMode   =   1
         BorderStyle     =   0
         Appearance      =   0
      End
      Begin VB.Image Fundo_Frame_Centro_Top_Esq 
         Enabled         =   0   'False
         Height          =   135
         Left            =   0
         Picture         =   "Form_Principal.frx":24CE6
         Top             =   0
         Width           =   1800
      End
      Begin VB.Image Fundo_Frame_Centro_Top_Dir 
         Enabled         =   0   'False
         Height          =   135
         Left            =   5520
         Picture         =   "Form_Principal.frx":250A2
         Top             =   0
         Width           =   1800
      End
      Begin VB.Image Fundo_Frame_Centro_Down_Esq 
         Enabled         =   0   'False
         Height          =   135
         Left            =   0
         Picture         =   "Form_Principal.frx":2545F
         Top             =   2040
         Width           =   1800
      End
      Begin VB.Image Fundo_Frame_Centro_Down_Dir 
         Enabled         =   0   'False
         Height          =   135
         Left            =   5520
         Picture         =   "Form_Principal.frx":25768
         Top             =   1920
         Width           =   1800
      End
      Begin VB.Shape Shape_Frame_Centro 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   1215
         Left            =   0
         Top             =   240
         Width           =   1545
      End
   End
   Begin VB.PictureBox Barra_Detalhes 
      Appearance      =   0  'Flat
      BackColor       =   &H00252525&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   456
      Left            =   120
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   833
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   7560
      Width           =   12492
      Begin VB.PictureBox Frame_Estrelas 
         Appearance      =   0  'Flat
         BackColor       =   &H002E2E2E&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   10440
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   92
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   0
         Width           =   1380
         Begin VB.TextBox Text_Classificacao 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1200
            TabIndex        =   14
            Text            =   "1"
            Top             =   105
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Image Image1 
            Height          =   240
            Left            =   15
            Picture         =   "Form_Principal.frx":25A6F
            Top             =   105
            Width           =   240
         End
         Begin VB.Image Image2 
            Height          =   240
            Left            =   255
            Picture         =   "Form_Principal.frx":25DB1
            Top             =   105
            Width           =   240
         End
         Begin VB.Image Image3 
            Height          =   240
            Left            =   495
            Picture         =   "Form_Principal.frx":260F3
            Top             =   105
            Width           =   240
         End
         Begin VB.Image Image4 
            Height          =   240
            Left            =   735
            Picture         =   "Form_Principal.frx":26435
            Top             =   105
            Width           =   240
         End
         Begin VB.Image Image5 
            Height          =   240
            Left            =   975
            Picture         =   "Form_Principal.frx":26777
            Top             =   105
            Width           =   240
         End
         Begin VB.Image Fundo_Frame_Estrelas 
            Enabled         =   0   'False
            Height          =   348
            Left            =   0
            Picture         =   "Form_Principal.frx":26AB9
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1440
         End
      End
      Begin VB.Label Label_Total 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H002E2E2E&
         BackStyle       =   0  'Transparent
         Caption         =   "Total de registos: nenhum registro encontrado"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   3990
      End
      Begin VB.Image Fundo_Barra_Detalhes 
         Enabled         =   0   'False
         Height          =   435
         Left            =   0
         Picture         =   "Form_Principal.frx":26DC9
         Top             =   0
         Width           =   600
      End
   End
   Begin VB.PictureBox Barra_ControlBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00252525&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   372
      Left            =   0
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   769
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   11535
      Begin VB.Image Image6 
         Enabled         =   0   'False
         Height          =   420
         Left            =   60
         Picture         =   "Form_Principal.frx":270D9
         Top             =   0
         Width           =   390
      End
      Begin VB.Label Label_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00272727&
         BackStyle       =   0  'Transparent
         Caption         =   "VBMovieManager - Nikyts software"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   525
         TabIndex        =   7
         Top             =   150
         Width           =   3390
      End
      Begin VB.Image Botao_Fechar 
         Height          =   225
         Left            =   8880
         Picture         =   "Form_Principal.frx":279DB
         ToolTipText     =   "Fechar"
         Top             =   120
         Width           =   240
      End
      Begin VB.Image Botao_Maximizar 
         Height          =   225
         Left            =   8520
         Picture         =   "Form_Principal.frx":27CED
         ToolTipText     =   "Maximizar"
         Top             =   120
         Width           =   240
      End
      Begin VB.Image Botao_Restaurar 
         Height          =   225
         Left            =   8160
         Picture         =   "Form_Principal.frx":27FFF
         ToolTipText     =   "Restaurar"
         Top             =   120
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Botao_Minimizar 
         Height          =   225
         Left            =   7800
         Picture         =   "Form_Principal.frx":28311
         ToolTipText     =   "Minimizar"
         Top             =   120
         Width           =   240
      End
      Begin VB.Image Fundo_Barra_ControlBox 
         Enabled         =   0   'False
         Height          =   465
         Left            =   0
         Picture         =   "Form_Principal.frx":28623
         Top             =   0
         Width           =   585
      End
   End
   Begin VB.PictureBox Barra_Ferramentas 
      Appearance      =   0  'Flat
      BackColor       =   &H00202020&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   408
      Left            =   0
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   945
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   912
      Width           =   14175
      Begin VB.PictureBox Pic_Letra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   26
         Left            =   11520
         Picture         =   "Form_Principal.frx":28968
         ScaleHeight     =   24
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   67
         TabStop         =   0   'False
         ToolTipText     =   "Ver todos os registos"
         Top             =   36
         Width           =   375
         Begin VB.Label Label_Letra 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H002E2E2E&
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   26
            Left            =   120
            TabIndex        =   68
            ToolTipText     =   "Ver todos os registos"
            Top             =   75
            Width           =   165
         End
      End
      Begin VB.PictureBox Pic_Letra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   25
         Left            =   11160
         Picture         =   "Form_Principal.frx":28C23
         ScaleHeight     =   24
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   36
         Width           =   375
         Begin VB.Label Label_Letra 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H002E2E2E&
            BackStyle       =   0  'Transparent
            Caption         =   "Z"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   25
            Left            =   120
            TabIndex        =   66
            Top             =   75
            Width           =   135
         End
      End
      Begin VB.PictureBox Pic_Letra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   24
         Left            =   10800
         Picture         =   "Form_Principal.frx":28EDE
         ScaleHeight     =   24
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   36
         Width           =   375
         Begin VB.Label Label_Letra 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H002E2E2E&
            BackStyle       =   0  'Transparent
            Caption         =   "Y"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   24
            Left            =   120
            TabIndex        =   64
            Top             =   75
            Width           =   135
         End
      End
      Begin VB.PictureBox Pic_Letra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   23
         Left            =   10440
         Picture         =   "Form_Principal.frx":29199
         ScaleHeight     =   24
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   36
         Width           =   375
         Begin VB.Label Label_Letra 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H002E2E2E&
            BackStyle       =   0  'Transparent
            Caption         =   "X"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   23
            Left            =   120
            TabIndex        =   62
            Top             =   75
            Width           =   135
         End
      End
      Begin VB.PictureBox Pic_Letra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   22
         Left            =   10080
         Picture         =   "Form_Principal.frx":29454
         ScaleHeight     =   24
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   36
         Width           =   375
         Begin VB.Label Label_Letra 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H002E2E2E&
            BackStyle       =   0  'Transparent
            Caption         =   "W"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   22
            Left            =   90
            TabIndex        =   60
            Top             =   75
            Width           =   195
         End
      End
      Begin VB.PictureBox Pic_Letra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   21
         Left            =   9720
         Picture         =   "Form_Principal.frx":2970F
         ScaleHeight     =   24
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   36
         Width           =   375
         Begin VB.Label Label_Letra 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H002E2E2E&
            BackStyle       =   0  'Transparent
            Caption         =   "V"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   21
            Left            =   120
            TabIndex        =   58
            Top             =   75
            Width           =   135
         End
      End
      Begin VB.PictureBox Pic_Letra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   20
         Left            =   9360
         Picture         =   "Form_Principal.frx":299CA
         ScaleHeight     =   24
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   36
         Width           =   375
         Begin VB.Label Label_Letra 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H002E2E2E&
            BackStyle       =   0  'Transparent
            Caption         =   "U"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   20
            Left            =   120
            TabIndex        =   56
            Top             =   75
            Width           =   135
         End
      End
      Begin VB.PictureBox Pic_Letra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   19
         Left            =   9000
         Picture         =   "Form_Principal.frx":29C85
         ScaleHeight     =   24
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   36
         Width           =   375
         Begin VB.Label Label_Letra 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H002E2E2E&
            BackStyle       =   0  'Transparent
            Caption         =   "T"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   19
            Left            =   120
            TabIndex        =   54
            Top             =   75
            Width           =   135
         End
      End
      Begin VB.PictureBox Pic_Letra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   18
         Left            =   8640
         Picture         =   "Form_Principal.frx":29F40
         ScaleHeight     =   24
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   36
         Width           =   375
         Begin VB.Label Label_Letra 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H002E2E2E&
            BackStyle       =   0  'Transparent
            Caption         =   "S"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   18
            Left            =   120
            TabIndex        =   52
            Top             =   75
            Width           =   135
         End
      End
      Begin VB.PictureBox Pic_Letra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   17
         Left            =   8280
         Picture         =   "Form_Principal.frx":2A1FB
         ScaleHeight     =   24
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   36
         Width           =   375
         Begin VB.Label Label_Letra 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H002E2E2E&
            BackStyle       =   0  'Transparent
            Caption         =   "R"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   17
            Left            =   120
            TabIndex        =   50
            Top             =   75
            Width           =   135
         End
      End
      Begin VB.PictureBox Pic_Letra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   16
         Left            =   7920
         Picture         =   "Form_Principal.frx":2A4B6
         ScaleHeight     =   24
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   36
         Width           =   375
         Begin VB.Label Label_Letra 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H002E2E2E&
            BackStyle       =   0  'Transparent
            Caption         =   "Q"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   16
            Left            =   105
            TabIndex        =   48
            Top             =   75
            Width           =   165
         End
      End
      Begin VB.PictureBox Pic_Letra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   15
         Left            =   7560
         Picture         =   "Form_Principal.frx":2A771
         ScaleHeight     =   24
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   36
         Width           =   375
         Begin VB.Label Label_Letra 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H002E2E2E&
            BackStyle       =   0  'Transparent
            Caption         =   "P"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   15
            Left            =   120
            TabIndex        =   46
            Top             =   75
            Width           =   135
         End
      End
      Begin VB.PictureBox Pic_Letra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   14
         Left            =   7200
         Picture         =   "Form_Principal.frx":2AA2C
         ScaleHeight     =   24
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   36
         Width           =   375
         Begin VB.Label Label_Letra 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H002E2E2E&
            BackStyle       =   0  'Transparent
            Caption         =   "O"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   14
            Left            =   105
            TabIndex        =   44
            Top             =   75
            Width           =   165
         End
      End
      Begin VB.PictureBox Pic_Letra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   13
         Left            =   6840
         Picture         =   "Form_Principal.frx":2ACE7
         ScaleHeight     =   24
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   36
         Width           =   375
         Begin VB.Label Label_Letra 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H002E2E2E&
            BackStyle       =   0  'Transparent
            Caption         =   "N"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   13
            Left            =   120
            TabIndex        =   42
            Top             =   75
            Width           =   135
         End
      End
      Begin VB.PictureBox Pic_Letra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   12
         Left            =   6480
         Picture         =   "Form_Principal.frx":2AFA2
         ScaleHeight     =   24
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   36
         Width           =   375
         Begin VB.Label Label_Letra 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H002E2E2E&
            BackStyle       =   0  'Transparent
            Caption         =   "M"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   12
            Left            =   105
            TabIndex        =   40
            Top             =   75
            Width           =   165
         End
      End
      Begin VB.PictureBox Pic_Letra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   11
         Left            =   6120
         Picture         =   "Form_Principal.frx":2B25D
         ScaleHeight     =   24
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   36
         Width           =   375
         Begin VB.Label Label_Letra 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H002E2E2E&
            BackStyle       =   0  'Transparent
            Caption         =   "L"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   11
            Left            =   150
            TabIndex        =   38
            Top             =   75
            Width           =   105
         End
      End
      Begin VB.PictureBox Pic_Letra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   10
         Left            =   5760
         Picture         =   "Form_Principal.frx":2B518
         ScaleHeight     =   24
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   36
         Width           =   375
         Begin VB.Label Label_Letra 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H002E2E2E&
            BackStyle       =   0  'Transparent
            Caption         =   "K"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   36
            Top             =   75
            Width           =   135
         End
      End
      Begin VB.PictureBox Pic_Letra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   9
         Left            =   5400
         Picture         =   "Form_Principal.frx":2B7D3
         ScaleHeight     =   24
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   36
         Width           =   375
         Begin VB.Label Label_Letra 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H002E2E2E&
            BackStyle       =   0  'Transparent
            Caption         =   "J"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   9
            Left            =   135
            TabIndex        =   34
            Top             =   75
            Width           =   105
         End
      End
      Begin VB.PictureBox Pic_Letra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   8
         Left            =   5040
         Picture         =   "Form_Principal.frx":2BA8E
         ScaleHeight     =   24
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   36
         Width           =   375
         Begin VB.Label Label_Letra 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H002E2E2E&
            BackStyle       =   0  'Transparent
            Caption         =   "I"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   8
            Left            =   135
            TabIndex        =   32
            Top             =   75
            Width           =   105
         End
      End
      Begin VB.PictureBox Pic_Letra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   7
         Left            =   4680
         Picture         =   "Form_Principal.frx":2BD49
         ScaleHeight     =   24
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   36
         Width           =   375
         Begin VB.Label Label_Letra 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H002E2E2E&
            BackStyle       =   0  'Transparent
            Caption         =   "H"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   30
            Top             =   75
            Width           =   135
         End
      End
      Begin VB.PictureBox Pic_Letra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   6
         Left            =   4290
         Picture         =   "Form_Principal.frx":2C004
         ScaleHeight     =   24
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   36
         Width           =   375
         Begin VB.Label Label_Letra 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H002E2E2E&
            BackStyle       =   0  'Transparent
            Caption         =   "G"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   6
            Left            =   105
            TabIndex        =   28
            Top             =   75
            Width           =   165
         End
      End
      Begin VB.PictureBox Pic_Letra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   5
         Left            =   3915
         Picture         =   "Form_Principal.frx":2C2BF
         ScaleHeight     =   24
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   36
         Width           =   375
         Begin VB.Label Label_Letra 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H002E2E2E&
            BackStyle       =   0  'Transparent
            Caption         =   "F"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   5
            Left            =   135
            TabIndex        =   26
            Top             =   75
            Width           =   105
         End
      End
      Begin VB.PictureBox Pic_Letra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   4
         Left            =   3540
         Picture         =   "Form_Principal.frx":2C57A
         ScaleHeight     =   24
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   36
         Width           =   375
         Begin VB.Label Label_Letra 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H002E2E2E&
            BackStyle       =   0  'Transparent
            Caption         =   "E"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   24
            Top             =   75
            Width           =   135
         End
      End
      Begin VB.PictureBox Pic_Letra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   3
         Left            =   3165
         Picture         =   "Form_Principal.frx":2C835
         ScaleHeight     =   24
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   36
         Width           =   375
         Begin VB.Label Label_Letra 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H002E2E2E&
            BackStyle       =   0  'Transparent
            Caption         =   "D"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   3
            Left            =   105
            TabIndex        =   22
            Top             =   75
            Width           =   165
         End
      End
      Begin VB.PictureBox Pic_Letra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   2
         Left            =   2790
         Picture         =   "Form_Principal.frx":2CAF0
         ScaleHeight     =   24
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   36
         Width           =   375
         Begin VB.Label Label_Letra 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H002E2E2E&
            BackStyle       =   0  'Transparent
            Caption         =   "C"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   2
            Left            =   105
            TabIndex        =   20
            Top             =   75
            Width           =   165
         End
      End
      Begin VB.PictureBox Pic_Letra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   1
         Left            =   2415
         Picture         =   "Form_Principal.frx":2CDAB
         ScaleHeight     =   24
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   36
         Width           =   375
         Begin VB.Label Label_Letra 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H002E2E2E&
            BackStyle       =   0  'Transparent
            Caption         =   "B"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   18
            Top             =   75
            Width           =   135
         End
      End
      Begin VB.PictureBox Pic_Letra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   0
         Left            =   2115
         Picture         =   "Form_Principal.frx":2D066
         ScaleHeight     =   24
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   36
         Width           =   375
         Begin VB.Label Label_Letra 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H002E2E2E&
            BackStyle       =   0  'Transparent
            Caption         =   "A"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   16
            Top             =   75
            Width           =   135
         End
      End
      Begin VB.PictureBox Barra_Pesquisa 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   12000
         Picture         =   "Form_Principal.frx":2D321
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   133
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Width           =   1992
         Begin VB.TextBox Text_Pesquisa 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   240
            TabIndex        =   5
            Top             =   84
            Width           =   1452
         End
      End
      Begin VB.Image Image_Inicio_Barra_Ferramentas 
         Enabled         =   0   'False
         Height          =   510
         Left            =   1875
         Picture         =   "Form_Principal.frx":2DAA8
         Top             =   0
         Width           =   420
      End
      Begin VB.Image Fundo_Barra_Ferramentas 
         Enabled         =   0   'False
         Height          =   510
         Left            =   0
         Picture         =   "Form_Principal.frx":2DDDC
         Top             =   0
         Width           =   1530
      End
   End
   Begin VB.Shape Shape_Contorno 
      BorderColor     =   &H00404040&
      Height          =   615
      Left            =   12120
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "Form_Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   MOVIE MANAGER
'   COPYRIGHT © 2010 Nikyts software ™  -  INFORMÁTICA & TECNOLOGIA
'   WWW.NIKYTS.COM.SAPO.PT
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Declaração das variáveis
Option Explicit
Dim bMoveFrom As Boolean, LastPoint As POINTAPI

'Ajusta o Form para sempre exibir a barra de tarefas do windows, full screen
Private Const SPI_GETWORKAREA = 48
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
Private Declare Function SystemParametersInfo Lib "user32" _
Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal CX As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'Variavel para verificar a janela do formulário
Dim Tela_Cheia As Boolean

'Variáveis para a base de dados
Dim Criterio As String
Dim Categoria_Selecionada As String
Dim i As Integer
Public Cnn_Filmes As New ADODB.Connection
Public Rs_Filmes As New ADODB.Recordset

'Variável para criar relatório
Dim cRelatório As New clsRelatórioHTML

'Variável saber quando deve carregar a lista de categorias
Public Iniciando As Boolean

'Variável para guardar o index da letra selecionada e activa
Public Letra_selecionada As Integer
Public Letra_Activa As Integer

'API para abrir web
Private Const SW_NORMAL = 1
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Redimensionar formulário
'Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal CX As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Const WM_NCLBUTTONDOWN = &HA1
Const HTBOTTOMRIGHT = 17
Const HTCAPTION = 2

'Variável para guardar as dimensões do formulário
Dim ALtura_Formulario As Long
Dim Largura_Formulario As Long

Public Function PosFormRelativeTaskBar(F As Form)
    'Função para ao maximizar o form seja visivel a barra do windows iniciar
    'Colocar o WindowsState=0 normal
    Dim WindowRect As RECT
    SystemParametersInfo SPI_GETWORKAREA, 0, WindowRect, 0
    
    SetWindowPos hWnd, 0, WindowRect.Left, WindowRect.Top, WindowRect.Right - WindowRect.Left, WindowRect.Bottom - WindowRect.Top, 0
    F.Top = WindowRect.Bottom * Screen.TwipsPerPixelY - F.Height
    F.Left = WindowRect.Right * Screen.TwipsPerPixelX - F.Width
End Function

Public Sub Conectar()
    'Procedimento para conectar á base de dados
    On Error GoTo Corrige_Erro
    Cnn_Filmes.CursorLocation = adUseClient
    Cnn_Filmes.Open "provider=microsoft.jet.oledb.4.0;persist security info = false; data source = " & App.path & "\Data\Biblioteca.mdb;Jet " & "OLEDB:Database Password=nikita;"
    
Exit Sub
Corrige_Erro:
Select Case Err.Number
    Case "-2147467259"
        Mensagem_de_Aviso "Erro", "A base de dados do programa não foi encontrada."
        End
End Select
End Sub

Public Sub Verifica_Rs_Filmes()
    'Procedimento para verificar o estado
   If Rs_Filmes.State = 1 Then Rs_Filmes.Close
End Sub

Private Sub Barra_Botoes_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Chamar o procedimento
    Repor_Imagens
End Sub

Private Sub Barra_ControlBox_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Capturar a posição de x e y
    Dim POINT As POINTAPI
    GetCursorPos POINT
    LastPoint.x = POINT.x
    LastPoint.y = POINT.y
    bMoveFrom = True
End Sub

Private Sub Barra_ControlBox_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Chamar o procedimento
    Repor_Imagens
    
    'Mover o formulário e obter a posição de x e y
    If Me.WindowState = 0 And Tela_Cheia = False Then
        Dim iDX As Long, iDY As Long
        Dim POINT As POINTAPI
        If Not bMoveFrom Then Exit Sub
        GetCursorPos POINT
        iDX& = (POINT.x - LastPoint.x) * iTPPX&
        iDY& = (POINT.y - LastPoint.y) * iTPPY&
        LastPoint.x = POINT.x
        LastPoint.y = POINT.y
        Me.Move Me.Left + iDX&, Me.Top + iDY&
    End If
End Sub

Private Sub Barra_ControlBox_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Largar o formulário para a posição final
    bMoveFrom = False
End Sub

Private Sub Barra_ControlBox_DblClick()
    'Atalho para
    Label_Titulo_DblClick
End Sub

Private Sub Barra_Detalhes_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Chamar o procedimento
    Repor_Imagens
End Sub

Private Sub Barra_Ferramentas_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Chamar o procedimento
    Repor_Imagens
End Sub

Private Sub Barra_Lateral_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Chamar o procedimento
    Repor_Imagens
End Sub

Private Sub Barra_Menu_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Chamar o procedimentor
    Repor_Imagens
End Sub

Private Sub Botao_Editar_Click()
    'Editar dados existentes
    Repor_Imagens
    
    If Lista_Filmes.Rows = 1 Then Exit Sub
    
    With Form_Dados
        .Text_Id.Text = Lista_Filmes.TextMatrix(Lista_Filmes.Row, 0)
        .Text_Nome.Text = Lista_Filmes.TextMatrix(Lista_Filmes.Row, 1)
        .Text_Categoria.Text = Lista_Filmes.TextMatrix(Lista_Filmes.Row, 2)
        .Text_Tipo.Text = Lista_Filmes.TextMatrix(Lista_Filmes.Row, 3)
        .Text_Classificacao.Text = Lista_Filmes.TextMatrix(Lista_Filmes.Row, 4)
        
        .Verificar_Estrelas
        
        .Text_Actores.Text = Lista_Filmes.TextMatrix(Lista_Filmes.Row, 5)
        .Text_Observacoes.Text = Lista_Filmes.TextMatrix(Lista_Filmes.Row, 6)
        .Text_Directorio.Text = Lista_Filmes.TextMatrix(Lista_Filmes.Row, 7)
        .Text_Capa.Text = Lista_Filmes.TextMatrix(Lista_Filmes.Row, 8)
        'Carrega imagem da capa, caso exista
        .Carregar_Capa
        
        .Mode = "editar"
        .Label_Titulo.Caption = "Editar registo"
        .Show vbModal
     End With
End Sub

Private Sub Botao_Editar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Animar a label
    If Label_Editar.FontUnderline = True Then Exit Sub
    Label_Novo.FontUnderline = False
    Label_Editar.FontUnderline = True
    Label_Eliminar.FontUnderline = False
    Label_Relatorio.FontUnderline = False
    Label_Play.FontUnderline = False
End Sub

Private Sub Botao_Eliminar_Click()
    'Eliminar registos da base de dados
    Repor_Imagens
    
    If Lista_Filmes.Rows = 1 Then Exit Sub
        
    Mensagem_de_Aviso "Questão", "Deseja eliminar o seguinte registo: " & Lista_Filmes.TextMatrix(Lista_Filmes.Row, 1)
    
    'Verificar o resultado da resposta
    If Resposta = "Sim" Then 'Confirmar eliminação do registo
        Cnn_Filmes.Execute "Delete From Tabela_Filmes Where Id = '" & Lista_Filmes.TextMatrix(Lista_Filmes.Row, 0) & "'"
        Rs_Filmes.Requery
        
        'Verifica se o arquivo de existe na pasta 'Capas'
        If Dir$(App.path & "\Covers\" & Lista_Filmes.TextMatrix(Lista_Filmes.Row, 0) & ".jpg") <> "" Then
            Kill (App.path & "\Covers\" & Lista_Filmes.TextMatrix(Lista_Filmes.Row, 0) & ".jpg")
        End If
    
        'Alterar o valor da variável para poder alctualizar a lista de pesquisa por categorias
        Formatar_Lista_Categorias
        Iniciando = True
        
        'Carrega os dados
        Preenche_Lista
        Limpar_Campos
    End If
End Sub

Private Sub Botao_Eliminar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Animar a label
    If Label_Eliminar.FontUnderline = True Then Exit Sub
    Label_Novo.FontUnderline = False
    Label_Editar.FontUnderline = False
    Label_Eliminar.FontUnderline = True
    Label_Relatorio.FontUnderline = False
    Label_Play.FontUnderline = False
End Sub

Public Sub Botao_Fechar_Click()
    'Fechar o programa
    Rs_Filmes.Close
    Form_Wmp.Wmp.Controls.stop
    Unload Form_Wmp
    Unload Me
    End
End Sub

Private Sub Botao_Maximizar_Click()
    'Maximixar formulário
    PosFormRelativeTaskBar Me
    Tela_Cheia = True
    Form_Opcoes.Text_Tela_Cheia.Text = "True"
'    Form_Opcoes.Actualizar_Opcoes
    Botao_Maximizar.Visible = False
    Botao_Restaurar.Visible = True
End Sub

Private Sub Botao_Novo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Animar a label
    If Label_Novo.FontUnderline = True Then Exit Sub
    Label_Novo.FontUnderline = True
    Label_Editar.FontUnderline = False
    Label_Eliminar.FontUnderline = False
    Label_Relatorio.FontUnderline = False
    Label_Play.FontUnderline = False
End Sub

Private Sub Botao_Play_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Animar a label
    If Label_Play.FontUnderline = True Then Exit Sub
    Label_Novo.FontUnderline = False
    Label_Editar.FontUnderline = False
    Label_Eliminar.FontUnderline = False
    Label_Relatorio.FontUnderline = False
    Label_Play.FontUnderline = True
End Sub

Private Sub Botao_Redimensionar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  'Redimensionar o formulário conforme as dimensões pretendidas
    If Button = vbLeftButton Then
        If Tela_Cheia = False Then
            ReleaseCapture
            SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, 0
            
            'Verificar se não exedeu os limites
            If Me.Height < 8616 Then
                Me.Height = "8616"
            End If
        
            If Me.Width < 14385 Then
                Me.Width = "14385"
            End If
            ALtura_Formulario = Me.Height
            Largura_Formulario = Me.Width
            Desenhar_Formulario
        End If
    End If
End Sub

Private Sub Botao_Redimensionar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Alterar o mousepointer
    If Tela_Cheia = True Then
        Botao_Redimensionar.MousePointer = vbDefault
    Else
        Botao_Redimensionar.MousePointer = 8
    End If
End Sub


Private Sub Botao_Relatorio_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Animar a label
    If Label_Relatorio.FontUnderline = True Then Exit Sub
    Label_Novo.FontUnderline = False
    Label_Editar.FontUnderline = False
    Label_Eliminar.FontUnderline = False
    Label_Relatorio.FontUnderline = True
    Label_Play.FontUnderline = False
End Sub

Private Sub Botao_Restaurar_Click()
    'Restaurar janela
    With Me
        .Height = 10000
        .Width = 14295
        .Top = (Screen.Height - Me.Height) / 2
        .Left = (Screen.Width - Me.Width) / 2
    End With
    Tela_Cheia = False
    Form_Opcoes.Text_Tela_Cheia.Text = "False"
'    Form_Opcoes.Actualizar_Opcoes
    Botao_Maximizar.Visible = True
    Botao_Restaurar.Visible = False
End Sub

Private Sub Botao_Minimizar_Click()
    'Minimizar o formulário
    Me.WindowState = 1
End Sub

Private Sub Botao_Novo_Click()
    'Adicionar um novo registo
    Repor_Imagens
    
    With Form_Dados
        .Label_Titulo.Caption = "Novo registo"
        
        'Adicionar novo registo
        If Rs_Filmes.RecordCount = 0 Then
            .Text_Id.Text = "0"
        Else
            Rs_Filmes.MoveLast
            .Text_Id.Text = Rs_Filmes(0).Value + 1
        End If
        .Mode = "novo"
        .Show vbModal
    End With
End Sub

Private Sub Botao_Relatorio_Click()
    'Gerar relatório através dos dados da Base de dados
    Repor_Imagens
    If Lista_Filmes.Rows = 1 Then Exit Sub
    
    Verifica_Rs_Filmes
    Rs_Filmes.Open "select * from Tabela_Filmes order by Id Asc", Cnn_Filmes
    
    Dim arrayTítulosColunas(6) As String
    Dim arrayAlinhamentoTítulos(6) As Integer
    Dim arrayAlinhamentoDetalhes(6) As Integer
    Dim arrayNegritoDetalhes(6) As Boolean
    Dim arrayItálicoDetalhes(6) As Boolean
    Dim arrayTamanhoColunas(6) As Integer
    Dim arrayFiltros(1) As String
    Dim arrayCamposDetalhe(6) As String

    cRelatório.Arquivo = App.path & "\Listagem de Filmes.htm"
    cRelatório.Data = Date
    cRelatório.NomeRelatório = App.ProductName & " " & App.Minor & "." & App.Revision & "." & App.Major
    cRelatório.Empresa = "Nikyts software"
    cRelatório.LinkEmpresa = "www.nikyts.com"
    cRelatório.Desenvolvedor = "Nelson do Carmo"
    cRelatório.LinkDesenvolvedor = "nikyts@hotmail.com"
    cRelatório.UsaFonteCourierNoCorpo = False
    cRelatório.MostraMensagemDeFinalização = True

    arrayTítulosColunas(1) = "Nome do Filme"
    arrayTítulosColunas(2) = "Categoria"
    arrayTítulosColunas(3) = "Tipo"
    arrayTítulosColunas(4) = "Classificação"
    arrayTítulosColunas(5) = "Actores"
    arrayTítulosColunas(6) = "Observações"

    arrayAlinhamentoTítulos(1) = cRelatório.AlinharEsquerda
    arrayAlinhamentoTítulos(2) = cRelatório.AlinharEsquerda
    arrayAlinhamentoTítulos(3) = cRelatório.AlinharEsquerda
    arrayAlinhamentoTítulos(4) = cRelatório.AlinharEsquerda
    arrayAlinhamentoTítulos(5) = cRelatório.AlinharEsquerda
    arrayAlinhamentoTítulos(6) = cRelatório.AlinharEsquerda

    arrayAlinhamentoDetalhes(1) = cRelatório.AlinharEsquerda
    arrayAlinhamentoDetalhes(2) = cRelatório.AlinharEsquerda
    arrayAlinhamentoDetalhes(3) = cRelatório.AlinharEsquerda
    arrayAlinhamentoDetalhes(4) = cRelatório.AlinharEsquerda
    arrayAlinhamentoDetalhes(5) = cRelatório.AlinharEsquerda
    arrayAlinhamentoDetalhes(6) = cRelatório.AlinharEsquerda

    arrayNegritoDetalhes(1) = False
    arrayNegritoDetalhes(2) = False
    arrayNegritoDetalhes(3) = False
    arrayNegritoDetalhes(4) = False
    arrayNegritoDetalhes(5) = False
    arrayNegritoDetalhes(6) = False

    arrayItálicoDetalhes(1) = False
    arrayItálicoDetalhes(2) = False
    arrayItálicoDetalhes(3) = False
    arrayItálicoDetalhes(4) = False
    arrayItálicoDetalhes(5) = False
    arrayItálicoDetalhes(6) = False

    arrayTamanhoColunas(1) = 20
    arrayTamanhoColunas(2) = 10
    arrayTamanhoColunas(3) = 10
    arrayTamanhoColunas(4) = 10
    arrayTamanhoColunas(5) = 15
    arrayTamanhoColunas(6) = 30

    arrayFiltros(1) = "Listagem de filmes (Total de registos " & Rs_Filmes.RecordCount & ")"

    cRelatório.TítulosColunas = arrayTítulosColunas
    cRelatório.AlinhamentoTítulos = arrayAlinhamentoTítulos
    cRelatório.AlinhamentoDetalhes = arrayAlinhamentoDetalhes
    cRelatório.NegritoDetalhes = arrayNegritoDetalhes
    cRelatório.ItálicoDetalhes = arrayItálicoDetalhes
    cRelatório.TamanhoColunas = arrayTamanhoColunas
    cRelatório.Filtros = arrayFiltros

    cRelatório.HoraInício = Time
    cRelatório.ImprimeCabeçalho
    While Not Rs_Filmes.EOF
        arrayCamposDetalhe(1) = Rs_Filmes("Video")
        arrayCamposDetalhe(2) = Rs_Filmes("Categoria")
        arrayCamposDetalhe(3) = Rs_Filmes("Tipo")
        arrayCamposDetalhe(4) = Rs_Filmes("Classificacao")
        arrayCamposDetalhe(5) = Rs_Filmes("Actores")
        arrayCamposDetalhe(6) = Rs_Filmes("Observacoes")
        cRelatório.CamposDetalhe = arrayCamposDetalhe
        cRelatório.ImprimeDetalhe
        Rs_Filmes.MoveNext
    Wend
    arrayNegritoDetalhes(1) = False
    arrayNegritoDetalhes(2) = False
    arrayNegritoDetalhes(3) = False
    arrayNegritoDetalhes(4) = False
    arrayNegritoDetalhes(5) = False
    arrayNegritoDetalhes(6) = False
    arrayCamposDetalhe(1) = ""
    arrayCamposDetalhe(2) = ""
    arrayCamposDetalhe(3) = ""
    arrayCamposDetalhe(4) = ""
    arrayCamposDetalhe(5) = ""
    arrayCamposDetalhe(6) = ""
    cRelatório.NegritoDetalhes = arrayNegritoDetalhes
    cRelatório.CamposDetalhe = arrayCamposDetalhe
    cRelatório.ImprimeDetalhe
    
    cRelatório.HoraFim = Time
    cRelatório.ImprimeRodapé
    Set cRelatório = Nothing
    
    'Abrir o relatório
    If Dir(App.path & "\Listagem de Filmes.htm") <> "" Then
        ShellExecute hWnd, vbNullString, App.path & "\Listagem de Filmes.htm", vbNullString, vbNullString, SW_SHOWMAXIMIZED
    Else
        Mensagem_de_Aviso "Erro", "O arquivo do Relatório não foi encontrado."
    End If
End Sub

Private Sub Form_Load()
    'Chamar o procedimento
    Desenhar_Formulario
    Ver_Opcoes
    
    Iniciando = True
    Letra_selecionada = -1
    Letra_Activa = -1
    
    iTPPX& = Screen.TwipsPerPixelX
    iTPPY& = Screen.TwipsPerPixelY
    
    'Ligação á base de dados
    Conectar
    Verifica_Rs_Filmes
    Rs_Filmes.Open "select * from Tabela_Filmes order by Id Asc", Cnn_Filmes
    
    'Chamar procedimentos
    Formatar_Lista_Categorias
    Preenche_Lista
    Limpar_Campos
    
    ALtura_Formulario = "8616"
    Largura_Formulario = "14385"
    
    'Actualizar a localização da pasta do programa
    Dim Localizacao_Ficheiro_Preferencias As String
    Localizacao_Ficheiro_Preferencias = App.path & "\Options\Properties.ini"
    Call WriteINI("Directório", "Localização do programa", App.path & "\", (Localizacao_Ficheiro_Preferencias))
        
    'Verificar possiveis actualizações do programa
    If Form_Opcoes.Check_Actualizacoes.Value = 1 Then
        Verificar_Actualizacoes
    End If
End Sub

Private Sub Verificar_Actualizacoes()
    On Error GoTo Corrige_Erro
    Dim servidor As XMLHTTP60: Set servidor = New XMLHTTP60
    servidor.Open "GET", "http://www.nikyts.com/vbmoviemanager/" & "verificarversao.asp?", False
    servidor.send 'envia o pedido para o servidor
    
    'Verificar os dados acesso
    If Not InStr(servidor.responseText, "HTTP Error") > 0 Then
        If servidor.readyState = 4 And servidor.Status = 200 Then
            'Ler os dados do xml referente aos dados do perfil do utilizador
            Dim versao_actual, nova_versao As String
            versao_actual = App.Major & App.Minor & App.Revision
            nova_versao = servidor.responseText 'CInt(responseText)
            
            'Verificar se existem versões novas
            If versao_actual < nova_versao Then
                Form_Actualizacoes.Show 'Indica que há uma nova versão caso a minha versao seja diferente á versão que está no servidor
            End If
        End If
    End If
    
    
Exit Sub
Corrige_Erro:
Select Case Err.Number
    Case -2146697211
        Mensagem_de_Aviso "Erro", "Ocorreu um erro ao tentar conectar-se ao servidor." & vbNewLine & "Verifique a sua ligação à internet."
        
    Case Else
        'Correção de outros erros que poderão surgir
        Mensagem_de_Aviso "Erro", "Ocorreu um erro durante a execução do programa." & vbNewLine & "Erro nº: " & Err.Number & vbNewLine & "Descrição: " & Err.Description
End Select
End Sub

Public Sub Preenche_Lista()
    'Procedimento para carregar a lista de filmes com os dados da base de dados
    'On Error Resume Next
    With Lista_Filmes
        If Rs_Filmes.RecordCount = 0 Then
            .Clear
            .Rows = 1
            Label_Total.Caption = "Total de registos: nenhum registo encontrado"
'''            Formatar_Lista_Categorias
            Formatar_Lista_Filmes
            Exit Sub

        Else
        'Variáveis para contas as categorias correspondentes
        Dim contar_accao, contar_animacao, contar_aventura, contar_comedia, contar_crime, contar_documentarios, _
        contar_desporto, contar_drama, contar_faroeste, contar_ficcao, contar_guerra, contar_musical, _
        contar_policial, contar_romance, contar_serie, contar_suspense, contar_terror, contar_outras As Integer
        
        'Iniciar os cantadores a '0'
        contar_accao = 0
        contar_animacao = 0
        contar_aventura = 0
        contar_comedia = 0
        contar_crime = 0
        contar_documentarios = 0
        contar_desporto = 0
        contar_drama = 0
        contar_faroeste = 0
        contar_ficcao = 0
        contar_guerra = 0
        contar_musical = 0
        contar_policial = 0
        contar_romance = 0
        contar_serie = 0
        contar_suspense = 0
        contar_terror = 0
        contar_outras = 0
        
            .Clear
            .Rows = 1
            i = 1
            Do While Not Rs_Filmes.EOF
                .Rows = Rs_Filmes.RecordCount + 1
                If Rs_Filmes(0).Value <> "" Then .TextMatrix(i, 0) = Rs_Filmes(0).Value
                If Rs_Filmes(1).Value <> "" Then .TextMatrix(i, 1) = Rs_Filmes(1).Value
                If Rs_Filmes(2).Value <> "" Then .TextMatrix(i, 2) = Rs_Filmes(2).Value
                    Select Case Rs_Filmes(2)
                        Case "Acção"
                            contar_accao = contar_accao + 1
                        Case "Animação"
                            contar_animacao = contar_animacao + 1
                        Case "Aventura"
                            contar_aventura = contar_aventura + 1
                        Case "Comédia"
                            contar_comedia = contar_comedia + 1
                        Case "Crime"
                            contar_crime = contar_crime + 1
                        Case "Documentário"
                            contar_documentarios = contar_documentarios + 1
                        Case "Desporto"
                            contar_desporto = contar_desporto + 1
                        Case "Drama"
                            contar_drama = contar_drama + 1
                        Case "Faroeste"
                            contar_faroeste = contar_faroeste + 1
                        Case "Ficção cientifica"
                            contar_ficcao = contar_ficcao + 1
                        Case "Guerra"
                            contar_guerra = contar_guerra + 1
                        Case "Musical"
                            contar_musical = contar_musical + 1
                        Case "Policial"
                            contar_policial = contar_policial + 1
                        Case "Romance"
                            contar_romance = contar_romance + 1
                        Case "Série"
                            contar_serie = contar_serie + 1
                        Case "Suspense"
                            contar_suspense = contar_suspense + 1
                        Case "Terror"
                            contar_terror = contar_terror + 1
                        Case "Outra"
                            contar_outras = contar_outras + 1
                    End Select
                If Rs_Filmes(3).Value <> "" Then .TextMatrix(i, 3) = Rs_Filmes(3).Value
                If Rs_Filmes(4).Value <> "" Then .TextMatrix(i, 4) = Rs_Filmes(4).Value
                If Rs_Filmes(5).Value <> "" Then .TextMatrix(i, 5) = Rs_Filmes(5).Value
                If Rs_Filmes(6).Value <> "" Then .TextMatrix(i, 6) = Rs_Filmes(6).Value
                If Rs_Filmes(7).Value <> "" Then .TextMatrix(i, 7) = Rs_Filmes(7).Value
                If Rs_Filmes(8).Value <> "" Then .TextMatrix(i, 8) = Rs_Filmes(8).Value
                i = i + 1
                Rs_Filmes.MoveNext
            Loop
            
            If Iniciando = True Then
                Lista_Categorias.TextMatrix(1, 1) = "Todas (" & Rs_Filmes.RecordCount & ")"
                Lista_Categorias.TextMatrix(2, 1) = "Acção (" & contar_accao & ")"
                Lista_Categorias.TextMatrix(3, 1) = "Animação (" & contar_animacao & ")"
                Lista_Categorias.TextMatrix(4, 1) = "Aventura (" & contar_aventura & ")"
                Lista_Categorias.TextMatrix(5, 1) = "Comédia (" & contar_comedia & ")"
                Lista_Categorias.TextMatrix(6, 1) = "Crime (" & contar_crime & ")"
                Lista_Categorias.TextMatrix(7, 1) = "Documentário (" & contar_documentarios & ")"
                Lista_Categorias.TextMatrix(8, 1) = "Desporto (" & contar_desporto & ")"
                Lista_Categorias.TextMatrix(9, 1) = "Drama (" & contar_drama & ")"
                Lista_Categorias.TextMatrix(10, 1) = "Faroeste (" & contar_faroeste & ")"
                Lista_Categorias.TextMatrix(11, 1) = "Ficção cientifica (" & contar_ficcao & ")"
                Lista_Categorias.TextMatrix(12, 1) = "Guerra (" & contar_guerra & ")"
                Lista_Categorias.TextMatrix(13, 1) = "Musical (" & contar_musical & ")"
                Lista_Categorias.TextMatrix(14, 1) = "Policial (" & contar_policial & ")"
                Lista_Categorias.TextMatrix(15, 1) = "Romance (" & contar_romance & ")"
                Lista_Categorias.TextMatrix(16, 1) = "Série (" & contar_serie & ")"
                Lista_Categorias.TextMatrix(17, 1) = "Suspense (" & contar_suspense & ")"
                Lista_Categorias.TextMatrix(18, 1) = "Terror (" & contar_terror & ")"
                Lista_Categorias.TextMatrix(19, 1) = "Outra (" & contar_outras & ")"
                Iniciando = False
            End If
            
            If Lista_Filmes.Rows = 2 Then
                Label_Total.Caption = "Total de registos: 1 registo encontrado"
            Else
                Label_Total.Caption = "Total de registos: " & Lista_Filmes.Rows - 1 & " registos encontrados"
            End If
            Formatar_Lista_Filmes
        End If
    End With
End Sub

Public Sub Desenhar_Formulario()
    'Procedimento para construir o formulario, ajustando os objectos
    On Error GoTo CORRIGIR_ERRO
    With Shape_Contorno
        .Height = Me.ScaleHeight
        .Top = 0
        .Width = Me.ScaleWidth
        .Left = 0
    End With
    
    With Barra_ControlBox
        .Height = Fundo_Barra_ControlBox.Height
        .Top = 0
        .Width = Me.ScaleWidth - 2
        .Left = 1
    End With
    
    With Fundo_Barra_ControlBox
        .Stretch = True
        .Top = 0
        .Width = Barra_ControlBox.Width
        .Left = 0
    End With
    
    With Botao_Fechar
        .Top = 8
        .Left = Barra_ControlBox.Width - .Width - 4
    End With
    
    With Botao_Maximizar
        .Top = 8
        .Left = Botao_Fechar.Left - .Width - 2
    End With
    
    With Botao_Restaurar
        .Top = 8
        .Left = Botao_Maximizar.Left
    End With
    
    With Botao_Minimizar
        .Top = 8
        .Left = Botao_Maximizar.Left - .Width - 2
    End With
    
    With Label_Titulo
        .Top = (Barra_ControlBox.ScaleHeight - .Height) / 2
        .Left = 35
    End With
    
    With Barra_Menu
        .Height = Fundo_Barra_Menu.Height
        .Top = Barra_ControlBox.Top + Barra_ControlBox.ScaleHeight
        .Width = Barra_ControlBox.ScaleWidth
        .Left = Barra_ControlBox.Left
    End With
    
    With Fundo_Barra_Menu
        .Stretch = True
        .Top = 0
        .Width = Barra_Menu.Width
        .Left = 0
    End With
    
    With Barra_Ferramentas
        .Height = Fundo_Barra_Ferramentas.Height
        .Top = Barra_Menu.Top + Barra_Menu.ScaleHeight
        .Width = Barra_ControlBox.ScaleWidth
        .Left = Barra_ControlBox.Left
    End With
    
    With Fundo_Barra_Ferramentas
        .Stretch = True
        .Top = 0
        .Width = Barra_Ferramentas.Width
        .Left = 0
    End With
    
    With Barra_Pesquisa
        .Height = Form_Skin.Caixa_de_Pesquisa.Height
        .Top = 0
        .Left = Barra_Menu.ScaleWidth - .ScaleWidth
        .Width = Form_Skin.Caixa_de_Pesquisa.Width
    End With
    
    With Text_Pesquisa
        .Height = 17
        .Top = 7
        .Width = 121
        .Left = 20
    End With
    
    With Barra_Botoes
        .Height = Fundo_Barra_Botoes.Height
        .Top = Me.ScaleHeight - .ScaleHeight - 1
        .Width = Me.ScaleWidth - 2
        .Left = 1
    End With
    
    With Fundo_Barra_Botoes
        .Stretch = True
        .Top = 0
        .Width = Barra_Botoes.Width
        .Left = 0
    End With
    
    With Barra_Detalhes
        .Height = Fundo_Barra_Detalhes.Height
        .Top = Barra_Botoes.Top - .ScaleHeight
        .Width = Barra_Botoes.ScaleWidth
        .Left = Barra_Botoes.Left
    End With
    
    With Fundo_Barra_Detalhes
        .Stretch = True
        .Width = Barra_Detalhes.Width
        .Left = 0
    End With
    
    With Pic_Gadgets
        .Height = Form_Skin.Image_Gratis.Height
        .Top = Barra_ControlBox.Top + Barra_ControlBox.ScaleHeight
        .Width = Form_Skin.Image_Gratis.Width
        .Left = 1
    End With
    
    With Botao_Redimensionar
        .Left = Barra_Detalhes.ScaleWidth - .Width
    End With
    
    With Label_Gadgets
        .Top = (Barra_Menu.ScaleHeight - .Height) / 2
        .Left = Barra_Menu.ScaleWidth - .Width - 8
    End With
    
    With Label_Sobre
        .Top = (Barra_Menu.ScaleHeight - .Height) / 2
        .Left = Label_Gadgets.Left - .Width - 16
    End With
    
    With Label_Suporte
        .Top = (Barra_Menu.ScaleHeight - .Height) / 2
        .Left = Label_Sobre.Left - .Width - 16
    End With
    
    With Label_Preferencias
        .Top = (Barra_Menu.ScaleHeight - .Height) / 2
        .Left = Label_Suporte.Left - .Width - 16
    End With
    
    '----------------------------------------------------------------------------------------------------------
    With Barra_Lateral
        .Top = Barra_Ferramentas.Top + Barra_Ferramentas.Height + 20
        .Height = Me.ScaleHeight - Barra_ControlBox.ScaleHeight - Barra_Menu.ScaleHeight - Barra_Ferramentas.ScaleHeight - Barra_Detalhes.ScaleHeight - Barra_Botoes.ScaleHeight - Frame_Capa.ScaleHeight - 10 - 20 - 20
        .Width = Fundo_Top_Barra_Lateral.Width
        .Left = 1 + 10
    End With
    
    With Fundo_Top_Barra_Lateral
        .Top = 0
        .Left = 0
    End With
    
    With Lista_Categorias
        '.Height = Barra_Lateral.ScaleHeight - .Top - 3
        .Height = Barra_Lateral.ScaleHeight - 4
        .Top = 2
        .Width = Barra_Lateral.ScaleWidth - 4
        .Left = 2
    End With
    
    With Fundo_Centro_Barra_Lateral
        .Stretch = True
        .Top = Fundo_Top_Barra_Lateral.Top + Fundo_Top_Barra_Lateral.Height
        .Height = Barra_Lateral.ScaleHeight - Fundo_Top_Barra_Lateral.Height - Fundo_Down_Barra_Lateral.Height
        .Left = 0
    End With
    
    With Fundo_Down_Barra_Lateral
        .Top = Barra_Lateral.ScaleHeight - .Height
        .Left = 0
    End With
    
    With frame_Centro
        .Height = Barra_Lateral.Height + Frame_Capa.ScaleHeight + 10
        .Top = Barra_Lateral.Top
        .Left = Barra_Lateral.Left + Barra_Lateral.Width + 10
        .Width = Me.ScaleWidth - Barra_Lateral.ScaleWidth - 2 - 10 - 10 - 10
    End With
    
    With Shape_Frame_Centro
        .Height = frame_Centro.ScaleHeight
        .Top = 0
        .Width = frame_Centro.ScaleWidth
        .Left = 0
    End With
    
    With Fundo_Frame_Centro_Top_Esq
        .Top = 0
        .Left = 0
    End With
    
    With Fundo_Frame_Centro_Top_Dir
        .Top = 0
        .Left = frame_Centro.ScaleWidth - .Width
    End With
        
    With Fundo_Frame_Centro_Down_Esq
        .Top = frame_Centro.ScaleHeight - .Height
        .Left = 0
    End With
    
    With Fundo_Frame_Centro_Down_Dir
        .Top = frame_Centro.ScaleHeight - .Height
        .Left = frame_Centro.ScaleWidth - .Width
    End With
    
    With Lista_Filmes
        .Height = frame_Centro.ScaleHeight - 4
        .Top = 2
        .Width = frame_Centro.ScaleWidth - 4
        .Left = 2
    End With
    
    With Frame_Capa
        .Height = Image_Capa.Height + (2 * Image_Capa.Top)
        .Top = Barra_Lateral.Top + Barra_Lateral.ScaleHeight + 10
        .Width = Fundo_Top_Barra_Lateral_2.Width
        .Left = Barra_Lateral.Left
    End With
    
    With Image_Sem_Capa
        .Top = (Frame_Capa.ScaleHeight - .Height) / 2
        .Left = (Frame_Capa.ScaleWidth - .Width) / 2
    End With

    With Image_Capa
        .Top = (Frame_Capa.ScaleHeight - .Height) / 2
        .Left = (Frame_Capa.ScaleWidth - .Width) / 2
    End With
    
    With Fundo_Top_Barra_Lateral_2
        .Top = 0
        .Left = 0
    End With
    
    With Fundo_Centro_Barra_Lateral_2
        .Stretch = True
        .Top = Fundo_Top_Barra_Lateral_2.Top + Fundo_Top_Barra_Lateral_2.Height
        .Height = Frame_Capa.ScaleHeight - Fundo_Top_Barra_Lateral_2.Height - Fundo_Down_Barra_Lateral_2.Height
        .Left = 0
    End With
    
    With Fundo_Down_Barra_Lateral_2
        .Top = Frame_Capa.ScaleHeight - .Height
        .Left = 0
    End With
    
    With Frame_Estrelas
        .Height = Form_Skin.Fundo_Frame_Estrelas.Height
        .Top = 0
        .Width = Form_Skin.Fundo_Frame_Estrelas.Width
        .Left = Barra_Detalhes.ScaleWidth - .ScaleWidth
    End With
    
    'Ajustar as letras de pesquisa
    Dim x, posicao_inicial As Integer
    posicao_inicial = Pic_Letra(0).Left
    Pic_Letra(0).Height = Form_Skin.Sombra_Letra_Normal.Height
    Pic_Letra(0).Top = ((Barra_Ferramentas.ScaleHeight - Pic_Letra(0).ScaleHeight) - 4) / 2
    Pic_Letra(0).Width = Form_Skin.Sombra_Letra_Normal.Width
    Label_Letra(0).Top = (Pic_Letra(0).ScaleHeight - Label_Letra(0).Height) / 2
    Label_Letra(0).Left = (Pic_Letra(0).ScaleWidth - Label_Letra(0).Width) / 2
    
    For x = 1 To 26
        Pic_Letra(x).Height = Form_Skin.Sombra_Letra_Normal.Height
        Pic_Letra(x).Top = ((Barra_Ferramentas.ScaleHeight - Pic_Letra(x).ScaleHeight) - 4) / 2
        Pic_Letra(x).Width = Form_Skin.Sombra_Letra_Normal.Width
        Pic_Letra(x).Left = posicao_inicial + Pic_Letra(0).ScaleWidth
        posicao_inicial = posicao_inicial + Pic_Letra(0).ScaleWidth
        Label_Letra(x).Top = (Pic_Letra(x).ScaleHeight - Label_Letra(x).Height) / 2
        Label_Letra(x).Left = (Pic_Letra(x).ScaleWidth - Label_Letra(x).Width) / 2
    Next x
        
    'Chamar o procedimento para alinhar as frame sconsoante as opções do programa
    Posicionar_Frames
    
    With Label_Total
        .Top = (Barra_Detalhes.ScaleHeight - .Height) / 2
        .Left = 10
    End With
    
    'Botões da barra de detalhes
    With Botao_Novo
        .Height = Form_Skin.Botao_Novo.Height
        .Top = (Barra_Botoes.ScaleHeight - .ScaleHeight) / 2
        .Width = Form_Skin.Botao_Novo.Width
        .Left = 8
        Label_Novo.Left = .Left + .ScaleWidth + 5
        Label_Novo.Top = (Barra_Botoes.ScaleHeight - Label_Play.Height) / 2
    End With
    
    With Botao_Editar
        .Height = Form_Skin.Botao_Novo.Height
        .Top = (Barra_Botoes.ScaleHeight - .ScaleHeight) / 2
        .Width = Form_Skin.Botao_Novo.Width
        .Left = Label_Novo.Left + Label_Novo.Width + 40
        Label_Editar.Left = .Left + .ScaleWidth + 5
        Label_Editar.Top = (Barra_Botoes.ScaleHeight - Label_Play.Height) / 2
    End With
    
    With Botao_Eliminar
        .Height = Form_Skin.Botao_Novo.Height
        .Top = (Barra_Botoes.ScaleHeight - .ScaleHeight) / 2
        .Width = Form_Skin.Botao_Novo.Width
        .Left = Label_Editar.Left + Label_Editar.Width + 40
        Label_Eliminar.Left = .Left + .ScaleWidth + 5
        Label_Eliminar.Top = (Barra_Botoes.ScaleHeight - Label_Play.Height) / 2
    End With
    
    With Botao_Relatorio
        .Height = Form_Skin.Botao_Novo.Height
        .Top = (Barra_Botoes.ScaleHeight - .ScaleHeight) / 2
        .Width = Form_Skin.Botao_Novo.Width
        .Left = Label_Eliminar.Left + Label_Eliminar.Width + 40
        Label_Relatorio.Left = .Left + .ScaleWidth + 5
        Label_Relatorio.Top = (Barra_Botoes.ScaleHeight - Label_Play.Height) / 2
    End With
        
    With Botao_Play
        .Height = Form_Skin.Botao_Novo.Height
        .Top = (Barra_Botoes.ScaleHeight - .ScaleHeight) / 2
        .Width = Form_Skin.Botao_Novo.Width
        .Left = Label_Relatorio.Left + Label_Relatorio.Width + 40
        Label_Play.Left = .Left + .ScaleWidth + 5
        Label_Play.Top = (Barra_Botoes.ScaleHeight - Label_Play.Height) / 2
    End With
    
    With Botao_Redimensionar
        .Top = Barra_Botoes.ScaleHeight - .Height
        .Left = Barra_Botoes.ScaleWidth - .Width
    End With
    
Exit Sub
CORRIGIR_ERRO:
    Me.Height = 8616
    Me.Width = 14385
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Chamar o procedimento
    Repor_Imagens
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    'Chamar o procedimento
    Desenhar_Formulario
End Sub

Private Sub Botao_Play_Click()
    'On Error Resume Next
    Repor_Imagens
    If Lista_Filmes.Rows = 1 Then Exit Sub
    
    If Lista_Filmes.TextMatrix(Lista_Filmes.Row, 7) = "" Then
        Mensagem_de_Aviso "Informação", "Este registo não contem a localização do filme."
        Exit Sub
        
    Else
        If Form_Opcoes.Opcao_Programa.Value = True Then
            Load Form_Wmp
            With Form_Wmp
                'Reproduzir o som
                .Label_Faixa.Caption = "A reproduzir: [" & Lista_Filmes.TextMatrix(Lista_Filmes.Row, 1) & "]"
                
                .Musica_Play = False
                .Label_Duracao.Caption = "00:00" & "  |  "
                .Tempo_Estimado.Caption = "00:00"
                .VideoDuration = 0
    '            .Slide_Som.Top = Text_Slide_Som.Text
    '            Verificar_Volume
                
                .Slide.Left = 0
                .Wmp.URL = Lista_Filmes.TextMatrix(Lista_Filmes.Row, 7)
                .Timer_Slider_Video.Enabled = True
                .Botao_Play_Click
                .Show
            End With
            
        Else
            Dim Video As Long
            Video = Shell("explorer.exe " & Lista_Filmes.TextMatrix(Lista_Filmes.Row, 7), vbMaximizedFocus)
        End If
    End If
End Sub

Private Sub Frame_Capa_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Chamar o procedimento
    Repor_Imagens
End Sub

Private Sub frame_Centro_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Chamar o procedimento
    Repor_Imagens
End Sub

Private Sub Label_Editar_Click()
    'Atalho para
    Botao_Editar_Click
End Sub

Private Sub Label_Editar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Animar a label
    If Label_Editar.FontUnderline = True Then Exit Sub
    Label_Novo.FontUnderline = False
    Label_Editar.FontUnderline = True
    Label_Eliminar.FontUnderline = False
    Label_Relatorio.FontUnderline = False
    Label_Play.FontUnderline = False
End Sub

Private Sub Label_Eliminar_Click()
    'Atalho para
    Botao_Eliminar_Click
End Sub

Private Sub Label_Eliminar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Animar a label
    If Label_Eliminar.FontUnderline = True Then Exit Sub
    Label_Novo.FontUnderline = False
    Label_Editar.FontUnderline = False
    Label_Eliminar.FontUnderline = True
    Label_Relatorio.FontUnderline = False
    Label_Play.FontUnderline = False
End Sub

Private Sub Label_Gadgets_Click()
    'Atalgho para
    Pic_Gadgets_Click
End Sub

Private Sub Label_Gadgets_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Animar a label
    If Label_Gadgets.ForeColor = vbWhite Then Exit Sub
    Label_Gadgets.ForeColor = vbWhite
    Label_Preferencias.ForeColor = &HC0C0C0
    Label_Sobre.ForeColor = &HC0C0C0
    Label_Suporte.ForeColor = &HC0C0C0
End Sub

Private Sub Label_Letra_Click(Index As Integer)
    'Efectuar a pesquisa por letra
    Repor_Labels
    
    'Verificar se a letra selecionada é o '#' (Todos os registos)
    If Index = 26 Then
        Verifica_Rs_Filmes
        Rs_Filmes.Open "select * from Tabela_Filmes order by Id Asc", Cnn_Filmes
        Preenche_Lista
    
    Else
        Criterio = Label_Letra(Index).Caption
        Text_Filtro
    End If
    
    'Activar a label de pesquisa escolhida
    If Letra_Activa > -1 Then Pic_Letra(Letra_Activa).Picture = Form_Skin.Sombra_Letra_Normal.Picture
    Pic_Letra(Index).Picture = Form_Skin.Sombra_Letra_Activa.Picture
    Letra_Activa = Index
End Sub

Private Sub Label_Letra_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    'Animar a label de pesquisa
    If Index = Letra_selecionada Then Exit Sub
    If Letra_selecionada > -1 And Letra_Activa <> Letra_selecionada Then Pic_Letra(Letra_selecionada).Picture = Form_Skin.Sombra_Letra_Normal.Picture
    If Letra_Activa <> Index Then
        Pic_Letra(Index).Picture = Form_Skin.Sombra_Letra_Over.Picture
        Letra_selecionada = Index
    End If
End Sub

Private Sub Label_Novo_Click()
    'Atalho para
    Botao_Novo_Click
End Sub

Private Sub Label_Novo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Animar a label
    If Label_Novo.FontUnderline = True Then Exit Sub
    Label_Novo.FontUnderline = True
    Label_Editar.FontUnderline = False
    Label_Eliminar.FontUnderline = False
    Label_Relatorio.FontUnderline = False
    Label_Play.FontUnderline = False
End Sub

Private Sub Label_Play_Click()
    'Atalho para
    Botao_Play_Click
End Sub

Private Sub Label_Play_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Animar a label
    If Label_Play.FontUnderline = True Then Exit Sub
    Label_Novo.FontUnderline = False
    Label_Editar.FontUnderline = False
    Label_Eliminar.FontUnderline = False
    Label_Relatorio.FontUnderline = False
    Label_Play.FontUnderline = True
End Sub

Public Sub Label_Preferencias_Click()
    'Ver formulários das opções
    Repor_Imagens
    Form_Opcoes.Show vbModal
End Sub

Private Sub Label_Preferencias_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Animar a label
    If Label_Preferencias.ForeColor = vbWhite Then Exit Sub
    Label_Gadgets.ForeColor = &HC0C0C0
    Label_Preferencias.ForeColor = vbWhite
    Label_Sobre.ForeColor = &HC0C0C0
    Label_Suporte.ForeColor = &HC0C0C0
End Sub

Private Sub Label_Relatorio_Click()
    'Atalho para
    Botao_Relatorio_Click
End Sub

Private Sub Label_Relatorio_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Animar a label
    If Label_Relatorio.FontUnderline = True Then Exit Sub
    Label_Novo.FontUnderline = False
    Label_Editar.FontUnderline = False
    Label_Eliminar.FontUnderline = False
    Label_Relatorio.FontUnderline = True
    Label_Play.FontUnderline = False
End Sub

Private Sub Label_Sobre_Click()
    'Ver formulário sobre
    Repor_Imagens
    Form_Sobre.Show vbModal
End Sub

Private Sub Label_Sobre_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Animar a label
    If Label_Sobre.ForeColor = vbWhite Then Exit Sub
    Label_Gadgets.ForeColor = &HC0C0C0
    Label_Preferencias.ForeColor = &HC0C0C0
    Label_Sobre.ForeColor = vbWhite
    Label_Suporte.ForeColor = &HC0C0C0
End Sub

Private Sub Label_Suporte_Click()
    'Ver o formulário suporte técnico
    Form_Reportar_Erro.Show vbModal
End Sub

Private Sub Label_Suporte_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Animar a label
    If Label_Suporte.ForeColor = vbWhite Then Exit Sub
    Label_Gadgets.ForeColor = &HC0C0C0
    Label_Suporte.ForeColor = vbWhite
    Label_Preferencias.ForeColor = &HC0C0C0
    Label_Sobre.ForeColor = &HC0C0C0
End Sub

Private Sub Label_Titulo_DblClick()
    'Maximixar/ Restaurar Formulários
    If Tela_Cheia = True Then
        Botao_Restaurar_Click
    Else
        Botao_Maximizar_Click
    End If
End Sub

Private Sub Label_Titulo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Capturar a posição de x e y
    Dim POINT As POINTAPI
    GetCursorPos POINT
    LastPoint.x = POINT.x
    LastPoint.y = POINT.y
    bMoveFrom = True
End Sub

Private Sub Label_Titulo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Chamar o procedimento
    Repor_Imagens
    
    'Mover o formulário e obter a posição de x e y
    If Me.WindowState = 0 And Tela_Cheia = False Then
        Dim iDX As Long, iDY As Long
        Dim POINT As POINTAPI
        If Not bMoveFrom Then Exit Sub
        GetCursorPos POINT
        iDX& = (POINT.x - LastPoint.x) * iTPPX&
        iDY& = (POINT.y - LastPoint.y) * iTPPY&
        LastPoint.x = POINT.x
        LastPoint.y = POINT.y
        Me.Move Me.Left + iDX&, Me.Top + iDY&
    End If
End Sub

Private Sub Label_Titulo_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Largar o formulário para a posição final
    bMoveFrom = False
End Sub

Private Sub Label_Total_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Chamar o procedimento
    Repor_Imagens
End Sub

Private Sub Lista_Categorias_Click()
    'Activar a label de pesquisa escolhida
    If Letra_Activa > -1 Then Pic_Letra(Letra_Activa).Picture = Form_Skin.Sombra_Letra_Normal.Picture
    Letra_Activa = -1
    
    'Selecionar a categoria pretendida para efectuar uma pesquisa personalizada
    Select Case Lista_Categorias.Row
        Case 1
            Repor_Labels
            Verifica_Rs_Filmes
            Rs_Filmes.Open "select * from Tabela_Filmes order by Id Asc", Cnn_Filmes
            Preenche_Lista
            
        Case 2
            Categoria_Selecionada = "Acção"
            Filtar_Categorias
            
        Case 3
            Categoria_Selecionada = "Animação"
            Filtar_Categorias
            
        Case 4
            Categoria_Selecionada = "Aventura"
            Filtar_Categorias
            
        Case 5
            Categoria_Selecionada = "Comédia"
            Filtar_Categorias
            
        Case 6
            Categoria_Selecionada = "Crime"
            Filtar_Categorias
            
        Case 7
            Categoria_Selecionada = "Documentário"
            Filtar_Categorias
            
        Case 8
            Categoria_Selecionada = "Desporto"
            Filtar_Categorias
            
        Case 9
            Categoria_Selecionada = "Drama"
            Filtar_Categorias
            
        Case 10
            Categoria_Selecionada = "Faroeste"
            Filtar_Categorias
            
        Case 11
            Categoria_Selecionada = "Ficção cientifica"
            Filtar_Categorias
            
        Case 12
            Categoria_Selecionada = "Guerra"
            Filtar_Categorias
            
        Case 13
            Categoria_Selecionada = "Musical"
            Filtar_Categorias
            
        Case 14
            Categoria_Selecionada = "Policial"
            Filtar_Categorias
            
        Case 15
            Categoria_Selecionada = "Romance"
            Filtar_Categorias
            
        Case 16
            Categoria_Selecionada = "Série"
            Filtar_Categorias
            
        Case 17
            Categoria_Selecionada = "Suspense"
            Filtar_Categorias
            
        Case 18
            Categoria_Selecionada = "Terror"
            Filtar_Categorias
            
        Case 19
            Categoria_Selecionada = "Outra"
            Filtar_Categorias
    End Select
End Sub

Public Sub Filtar_Categorias()
    'Procedimento para pesquisar por categoria
    Verifica_Rs_Filmes
    Rs_Filmes.Open "select * from Tabela_Filmes where Categoria like '" & Categoria_Selecionada & "%' order by Id Asc", Cnn_Filmes
    Preenche_Lista
End Sub

Private Sub Skin_Top_Centro_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Capturar a posição de x e y
    Dim POINT As POINTAPI
    GetCursorPos POINT
    LastPoint.x = POINT.x
    LastPoint.y = POINT.y
    bMoveFrom = True
End Sub

Private Sub Skin_Top_Centro_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Mover o formulário e obter a posição de x e y
    If Me.WindowState = 0 And Tela_Cheia = False Then
        Dim iDX As Long, iDY As Long
        Dim POINT As POINTAPI
        If Not bMoveFrom Then Exit Sub
        GetCursorPos POINT
        iDX& = (POINT.x - LastPoint.x) * iTPPX&
        iDY& = (POINT.y - LastPoint.y) * iTPPY&
        LastPoint.x = POINT.x
        LastPoint.y = POINT.y
        Me.Move Me.Left + iDX&, Me.Top + iDY&
    End If
End Sub

Private Sub Skin_Top_Centro_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Largar o formulário para a posição final
    bMoveFrom = False
End Sub

Private Sub Skin_Top_Centro_DblClick()
    'Atalho para
    Label_Titulo_DblClick
End Sub

Private Sub Lista_Categorias_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Chamar o procedimento
    Repor_Imagens
End Sub

Private Sub Lista_Categorias_SelChange()
    'Atalho para
    Lista_Categorias_Click
End Sub

Public Sub Lista_Filmes_Click()
    'Selecionar linha pretendida da lista
    'On Error GoTo Corrige_Erro
    If Lista_Filmes.Rows = 1 Then Exit Sub
    
    With Lista_Filmes
        Text_Classificacao.Text = .TextMatrix(.Row, 4)

        'Verificar as estrelas
        If Text_Classificacao.Text = "1" Then
            Image1.Picture = Form_Skin.Estrela_Over_2.Picture
            Image2.Picture = Form_Skin.Estrela_Normal_2.Picture
            Image3.Picture = Form_Skin.Estrela_Normal_2.Picture
            Image4.Picture = Form_Skin.Estrela_Normal_2.Picture
            Image5.Picture = Form_Skin.Estrela_Normal_2.Picture
        ElseIf Text_Classificacao.Text = "2" Then
            Image1.Picture = Form_Skin.Estrela_Over_2.Picture
            Image2.Picture = Form_Skin.Estrela_Over_2.Picture
            Image3.Picture = Form_Skin.Estrela_Normal_2.Picture
            Image4.Picture = Form_Skin.Estrela_Normal_2.Picture
            Image5.Picture = Form_Skin.Estrela_Normal_2.Picture
        ElseIf Text_Classificacao.Text = "3" Then
            Image1.Picture = Form_Skin.Estrela_Over_2.Picture
            Image2.Picture = Form_Skin.Estrela_Over_2.Picture
            Image3.Picture = Form_Skin.Estrela_Over_2.Picture
            Image4.Picture = Form_Skin.Estrela_Normal_2.Picture
            Image5.Picture = Form_Skin.Estrela_Normal_2.Picture
        ElseIf Text_Classificacao.Text = "4" Then
            Image1.Picture = Form_Skin.Estrela_Over_2.Picture
            Image2.Picture = Form_Skin.Estrela_Over_2.Picture
            Image3.Picture = Form_Skin.Estrela_Over_2.Picture
            Image4.Picture = Form_Skin.Estrela_Over_2.Picture
            Image5.Picture = Form_Skin.Estrela_Normal_2.Picture
        ElseIf Text_Classificacao.Text = "5" Then
            Image1.Picture = Form_Skin.Estrela_Over_2.Picture
            Image2.Picture = Form_Skin.Estrela_Over_2.Picture
            Image3.Picture = Form_Skin.Estrela_Over_2.Picture
            Image4.Picture = Form_Skin.Estrela_Over_2.Picture
            Image5.Picture = Form_Skin.Estrela_Over_2.Picture
        End If
    End With
    
    'Carrega imagem da capa, caso exista
    On Error GoTo Corrige_Erro
    If Lista_Filmes.TextMatrix(Lista_Filmes.Row, 8) <> Empty Then
        Image_Capa.Picture = Form_Skin.Image_Sem_Capa.Picture
        Image_Capa.Visible = True
        Image_Capa.Picture = LoadPicture(Lista_Filmes.TextMatrix(Lista_Filmes.Row, 8))
        
    Else
        Image_Capa.Picture = Form_Skin.Image_Sem_Capa.Picture
        Image_Capa.Visible = False
    End If

Exit Sub
Corrige_Erro:
    Select Case Err.Number
        Case "76"
            'Image_Capa.Visible = False
    End Select
End Sub

Private Sub Lista_Filmes_DblClick()
    'Atalho para
    Botao_Play_Click
End Sub

Private Sub Lista_Filmes_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Chamar o procedimento
    Repor_Imagens
End Sub

Private Sub Lista_Filmes_SelChange()
    'Atalho para
    Lista_Filmes_Click
End Sub

Private Sub Pic_Gadgets_Click()
    'Abrir página pessoal
    Call ShellExecute(0, "open", "http://www.nikyts.com/gadgets.html", vbNullString, vbNullString, SW_NORMAL)
End Sub

Private Sub Pic_Gadgets_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Chamar o procedimento
    Repor_Imagens
End Sub

Private Sub Pic_Letra_Click(Index As Integer)
    'Efectuar a pesquisa por letra
    Repor_Labels
    
    'Verificar se a letra selecionada é o '#' (Todos os registos)
    If Index = 26 Then
        Verifica_Rs_Filmes
        Rs_Filmes.Open "select * from Tabela_Filmes order by Id Asc", Cnn_Filmes
        Preenche_Lista
    
    Else
        Criterio = Label_Letra(Index).Caption
        Text_Filtro
    End If
    
    'Activar a label de pesquisa escolhida
    If Letra_Activa > -1 Then Pic_Letra(Letra_Activa).Picture = Form_Skin.Sombra_Letra_Normal.Picture
    Pic_Letra(Index).Picture = Form_Skin.Sombra_Letra_Activa.Picture
    Letra_Activa = Index
End Sub

Private Sub Pic_Letra_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    'Animar a label de pesquisa
    If Index = Letra_selecionada Then Exit Sub
    If Letra_selecionada > -1 And Letra_Activa <> Letra_selecionada Then Pic_Letra(Letra_selecionada).Picture = Form_Skin.Sombra_Letra_Normal.Picture
    If Letra_Activa <> Index Then
        Pic_Letra(Index).Picture = Form_Skin.Sombra_Letra_Over.Picture
        Letra_selecionada = Index
    End If
End Sub

Private Sub Picture2_Click()

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Chamar o procedimento
    Repor_Imagens
End Sub

Private Sub Text_Pesquisa_Change()
    'Efectuar pesquisa pela caixa de texto
    If Letra_Activa > -1 Then Pic_Letra(Letra_Activa).Picture = Form_Skin.Sombra_Letra_Normal.Picture
    Letra_Activa = -1
    Criterio = Text_Pesquisa.Text
    Text_Filtro
End Sub

Public Sub Text_Filtro()
    Verifica_Rs_Filmes
    Rs_Filmes.Open "select * from Tabela_Filmes where Video like '" & Criterio & "%' order by Id Asc", Cnn_Filmes
    Preenche_Lista
End Sub

Public Sub Formatar_Lista_Filmes()
    'Procedimento para formatar as msflexgrid
    With Lista_Filmes
        .AllowUserResizing = flexResizeColumns
        .RowHeight(0) = 270
        .TextArray(0) = "Id"
        .ColAlignment(0) = 0
        .ColWidth(0) = 0
        .TextArray(1) = "Nome do filme"
        .ColAlignment(1) = 0
        .ColWidth(1) = 4000
        .TextArray(2) = "Categoria"
        .ColAlignment(2) = 0
        .ColWidth(2) = 2300
        .TextArray(3) = "Tipo"
        .ColAlignment(3) = 0
        .ColWidth(3) = 1500
        .TextArray(4) = "Classificacao"
        .ColAlignment(4) = 0
        .ColWidth(4) = 1300
        .TextArray(5) = "Actores"
        .ColAlignment(5) = 0
        .ColWidth(5) = 3000
        .TextArray(6) = "Observações"
        .ColAlignment(6) = 0
        .ColWidth(6) = 5000
        .TextArray(7) = "Directório"
        .ColAlignment(7) = 0
        .ColWidth(7) = 2000
        .TextArray(8) = "Capa"
        .ColAlignment(8) = 0
        .ColWidth(8) = 2000
    End With
End Sub

Public Sub Formatar_Lista_Categorias()
    'Procedimento para formatar as msflexgrid
    With Lista_Categorias
        .TextMatrix(0, 1) = "Categorias"
        .ColWidth(0) = 0
        .ColWidth(1) = 6000
        
        .TextMatrix(1, 1) = "Todas (0)"
        .TextMatrix(2, 1) = "Acção (0)"
        .TextMatrix(3, 1) = "Animação (0)"
        .TextMatrix(4, 1) = "Aventura (0)"
        .TextMatrix(5, 1) = "Comédia (0)"
        .TextMatrix(6, 1) = "Crime (0)"
        .TextMatrix(7, 1) = "Documentário (0)"
        .TextMatrix(8, 1) = "Desporto (0)"
        .TextMatrix(9, 1) = "Drama (0)"
        .TextMatrix(10, 1) = "Faroeste (0)"
        .TextMatrix(11, 1) = "Ficção cientifica (0)"
        .TextMatrix(12, 1) = "Guerra (0)"
        .TextMatrix(13, 1) = "Musical (0)"
        .TextMatrix(14, 1) = "Policial (0)"
        .TextMatrix(15, 1) = "Romance (0)"
        .TextMatrix(16, 1) = "Série (0)"
        .TextMatrix(17, 1) = "Suspense (0)"
        .TextMatrix(18, 1) = "Terror (0)"
        .TextMatrix(19, 1) = "Outra (0)"
    End With
End Sub

Public Sub Limpar_Campos()
    'Procedimento para limpar as textbox e capa
    Text_Classificacao.Text = ""
    Image_Capa.Visible = False
End Sub

Public Sub Repor_Labels()
    'Repor a cor das labels de pesquisa personalizada
    Text_Pesquisa.Text = ""
    Lista_Categorias.Row = 1
End Sub

Public Sub Ver_Opcoes()
    'Procedimento ver o estado da janela
    If Form_Opcoes.Text_Tela_Cheia.Text = "True" Then
        Botao_Maximizar_Click
        Tela_Cheia = True
    Else
        Botao_Restaurar_Click
        Tela_Cheia = False
    End If
End Sub

Public Sub Repor_Imagens()
    'Procedimento para repor imagens originais ou ocultar não desejadas    Shape_Over_Letras.Visible = False
    If Letra_selecionada <> -1 And Letra_selecionada <> Letra_Activa Then
        Pic_Letra(Letra_selecionada).Picture = Form_Skin.Sombra_Letra_Normal.Picture
        Letra_selecionada = -1
    End If
    
    If Label_Gadgets.ForeColor <> &HC0C0C0 Then Label_Gadgets.ForeColor = &HC0C0C0
    If Label_Preferencias.ForeColor <> &HC0C0C0 Then Label_Preferencias.ForeColor = &HC0C0C0
    If Label_Suporte.ForeColor <> &HC0C0C0 Then Label_Suporte.ForeColor = &HC0C0C0
    If Label_Sobre.ForeColor <> &HC0C0C0 Then Label_Sobre.ForeColor = &HC0C0C0
    
    If Label_Novo.FontUnderline <> False Then Label_Novo.FontUnderline = False
    If Label_Editar.FontUnderline <> False Then Label_Editar.FontUnderline = False
    If Label_Eliminar.FontUnderline <> False Then Label_Eliminar.FontUnderline = False
    If Label_Relatorio.FontUnderline <> False Then Label_Relatorio.FontUnderline = False
    If Label_Play.FontUnderline <> False Then Label_Play.FontUnderline = False
End Sub

Public Sub Posicionar_Frames()
    'Procedimento para posicionar a barra lateral e centro consoante as opções selecionadas pelo utilizador
    If Form_Opcoes.Opcao_Esquerda.Value = True Then
        Barra_Lateral.Left = 1 + 10
        frame_Centro.Left = Barra_Lateral.Left + Barra_Lateral.Width + 10
    Else
        frame_Centro.Left = 1 + 10
        Barra_Lateral.Left = frame_Centro.Left + frame_Centro.Width + 10
    End If
    
    Frame_Capa.Left = Barra_Lateral.Left
End Sub
