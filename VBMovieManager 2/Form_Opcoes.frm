VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form_Opcoes 
   Appearance      =   0  'Flat
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   0  'None
   ClientHeight    =   5850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9645
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   390
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   643
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Frame_Actualizacoes 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   120
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   417
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   9120
      Visible         =   0   'False
      Width           =   6255
      Begin VB.PictureBox Pic_Actualizacoes 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         Picture         =   "Form_Opcoes.frx":0000
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   60
         Width           =   195
      End
      Begin VB.CheckBox Check_Actualizacoes 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Verificar actualizações automaticamente"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   0
         TabIndex        =   49
         Top             =   0
         Value           =   1  'Checked
         Width           =   6000
      End
   End
   Begin VB.PictureBox Frame_Informacao 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   3000
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   417
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   6255
      Begin VB.Label Label_Close 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "x"
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   6000
         TabIndex        =   41
         ToolTipText     =   "Ocultar"
         Top             =   96
         Width           =   108
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "As opções do programa foram actualizadas com sucesso"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   105
         Width           =   5475
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00F0E7D7&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00E6964D&
         Height          =   372
         Left            =   0
         Top             =   0
         Width           =   6252
      End
   End
   Begin VB.PictureBox Frame_Capas 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   9960
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   417
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   7920
      Visible         =   0   'False
      Width           =   6255
      Begin VB.PictureBox Pic_Copiar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         Picture         =   "Form_Opcoes.frx":034C
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   60
         Width           =   195
      End
      Begin VB.CheckBox Check_Copiar 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Copiar a capa do filme para a pasta 'Capas' ao adicionar ou editar um registo"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   0
         TabIndex        =   38
         Top             =   0
         Width           =   6000
      End
   End
   Begin VB.PictureBox Frame_Reiniciar 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   120
      ScaleHeight     =   193
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   417
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   6255
      Begin VB.PictureBox Botao_Reset 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         Picture         =   "Form_Opcoes.frx":0596
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   181
         TabIndex        =   42
         Top             =   360
         Width           =   2715
         Begin VB.Label Label_Reset 
            Alignment       =   2  'Center
            BackColor       =   &H00272727&
            BackStyle       =   0  'Transparent
            Caption         =   "Iniciar o processo"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   0
            TabIndex        =   43
            Top             =   135
            Width           =   2715
         End
         Begin VB.Shape Contorno_Reset 
            BorderColor     =   &H00E6964D&
            Height          =   495
            Left            =   240
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Visible         =   0   'False
            Width           =   2715
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Eliminar todos os registos da minha base de dados"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   0
         TabIndex        =   44
         Top             =   0
         Width           =   4365
      End
   End
   Begin VB.PictureBox Barra_ControlBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00313131&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   0
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   641
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   0
      Width           =   9615
      Begin VB.Label Label_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00272727&
         BackStyle       =   0  'Transparent
         Caption         =   "Preferências"
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
         Left            =   120
         TabIndex        =   24
         Top             =   120
         Width           =   1245
      End
      Begin VB.Image Botao_Fechar 
         Height          =   225
         Left            =   9240
         Picture         =   "Form_Opcoes.frx":0D07
         ToolTipText     =   "Fechar"
         Top             =   120
         Width           =   240
      End
      Begin VB.Image Fundo_Barra_ControlBox 
         Enabled         =   0   'False
         Height          =   465
         Left            =   0
         Picture         =   "Form_Opcoes.frx":1019
         Top             =   0
         Width           =   585
      End
   End
   Begin VB.PictureBox Barra_Botoes 
      Appearance      =   0  'Flat
      BackColor       =   &H00D9D9D9&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   900
      Left            =   0
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   641
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   4920
      Width           =   9615
      Begin VB.PictureBox Botao_Aplicar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   5760
         Picture         =   "Form_Opcoes.frx":135E
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   61
         TabIndex        =   8
         Top             =   240
         Width           =   915
         Begin VB.Shape Contorno_Aplicar 
            BorderColor     =   &H00E6964D&
            Height          =   495
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label Label_Aplicar 
            Alignment       =   2  'Center
            BackColor       =   &H00272727&
            BackStyle       =   0  'Transparent
            Caption         =   "Aplicar"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   0
            TabIndex        =   35
            Top             =   135
            Width           =   915
         End
      End
      Begin VB.PictureBox Botao_Cancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   7920
         Picture         =   "Form_Opcoes.frx":2B58
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   61
         TabIndex        =   10
         Top             =   240
         Width           =   915
         Begin VB.Shape Contorno_Cancelar 
            BorderColor     =   &H00E6964D&
            Height          =   495
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label Label_Cancelar 
            Alignment       =   2  'Center
            BackColor       =   &H00272727&
            BackStyle       =   0  'Transparent
            Caption         =   "Cancelar"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   0
            TabIndex        =   25
            Top             =   135
            Width           =   915
         End
      End
      Begin VB.PictureBox Botao_Ok 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   6840
         Picture         =   "Form_Opcoes.frx":4352
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   61
         TabIndex        =   9
         Top             =   240
         Width           =   915
         Begin VB.Label Label_Ok 
            Alignment       =   2  'Center
            BackColor       =   &H00272727&
            BackStyle       =   0  'Transparent
            Caption         =   "Ok"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   0
            TabIndex        =   22
            Top             =   135
            Width           =   915
         End
         Begin VB.Shape Contorno_Ok 
            BorderColor     =   &H00E6964D&
            Height          =   495
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Visible         =   0   'False
            Width           =   915
         End
      End
      Begin VB.Image Fundo_Barra_Botoes 
         Enabled         =   0   'False
         Height          =   900
         Left            =   0
         Picture         =   "Form_Opcoes.frx":5B4C
         Top             =   0
         Width           =   585
      End
   End
   Begin VB.TextBox Text_Tela_Cheia 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF80FF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   840
      TabIndex        =   20
      Text            =   "True"
      Top             =   4320
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.PictureBox Frame_Relatorio 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   10200
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   417
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4080
      Visible         =   0   'False
      Width           =   6255
      Begin VB.PictureBox Pic_Abrir_Relatorio 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         Picture         =   "Form_Opcoes.frx":5EB7
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   30
         Width           =   195
      End
      Begin VB.CheckBox Check_Abrir_Relatorio 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Abrir relatório automaticamente após este ser criado"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   5040
      End
   End
   Begin VB.PictureBox Frame_Video 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   10080
      ScaleHeight     =   129
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   417
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5760
      Visible         =   0   'False
      Width           =   6255
      Begin VB.PictureBox Pic_Programa 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   0
         Picture         =   "Form_Opcoes.frx":6101
         ScaleHeight     =   12
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   12
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   390
         Width           =   180
      End
      Begin VB.PictureBox Pic_Wmp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   0
         Picture         =   "Form_Opcoes.frx":62F3
         ScaleHeight     =   12
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   12
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   750
         Width           =   180
      End
      Begin VB.OptionButton Opcao_Wmp 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Windows media player"
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   720
         Width           =   2295
      End
      Begin VB.OptionButton Opcao_Programa 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Tela de video do programa"
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   2655
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Visualiar os filmes com:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   2055
      End
      Begin VB.Label Label_Video 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FF80&
         Caption         =   "Programa"
         Height          =   195
         Left            =   2280
         TabIndex        =   18
         Top             =   0
         Visible         =   0   'False
         Width           =   1080
      End
   End
   Begin VB.PictureBox Frame_Visualizacao 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   10080
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   417
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2040
      Visible         =   0   'False
      Width           =   6255
      Begin VB.PictureBox Pic_Esquerda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   0
         Picture         =   "Form_Opcoes.frx":64E5
         ScaleHeight     =   12
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   12
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   390
         Width           =   180
      End
      Begin VB.PictureBox Pic_Direita 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   0
         Picture         =   "Form_Opcoes.frx":66D7
         ScaleHeight     =   12
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   12
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   750
         Width           =   180
      End
      Begin VB.OptionButton Opcao_Direita 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Direita"
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton Opcao_Esquerda 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Esquerda"
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.Label Label_Posicionamento 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FF80&
         Caption         =   "Esquerda"
         Height          =   195
         Left            =   3480
         TabIndex        =   15
         Top             =   0
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Posicionamento da barra de categorias:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   3420
      End
   End
   Begin VB.PictureBox Frame_Base_Dados 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   3000
      ScaleHeight     =   209
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   417
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1200
      Width           =   6255
      Begin VB.TextBox Text_Destino 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   0
         TabIndex        =   46
         Top             =   1320
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox Text_Origem 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3120
         TabIndex        =   45
         Top             =   1320
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.PictureBox Botao_Copia_Seguranca 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         Picture         =   "Form_Opcoes.frx":68C9
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   181
         TabIndex        =   2
         Top             =   720
         Width           =   2715
         Begin VB.Shape Contorno_Copia_Seguranca 
            BorderColor     =   &H00E6964D&
            Height          =   495
            Left            =   240
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Visible         =   0   'False
            Width           =   2715
         End
         Begin VB.Label Label_Copia_Seguranca 
            Alignment       =   2  'Center
            BackColor       =   &H00272727&
            BackStyle       =   0  'Transparent
            Caption         =   "Efectuar copia de segurança"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   0
            TabIndex        =   28
            Top             =   120
            Width           =   2715
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   0
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   345
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   240
         Width           =   5175
         Begin VB.TextBox Text_Localizacao_BD 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   15
            Locked          =   -1  'True
            TabIndex        =   1
            Top             =   15
            Width           =   5025
         End
         Begin VB.Shape Shape_Localizacao_BD 
            BorderColor     =   &H00C0C0C0&
            Height          =   315
            Left            =   0
            Top             =   0
            Width           =   5055
         End
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Localização da base de dados"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   2550
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Lista_Opcoes 
      Height          =   2772
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   2532
      _ExtentX        =   4471
      _ExtentY        =   4895
      _Version        =   393216
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
      ScrollBars      =   0
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
      X1              =   250
      X2              =   666
      Y1              =   80
      Y2              =   80
   End
   Begin VB.Label Label_Topico_Selecionado 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Base de dados"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   270
      Left            =   3000
      TabIndex        =   26
      Top             =   600
      Width           =   1815
   End
   Begin VB.Line Linha_Vertical 
      BorderColor     =   &H00E0E0E0&
      X1              =   176
      X2              =   176
      Y1              =   24
      Y2              =   312
   End
   Begin VB.Shape Shape_Contorno 
      BorderColor     =   &H00C0C0C0&
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "Form_Opcoes"
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
    'Mover o formulário e obter a posição de x e y
    Dim iDX As Long, iDY As Long
    Dim POINT As POINTAPI
    If Not bMoveFrom Then Exit Sub
    GetCursorPos POINT
    iDX& = (POINT.x - LastPoint.x) * iTPPX&
    iDY& = (POINT.y - LastPoint.y) * iTPPY&
    LastPoint.x = POINT.x
    LastPoint.y = POINT.y
    Me.Move Me.Left + iDX&, Me.Top + iDY&
End Sub

Private Sub Barra_ControlBox_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Largar o formulário para a posição final
    bMoveFrom = False
End Sub

Private Sub Botao_Aplicar_Click()
    'Atalho para
    Label_Aplicar_Click
End Sub

Private Sub Botao_Aplicar_GotFocus()
    'Colocar o focus no botao
    Contorno_Aplicar.Visible = True
End Sub

Private Sub Botao_Aplicar_KeyDown(KeyCode As Integer, Shift As Integer)
    'Atalho de teclas
    If KeyCode = vbKeyReturn Then Botao_Aplicar_Click
End Sub

Private Sub Botao_Aplicar_LostFocus()
    'Remover o focus no botao
    Contorno_Aplicar.Visible = False
End Sub

Private Sub Botao_Aplicar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Animar o botão
    If Botao_Aplicar.Picture = Form_Skin.Botao_Over.Picture Then Exit Sub
    Repor_Imagens
    Botao_Aplicar.Picture = Form_Skin.Botao_Over.Picture
End Sub

Private Sub Botao_Cancelar_Click()
    'Atalho para
    Label_Cancelar_Click
End Sub

Private Sub Botao_Cancelar_GotFocus()
    'Colocar o focus no botao
    Contorno_Cancelar.Visible = True
End Sub

Private Sub Botao_Cancelar_KeyDown(KeyCode As Integer, Shift As Integer)
    'Atalho de teclas
    If KeyCode = vbKeyReturn Then Botao_Cancelar_Click
End Sub

Private Sub Botao_Cancelar_LostFocus()
    'Remover o focus no botao
    Contorno_Cancelar.Visible = False
End Sub

Private Sub Botao_Cancelar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Animar o botão
    If Botao_Cancelar.Picture = Form_Skin.Botao_Over.Picture Then Exit Sub
    Repor_Imagens
    Botao_Cancelar.Picture = Form_Skin.Botao_Over.Picture
End Sub

Private Sub Botao_Copia_Seguranca_Click()
    'Atalho para
    Label_Copia_Seguranca_Click
End Sub

Private Sub Botao_Copia_Seguranca_GotFocus()
    'Colocar o focus no botao
    Contorno_Copia_Seguranca.Visible = True
End Sub

Private Sub Botao_Copia_Seguranca_KeyDown(KeyCode As Integer, Shift As Integer)
    'Atalho de teclas
    If KeyCode = vbKeyReturn Then Botao_Copia_Seguranca_Click
End Sub

Private Sub Botao_Copia_Seguranca_LostFocus()
    'Remover o focus no botao
    Contorno_Copia_Seguranca.Visible = False
End Sub

Private Sub Botao_Fechar_Click()
    'Fechar o formulário
    Unload Me
End Sub

Private Sub Botao_Ok_Click()
    'Atalho para
    Label_Ok_Click
End Sub

Private Sub Botao_Ok_GotFocus()
    'Colocar o focus no botao
    Contorno_Ok.Visible = True
End Sub

Private Sub Botao_Ok_KeyDown(KeyCode As Integer, Shift As Integer)
    'Atalho de teclas
    If KeyCode = vbKeyReturn Then Botao_Ok_Click
End Sub

Private Sub Botao_Ok_LostFocus()
    'Remover o focus no botao
    Contorno_Ok.Visible = False
End Sub

Private Sub Botao_Ok_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Animar o botão
    If Botao_Ok.Picture = Form_Skin.Botao_Over.Picture Then Exit Sub
    Repor_Imagens
    Botao_Ok.Picture = Form_Skin.Botao_Over.Picture
End Sub

Private Sub Botao_Reset_Click()
    'Atalho para
    Label_Reset_Click
End Sub

Private Sub Botao_Reset_GotFocus()
    'Colocar o focus no botao
    Contorno_Reset.Visible = True
End Sub

Private Sub Botao_Reset_KeyDown(KeyCode As Integer, Shift As Integer)
    'Atalho de teclas
    If KeyCode = vbKeyReturn Then Botao_Reset_Click
End Sub

Private Sub Botao_Reset_LostFocus()
    'Remover o focus no botao
    Contorno_Reset.Visible = False
End Sub

Private Sub Check_Abrir_Relatorio_Click()
    'Des/Activar a opcção
    If Check_Abrir_Relatorio.Value = 1 Then
        Pic_Abrir_Relatorio.Picture = Form_Skin.Check_Over.Picture
        
    Else
        Pic_Abrir_Relatorio.Picture = Form_Skin.Check_Normal.Picture
    End If
End Sub

Private Sub Check_Actualizacoes_Click()
    'Des/Activar a opcção de "Lembrar"
    If Check_Actualizacoes.Value = 1 Then
        Pic_Actualizacoes.Picture = Form_Skin.Check_Over.Picture
        
    Else
        Pic_Actualizacoes.Picture = Form_Skin.Check_Normal.Picture
    End If
End Sub

Private Sub Check_Copiar_Click()
    'Des/Activar a opcção
    If Check_Copiar.Value = 1 Then
        Pic_Copiar.Picture = Form_Skin.Check_Over.Picture
        
    Else
        Pic_Copiar.Picture = Form_Skin.Check_Normal.Picture
    End If
End Sub

Private Sub Form_Load()
    'Propriedades iniciais do formulário
    Desenhar_Formulario
    
    'Variáveis para identificar as pastas e ficheiros utilizados pelo programa
    Dim Localizacao_Ficheiro_Preferencias As String
    Localizacao_Ficheiro_Preferencias = App.path & "\Options\Properties.ini"
    
    iTPPX& = Screen.TwipsPerPixelX
    iTPPY& = Screen.TwipsPerPixelY
    Text_Localizacao_BD.Text = App.path & "\Data\Biblioteca.mdb"
    
    'Formatar a lista de opções
    With Lista_Opcoes
        .RowHeight(0) = 0
        .ColWidth(0) = 0
        .ColWidth(1) = 6000
        .Rows = 8
        .TextMatrix(1, 1) = "Base de dados"
        .TextMatrix(2, 1) = "Visualização"
        .TextMatrix(3, 1) = "Relatório"
        .TextMatrix(4, 1) = "Tela de video"
        .TextMatrix(5, 1) = "Reiniciar programa"
        .TextMatrix(6, 1) = "Capas dos filmes"
        .TextMatrix(7, 1) = "Actualizações"
    End With
    
    'Chamar o procedimento
    Posicionar_Frames

    'Frame visualização
    Label_Posicionamento.Caption = ReadINI("Visualização", "Posição da barra lateral", Localizacao_Ficheiro_Preferencias)
    If Label_Posicionamento.Caption = "Esquerda" Then
        Pic_Esquerda.Picture = Form_Skin.Opcao_Over.Picture
        Pic_Direita.Picture = Form_Skin.Opcao_Normal.Picture
        Opcao_Esquerda.Value = True
    Else
        Pic_Esquerda.Picture = Form_Skin.Opcao_Normal.Picture
        Pic_Direita.Picture = Form_Skin.Opcao_Over.Picture
        Opcao_Direita.Value = True
    End If

    'Frame video
    Label_Video.Caption = ReadINI("Tela de video", "Visualizar os filmes com", Localizacao_Ficheiro_Preferencias)
    If Label_Video.Caption = "Programa" Then
        Pic_Programa.Picture = Form_Skin.Opcao_Over.Picture
        Pic_Wmp.Picture = Form_Skin.Opcao_Normal.Picture
        Opcao_Programa.Value = True
    Else
        Pic_Programa.Picture = Form_Skin.Opcao_Normal.Picture
        Pic_Wmp.Picture = Form_Skin.Opcao_Over.Picture
        Opcao_Wmp.Value = True
    End If

    'Frame relatório
    Check_Abrir_Relatorio.Value = ReadINI("Relatório", "Abrir relatório automáticamente", Localizacao_Ficheiro_Preferencias)
    If Check_Abrir_Relatorio.Value = "1" Then
        Pic_Abrir_Relatorio.Picture = Form_Skin.Check_Over.Picture
    Else
        Pic_Abrir_Relatorio.Picture = Form_Skin.Check_Normal.Picture
    End If
    
    'Frame capas dos filmes
    Check_Copiar.Value = ReadINI("Capas dos filmes", "Copiar as imagens para a pasta do programa", Localizacao_Ficheiro_Preferencias)
    If Check_Copiar.Value = "1" Then
        Pic_Copiar.Picture = Form_Skin.Check_Over.Picture
    Else
        Pic_Copiar.Picture = Form_Skin.Check_Normal.Picture
    End If
    
    'Frame actualizações
    Check_Actualizacoes.Value = ReadINI("Actualizações", "Verificar actualizações automaticamente", Localizacao_Ficheiro_Preferencias)
    If Check_Actualizacoes.Value = "1" Then
        Pic_Actualizacoes.Picture = Form_Skin.Check_Over.Picture
    Else
        Pic_Actualizacoes.Picture = Form_Skin.Check_Normal.Picture
    End If
End Sub

Private Sub Form_Resize()
    'Chamar o procedimento
    If Me.WindowState = 1 Then Exit Sub
    Desenhar_Formulario
End Sub

Private Sub Frame_Base_Dados_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Chamar o procedimento
    Repor_Imagens
End Sub

Private Sub Frame_Reiniciar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Chamar o procedimento
    Repor_Imagens
End Sub

Private Sub Frame_Relatorio_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Chamar o procedimento
    Repor_Imagens
End Sub

Private Sub Frame_Video_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Chamar o procedimento
    Repor_Imagens
End Sub

Private Sub Frame_Visualizacao_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Chamar o procedimento
    Repor_Imagens
End Sub

Private Sub Label_Aplicar_Click()
    'Actualizar as opções do programa
    Dim Localizacao_Ficheiro_Preferencias As String
    Localizacao_Ficheiro_Preferencias = App.path & "\Options\Properties.ini"
    
    Call WriteINI("Relatório", "Abrir relatório automáticamente", Check_Abrir_Relatorio.Value, (Localizacao_Ficheiro_Preferencias))
    Call WriteINI("Tela de video", "Visualizar os filmes com", Label_Video.Caption, (Localizacao_Ficheiro_Preferencias))
    Call WriteINI("Visualização", "Posição da barra lateral", Label_Posicionamento.Caption, (Localizacao_Ficheiro_Preferencias))
    Call WriteINI("Capas dos filmes", "Copiar as imagens para a pasta do programa", Check_Copiar.Value, (Localizacao_Ficheiro_Preferencias))
    Call WriteINI("Actualizações", "Verificar actualizações automaticamente", Check_Actualizacoes.Value, (Localizacao_Ficheiro_Preferencias))
    
    Frame_Informacao.Visible = True
    Posicionar_Frames
    
    'Chamar o procedimento para alinhar as frame sconsoante as opções do programa
    Form_Principal.Posicionar_Frames
End Sub

Private Sub Label_Aplicar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Animar o botão
    If Botao_Aplicar.Picture = Form_Skin.Botao_Over.Picture Then Exit Sub
    Repor_Imagens
    Botao_Aplicar.Picture = Form_Skin.Botao_Over.Picture
End Sub

Private Sub Label_Cancelar_Click()
    'Fechar o formulário
    Unload Me
End Sub

Private Sub Label_Cancelar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Animar o botão
    If Botao_Cancelar.Picture = Form_Skin.Botao_Over.Picture Then Exit Sub
    Repor_Imagens
    Botao_Cancelar.Picture = Form_Skin.Botao_Over.Picture
End Sub

Private Sub Label_Close_Click()
    'Ocultar a frame informação
    Frame_Informacao.Visible = False
    Posicionar_Frames
End Sub

Private Sub Label_Copia_Seguranca_Click()
    'Efectuar copia de segurança da base de dados
    Text_Origem.Text = Text_Localizacao_BD.Text
    Text_Destino.Text = App.path & "\Data\Backups\Biblioteca" & "(" & Date & ")" & ".mdb"
    CopiarArquivo Text_Origem.Text, Text_Destino.Text
End Sub

Private Sub Label_Ok_Click()
    'Fechar o formulário
    Unload Me
End Sub

Private Sub Label_Ok_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Animar o botão
    If Botao_Ok.Picture = Form_Skin.Botao_Over.Picture Then Exit Sub
    Repor_Imagens
    Botao_Ok.Picture = Form_Skin.Botao_Over.Picture
End Sub

Private Sub Label_Reset_Click()
    'Eliminar todos os registos da tabela filmes
    Mensagem_de_Aviso "Questão", "Esta opção vai eliminar todos os registos da sua base de dados." & vbNewLine & "Pretende continuar o processo?"
    If Resposta = "Sim" Then
        With Form_Principal
            Me.MousePointer = 11
            .Cnn_Filmes.Execute "DELETE FROM Tabela_Filmes"
        
            .Rs_Filmes.Close
            .Conectar
            .Verifica_Rs_Filmes
            .Rs_Filmes.Open "select * from Tabela_Filmes order by Id Asc", .Cnn_Filmes

            .Formatar_Lista_Categorias
            .Preenche_Lista
            .Limpar_Campos
            .Image_Capa.Picture = Form_Skin.Image_Sem_Capa.Picture
            If .Letra_Activa > -1 Then .Pic_Letra(.Letra_Activa).Picture = Form_Skin.Sombra_Letra_Normal.Picture
            .Letra_Activa = -1
            
            Me.MousePointer = 1
            Mensagem_de_Aviso "Informação", "Processo concluido."
        End With
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
    'Mover o formulário e obter a posição de x e y
    Dim iDX As Long, iDY As Long
    Dim POINT As POINTAPI
    If Not bMoveFrom Then Exit Sub
    GetCursorPos POINT
    iDX& = (POINT.x - LastPoint.x) * iTPPX&
    iDY& = (POINT.y - LastPoint.y) * iTPPY&
    LastPoint.x = POINT.x
    LastPoint.y = POINT.y
    Me.Move Me.Left + iDX&, Me.Top + iDY&
End Sub

Private Sub Label_Titulo_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Largar o formulário para a posição final
    bMoveFrom = False
End Sub

Private Sub Label17_Click()

End Sub

Private Sub Label18_Click()

End Sub

Private Sub Lista_Opcoes_Click()
    'Selecionar os tópicos
    Ocultar_Frames
    Label_Close_Click
    Label_Topico_Selecionado.Caption = Lista_Opcoes.TextMatrix(Lista_Opcoes.Row, 1)
    
    Select Case Lista_Opcoes.Row
        Case 1
            Frame_Base_Dados.Visible = True
        
        Case 2
            Frame_Visualizacao.Visible = True
        
        Case 3
            Frame_Relatorio.Visible = True
        
        Case 4
            Frame_Video.Visible = True
        
        Case 5
            Frame_Reiniciar.Visible = True
            
        Case 6
            Frame_Capas.Visible = True
            
        Case 7
            Frame_Actualizacoes.Visible = True
    End Select
End Sub

Private Sub Lista_Opcoes_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Chamar o procedimento
    Repor_Imagens
End Sub

Private Sub Lista_Opcoes_SelChange()
    'Atalho para
    Lista_Opcoes_Click
End Sub

Private Sub Opcao_Direita_Click()
    'Opção escolhida
    Label_Posicionamento.Caption = "Direita"
    Pic_Esquerda.Picture = Form_Skin.Opcao_Normal.Picture
    Pic_Direita.Picture = Form_Skin.Opcao_Over.Picture
End Sub

Private Sub Opcao_Esquerda_Click()
    'Opção escolhida
    Label_Posicionamento.Caption = "Esquerda"
    Pic_Esquerda.Picture = Form_Skin.Opcao_Over.Picture
    Pic_Direita.Picture = Form_Skin.Opcao_Normal.Picture
End Sub

Private Sub Opcao_Programa_Click()
    'Opção escolhida
    Label_Video.Caption = "Programa"
    Pic_Programa.Picture = Form_Skin.Opcao_Over.Picture
    Pic_Wmp.Picture = Form_Skin.Opcao_Normal.Picture
End Sub

Private Sub Opcao_Wmp_Click()
    'Opção escolhida
    Label_Video.Caption = "Windows media player"
    Pic_Programa.Picture = Form_Skin.Opcao_Normal.Picture
    Pic_Wmp.Picture = Form_Skin.Opcao_Over.Picture
End Sub

Private Sub Pic_Abrir_Relatorio_Click()
    'Des/Activar a opcção
    If Check_Abrir_Relatorio.Value = 0 Then
        Check_Abrir_Relatorio.Value = 1
        Pic_Abrir_Relatorio.Picture = Form_Skin.Check_Over.Picture
        
    Else
        Check_Abrir_Relatorio.Value = 0
        Pic_Abrir_Relatorio.Picture = Form_Skin.Check_Normal.Picture
    End If
End Sub

Private Sub Pic_Actualizacoes_Click()
    'Des/Activar a opcção de "Lembrar"
    If Check_Actualizacoes.Value = 0 Then
        Check_Actualizacoes.Value = 1
        Pic_Actualizacoes.Picture = Form_Skin.Check_Over.Picture
        
    Else
        Check_Actualizacoes.Value = 0
        Pic_Actualizacoes.Picture = Form_Skin.Check_Normal.Picture
    End If
End Sub

Private Sub Pic_Copiar_Click()
    'Des/Activar a opcção
    If Check_Copiar.Value = 0 Then
        Check_Copiar.Value = 1
        Pic_Copiar.Picture = Form_Skin.Check_Over.Picture
        
    Else
        Check_Copiar.Value = 0
        Pic_Copiar.Picture = Form_Skin.Check_Normal.Picture
    End If
End Sub

Private Sub Pic_Direita_Click()
    'Opção escolhida
    Label_Posicionamento.Caption = "Direita"
    Pic_Esquerda.Picture = Form_Skin.Opcao_Normal.Picture
    Pic_Direita.Picture = Form_Skin.Opcao_Over.Picture
    Opcao_Direita.Value = True
End Sub

Private Sub Pic_Esquerda_Click()
    'Opção escolhida
    Label_Posicionamento.Caption = "Esquerda"
    Pic_Esquerda.Picture = Form_Skin.Opcao_Over.Picture
    Pic_Direita.Picture = Form_Skin.Opcao_Normal.Picture
    Opcao_Esquerda.Value = True
End Sub

Private Sub Pic_Programa_Click()
    'Opção escolhida
    Label_Video.Caption = "Programa"
    Pic_Programa.Picture = Form_Skin.Opcao_Over.Picture
    Pic_Wmp.Picture = Form_Skin.Opcao_Normal.Picture
    Opcao_Programa.Value = True
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
    Dim iDX As Long, iDY As Long
    Dim POINT As POINTAPI
    If Not bMoveFrom Then Exit Sub
    GetCursorPos POINT
    iDX& = (POINT.x - LastPoint.x) * iTPPX&
    iDY& = (POINT.y - LastPoint.y) * iTPPY&
    LastPoint.x = POINT.x
    LastPoint.y = POINT.y
    Me.Move Me.Left + iDX&, Me.Top + iDY&
End Sub

Private Sub Skin_Top_Centro_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Largar o formulário para a posição final
    bMoveFrom = False
End Sub

Public Sub Desenhar_Formulario()
    'Procedimento para ajustar os objectos
    With Shape_Contorno
        .Height = Me.ScaleHeight
        .Top = 0
        .Width = Me.ScaleWidth
        .Left = 0
    End With
    
    With Barra_ControlBox
        .Height = Fundo_Barra_ControlBox.Height
        .Top = 0
        .Width = Me.ScaleWidth
        .Left = 0
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
    
    With Line1
        .Y1 = Label_Topico_Selecionado.Top + Label_Topico_Selecionado.Height + 5
        .Y2 = Label_Topico_Selecionado.Top + Label_Topico_Selecionado.Height + 5
        .X1 = Label_Topico_Selecionado.Left
        '.X2 = Label_Topico_Selecionado.Left
    End With
    
    With Barra_Botoes
        .Height = Fundo_Barra_Botoes.Height
        .Top = Me.ScaleHeight - .Height - 1
        .Width = Me.ScaleWidth - 2
        .Left = 1
    End With
    
    With Fundo_Barra_Botoes
        .Stretch = True
        .Top = 0
        .Width = Barra_Botoes.Width
        .Left = 0
    End With
    
    With Botao_Cancelar
        .Top = 16
        .Height = Form_Skin.Botao_Normal.Height
        .Width = Form_Skin.Botao_Normal.Width
        .Left = Barra_Botoes.Width - .Width - 9
    End With
    
    With Contorno_Cancelar
        .Top = 0
        .Height = Botao_Cancelar.ScaleHeight
        .Left = 0
        .Width = Botao_Cancelar.ScaleWidth
    End With
    
    With Label_Cancelar
        .Top = (Botao_Cancelar.ScaleHeight - .Height) / 2
        .Width = Botao_Cancelar.ScaleWidth
    End With
    
    With Botao_Ok
        .Top = 16
        .Height = Form_Skin.Botao_Normal.Height
        .Width = Form_Skin.Botao_Normal.Width
        .Left = Botao_Cancelar.Left - .Width - 9
    End With
    
    With Contorno_Ok
        .Top = 0
        .Height = Botao_Ok.ScaleHeight
        .Left = 0
        .Width = Botao_Ok.ScaleWidth
    End With
    
    With Label_Ok
        .Top = (Botao_Ok.ScaleHeight - .Height) / 2
        .Width = Botao_Ok.ScaleWidth
    End With
    
    With Botao_Aplicar
        .Top = 16
        .Height = Form_Skin.Botao_Normal.Height
        .Width = Form_Skin.Botao_Normal.Width
        .Left = Botao_Ok.Left - .Width - 9
    End With
    
    With Contorno_Aplicar
        .Top = 0
        .Height = Botao_Aplicar.ScaleHeight
        .Left = 0
        .Width = Botao_Aplicar.ScaleWidth
    End With
    
    With Label_Aplicar
        .Top = (Botao_Aplicar.ScaleHeight - .Height) / 2
        .Width = Botao_Aplicar.ScaleWidth
    End With
    
    With Lista_Opcoes
        .Height = Me.ScaleHeight - Barra_ControlBox.Height - Barra_Botoes.Height - 2
        .Top = Barra_ControlBox.Top + Barra_ControlBox.Height
        .Left = 1
    End With
    
    With Linha_Vertical
        .Y2 = Me.ScaleHeight
        .Y1 = Lista_Opcoes.Top
        .X1 = Lista_Opcoes.Left + Lista_Opcoes.Width
        .X2 = Lista_Opcoes.Left + Lista_Opcoes.Width
    End With
    
    'Chamar o procedimento
    Posicionar_Frames
    
    With Pic_Actualizacoes
        .Height = Form_Skin.Check_Normal.Height
        .Width = Form_Skin.Check_Normal.Width
    End With
    
    With Pic_Abrir_Relatorio
        .Height = Form_Skin.Check_Normal.Height
        .Width = Form_Skin.Check_Normal.Width
    End With
    
    With Pic_Copiar
        .Height = Form_Skin.Check_Normal.Height
        .Width = Form_Skin.Check_Normal.Width
    End With
    
    With Pic_Esquerda
        .Height = Form_Skin.Opcao_Normal.Height
        .Width = Form_Skin.Opcao_Normal.Width
    End With
    
    With Pic_Direita
        .Height = Form_Skin.Opcao_Normal.Height
        .Width = Form_Skin.Opcao_Normal.Width
    End With
    
    With Pic_Programa
        .Height = Form_Skin.Opcao_Normal.Height
        .Width = Form_Skin.Opcao_Normal.Width
    End With
    
    With Pic_Wmp
        .Height = Form_Skin.Opcao_Normal.Height
        .Width = Form_Skin.Opcao_Normal.Width
    End With
    
    With Botao_Copia_Seguranca
        .Height = Form_Skin.Botao_Normal_2.Height
        .Width = Form_Skin.Botao_Normal_2.Width
        .Left = 0
    End With
    
    With Contorno_Copia_Seguranca
        .Top = 0
        .Height = Botao_Copia_Seguranca.ScaleHeight
        .Left = 0
        .Width = Botao_Copia_Seguranca.ScaleWidth
    End With
    
    With Label_Copia_Seguranca
        .Top = (Botao_Copia_Seguranca.ScaleHeight - .Height) / 2
        .Width = Botao_Copia_Seguranca.ScaleWidth
    End With
    
    With Botao_Reset
        .Height = Form_Skin.Botao_Normal_2.Height
        .Width = Form_Skin.Botao_Normal_2.Width
        .Left = 0
    End With
    
    With Contorno_Reset
        .Top = 0
        .Height = Botao_Reset.ScaleHeight
        .Left = 0
        .Width = Botao_Reset.ScaleWidth
    End With
    
    With Label_Reset
        .Top = (Botao_Reset.ScaleHeight - .Height) / 2
        .Width = Botao_Reset.ScaleWidth
    End With
End Sub

Public Sub Ocultar_Frames()
    'Procedimento para ocultar as frames
    Frame_Base_Dados.Visible = False
    Frame_Visualizacao.Visible = False
    Frame_Video.Visible = False
    Frame_Relatorio.Visible = False
    Frame_Reiniciar.Visible = False
    Frame_Capas.Visible = False
    Frame_Actualizacoes.Visible = False
End Sub

Public Sub Posicionar_Frames()
    'Procedimento para posicionar as frames
    If Frame_Informacao.Visible = True Then
        Frame_Base_Dados.Top = Frame_Informacao.Top + Frame_Informacao.ScaleHeight
    Else
        Frame_Base_Dados.Top = Frame_Informacao.Top
    End If
    
    'Posicionar as restantes frames consoante a altura da frame senha
    Frame_Visualizacao.Top = Frame_Base_Dados.Top
    Frame_Visualizacao.Left = Frame_Base_Dados.Left
    Frame_Video.Top = Frame_Base_Dados.Top
    Frame_Video.Left = Frame_Base_Dados.Left
    Frame_Relatorio.Top = Frame_Base_Dados.Top
    Frame_Relatorio.Left = Frame_Base_Dados.Left
    Frame_Reiniciar.Top = Frame_Base_Dados.Top
    Frame_Reiniciar.Left = Frame_Base_Dados.Left
    Frame_Capas.Top = Frame_Base_Dados.Top
    Frame_Capas.Left = Frame_Base_Dados.Left
    Frame_Actualizacoes.Top = Frame_Base_Dados.Top
    Frame_Actualizacoes.Left = Frame_Base_Dados.Left
End Sub

Private Sub Pic_Wmp_Click()
    'Opção escolhida
    Label_Video.Caption = "Windows media player"
    Pic_Programa.Picture = Form_Skin.Opcao_Normal.Picture
    Pic_Wmp.Picture = Form_Skin.Opcao_Over.Picture
    Opcao_Wmp.Value = True
End Sub

Public Sub Repor_Imagens()
    'Procedimento para repor as imagens originais dos botões após o over
    If Botao_Ok.Picture <> Form_Skin.Botao_Normal.Picture Then
        Botao_Ok.Picture = Form_Skin.Botao_Normal.Picture
    End If
    
    If Botao_Cancelar.Picture <> Form_Skin.Botao_Normal.Picture Then
        Botao_Cancelar.Picture = Form_Skin.Botao_Normal.Picture
    End If
    
    If Botao_Aplicar.Picture <> Form_Skin.Botao_Normal.Picture Then
        Botao_Aplicar.Picture = Form_Skin.Botao_Normal.Picture
    End If
End Sub

Private Sub Text_Localizacao_BD_GotFocus()
    'Contorno da text box ao receber o focus
    Shape_Localizacao_BD.BorderColor = &HE6964D    'Azul
End Sub

Private Sub Text_Localizacao_BD_LostFocus()
    'Contorno da text box ao perder o focus
    Shape_Localizacao_BD.BorderColor = &HC0C0C0      'Cinzento
End Sub

Function CopiarArquivo(Origem As String, Destino As String) As Single
    'Função para iniciar a cópia de arquivos
    Static Buf As String
    Dim BTest As Long
    Dim FSize As Long
    Dim Chunk As Integer
    Dim F1 As Integer
    Dim F2 As Integer
    
    Const BUFSIZE = 1024       'define o tamanho do buffer
    
    If Len(Dir(Destino)) Then      'verifica se o arquivo de destino ja existe
       'Resposta = MsgBox(Destino + Chr(10) + Chr(10) + "Arquivo já existe. Deseja sobrescrever o arquivo existente ?", vbYesNo + vbQuestion) 'exibe ao usuário uma caixa de mensagem
       Mensagem_de_Aviso "Questão", "Uma cópia de segurança já foi efectuada hoje." & vbNewLine & "Pretende substitui-la?"
       If Resposta = "Nao" Then 'Se clicou no botão Não
          Exit Function        'sai da rotina
       Else                    'senao
          Kill Destino             'exclui o arquivo existente e continua a executar o codigo
       End If
    End If
     
    'On Error GoTo FileCopyError  'se houver erro trata aqui
    F1 = FreeFile                'retorna o numero do arquivo disponivel
    Open Origem For Binary As F1    'abre o arquivo de destino
    F2 = FreeFile                'retorna o numero do arquivo disponivel
    Open Destino For Binary As F2    'abre o arquivo de destino
     
    FSize = LOF(F1)
    BTest = FSize - LOF(F2)
    
    Do
    If BTest < BUFSIZE Then
       Chunk = BTest
    Else
       Chunk = BUFSIZE
    End If
          
    Buf = String(Chunk, " ")
    Get F1, , Buf
    Put F2, , Buf
    BTest = FSize - LOF(F2)
    
    '''pbCopiaArquivos.Value = (100 - Int(100 * BTest / FSize)) 'avanca com a barra de progrossse durante a copia
    
    Loop Until BTest = 0
    Close F1 'fecha o fonte
    Close F2 'fecha o destino
    CopiarArquivo = FSize
    
    Mensagem_de_Aviso "Informação", "A cópia de segurança foi efectuada com sucesso."
        
    '''pbCopiaArquivos.Value = 0 'retorna a barra de progresso para o valor zero
    Exit Function      'sai da rotina
    
FileCopyError:         'trata o erro aqui
    Mensagem_de_Aviso "Erro", "Ocorreu um erro ao efectuar a cópia de segurança."
    Close F1           'fecha a fonte
    Close F2           'fecha o destino
    Unload Me
    Exit Function      'sai da rotina
End Function

