VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form_Dados 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   0  'None
   ClientHeight    =   6900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10455
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
   Icon            =   "Form_Dados.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   460
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   697
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text_Origem 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF80FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   6720
      TabIndex        =   40
      Top             =   5520
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.TextBox Text_Destino 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF80FF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3600
      TabIndex        =   39
      Top             =   5520
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.PictureBox Pic_Capa 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1920
      Left            =   720
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1920
      Begin VB.Image Imagem_Capa 
         Height          =   1920
         Left            =   0
         Picture         =   "Form_Dados.frx":000C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1920
      End
   End
   Begin VB.PictureBox Frame_Erro 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3720
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   433
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   6495
      Begin VB.Image Image_Erro 
         Enabled         =   0   'False
         Height          =   210
         Left            =   0
         Picture         =   "Form_Dados.frx":C04E
         Top             =   45
         Width           =   210
      End
      Begin VB.Label Label_Erro 
         AutoSize        =   -1  'True
         BackColor       =   &H00F5F5F5&
         BackStyle       =   0  'Transparent
         Caption         =   "Campo de preenchimento obrigatório."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   360
         TabIndex        =   35
         Top             =   60
         Width           =   3270
      End
      Begin VB.Shape Shape_Erro 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H008080FF&
         Height          =   315
         Left            =   0
         Top             =   0
         Width           =   6375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Lista_Categorias 
      Height          =   1440
      Left            =   5040
      TabIndex        =   19
      Top             =   1755
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   2540
      _Version        =   393216
      Rows            =   19
      Cols            =   16
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   16777215
      ForeColorFixed  =   0
      BackColorSel    =   14200408
      ForeColorSel    =   16777215
      BackColorBkg    =   16777215
      GridColor       =   14737632
      GridColorFixed  =   14737632
      Redraw          =   -1  'True
      FocusRect       =   0
      GridLines       =   0
      GridLinesFixed  =   0
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Lista_Tipo 
      Height          =   990
      Left            =   5040
      TabIndex        =   22
      Top             =   2115
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1746
      _Version        =   393216
      Rows            =   5
      Cols            =   16
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   16777215
      ForeColorFixed  =   0
      BackColorSel    =   14200408
      ForeColorSel    =   16777215
      BackColorBkg    =   16777215
      GridColor       =   14737632
      GridColorFixed  =   14737632
      Redraw          =   -1  'True
      FocusRect       =   0
      GridLines       =   0
      GridLinesFixed  =   0
      ScrollBars      =   0
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin VB.TextBox Text_Id 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
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
      Height          =   315
      Left            =   0
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text_Capa 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   120
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   4560
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2115
      Left            =   5040
      ScaleHeight     =   141
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   345
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   3240
      Width           =   5175
      Begin VB.TextBox Text_Observacoes 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   2085
         IMEMode         =   3  'DISABLE
         Left            =   15
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   15
         Width           =   5025
      End
      Begin VB.Shape Shape_Observacoes 
         BorderColor     =   &H00C0C0C0&
         Height          =   2115
         Left            =   0
         Top             =   0
         Width           =   5055
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   5040
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   345
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2880
      Width           =   5175
      Begin VB.PictureBox Seta_Directorio 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4755
         Picture         =   "Form_Dados.frx":C2F8
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   15
         Width           =   285
      End
      Begin VB.TextBox Text_Directorio 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   15
         TabIndex        =   4
         Top             =   15
         Width           =   4740
      End
      Begin VB.Shape Shape_Directorio 
         BorderColor     =   &H00C0C0C0&
         Height          =   315
         Left            =   0
         Top             =   0
         Width           =   5055
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   5040
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   345
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2520
      Width           =   5175
      Begin VB.TextBox Text_Actores 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   15
         TabIndex        =   3
         Top             =   15
         Width           =   5025
      End
      Begin VB.Shape Shape_Actores 
         BorderColor     =   &H00C0C0C0&
         Height          =   315
         Left            =   0
         Top             =   0
         Width           =   5055
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   5040
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   201
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1800
      Width           =   3015
      Begin VB.PictureBox Seta_Tipo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2475
         Picture         =   "Form_Dados.frx":C61D
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   15
         Width           =   285
      End
      Begin VB.TextBox Text_Tipo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   15
         TabIndex        =   2
         Top             =   15
         Width           =   2505
      End
      Begin VB.Shape Shape_Tipo 
         BorderColor     =   &H00C0C0C0&
         Height          =   315
         Left            =   0
         Top             =   0
         Width           =   2775
      End
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   5040
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   201
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1440
      Width           =   3015
      Begin VB.PictureBox Seta_Categoria 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2475
         Picture         =   "Form_Dados.frx":CB0F
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   15
         Width           =   285
      End
      Begin VB.TextBox Text_Categoria 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   2460
      End
      Begin VB.Shape Shape_Categoria 
         BorderColor     =   &H00C0C0C0&
         Height          =   315
         Left            =   0
         Top             =   0
         Width           =   2775
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   5040
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   345
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1080
      Width           =   5175
      Begin VB.TextBox Text_Nome 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   15
         TabIndex        =   0
         Top             =   15
         Width           =   5025
      End
      Begin VB.Shape Shape_Nome 
         BorderColor     =   &H00C0C0C0&
         Height          =   315
         Left            =   0
         Top             =   0
         Width           =   5055
      End
   End
   Begin VB.PictureBox Barra_ControlBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00313131&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   0
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   681
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   0
      Width           =   10215
      Begin VB.Label Label_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00272727&
         BackStyle       =   0  'Transparent
         Caption         =   "Novo registo"
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
         TabIndex        =   16
         Top             =   120
         Width           =   1245
      End
      Begin VB.Image Botao_Fechar 
         Height          =   225
         Left            =   9480
         Picture         =   "Form_Dados.frx":D001
         ToolTipText     =   "Fechar"
         Top             =   120
         Width           =   240
      End
      Begin VB.Image Fundo_Barra_ControlBox 
         Enabled         =   0   'False
         Height          =   465
         Left            =   0
         Picture         =   "Form_Dados.frx":D313
         Top             =   0
         Width           =   585
      End
   End
   Begin VB.PictureBox Barra_Botoes 
      Appearance      =   0  'Flat
      BackColor       =   &H00D9D9D9&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   900
      Left            =   0
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   697
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6000
      Width           =   10455
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
         Left            =   8520
         Picture         =   "Form_Dados.frx":D658
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   61
         TabIndex        =   7
         Top             =   240
         Width           =   915
         Begin VB.Label Label_Cancelar 
            Alignment       =   2  'Center
            BackColor       =   &H00272727&
            BackStyle       =   0  'Transparent
            Caption         =   "Cancelar"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   0
            TabIndex        =   17
            Top             =   135
            Width           =   915
         End
         Begin VB.Shape Contorno_Cancelar 
            BorderColor     =   &H00E6964D&
            Height          =   495
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Visible         =   0   'False
            Width           =   915
         End
      End
      Begin VB.PictureBox Botao_Ok 
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
         Left            =   7320
         Picture         =   "Form_Dados.frx":EE52
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   61
         TabIndex        =   6
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
            TabIndex        =   14
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
         Picture         =   "Form_Dados.frx":1064C
         Top             =   0
         Width           =   585
      End
   End
   Begin VB.TextBox Text_classificacao 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
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
      Height          =   315
      Left            =   6360
      TabIndex        =   11
      Text            =   "0"
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00E0E0E0&
      Height          =   2175
      Left            =   600
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label_Remover 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remover"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1200
      TabIndex        =   37
      Top             =   3840
      Width           =   780
   End
   Begin VB.Label Label_Selecionar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selecionar"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1200
      TabIndex        =   36
      Top             =   3480
      Width           =   900
   End
   Begin VB.Image Image7 
      Enabled         =   0   'False
      Height          =   240
      Left            =   840
      Picture         =   "Form_Dados.frx":109B7
      Top             =   3480
      Width           =   240
   End
   Begin VB.Image Image6 
      Enabled         =   0   'False
      Height          =   240
      Left            =   840
      Picture         =   "Form_Dados.frx":10CF9
      Top             =   3840
      Width           =   240
   End
   Begin VB.Label Label_Campo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Observações:"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   8
      Left            =   3720
      TabIndex        =   30
      Top             =   3315
      Width           =   1185
   End
   Begin VB.Label Label_Campo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Directório:"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   3720
      TabIndex        =   28
      Top             =   2955
      Width           =   915
   End
   Begin VB.Label Label_Campo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Actores:"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   6
      Left            =   3720
      TabIndex        =   26
      Top             =   2595
      Width           =   720
   End
   Begin VB.Line Linha_Vertical 
      BorderColor     =   &H00E0E0E0&
      X1              =   224
      X2              =   224
      Y1              =   0
      Y2              =   288
   End
   Begin VB.Shape Shape_Contorno 
      BorderColor     =   &H00C0C0C0&
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   5070
      Picture         =   "Form_Dados.frx":1103B
      Top             =   2205
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   5325
      Picture         =   "Form_Dados.frx":1137D
      Top             =   2205
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   5565
      Picture         =   "Form_Dados.frx":116BF
      Top             =   2205
      Width           =   240
   End
   Begin VB.Image Image4 
      Height          =   240
      Left            =   5805
      Picture         =   "Form_Dados.frx":11A01
      Top             =   2205
      Width           =   240
   End
   Begin VB.Image Image5 
      Height          =   240
      Left            =   6045
      Picture         =   "Form_Dados.frx":11D43
      Top             =   2205
      Width           =   240
   End
   Begin VB.Label Label_Campo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Classificacao:"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   7
      Left            =   3720
      TabIndex        =   12
      Top             =   2235
      Width           =   1185
   End
   Begin VB.Label Label_Campo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome:"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   3720
      TabIndex        =   10
      Top             =   1155
      Width           =   570
   End
   Begin VB.Label Label_Campo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Categoria:"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   3720
      TabIndex        =   9
      Top             =   1515
      Width           =   915
   End
   Begin VB.Label Label_Campo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo:"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   3720
      TabIndex        =   8
      Top             =   1875
      Width           =   435
   End
   Begin VB.Shape Shape_Foto 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   5535
      Left            =   0
      Top             =   480
      Width           =   3300
   End
End
Attribute VB_Name = "Form_Dados"
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

Public Mode As String

'Variável para saber se o utilizador está a criar ou editar um registo
Dim Janela_Dialogo As New Class_CommDialog

'Variáveis do commondialog
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Boolean
Const DI_MASK = &H1
Const DI_IMAGE = &H2
Const DI_NORMAL = &H3
Const DI_COMPAT = &H4
Const DI_DEFAULTSIZE = &H8

Private Sub Barra_Botoes_Click()
    'Atalho para
    Ocultar_Objectos
End Sub

Private Sub Barra_Botoes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Chamar o procedimento
    Repor_Imagens
End Sub

Private Sub Barra_ControlBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Dim POINT As POINTAPI
    GetCursorPos POINT
    LastPoint.X = POINT.X
    LastPoint.Y = POINT.Y
    bMoveFrom = True
End Sub

Private Sub Barra_ControlBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Dim iDX As Long, iDY As Long
    Dim POINT As POINTAPI
    If Not bMoveFrom Then Exit Sub
    GetCursorPos POINT
    iDX& = (POINT.X - LastPoint.X) * iTPPX&
    iDY& = (POINT.Y - LastPoint.Y) * iTPPY&
    LastPoint.X = POINT.X
    LastPoint.Y = POINT.Y
    Me.Move Me.Left + iDX&, Me.Top + iDY&
End Sub

Private Sub Barra_ControlBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    bMoveFrom = False
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

Private Sub Botao_Cancelar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Animar o botão
    If Botao_Cancelar.Picture = Form_Skin.Botao_Over.Picture Then Exit Sub
    Repor_Imagens
    Botao_Cancelar.Picture = Form_Skin.Botao_Over.Picture
End Sub

Private Sub Botao_Fechar_Click()
    'Atalho para
    Botao_Cancelar_Click
End Sub

Private Sub Botao_Ok_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Animar o botão
    If Botao_Ok.Picture = Form_Skin.Botao_Over.Picture Then Exit Sub
    Repor_Imagens
    Botao_Ok.Picture = Form_Skin.Botao_Over.Picture
End Sub

Private Sub Form_Click()
    'Atalho para
    Ocultar_Objectos
End Sub

Private Sub Form_Load()
    'Propriedades iniciais do formulário
    iTPPX& = Screen.TwipsPerPixelX
    iTPPY& = Screen.TwipsPerPixelY
    
    'Chamar procedimento
    Desenhar_Formulario
    Formatar_Lista_Categorias
    Preenche_Lista
    Formatar_Lista_Tipo
    Preenche_Lista_Tipo
End Sub

Public Sub Preenche_Lista()
    'Preencher lista de categorias
    With Lista_Categorias
        .Clear
        .TextMatrix(1, 0) = "Acção"
        .TextMatrix(2, 0) = "Animação"
        .TextMatrix(3, 0) = "Aventura"
        .TextMatrix(4, 0) = "Comédia"
        .TextMatrix(5, 0) = "Crime"
        .TextMatrix(6, 0) = "Documentário"
        .TextMatrix(7, 0) = "Desporto"
        .TextMatrix(8, 0) = "Drama"
        .TextMatrix(9, 0) = "Faroeste"
        .TextMatrix(10, 0) = "Ficção cientifica"
        .TextMatrix(11, 0) = "Guerra"
        .TextMatrix(12, 0) = "Musical"
        .TextMatrix(13, 0) = "Policial"
        .TextMatrix(14, 0) = "Romance"
        .TextMatrix(15, 0) = "Série"
        .TextMatrix(16, 0) = "Suspense"
        .TextMatrix(17, 0) = "Terror"
        .TextMatrix(18, 0) = "Outra"
    End With
End Sub

Public Sub Preenche_Lista_Tipo()
    'Preencher lista de categorias
    With Lista_Tipo
        .Clear
        .TextMatrix(1, 0) = "Clip"
        .TextMatrix(2, 0) = "Filme"
        .TextMatrix(3, 0) = "Video clip"
        .TextMatrix(4, 0) = "Outro"
    End With
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Chamar o procedimento
    Repor_Imagens
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    'Chamar o procedimento
    Desenhar_Formulario
End Sub

Private Sub Image1_Click()
    'Escolher classificacao
    Image1.Picture = Form_Skin.Estrela_Over.Picture
    Image2.Picture = Form_Skin.Estrela_Normal.Picture
    Image3.Picture = Form_Skin.Estrela_Normal.Picture
    Image4.Picture = Form_Skin.Estrela_Normal.Picture
    Image5.Picture = Form_Skin.Estrela_Normal.Picture
    Text_Classificacao.Text = "1"
End Sub

Private Sub Image2_Click()
    'Escolher classificacao
    Image1.Picture = Form_Skin.Estrela_Over.Picture
    Image2.Picture = Form_Skin.Estrela_Over.Picture
    Image3.Picture = Form_Skin.Estrela_Normal.Picture
    Image4.Picture = Form_Skin.Estrela_Normal.Picture
    Image5.Picture = Form_Skin.Estrela_Normal.Picture
    Text_Classificacao.Text = "2"
End Sub

Private Sub Image3_Click()
    'Escolher classificacao
    Image1.Picture = Form_Skin.Estrela_Over.Picture
    Image2.Picture = Form_Skin.Estrela_Over.Picture
    Image3.Picture = Form_Skin.Estrela_Over.Picture
    Image4.Picture = Form_Skin.Estrela_Normal.Picture
    Image5.Picture = Form_Skin.Estrela_Normal.Picture
    Text_Classificacao.Text = "3"
End Sub

Private Sub Image4_Click()
    'Escolher classificacao
    Image1.Picture = Form_Skin.Estrela_Over.Picture
    Image2.Picture = Form_Skin.Estrela_Over.Picture
    Image3.Picture = Form_Skin.Estrela_Over.Picture
    Image4.Picture = Form_Skin.Estrela_Over.Picture
    Image5.Picture = Form_Skin.Estrela_Normal.Picture
    Text_Classificacao.Text = "4"
End Sub

Private Sub Image5_Click()
    'Escolher classificacao
    Image1.Picture = Form_Skin.Estrela_Over.Picture
    Image2.Picture = Form_Skin.Estrela_Over.Picture
    Image3.Picture = Form_Skin.Estrela_Over.Picture
    Image4.Picture = Form_Skin.Estrela_Over.Picture
    Image5.Picture = Form_Skin.Estrela_Over.Picture
    Text_Classificacao.Text = "5"
End Sub

Private Sub Imagem_Capa_Click()
    'Escolher a capa do filme
'    With Form_Pesquisa_Capas
'        .Show vbModal
'    End With
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

Private Sub Label_Capa_Click()
    'Atalho para
    Imagem_Capa_Click
End Sub


Private Sub Label_Cancelar_Click()
    'Fechar formulario
    Unload Me
End Sub

Private Sub Label_Cancelar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Animar o botão
    If Botao_Cancelar.Picture = Form_Skin.Botao_Over.Picture Then Exit Sub
    Repor_Imagens
    Botao_Cancelar.Picture = Form_Skin.Botao_Over.Picture
End Sub

Private Sub Label_Ok_Click()
    'Guardar dados
    Shape_Nome.BorderColor = &HC0C0C0
    Shape_Actores.BorderColor = &HC0C0C0
    Shape_Categoria.BorderColor = &HC0C0C0
    Frame_Erro.Visible = False
    
    'Remover os espaços á esquerda e á direita das strings
    Text_Nome.Text = Trim(Text_Nome.Text)
    Text_Categoria.Text = Trim(Text_Categoria.Text)
    Text_Tipo.Text = Trim(Text_Tipo.Text)
    Text_Actores.Text = Trim(Text_Actores.Text)
    Text_Directorio.Text = Trim(Text_Directorio.Text)
    'Text_Observacoes.Text = Trim(Text_Observacoes.Text)
    
    If Text_Nome.Text = Empty Then
        Frame_Erro.Visible = True
        Label_Erro.Caption = "O campo 'Nome' é de preenchimento obrigatório."
        Text_Nome.SetFocus
        Exit Sub
        
    ElseIf Text_Tipo.Text = Empty Then
        Frame_Erro.Visible = True
        Label_Erro.Caption = "O campo 'Tipo' é de preenchimento obrigatório."
        Text_Tipo.SetFocus
        Exit Sub
    
    ElseIf Text_Categoria.Text = Empty Then
        Frame_Erro.Visible = True
        Label_Erro.Caption = "O campo 'Categoria' é de preenchimento obrigatório."
        Text_Categoria.SetFocus
        Exit Sub
    End If
    
    'Verificar se é para copiar a capa do filme para a pasta do programa
    If Mode = "novo" Then
        If Text_Capa.Text <> Empty Then
            If Form_Opcoes.Check_Copiar.Value = 1 Then
                Text_Origem.Text = Text_Capa.Text
                Text_Destino.Text = App.path & "\Covers\" & Text_Id.Text & ".jpg"
                CopiarArquivo Text_Origem.Text, Text_Destino.Text
                Text_Capa.Text = Text_Destino.Text
            End If
        End If
    End If
    
    'Caso esteja tudo in guarda ou edita o registo
    With Form_Principal
        If Mode = "editar" Then
            .Cnn_Filmes.Execute "Update Tabela_Filmes Set Id = '" & Text_Id.Text & "', Video = '" & Text_Nome.Text & "', Categoria = '" & Text_Categoria.Text & "', Tipo = '" & Text_Tipo.Text & "', Classificacao = '" & Text_Classificacao.Text & "', Actores = '" & Text_Actores.Text & "', Observacoes = '" & Text_Observacoes.Text & "', Directorio = '" & Text_Directorio.Text & "', Capa = '" & Text_Capa.Text & "'   where Id = '" & Text_Id.Text & "'"
            With Form_Principal.Lista_Filmes
                .TextMatrix(.Row, 0) = Text_Id.Text
                .TextMatrix(.Row, 1) = Text_Nome.Text
                .TextMatrix(.Row, 2) = Text_Categoria.Text
                .TextMatrix(.Row, 3) = Text_Tipo.Text
                .TextMatrix(.Row, 4) = Text_Classificacao.Text
                .TextMatrix(.Row, 5) = Text_Actores.Text
                .TextMatrix(.Row, 6) = Text_Observacoes.Text
                .TextMatrix(.Row, 7) = Text_Directorio.Text
                .TextMatrix(.Row, 8) = Text_Capa.Text
            End With
            
        ElseIf Mode = "novo" Then
            .Cnn_Filmes.Execute "Insert Into Tabela_Filmes Values('" & Text_Id.Text & "','" & Text_Nome.Text & "', '" & Text_Categoria.Text & "', '" & Text_Tipo.Text & "', '" & Text_Classificacao.Text & "', '" & Text_Actores.Text & "', '" & Text_Observacoes.Text & "', '" & Text_Directorio.Text & "', '" & Text_Capa.Text & "')"
            .Rs_Filmes.Requery 1
            .Iniciando = True
            .Preenche_Lista
        End If
        
        Form_Principal.Lista_Filmes_Click
    End With
    Unload Me
End Sub

Private Sub Label_Ok_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Animar o botão
    If Botao_Ok.Picture = Form_Skin.Botao_Over.Picture Then Exit Sub
    Repor_Imagens
    Botao_Ok.Picture = Form_Skin.Botao_Over.Picture
End Sub

Private Sub Label_Remover_Click()
    'Remover a capa do filme
    Text_Capa.Text = Empty
    Imagem_Capa.Picture = Form_Skin.Image_Sem_Capa.Picture
End Sub

Private Sub Label_Remover_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Animar a label
    If Label_Remover.FontUnderline = True Then Exit Sub
    Label_Selecionar.FontUnderline = False
    Label_Remover.FontUnderline = True
End Sub

Private Sub Label_Selecionar_Click()
    'Selecionar a capa do filme
    With Janela_Dialogo
        .DialogTitle = "Selecionar capa"
        .CancelError = False
        .Filter = "Imagens (*.jpg;*.jpeg;*.bmp)|*.jpg;*.jpeg;*.bmp|"
        .ShowOpen
        .hIcon = Me.Icon
        If Len(.FileName) <> 0 Then
            Text_Capa.Text = .FileName
            Carregar_Capa
        End If
    End With
End Sub

Private Sub Label_Selecionar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Animar a label
    If Label_Selecionar.FontUnderline = True Then Exit Sub
    Label_Selecionar.FontUnderline = True
    Label_Remover.FontUnderline = False
End Sub

Private Sub Label_Titulo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Dim POINT As POINTAPI
    GetCursorPos POINT
    LastPoint.X = POINT.X
    LastPoint.Y = POINT.Y
    bMoveFrom = True
End Sub

Private Sub Label_Titulo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Dim iDX As Long, iDY As Long
    Dim POINT As POINTAPI
    If Not bMoveFrom Then Exit Sub
    GetCursorPos POINT
    iDX& = (POINT.X - LastPoint.X) * iTPPX&
    iDY& = (POINT.Y - LastPoint.Y) * iTPPY&
    LastPoint.X = POINT.X
    LastPoint.Y = POINT.Y
    Me.Move Me.Left + iDX&, Me.Top + iDY&
End Sub

Private Sub Label_Titulo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    bMoveFrom = False
End Sub

Private Sub Lista_Categorias_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Selecionar as linhas da lista com o mouse
    If Lista_Categorias.Rows > 1 Then
        If Lista_Categorias.Row <> Lista_Categorias.MouseRow And Lista_Categorias.MouseRow > 0 Then
            Lista_Categorias.Col = 0
            Lista_Categorias.Row = Lista_Categorias.MouseRow
            Lista_Categorias.ColSel = Lista_Categorias.Cols - 1
        End If
    End If
End Sub

Private Sub Lista_Tipo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Selecionar as linhas da lista com o mouse
    If Lista_Tipo.Rows > 1 Then
        If Lista_Tipo.Row <> Lista_Tipo.MouseRow And Lista_Tipo.MouseRow > 0 Then
            Lista_Tipo.Col = 0
            Lista_Tipo.Row = Lista_Tipo.MouseRow
            Lista_Tipo.ColSel = Lista_Tipo.Cols - 1
        End If
    End If
End Sub

Private Sub Skin_Top_Centro_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Capturar a posição de x e y
    Dim POINT As POINTAPI
    GetCursorPos POINT
    LastPoint.X = POINT.X
    LastPoint.Y = POINT.Y
    bMoveFrom = True
End Sub

Private Sub Skin_Top_Centro_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Mover o formulário e obter a posição de x e y
    Dim iDX As Long, iDY As Long
    Dim POINT As POINTAPI
    If Not bMoveFrom Then Exit Sub
    GetCursorPos POINT
    iDX& = (POINT.X - LastPoint.X) * iTPPX&
    iDY& = (POINT.Y - LastPoint.Y) * iTPPY&
    LastPoint.X = POINT.X
    LastPoint.Y = POINT.Y
    Me.Move Me.Left + iDX&, Me.Top + iDY&
End Sub

Private Sub Skin_Top_Centro_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    bMoveFrom = False
End Sub

Public Sub Desenhar_Formulario()
    'Procedimento para construir o formulario, ajustando os objectos
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
    
    With Barra_Botoes
        .Height = Fundo_Barra_Botoes.Height
        .Top = Me.ScaleHeight - .Height - 1
        .Width = Barra_ControlBox.ScaleWidth
        .Left = Barra_ControlBox.Left
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
    
    With Shape_Foto
        .Height = Me.ScaleHeight - Barra_ControlBox.Height - Barra_Botoes.Height - 1
        .Top = Barra_ControlBox.Top + Barra_ControlBox.Height
        .Left = 1
    End With
    
    With Linha_Vertical
        .Y2 = Me.ScaleHeight
        .Y1 = Shape_Foto.Top
        .X1 = Shape_Foto.Left + Shape_Foto.Width
        .X2 = Shape_Foto.Left + Shape_Foto.Width
    End With
End Sub

Private Sub Picture_Categoria_Click()
    'Atalho para
    Ocultar_Objectos
End Sub

Private Sub Lista_Categorias_Click()
    'Selecionar o dia
    Text_Categoria.Text = Lista_Categorias.TextMatrix(Lista_Categorias.Row, 0)
    Ocultar_Objectos
End Sub

Private Sub Lista_Tipo_Click()
    'Selecionar o dia
    Text_Tipo.Text = Lista_Tipo.TextMatrix(Lista_Tipo.Row, 0)
    Ocultar_Objectos
End Sub

Private Sub Seta_Categoria_Click()
    'Ver/ocultar lista
    Lista_Tipo.Visible = False
    
    If Lista_Categorias.Visible = True Then
        Lista_Categorias.Visible = False
    Else
        Lista_Categorias.Visible = True
    End If
End Sub

Private Sub Seta_Tipo_Click()
    'Ver/ocultar lista
    Lista_Categorias.Visible = False
    
    If Lista_Tipo.Visible = True Then
        Lista_Tipo.Visible = False
    Else
        Lista_Tipo.Visible = True
    End If
End Sub

Private Sub Seta_Directorio_Click()
    'Escolher o directório do filme
    With Janela_Dialogo
        .DialogTitle = "Selecionar filme"
        .CancelError = False
        .Filter = "Ficheiros de video (*.avi;*.mpg;*.mpeg;*.mpeg;*.wmv)|*.avi;*.mpg;*.mpeg;*.mpeg;*.wmv|"
        .ShowOpen
        .hIcon = Me.Icon
        If Len(.FileName) <> 0 Then
            Text_Directorio.Text = .FileName
        End If
    End With
End Sub

Public Sub Ocultar_Objectos()
    'Ocultar objectos nao desejados
    Lista_Categorias.Visible = False
    Lista_Tipo.Visible = False
End Sub

Public Sub Formatar_Lista_Categorias()
    'Procedimento para formatar as msflexgrid
    With Lista_Categorias
        .RowHeight(0) = 0
        .ColWidth(0) = 3100
        .ColWidth(1) = 0
    End With
End Sub

Public Sub Formatar_Lista_Tipo()
    'Procedimento para formatar as msflexgrid
    With Lista_Tipo
        .RowHeight(0) = 0
        .ColWidth(0) = 3100
        .ColWidth(1) = 0
    End With
End Sub

Public Sub Carregar_Capa()
    'Carrega imagem da capa, caso exista
    On Error GoTo Corrige_Erro
    If Text_Capa.Text <> "" Then
        Imagem_Capa.Picture = LoadPicture(Text_Capa.Text)
    End If

Exit Sub
Corrige_Erro:
    Select Case Err.Number
        Case "76"
            Imagem_Capa.Picture = Form_Skin.Image_Sem_Capa.Picture
    End Select
End Sub

Private Sub Text_Actores_Click()
    'Atalho para
    Ocultar_Objectos
End Sub

Private Sub Text_Actores_GotFocus()
    'Contorno da text box ao receber o focus
    Shape_Actores.BorderColor = &HE6964D    'Azul
End Sub

Private Sub Text_Actores_LostFocus()
    'Contorno da text box ao perder o focus
    Text_Actores.Text = StrConv(Text_Actores.Text, vbProperCase)
    Shape_Actores.BorderColor = &HC0C0C0      'Cinzento
End Sub

Private Sub Text_Categoria_Click()
    'Atalho para
    Ocultar_Objectos
End Sub

Private Sub Text_Categoria_GotFocus()
    'Contorno da text box ao receber o focus
    Shape_Categoria.BorderColor = &HE6964D    'Azul
End Sub

Private Sub Text_Categoria_KeyDown(KeyCode As Integer, Shift As Integer)
    'Altalho para percorrer as linha da combo de imagens
    If KeyCode = vbKeyUp Then 'Para cima
        If Lista_Categorias.Row <> 1 Then
            Lista_Categorias.Row = Lista_Categorias.Row - 1
            Text_Categoria.Text = Lista_Categorias.TextMatrix(Lista_Categorias.Row, 0)
        End If
    End If
    If KeyCode = vbKeyDown Then 'Para baixo
        If Lista_Categorias.Row <> Lista_Categorias.Rows - 1 Then
            Lista_Categorias.Row = Lista_Categorias.Row + 1
            Text_Categoria.Text = Lista_Categorias.TextMatrix(Lista_Categorias.Row, 0)
        End If
    End If
End Sub

Private Sub Text_Categoria_LostFocus()
    'Contorno da text box ao perder o focus
    Text_Categoria.Text = StrConv(Text_Categoria.Text, vbProperCase)
    Shape_Categoria.BorderColor = &HC0C0C0      'Cinzento
End Sub

Private Sub Text_Directorio_Click()
    'Atalho para
    Ocultar_Objectos
End Sub

Private Sub Text_Directorio_GotFocus()
    'Contorno da text box ao receber o focus
    Shape_Directorio.BorderColor = &HE6964D    'Azul
End Sub

Private Sub Text_Directorio_LostFocus()
    'Contorno da text box ao perder o focus
    Text_Directorio.Text = StrConv(Text_Directorio.Text, vbProperCase)
    Shape_Directorio.BorderColor = &HC0C0C0      'Cinzento
End Sub

Private Sub Text_Nome_Click()
    'Atalho para
    Ocultar_Objectos
End Sub

Private Sub Text_Nome_GotFocus()
    'Contorno da text box ao receber o focus
    Shape_Nome.BorderColor = &HE6964D    'Azul
End Sub

Private Sub Text_Nome_LostFocus()
    'Contorno da text box ao perder o focus
    Text_Nome.Text = StrConv(Text_Nome.Text, vbProperCase)
    Shape_Nome.BorderColor = &HC0C0C0      'Cinzento
End Sub

Private Sub Text_Observacoes_Click()
    'Atalho para
    Ocultar_Objectos
End Sub

Private Sub Text_Observacoes_GotFocus()
    'Contorno da text box ao receber o focus
    Shape_Observacoes.BorderColor = &HE6964D    'Azul
End Sub

Private Sub Text_Observacoes_LostFocus()
    'Contorno da text box ao perder o focus
    Text_Observacoes.Text = StrConv(Text_Observacoes.Text, vbProperCase)
    Shape_Observacoes.BorderColor = &HC0C0C0      'Cinzento
End Sub

Private Sub Text_Tipo_Click()
    'Atalho para
    Ocultar_Objectos
End Sub

Private Sub Text_Tipo_GotFocus()
    'Contorno da text box ao receber o focus
    Shape_Tipo.BorderColor = &HE6964D    'Azul
End Sub

Private Sub Text_Tipo_KeyDown(KeyCode As Integer, Shift As Integer)
    'Altalho para percorrer as linha da combo de imagens
    If KeyCode = vbKeyUp Then 'Para cima
        If Lista_Tipo.Row <> 1 Then
            Lista_Tipo.Row = Lista_Tipo.Row - 1
            Text_Tipo.Text = Lista_Tipo.TextMatrix(Lista_Tipo.Row, 0)
        End If
    End If
    If KeyCode = vbKeyDown Then 'Para baixo
        If Lista_Tipo.Row <> Lista_Tipo.Rows - 1 Then
            Lista_Tipo.Row = Lista_Tipo.Row + 1
            Text_Tipo.Text = Lista_Tipo.TextMatrix(Lista_Tipo.Row, 0)
        End If
    End If
End Sub

Private Sub Text_Tipo_LostFocus()
    'Contorno da text box ao perder o focus
    Text_Tipo.Text = StrConv(Text_Tipo.Text, vbProperCase)
    Shape_Tipo.BorderColor = &HC0C0C0      'Cinzento
End Sub

Public Sub Repor_Imagens()
    'Procedimento para repor as imagens originais dos botões após o over
    If Botao_Ok.Picture <> Form_Skin.Botao_Normal.Picture Then
        Botao_Ok.Picture = Form_Skin.Botao_Normal.Picture
    End If
    
    If Botao_Cancelar.Picture <> Form_Skin.Botao_Normal.Picture Then
        Botao_Cancelar.Picture = Form_Skin.Botao_Normal.Picture
    End If
    
    If Label_Selecionar.FontUnderline = True Then Label_Selecionar.FontUnderline = False
    If Label_Remover.FontUnderline = True Then Label_Remover.FontUnderline = False
End Sub

Public Sub Verificar_Estrelas()
    'Procedimento para verificar a classificação
    'Verificar as estrelas
    If Text_Classificacao.Text = "1" Then
        Image1.Picture = Form_Skin.Estrela_Over.Picture
        Image2.Picture = Form_Skin.Estrela_Normal.Picture
        Image3.Picture = Form_Skin.Estrela_Normal.Picture
        Image4.Picture = Form_Skin.Estrela_Normal.Picture
        Image5.Picture = Form_Skin.Estrela_Normal.Picture
    ElseIf Text_Classificacao.Text = "2" Then
        Image1.Picture = Form_Skin.Estrela_Over.Picture
        Image2.Picture = Form_Skin.Estrela_Over.Picture
        Image3.Picture = Form_Skin.Estrela_Normal.Picture
        Image4.Picture = Form_Skin.Estrela_Normal.Picture
        Image5.Picture = Form_Skin.Estrela_Normal.Picture
    ElseIf Text_Classificacao.Text = "3" Then
        Image1.Picture = Form_Skin.Estrela_Over.Picture
        Image2.Picture = Form_Skin.Estrela_Over.Picture
        Image3.Picture = Form_Skin.Estrela_Over.Picture
        Image4.Picture = Form_Skin.Estrela_Normal.Picture
        Image5.Picture = Form_Skin.Estrela_Normal.Picture
    ElseIf Text_Classificacao.Text = "4" Then
        Image1.Picture = Form_Skin.Estrela_Over.Picture
        Image2.Picture = Form_Skin.Estrela_Over.Picture
        Image3.Picture = Form_Skin.Estrela_Over.Picture
        Image4.Picture = Form_Skin.Estrela_Over.Picture
        Image5.Picture = Form_Skin.Estrela_Normal.Picture
    ElseIf Text_Classificacao.Text = "5" Then
        Image1.Picture = Form_Skin.Estrela_Over.Picture
        Image2.Picture = Form_Skin.Estrela_Over.Picture
        Image3.Picture = Form_Skin.Estrela_Over.Picture
        Image4.Picture = Form_Skin.Estrela_Over.Picture
        Image5.Picture = Form_Skin.Estrela_Over.Picture
    End If
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
       Mensagem_de_Aviso "Questão", "A imagem selecionada já existe na pasta do programa." & vbNewLine & "Pretende substituir o ficheiro existente?"
       If Resposta = "Nao" Then 'Se clicou no botão Não
          Exit Function        'sai da rotina
       Else                    'senao
          Kill Destino             'exclui o arquivo existente e continua a executar o codigo
       End If
    End If
     
    On Error GoTo FileCopyError  'se houver erro trata aqui
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
    
    'MsgBox "Arquivo copiado com sucesso.", vbInformation, "Copia com sucesso"
    Form_Dados.Text_Capa.Text = Text_Destino.Text
        
    '''pbCopiaArquivos.Value = 0 'retorna a barra de progresso para o valor zero
    Exit Function      'sai da rotina
    
FileCopyError:         'trata o erro aqui
    'MsgBox "Erro durante a copia...!, Tente novamente..." 'exibe mensagem de erro
    Form_Dados.Text_Capa.Text = Text_Origem.Text
    Close F1           'fecha a fonte
    Close F2           'fecha o destino
    Unload Me
    Exit Function      'sai da rotina
End Function
