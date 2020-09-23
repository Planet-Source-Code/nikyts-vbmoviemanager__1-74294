VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form_Wmp 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Tela de video"
   ClientHeight    =   7140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11760
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
   Icon            =   "Form_Wmp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   476
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   784
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Frame_Video 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   3375
      Left            =   0
      ScaleHeight     =   225
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   417
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   840
      Width           =   6255
      Begin VB.PictureBox Botao_Amplicar 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   1080
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   136
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   2040
         Begin VB.Label Label_Ampliar 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Ver tela cheia"
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Left            =   0
            TabIndex        =   5
            Top             =   120
            Width           =   2040
         End
      End
      Begin VB.PictureBox Contorno_Down 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
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
         Height          =   45
         Left            =   4440
         ScaleHeight     =   3
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   59
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   2520
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.PictureBox Contorno_Right 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
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
         Height          =   1335
         Left            =   5400
         ScaleHeight     =   89
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   3
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1200
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.PictureBox Contorno_Top 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
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
         Height          =   45
         Left            =   4440
         ScaleHeight     =   3
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   59
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1200
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.PictureBox Contorno_Left 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
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
         Height          =   1335
         Left            =   4320
         ScaleHeight     =   89
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   3
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1200
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Image Image_Contorno 
         Height          =   480
         Left            =   4200
         Top             =   360
         Width           =   1620
      End
      Begin WMPLibCtl.WindowsMediaPlayer Wmp 
         Height          =   2640
         Left            =   0
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Width           =   3975
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   1
         autoStart       =   0   'False
         currentMarker   =   0
         invokeURLs      =   0   'False
         baseURL         =   ""
         volume          =   25
         mute            =   0   'False
         uiMode          =   "none"
         stretchToFit    =   -1  'True
         windowlessVideo =   0   'False
         enabled         =   0   'False
         enableContextMenu=   -1  'True
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   7011
         _cy             =   4657
      End
   End
   Begin VB.PictureBox Barra_Janelas 
      Appearance      =   0  'Flat
      BackColor       =   &H00282828&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   0
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   785
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   6120
      Width           =   11775
      Begin VB.Image Boato_Full_screen 
         Height          =   180
         Left            =   5280
         Picture         =   "Form_Wmp.frx":57E2
         ToolTipText     =   "Ver tela cheia"
         Top             =   120
         Width           =   180
      End
      Begin VB.Label Label_Faixa 
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   180
      End
      Begin VB.Image Fundo_Barra_Janelas 
         Enabled         =   0   'False
         Height          =   435
         Left            =   0
         Picture         =   "Form_Wmp.frx":59D4
         Top             =   0
         Width           =   600
      End
   End
   Begin VB.PictureBox Barra_Player 
      Appearance      =   0  'Flat
      BackColor       =   &H00282828&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   785
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6555
      Width           =   11775
      Begin VB.Timer Timer_Slider_Video 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   5760
         Top             =   120
      End
      Begin VB.Timer Timer_Duracao 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   5280
         Top             =   120
      End
      Begin VB.PictureBox Barra_Progresso 
         Appearance      =   0  'Flat
         BackColor       =   &H00282828&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2880
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   452
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   150
         Width           =   6780
         Begin VB.PictureBox Slide 
            Appearance      =   0  'Flat
            BackColor       =   &H00F0F1F2&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   0
            Picture         =   "Form_Wmp.frx":5CE4
            ScaleHeight     =   18
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   11
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   0
            Width           =   165
         End
         Begin VB.Image Image_Progresso 
            Enabled         =   0   'False
            Height          =   135
            Left            =   45
            Picture         =   "Form_Wmp.frx":5FAE
            Stretch         =   -1  'True
            Top             =   90
            Width           =   60
         End
         Begin VB.Image Image_Progresso_Direita 
            Enabled         =   0   'False
            Height          =   300
            Left            =   6690
            Picture         =   "Form_Wmp.frx":605C
            Top             =   0
            Width           =   90
         End
         Begin VB.Image Image_Progresso_Inicio 
            Enabled         =   0   'False
            Height          =   135
            Left            =   15
            Picture         =   "Form_Wmp.frx":622E
            Stretch         =   -1  'True
            Top             =   90
            Width           =   60
         End
         Begin VB.Image Image_Progresso_Esquerda 
            Enabled         =   0   'False
            Height          =   300
            Left            =   0
            Picture         =   "Form_Wmp.frx":62DC
            Top             =   0
            Width           =   90
         End
         Begin VB.Image Image_Barra_Slide 
            Enabled         =   0   'False
            Height          =   300
            Left            =   0
            Picture         =   "Form_Wmp.frx":64AE
            Stretch         =   -1  'True
            Top             =   0
            Width           =   6780
         End
      End
      Begin VB.Image Botao_Play 
         Height          =   240
         Left            =   120
         Picture         =   "Form_Wmp.frx":75C5
         ToolTipText     =   "Reproduzir"
         Top             =   180
         Width           =   180
      End
      Begin VB.Image Botao_Mudo 
         Height          =   270
         Left            =   1920
         Picture         =   "Form_Wmp.frx":7847
         ToolTipText     =   "Mudo"
         Top             =   180
         Width           =   330
      End
      Begin VB.Label Label_Duracao 
         AutoSize        =   -1  'True
         BackColor       =   &H003E3F3F&
         BackStyle       =   0  'Transparent
         Caption         =   "00:00  |  "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   9840
         TabIndex        =   17
         Top             =   195
         Width           =   735
      End
      Begin VB.Label Tempo_Estimado 
         AutoSize        =   -1  'True
         BackColor       =   &H001E1F1D&
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   10680
         TabIndex        =   16
         Top             =   195
         Width           =   480
      End
      Begin VB.Image Botao_Seguinte 
         Height          =   210
         Left            =   1200
         Picture         =   "Form_Wmp.frx":7D51
         ToolTipText     =   "Faixa seguinte"
         Top             =   180
         Width           =   165
      End
      Begin VB.Image Botao_Anterior 
         Height          =   210
         Left            =   720
         Picture         =   "Form_Wmp.frx":7F8B
         ToolTipText     =   "Faixa anteiro"
         Top             =   180
         Width           =   165
      End
      Begin VB.Image Botao_Redimensionar 
         Height          =   165
         Left            =   10800
         Picture         =   "Form_Wmp.frx":81C5
         Top             =   465
         Width           =   135
      End
      Begin VB.Image Botao_Pausa 
         Height          =   240
         Left            =   120
         Picture         =   "Form_Wmp.frx":833B
         ToolTipText     =   "Pausa"
         Top             =   180
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Image Fundo_Barra_Player 
         Enabled         =   0   'False
         Height          =   615
         Left            =   0
         Picture         =   "Form_Wmp.frx":863D
         Top             =   0
         Width           =   600
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
      ScaleWidth      =   785
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   11775
      Begin VB.Image Botao_Minimizar 
         Height          =   225
         Left            =   10080
         Picture         =   "Form_Wmp.frx":8954
         ToolTipText     =   "Minimizar"
         Top             =   120
         Width           =   240
      End
      Begin VB.Image Botao_Restaurar 
         Height          =   225
         Left            =   10440
         Picture         =   "Form_Wmp.frx":8C66
         ToolTipText     =   "Restaurar"
         Top             =   120
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Botao_Maximizar 
         Height          =   225
         Left            =   10800
         Picture         =   "Form_Wmp.frx":8F78
         ToolTipText     =   "Maximizar"
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00272727&
         BackStyle       =   0  'Transparent
         Caption         =   "Tela de video"
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
         TabIndex        =   12
         Top             =   120
         Width           =   1320
      End
      Begin VB.Image Botao_Fechar 
         Height          =   225
         Left            =   11160
         Picture         =   "Form_Wmp.frx":928A
         ToolTipText     =   "Fechar"
         Top             =   120
         Width           =   240
      End
      Begin VB.Image Fundo_Barra_ControlBox 
         Enabled         =   0   'False
         Height          =   465
         Left            =   0
         Picture         =   "Form_Wmp.frx":959C
         Top             =   0
         Width           =   585
      End
   End
   Begin VB.PictureBox Barra_Menu 
      Appearance      =   0  'Flat
      BackColor       =   &H00333333&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   785
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   465
      Width           =   11775
      Begin VB.TextBox Text_Teclas 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
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
         Height          =   375
         Left            =   -150
         TabIndex        =   6
         Top             =   0
         Width           =   135
      End
      Begin VB.Image Image_Luz 
         Height          =   240
         Left            =   120
         Picture         =   "Form_Wmp.frx":98E1
         Top             =   75
         Width           =   240
      End
      Begin VB.Label Label_Placa 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Placa de som não detectada"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   420
         TabIndex        =   1
         Top             =   90
         Width           =   2115
      End
   End
   Begin VB.Shape Shape_Contorno 
      BorderColor     =   &H00404040&
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "Form_Wmp"
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

'VARIÁVERIS DO SLIDER VIDEO
Dim tx As Integer, Ty As Integer, DN As Boolean
Dim Txa As Integer, DNa As Boolean
Dim Tyb, DNb As Boolean
Dim NewLeft As Integer

'Redimensionar formulário
Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal CX As Long, ByVal cy As Long, ByVal wFlags As Long)
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

'Com/ Sem som
Public Mudo As Boolean

'Variavel para verificar a janela do formulário
Dim Tela_Cheia As Boolean

'Variavel para ver a duracao do ficheiro a reproduzir
Public VideoDuration As Double

'DETECTAR PLACA DE SOM
Private Declare Function waveOutGetNumDevs Lib "winmm" () As Long

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
'Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'Faixa em Reproduzindo
Public Faixa_em_Reproduzindo As String
Public Musica_Play As Boolean

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

Private Sub Boato_Full_screen_Click()
    'Atalho para
    Botao_Amplicar_Click
End Sub

Private Sub Botao_Anterior_Click()
    'Reproduzir a faixa anterior
    With Form_Principal
        If .Lista_Filmes.Rows = 1 Then Exit Sub
        
        'Caso esteja na 1ºlinha não recua mais
        If .Lista_Filmes.Row = 1 Then
            Wmp.Controls.stop
            Botao_Play.Visible = True
            Botao_Pausa.Visible = False
            Slide.Left = 0
            Image_Progresso.Width = 1
            Image_Progresso.Left = 3
            VideoDuration = 0
            Label_Duracao.Caption = "00:00" & "  |  "
            Tempo_Estimado.Caption = "00:00"
            Exit Sub
            
        Else
            'Selecionar a linha anterior
            With .Lista_Filmes
                .Row = .Row - 1
                .Col = 0
                .ColSel = .Cols - 1 'Selecionar a linha por inteiro
                Faixa_em_Reproduzindo = .TextMatrix(.Row, 7)
            End With
            Tocar_Media
        End If
    End With
End Sub

Private Sub Botao_Pausa_Click()
    'Pausa do media player
    Wmp.Controls.pause
    Timer_Duracao.Enabled = False
    Botao_Play.Visible = True
    Botao_Pausa.Visible = False
End Sub

Private Sub Botao_Seguinte_Click()
    'Passar para a faixa seguinte
    With Form_Principal
        If .Lista_Filmes.Rows = 1 Then Exit Sub
        
        'Caso esteja na última linha não recua mais
        If .Lista_Filmes.Row = .Lista_Filmes.Rows - 1 Then
            Wmp.Controls.stop
            Botao_Play.Visible = True
            Botao_Pausa.Visible = False
            Slide.Left = 0
            Image_Progresso.Width = 1
            Image_Progresso.Left = 3
            VideoDuration = 0
            Label_Duracao.Caption = "00:00" & "  |  "
            Tempo_Estimado.Caption = "00:00"
            Exit Sub
            
        Else
            'Selecionar a linha seguinte
            With .Lista_Filmes
                .Row = .Row + 1
                .Col = 0
                .ColSel = .Cols - 1 'Selecionar a linha por inteiro
                Faixa_em_Reproduzindo = .TextMatrix(.Row, 7)
            End With
            Tocar_Media
        End If
    End With
End Sub

Private Sub Label_Faixa_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Chamar o procedimento
    Ocultar_Objectos
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

Private Sub Barra_Menu_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Chamar o procedimento
    Ocultar_Objectos
End Sub

Private Sub Barra_Slider_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Chamar o procedimento
    Ocultar_Objectos
End Sub

Private Sub Botao_Amplicar_Click()
    'Ampliar/ reduzir janela de video
    If Label_Ampliar.Caption = "Ver tela cheia" Then
        Label_Ampliar.Caption = "Ver tela normal"
        'Colocar em modo fullscreen
        Me.WindowState = 2
        Ocultar_Objectos
        
        'Frame_Video
        Frame_Video.Height = Me.ScaleHeight
        Frame_Video.Top = 0
        Frame_Video.Width = Me.ScaleWidth
        Frame_Video.Left = 0
        
        With Wmp
            .Height = Frame_Video.ScaleHeight
            .Top = 0
            .Width = Frame_Video.ScaleWidth
            .Left = 0
        End With
        
        'Botao_Amplicar
        Botao_Amplicar.Left = (Frame_Video.ScaleWidth - Botao_Amplicar.ScaleWidth) / 2
        Botao_Amplicar.Top = 0
        
        'Image_Contorno
        Image_Contorno.Height = Frame_Video.ScaleHeight
        Image_Contorno.Top = 0
        Image_Contorno.Width = Frame_Video.ScaleWidth - 200
        Image_Contorno.Left = (Frame_Video.ScaleWidth - Image_Contorno.Width) / 2
            
        'Contornos do ampliar video
        Contorno_Top.Width = Image_Contorno.Width
        Contorno_Top.Top = 0
        Contorno_Top.Left = Image_Contorno.Left
        Contorno_Down.Width = Image_Contorno.Width
        Contorno_Down.Top = Image_Contorno.Top + Image_Contorno.Height - Contorno_Down.ScaleHeight
        Contorno_Down.Left = Image_Contorno.Left
        Contorno_Left.Height = Image_Contorno.Height
        Contorno_Left.Top = 0
        Contorno_Left.Left = Image_Contorno.Left
        Contorno_Right.Height = Image_Contorno.Height
        Contorno_Right.Top = 0
        Contorno_Right.Left = Image_Contorno.Left + Image_Contorno.Width - Contorno_Right.ScaleWidth
        
    Else
        Label_Ampliar.Caption = "Ver tela cheia"
        Me.WindowState = 0
        
        With Frame_Video
            .Top = Barra_Menu.Top + Barra_Menu.ScaleHeight
            .Height = Me.ScaleHeight - Barra_ControlBox.ScaleHeight - Barra_Menu.ScaleHeight - Barra_Janelas.ScaleHeight - Barra_Player.ScaleHeight
            .Width = Barra_ControlBox.ScaleWidth
            .Left = Barra_ControlBox.Left
        End With
        
        With Wmp
            .Height = Frame_Video.ScaleHeight
            .Top = 0
            .Width = Frame_Video.ScaleWidth
            .Left = 0
        End With
        
        'Botao_Amplicar
        Botao_Amplicar.Left = (Frame_Video.ScaleWidth - Botao_Amplicar.ScaleWidth) / 2
        Botao_Amplicar.Top = 0
        
        'Image_Contorno
        Image_Contorno.Height = Frame_Video.ScaleHeight
        Image_Contorno.Top = 0
        Image_Contorno.Width = Frame_Video.ScaleWidth - 200
        Image_Contorno.Left = (Frame_Video.ScaleWidth - Image_Contorno.Width) / 2
            
        'Contornos do ampliar video
        Contorno_Top.Width = Image_Contorno.Width
        Contorno_Top.Top = 0
        Contorno_Top.Left = Image_Contorno.Left
        Contorno_Down.Width = Image_Contorno.Width
        Contorno_Down.Top = Image_Contorno.Top + Image_Contorno.Height - Contorno_Down.ScaleHeight
        Contorno_Down.Left = Image_Contorno.Left
        Contorno_Left.Height = Image_Contorno.Height
        Contorno_Left.Top = 0
        Contorno_Left.Left = Image_Contorno.Left
        Contorno_Right.Height = Image_Contorno.Height
        Contorno_Right.Top = 0
        Contorno_Right.Left = Image_Contorno.Left + Image_Contorno.Width - Contorno_Right.ScaleWidth
    End If
    Ocultar_Objectos
End Sub

Private Sub Botao_Amplicar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Chamar o prodecimento
    Ver_Objectos
End Sub

Private Sub Botao_Fechar_Click()
    'Fechar formulário
    Timer_Duracao.Enabled = False
    Timer_Slider_Video.Enabled = False
    Wmp.Controls.stop
    Unload Me
End Sub

Private Sub Botao_Maximizar_Click()
    'Maximixar formulário
    PosFormRelativeTaskBar Me
    Tela_Cheia = True
    Botao_Maximizar.Visible = False
    Botao_Restaurar.Visible = True
    
    'Chamar o procedimento
    Desenhar_Formulario
End Sub

Private Sub Botao_Minimizar_Click()
    'Minimizar o formulário
    Me.WindowState = 1
End Sub

Private Sub Botao_Mudo_Click()
    'Colocar o media player como mudo ou ouvir
    If Mudo = False Then
        Wmp.settings.mute = True
        Mudo = True
        Botao_Mudo.Picture = Form_Skin.Som_Over.Picture
        Botao_Mudo.ToolTipText = "Ouvir"
        
    Else
        Wmp.settings.mute = False
        Mudo = False
        Botao_Mudo.Picture = Form_Skin.Som_Normal.Picture
        Botao_Mudo.ToolTipText = "Mudo"
    End If
End Sub

Private Sub Botao_Pause_Click()
    Wmp.Controls.pause
    Timer_Duracao.Enabled = False
    Musica_Play = False
End Sub

Public Sub Botao_Play_Click()
    'Verificar a faixa de reprodução
    With Form_Principal
        If .Lista_Filmes.Rows = 1 Then Exit Sub
        'Associar ficheiro á faixa de reproduçºao
        If Faixa_em_Reproduzindo = "" Then
            Faixa_em_Reproduzindo = .Lista_Filmes.TextMatrix(.Lista_Filmes.Row, 7)
            Tocar_Media
        End If
    End With
    
    'Reproduzir o ficheiro existente no player
    Wmp.Controls.play
    Timer_Duracao.Enabled = True
    Botao_Play.Visible = False
    Botao_Pausa.Visible = True
End Sub

Private Sub Botao_Restaurar_Click()
    'Restaurar janela
    With Me
        .Height = ALtura_Formulario
        .Width = Largura_Formulario
        .Top = (Screen.Height - Me.Height) / 2
        .Left = (Screen.Width - Me.Width) / 2
    End With
    Tela_Cheia = False
    Botao_Maximizar.Visible = True
    Botao_Restaurar.Visible = False
End Sub

Private Sub Form_Load()
    'Chamar o procedimento para contruir o formulário
    ALtura_Formulario = "7140"
    Largura_Formulario = "11760"
    Desenhar_Formulario

    iTPPX& = Screen.TwipsPerPixelX
    iTPPY& = Screen.TwipsPerPixelY
    Tela_Cheia = False
    Mudo = False
    Wmp.settings.volume = 100

    'Detectar placa de som
    Dim Placa As Long
    Placa = waveOutGetNumDevs()
    If Placa > 0 Then
       Label_Placa.Caption = "Placa de som detectada"
       Image_Luz.Picture = Form_Skin.Luz_Verde.Picture
    Else
       Label_Placa.Caption = "Placa de som não detectada"
       Image_Luz.Picture = Form_Skin.Luz_Vermelha.Picture
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Chamar o procedimento
    Ocultar_Objectos
End Sub

Private Sub Form_Resize()
    'Chamar o procedimento para contruir o formulário
    Desenhar_Formulario
End Sub

Private Sub Image_Contorno_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Chamar o prodecimento
    'If Me.WindowState = 2 Then Exit Sub
    Ver_Objectos
End Sub

Private Sub Label_Ampliar_Click()
    'Atalho para
    Botao_Amplicar_Click
End Sub

Private Sub Label_Ampliar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Chamar o prodecimento
    Ver_Objectos
End Sub

Private Sub Label_Percentagem_Volume_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Chamar o procedimento
    Ocultar_Objectos
End Sub

Private Sub Label_Placa_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Chamar o procedimento
    Ocultar_Objectos
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
    'Chamar o procedimento
    Ocultar_Objectos
End Sub

Private Sub Skin_Top_Centro_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Largar o formulário para a posição final
    bMoveFrom = False
End Sub

Private Sub Skin_Top_Centro_DblClick()
    'Atalho para
    Label_Titulo_DblClick
End Sub

Private Sub Botao_Redimensionar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  'Redimensionar o formulário conforme as dimensões pretendidas
    If Button = vbLeftButton Then
        If Tela_Cheia = False Then
            ReleaseCapture
            SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, 0
            
            'Verificar se não exedeu os limites
            If Me.Height < 7140 Then
                Me.Height = "7140"
            End If
        
            If Me.Width < 11760 Then
                Me.Width = "11760"
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
    
    'Chamar o procedimento
    Ocultar_Objectos
End Sub

Private Sub Slide_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    DNa = True
    Txa = x
End Sub

Private Sub Slide_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If DNa Then
        NewLeft = Slide.Left + x - Txa
        If NewLeft < Image_Barra_Slide.Left + 5 Then
            NewLeft = Image_Barra_Slide.Left + 5
        End If
        If NewLeft > Image_Barra_Slide.Width + Image_Barra_Slide.Left - 8 - Slide.Width Then
            NewLeft = Image_Barra_Slide.Width + Image_Barra_Slide.Left - 8 - Slide.Width
        End If
        Slide.Left = NewLeft
        Image_Progresso.Width = Slide.Left
    End If
    'Chamar o procedimento
    Ocultar_Objectos
End Sub

Private Sub Slide_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'On Error Resume Next
    Dim offseti As Single
    DNa = False
    offseti = (Slide.Left - Image_Barra_Slide.Left - 3) / (Image_Barra_Slide.Width - 10 - Slide.Width)
    Wmp.Controls.CurrentPosition = Int(Wmp.currentMedia.Duration * offseti)
    Image_Progresso.Width = Slide.Left
End Sub

Public Sub Desenhar_Formulario()
    'Procedimento para ajustar os objectos
    On Error GoTo CORRIGIR_ERRO
    If Me.WindowState = 1 Then Exit Sub
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
        .Width = Barra_ControlBox.ScaleWidth
        .Left = 0
    End With
    
    With Botao_Fechar
        .Top = 8
        .Left = Barra_ControlBox.ScaleWidth - .Width - 4
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
    
    With Barra_Player
        .Height = Fundo_Barra_Player.Height
        .Top = Me.ScaleHeight - .Height
        .Width = Me.ScaleWidth
        .Left = 0
    End With
    
    With Fundo_Barra_Player
        .Stretch = True
        .Top = 0
        .Width = Barra_Player.Width
        .Left = 0
    End With
    
    With Barra_Janelas
        .Height = Fundo_Barra_Janelas.Height
        .Top = Barra_Player.Top - .Height
        .Width = Barra_Player.Width
        .Left = Barra_Player.Left
    End With
    
    With Fundo_Barra_Janelas
        .Stretch = True
        .Top = 0
        .Width = Barra_Janelas.Width
        .Left = 0
    End With
    
    With Barra_Menu
        .Top = Barra_ControlBox.Top + Barra_ControlBox.ScaleHeight
        .Left = Barra_ControlBox.Left
        .Width = Barra_ControlBox.ScaleWidth
    End With
    
    With Botao_Redimensionar
        .Left = Barra_Player.ScaleWidth - .Width
    End With
    
    With Frame_Video
        .Top = Barra_Menu.Top + Barra_Menu.ScaleHeight
        .Height = Me.ScaleHeight - Barra_ControlBox.ScaleHeight - Barra_Menu.ScaleHeight - Barra_Janelas.ScaleHeight - Barra_Player.ScaleHeight
        .Width = Barra_ControlBox.ScaleWidth
        .Left = Barra_ControlBox.Left
    End With
    
    With Wmp
        .Height = Frame_Video.ScaleHeight
        .Top = 0
        .Width = Frame_Video.ScaleWidth
        .Left = 0
    End With
    
    'Botao_Amplicar
    Botao_Amplicar.Left = (Frame_Video.ScaleWidth - Botao_Amplicar.ScaleWidth) / 2
    Botao_Amplicar.Top = 0
    
    'Image_Contorno
    Image_Contorno.Height = Frame_Video.ScaleHeight
    Image_Contorno.Top = 0
    Image_Contorno.Width = Frame_Video.ScaleWidth - 200
    Image_Contorno.Left = (Frame_Video.ScaleWidth - Image_Contorno.Width) / 2
        
    'Contornos do ampliar video
    Contorno_Top.Width = Image_Contorno.Width
    Contorno_Top.Top = 0
    Contorno_Top.Left = Image_Contorno.Left
    Contorno_Down.Width = Image_Contorno.Width
    Contorno_Down.Top = Image_Contorno.Top + Image_Contorno.Height - Contorno_Down.ScaleHeight
    Contorno_Down.Left = Image_Contorno.Left
    Contorno_Left.Height = Image_Contorno.Height
    Contorno_Left.Top = 0
    Contorno_Left.Left = Image_Contorno.Left
    Contorno_Right.Height = Image_Contorno.Height
    Contorno_Right.Top = 0
    Contorno_Right.Left = Image_Contorno.Left + Image_Contorno.Width - Contorno_Right.ScaleWidth
    
    With Image_Luz
        .Top = (Barra_Menu.ScaleHeight - .Height) / 2
    End With
    
    With Label_Placa
        .Top = (Barra_Menu.ScaleHeight - .Height) / 2
    End With
    
    With Slide
        .Height = Form_Skin.Image_Slider.Height
        .Top = 3
        .Width = Form_Skin.Image_Slider.Width
        .Left = 0
    End With
    
    With Barra_Progresso
        .Height = Slide.Height + 2
        .Top = 12
        '.Width = Image_Progresso_Centro.Width
    End With
    
    With Botao_Play
        .Top = (Barra_Player.ScaleHeight - .Height) / 2
    End With
    
    With Botao_Pausa
        .Top = (Barra_Player.ScaleHeight - .Height) / 2
    End With
    
    With Botao_Anterior
        .Top = (Barra_Player.ScaleHeight - .Height) / 2
    End With
    
    With Botao_Seguinte
        .Top = (Barra_Player.ScaleHeight - .Height) / 2
    End With
    
    With Botao_Mudo
        .Top = (Barra_Player.ScaleHeight - .Height) / 2
    End With
    
    With Barra_Progresso
        .Top = (Barra_Player.ScaleHeight - .Height) / 2
    End With
    
    With Label_Duracao
        .Top = (Barra_Player.ScaleHeight - .Height) / 2
    End With
    
    With Tempo_Estimado
        .Top = (Barra_Player.ScaleHeight - .Height) / 2
    End With
    
    With Label_Faixa
        .Top = (Barra_Janelas.ScaleHeight - .Height) / 2
        .Width = Barra_Janelas.ScaleWidth - Boato_Full_screen.Width - 8 - 8 - 8
        .Left = 8
    End With
    
    With Boato_Full_screen
        .Top = (Barra_Janelas.ScaleHeight - .Height) / 2
        .Left = Barra_Janelas.ScaleWidth - .Width - 8
    End With
    
Exit Sub
CORRIGIR_ERRO:
    Me.Height = 7140
    Me.Width = 11760
End Sub

Private Sub SliderBar_Mascara_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Chamar o procedimento
    Ocultar_Objectos
End Sub

Private Sub Text_Teclas_KeyDown(KeyCode As Integer, Shift As Integer)
    'Detectar teclas digitadas
    'Colocar o video em modo fullscreen
    If KeyCode = vbKeyReturn And Shift = vbAltMask Then
        'Colocar em modo fullscreen
        Label_Ampliar.Caption = "Ver tela normal"
        Me.WindowState = 2
        Ocultar_Objectos
        
        'Frame_Video
        Frame_Video.Height = Me.ScaleHeight
        Frame_Video.Top = 0
        Frame_Video.Width = Me.ScaleWidth
        Frame_Video.Left = 0
        
        With Wmp
            .Height = Frame_Video.ScaleHeight
            .Top = 0
            .Width = Frame_Video.ScaleWidth
            .Left = 0
        End With
        
        'Botao_Amplicar
        Botao_Amplicar.Left = (Frame_Video.ScaleWidth - Botao_Amplicar.ScaleWidth) / 2
        Botao_Amplicar.Top = 0
        
        'Image_Contorno
        Image_Contorno.Height = Frame_Video.ScaleHeight
        Image_Contorno.Top = 0
        Image_Contorno.Width = Frame_Video.ScaleWidth - 200
        Image_Contorno.Left = (Frame_Video.ScaleWidth - Image_Contorno.Width) / 2
            
        'Contornos do ampliar video
        Contorno_Top.Width = Image_Contorno.Width
        Contorno_Top.Top = 0
        Contorno_Top.Left = Image_Contorno.Left
        Contorno_Down.Width = Image_Contorno.Width
        Contorno_Down.Top = Image_Contorno.Top + Image_Contorno.Height - Contorno_Down.ScaleHeight
        Contorno_Down.Left = Image_Contorno.Left
        Contorno_Left.Height = Image_Contorno.Height
        Contorno_Left.Top = 0
        Contorno_Left.Left = Image_Contorno.Left
        Contorno_Right.Height = Image_Contorno.Height
        Contorno_Right.Top = 0
        Contorno_Right.Left = Image_Contorno.Left + Image_Contorno.Width - Contorno_Right.ScaleWidth


    End If
    
    'Colocar o video em modo normal
    If KeyCode = vbKeyEscape Then
        Label_Ampliar.Caption = "Ver tela cheia"
        Me.WindowState = 0
        
        With Frame_Video
            .Top = Barra_Menu.Top + Barra_Menu.ScaleHeight
            .Height = Me.ScaleHeight - Barra_ControlBox.ScaleHeight - Barra_Menu.ScaleHeight - Barra_Janelas.ScaleHeight - Barra_Player.ScaleHeight
            .Width = Barra_ControlBox.ScaleWidth
            .Left = Barra_ControlBox.Left
        End With
        
        With Wmp
            .Height = Frame_Video.ScaleHeight
            .Top = 0
            .Width = Frame_Video.ScaleWidth
            .Left = 0
        End With
        
        'Botao_Amplicar
        Botao_Amplicar.Left = (Frame_Video.ScaleWidth - Botao_Amplicar.ScaleWidth) / 2
        Botao_Amplicar.Top = 0
        
        'Image_Contorno
        Image_Contorno.Height = Frame_Video.ScaleHeight
        Image_Contorno.Top = 0
        Image_Contorno.Width = Frame_Video.ScaleWidth - 200
        Image_Contorno.Left = (Frame_Video.ScaleWidth - Image_Contorno.Width) / 2
            
        'Contornos do ampliar video
        Contorno_Top.Width = Image_Contorno.Width
        Contorno_Top.Top = 0
        Contorno_Top.Left = Image_Contorno.Left
        Contorno_Down.Width = Image_Contorno.Width
        Contorno_Down.Top = Image_Contorno.Top + Image_Contorno.Height - Contorno_Down.ScaleHeight
        Contorno_Down.Left = Image_Contorno.Left
        Contorno_Left.Height = Image_Contorno.Height
        Contorno_Left.Top = 0
        Contorno_Left.Left = Image_Contorno.Left
        Contorno_Right.Height = Image_Contorno.Height
        Contorno_Right.Top = 0
        Contorno_Right.Left = Image_Contorno.Left + Image_Contorno.Width - Contorno_Right.ScaleWidth
    End If
    
    'Reproduzir/ pausa do filme
    If KeyCode = vbKeySpace Then
        If Musica_Play = False Then
            Musica_Play = True
            Botao_Play_Click
            Exit Sub
        Else
            Musica_Play = False
            Botao_Pause_Click
            Exit Sub
        End If
    End If
End Sub

Private Sub Timer_Duracao_Timer()
    'On Error Resume Next
    Label_Duracao.Caption = Duration(Wmp.Controls.CurrentPosition) & "  |  "
    Wmp_PositionChange 0, 1
    If VideoDuration >= 1 Then
        'Wmp.Controls.stop
        Wmp.Controls.play
    End If
    Botao_Play.Visible = False
    Botao_Pausa.Visible = True
    
    Tempo_Estimado.Caption = Wmp.Controls.currentItem.durationString

    Label_Duracao.Left = Barra_Progresso.Left + Barra_Progresso.Width + 20
    Tempo_Estimado.Left = Label_Duracao.Left + Label_Duracao.Width
End Sub

Private Sub Timer_Slider_Video_Timer()
    On Error Resume Next
    Dim tm As Integer, tt As Integer, tp As Single, offset As Integer
    
    tm = Int(Wmp.Controls.CurrentPosition)
    tt = Int(Wmp.currentMedia.Duration)
    
    If tm <> -1 Then
        tp = tm / tt
        
        offset = Int((Image_Barra_Slide.Width - 5 - Slide.Width) * tp)
        
        If Not DNa Then
            Slide.Left = offset + Image_Barra_Slide.Left + 3
            Image_Progresso.Width = Slide.Left
        End If
        If Slide.Left >= 4320 Then
            Wmp.Controls.pause
            Timer_Duracao.Enabled = False
            Timer_Slider_Video.Enabled = False
            Label_Duracao.Caption = "00:00" & "  |  "
            Tempo_Estimado.Caption = "00:00"
            VideoDuration = 0
            Slide.Left = 0
            Botao_Play.Visible = True
            Botao_Pausa.Visible = False
        End If
    End If
End Sub

Private Sub Wmp_MouseMove(ByVal nButton As Integer, ByVal nShiftState As Integer, ByVal fX As Long, ByVal fY As Long)
    'Chamar o procedimento
    Ocultar_Objectos
End Sub

Private Sub Wmp_PositionChange(ByVal oldPosition As Double, ByVal newPosition As Double)
    'Ver posicao do slider do video
    On Error Resume Next
    Dim ComputeDuration
    VideoDuration = ComputeDuration(Trim(Wmp.Controls.currentItem.durationString))
End Sub

Public Sub Ocultar_Objectos()
    'Procedimento para ocultar objectos não desejados
    Botao_Amplicar.Visible = False
    Contorno_Top.Visible = False
    Contorno_Down.Visible = False
    Contorno_Left.Visible = False
    Contorno_Right.Visible = False
End Sub

Public Sub Ver_Objectos()
    'Ver contorno de ampliação
    Botao_Amplicar.Visible = True
    Contorno_Top.Visible = True
    Contorno_Down.Visible = True
    Contorno_Left.Visible = True
    Contorno_Right.Visible = True
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

Public Sub Tocar_Media()
    'Procedimento para reproduzir os ficheiros
    Slide.Left = 0
    Image_Progresso.Width = 1
    Image_Progresso.Left = 3
    VideoDuration = 0
    Wmp.Controls.stop

    'Reproduzir o som
    Label_Duracao.Caption = "00:00" & "  |  "
    Tempo_Estimado.Caption = "00:00"

    Wmp.URL = Faixa_em_Reproduzindo
    
    Timer_Slider_Video.Enabled = True
    Timer_Duracao.Enabled = True
    Wmp.Controls.play
    Slide.Visible = True
    Image_Progresso_Inicio.Visible = True
    Image_Progresso.Visible = True
    
    'Indicar a faixa que está a ser reproduzida
    Label_Faixa.Caption = "A reproduzir: [" & Form_Principal.Lista_Filmes.TextMatrix(Form_Principal.Lista_Filmes.Row, 1) & "]"
    Label_Faixa.Width = Barra_Janelas.Width - 150
    
    Botao_Play.Visible = False
    Botao_Pausa.Visible = True
    
    Musica_Play = True
    
    If Mudo = True Then Wmp.settings.mute = True Else Wmp.settings.mute = False

    Label_Duracao.Left = Barra_Progresso.Left + Barra_Progresso.Width + 20
    Tempo_Estimado.Left = Label_Duracao.Left + Label_Duracao.Width
End Sub

