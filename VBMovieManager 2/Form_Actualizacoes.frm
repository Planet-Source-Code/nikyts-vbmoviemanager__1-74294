VERSION 5.00
Begin VB.Form Form_Actualizacoes 
   Appearance      =   0  'Flat
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   0  'None
   ClientHeight    =   1980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5700
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   132
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      ScaleWidth      =   409
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   6135
      Begin VB.Image Botao_Fechar 
         Height          =   225
         Left            =   4320
         Picture         =   "Form_Actualizacoes.frx":0000
         ToolTipText     =   "Fechar"
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00272727&
         BackStyle       =   0  'Transparent
         Caption         =   "Actualizacões"
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
         TabIndex        =   2
         Top             =   120
         Width           =   1350
      End
      Begin VB.Image Fundo_Barra_ControlBox 
         Enabled         =   0   'False
         Height          =   465
         Left            =   0
         Picture         =   "Form_Actualizacoes.frx":0312
         Top             =   0
         Width           =   585
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   5040
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   15
      Left            =   4560
      Top             =   480
   End
   Begin VB.Label Label_Actualizar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Actualizar programa"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E6964D&
      Height          =   195
      Left            =   1080
      TabIndex        =   3
      Top             =   1320
      Width           =   1740
   End
   Begin VB.Label Label_Mensagem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Existe uma nova versão disponivel."
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   3045
   End
   Begin VB.Image Icon_Info 
      Enabled         =   0   'False
      Height          =   480
      Left            =   360
      Picture         =   "Form_Actualizacoes.frx":0657
      Top             =   840
      Width           =   480
   End
   Begin VB.Shape Shape_Contorno 
      BorderColor     =   &H00C0C0C0&
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "Form_Actualizacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SPI_GETWORKAREA = 48
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT = &H20&
Private Const LWA_ALPHA = &H2&

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal CX As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
Private i As Integer
Private NormalWindowStyle As Long
Private TaskBar As Long

Private Sub Botao_Fechar_Click()
    'Fechar o formulário
    Unload Me
End Sub

Private Sub Form_Load()
    'Propriedades iniciais do formulário
    Desenhar_Formulario

'''    Me.Show
    Me.Top = Screen.Height
    Me.Left = Screen.Width - Me.Width - 50
    NormalWindowStyle = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
    SetWindowLong Me.hWnd, GWL_EXSTYLE, NormalWindowStyle Or WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hWnd, 0, 255, LWA_ALPHA
    i = 100
    
    'Colocar o form por cima dos outros
    AlwaysOnTop Me, -1
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
        .Width = Me.ScaleWidth
        .Left = 0
    End With
    
    With Fundo_Barra_ControlBox
        .Stretch = True
        .Top = 0
        .Width = Barra_ControlBox.Width
        .Left = 0
    End With
    
    With Label_Titulo
        .Top = (Barra_ControlBox.ScaleHeight - .Height) / 2
        .Left = 8
    End With
    
    With Botao_Fechar
        .Top = 8
        .Left = Barra_ControlBox.Width - .Width - 4
    End With
End Sub

Private Sub Label_Actualizar_Click()
    'Efectuar actualizões do programa
    Shell App.path & "\Options\Actualizar.exe"
    Form_Principal.Botao_Fechar_Click
End Sub

Private Sub Timer1_Timer()
    Dim WindowRect As RECT
    SystemParametersInfo SPI_GETWORKAREA, 0, WindowRect, 0
    TaskBar = ((Screen.Height / Screen.TwipsPerPixelX) - WindowRect.Bottom) * Screen.TwipsPerPixelX
    If (Me.Top + Me.Height + TaskBar) > Screen.Height Then
        Me.Top = Me.Top - 30
    Else
        SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
        Timer2.Enabled = True
    End If
End Sub

Private Sub Timer2_Timer()
    Timer1.Enabled = False
    SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub
