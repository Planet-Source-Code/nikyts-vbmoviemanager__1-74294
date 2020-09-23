VERSION 5.00
Begin VB.Form Form_Sobre 
   Appearance      =   0  'Flat
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   0  'None
   ClientHeight    =   5265
   ClientLeft      =   5685
   ClientTop       =   2310
   ClientWidth     =   6735
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
   ScaleHeight     =   351
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   449
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Barra_Botoes 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   120
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   441
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4320
      Width           =   6615
      Begin VB.PictureBox Botao_Ok 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   396
         Left            =   2760
         Picture         =   "Form_Sobre.frx":0000
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   49
         TabIndex        =   0
         Top             =   240
         Width           =   732
         Begin VB.Shape Contorno_Ok 
            BorderColor     =   &H00E6964D&
            Height          =   396
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Width           =   732
         End
         Begin VB.Label Label_Ok 
            Alignment       =   2  'Center
            BackColor       =   &H00272727&
            BackStyle       =   0  'Transparent
            Caption         =   "Ok"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   0
            TabIndex        =   8
            Top             =   192
            Width           =   732
         End
      End
      Begin VB.Image Fundo_Barra_Botoes 
         Enabled         =   0   'False
         Height          =   900
         Left            =   0
         Picture         =   "Form_Sobre.frx":17FA
         Top             =   0
         Width           =   585
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
      ScaleWidth      =   449
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   6735
      Begin VB.Image Botao_Fechar 
         Height          =   225
         Left            =   5760
         Picture         =   "Form_Sobre.frx":1B65
         ToolTipText     =   "Fechar"
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00272727&
         BackStyle       =   0  'Transparent
         Caption         =   "Sobre"
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
         TabIndex        =   6
         Top             =   120
         Width           =   570
      End
      Begin VB.Image Fundo_Barra_ControlBox 
         Enabled         =   0   'False
         Height          =   465
         Left            =   0
         Picture         =   "Form_Sobre.frx":1E77
         Top             =   0
         Width           =   585
      End
   End
   Begin VB.Label Label_Site_Programa 
      AutoSize        =   -1  'True
      BackColor       =   &H00E7EBEF&
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.vbmoviemanager.comule.com"
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
      Left            =   2160
      TabIndex        =   13
      Top             =   3660
      Width           =   3615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00E7EBEF&
      BackStyle       =   0  'Transparent
      Caption         =   "Site do programa:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   480
      TabIndex        =   12
      Top             =   3660
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E7EBEF&
      BackStyle       =   0  'Transparent
      Caption         =   "Site Oficial:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   480
      TabIndex        =   11
      Top             =   3900
      Width           =   990
   End
   Begin VB.Label Label_Site 
      AutoSize        =   -1  'True
      BackColor       =   &H00E7EBEF&
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.nikyts.com"
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
      Left            =   2160
      TabIndex        =   10
      Top             =   3900
      Width           =   1950
   End
   Begin VB.Label Label_Versao 
      AutoSize        =   -1  'True
      BackColor       =   &H00E7EBEF&
      BackStyle       =   0  'Transparent
      Caption         =   "versão"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   2280
      TabIndex        =   9
      Top             =   1680
      Width           =   585
   End
   Begin VB.Shape Shape_Contorno 
      BorderColor     =   &H00C0C0C0&
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label_Contacto 
      AutoSize        =   -1  'True
      BackColor       =   &H00E7EBEF&
      BackStyle       =   0  'Transparent
      Caption         =   "Contacto: nikyts@hotmail.com"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   480
      TabIndex        =   4
      Top             =   3240
      Width           =   2610
   End
   Begin VB.Label Label_Autor 
      AutoSize        =   -1  'True
      BackColor       =   &H00E7EBEF&
      BackStyle       =   0  'Transparent
      Caption         =   "Desenvolvido por: Nelson do Carmo"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   3000
      Width           =   3120
   End
   Begin VB.Label Label_Direitos 
      AutoSize        =   -1  'True
      BackColor       =   &H00E7EBEF&
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2011 Nikyts Software - Informática e tecnologia"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   480
      TabIndex        =   2
      Top             =   2760
      Width           =   5160
   End
   Begin VB.Label Label_Programa 
      AutoSize        =   -1  'True
      BackColor       =   &H00E7EBEF&
      BackStyle       =   0  'Transparent
      Caption         =   "VBMovieManager"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   2280
      TabIndex        =   1
      Top             =   1440
      Width           =   1890
   End
   Begin VB.Image Image_Logo 
      Enabled         =   0   'False
      Height          =   1920
      Left            =   240
      Picture         =   "Form_Sobre.frx":21BC
      Top             =   600
      Width           =   1920
   End
End
Attribute VB_Name = "Form_Sobre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   MusicLink
'   COPYRIGHT © 2011 Nikyts software ™  -  INFORMÁTICA & TECNOLOGIA
'   WWW.NIKYTS.COM.SAPO.PT
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Declaração das variáveis
Dim bMoveFrom As Boolean, LastPoint As POINTAPI

'API para abrir web
Private Const SW_NORMAL = 1
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Barra_Botoes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Chamar o procedimento
    Repor_Imagens
End Sub

Private Sub Botao_Fechar_Click()
    'Fechar o formulario
    Unload Me
End Sub

Private Sub Botao_Ok_Click()
    'Fechar o formulario
    Unload Me
End Sub

Private Sub Botao_Ok_KeyDown(KeyCode As Integer, Shift As Integer)
    'Atalho de teclas
    If KeyCode = vbKeyReturn Then Botao_Ok_Click
    If KeyCode = vbKeyEscape Then Botao_Ok_Click
End Sub

Private Sub Botao_Ok_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Animar o botão
    If Botao_Ok.Picture = Form_Skin.Botao_Over.Picture Then Exit Sub
    Repor_Imagens
    Botao_Ok.Picture = Form_Skin.Botao_Over.Picture
End Sub

Private Sub Form_Load()
    'Propriedades iniciais do formulário
    Desenhar_Formulario
    
    'Definir os valores de x e y para poder mover o formulário
    iTPPX& = Screen.TwipsPerPixelX
    iTPPY& = Screen.TwipsPerPixelY
    
    'Propriedades iniciais do formulário
    Label_Programa.Caption = App.ProductName
    Label_Versao.Caption = "versão " & App.Major & "." & App.Minor & "." & App.Revision
    Label_Direitos.Caption = App.LegalCopyright
    Label_Autor.Caption = "Desenvolvido por: " & "Nelson do Carmo"
    Label_Contacto.Caption = "Contacto: " & "nikyts@hotmail.com"
End Sub

Public Sub Desenhar_Formulario()
    'Procedimento para ajustar os objectos
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
        .Width = Barra_ControlBox.Width
        .Left = 0
    End With
    
    With Label_Titulo
        .Top = (Barra_ControlBox.ScaleHeight - .Height) / 2
        .Left = 10
    End With
    
    With Botao_Fechar
        .Height = Barra_Botoes.Height
        .Top = 8
        .Left = Barra_ControlBox.Width - .Width - 4
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
    
    With Botao_Ok
        .Top = 16
        .Height = Form_Skin.Botao_Normal.Height
        .Width = Form_Skin.Botao_Normal.Width
        .Left = (Barra_Botoes.Width - .Width) / 2
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
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Chamar o procedimento
    Repor_Imagens
End Sub

Private Sub Label_Ok_Click()
    'Atalho para
    Botao_Ok_Click
End Sub

Private Sub Label_Ok_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Animar o botão
    If Botao_Ok.Picture = Form_Skin.Botao_Over.Picture Then Exit Sub
    Repor_Imagens
    Botao_Ok.Picture = Form_Skin.Botao_Over.Picture
End Sub

Private Sub Label_Site_Click()
    'Abrir página pessoal
    Call ShellExecute(0, "open", Label_Site.Caption, vbNullString, vbNullString, SW_NORMAL)
End Sub

Private Sub Label_Site_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Animar a label
    If Label_Site.ForeColor <> &H80FF& Then
        Label_Site.ForeColor = &H80FF& 'Laranja
    End If
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
    If Me.WindowState = 0 And Tela_Cheia = False Then
        Dim iDX As Long, iDY As Long
        Dim POINT As POINTAPI
        If Not bMoveFrom Then Exit Sub
        GetCursorPos POINT
        iDX& = (POINT.X - LastPoint.X) * iTPPX&
        iDY& = (POINT.Y - LastPoint.Y) * iTPPY&
        LastPoint.X = POINT.X
        LastPoint.Y = POINT.Y
        Me.Move Me.Left + iDX&, Me.Top + iDY&
    End If
End Sub

Private Sub Label_Titulo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    bMoveFrom = False
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
    If Me.WindowState = 0 And Tela_Cheia = False Then
        Dim iDX As Long, iDY As Long
        Dim POINT As POINTAPI
        If Not bMoveFrom Then Exit Sub
        GetCursorPos POINT
        iDX& = (POINT.X - LastPoint.X) * iTPPX&
        iDY& = (POINT.Y - LastPoint.Y) * iTPPY&
        LastPoint.X = POINT.X
        LastPoint.Y = POINT.Y
        Me.Move Me.Left + iDX&, Me.Top + iDY&
    End If
End Sub

Private Sub Barra_ControlBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Largar o formulário para a posição final
    bMoveFrom = False
End Sub

Public Sub Repor_Imagens()
    'Procedimento para repor as imagens originais dos botões após o over
    If Botao_Ok.Picture <> Form_Skin.Botao_Normal.Picture Then
        Botao_Ok.Picture = Form_Skin.Botao_Normal.Picture
    End If
    
    If Label_Site.ForeColor <> &HE6964D Then
        Label_Site.ForeColor = &HE6964D 'Azul
    End If
    
    If Label_Site_Programa.ForeColor <> &HE6964D Then
        Label_Site_Programa.ForeColor = &HE6964D 'Azul
    End If
End Sub

Private Sub Label_Site_Programa_Click()
    'Ver site do programa
    Call ShellExecute(0, "open", Label_Site_Programa.Caption, vbNullString, vbNullString, SW_NORMAL)
End Sub

Private Sub Label_Site_Programa_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Animar a label
    If Label_Site_Programa.ForeColor <> &H80FF& Then
        Label_Site_Programa.ForeColor = &H80FF& 'Laranja
    End If
End Sub
