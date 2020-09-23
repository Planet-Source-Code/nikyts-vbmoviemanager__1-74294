VERSION 5.00
Begin VB.Form Form_Mensagem 
   Appearance      =   0  'Flat
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   0  'None
   ClientHeight    =   3180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7020
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
   ScaleHeight     =   212
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   468
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Botao_Nao 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3840
      Picture         =   "Form_Mensagem.frx":0000
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   1
      Top             =   1920
      Visible         =   0   'False
      Width           =   915
      Begin VB.Shape Contorno_Nao 
         BorderColor     =   &H00E6964D&
         Height          =   495
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label_Nao 
         Alignment       =   2  'Center
         BackColor       =   &H00272727&
         BackStyle       =   0  'Transparent
         Caption         =   "Não"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   0
         TabIndex        =   8
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
      Left            =   2640
      Picture         =   "Form_Mensagem.frx":17FA
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   2
      Top             =   1920
      Visible         =   0   'False
      Width           =   915
      Begin VB.Shape Contorno_Ok 
         BorderColor     =   &H00E6964D&
         Height          =   495
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label_Ok 
         Alignment       =   2  'Center
         BackColor       =   &H00272727&
         BackStyle       =   0  'Transparent
         Caption         =   "Ok"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   0
         TabIndex        =   7
         Top             =   135
         Width           =   915
      End
   End
   Begin VB.PictureBox Botao_Sim 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1560
      Picture         =   "Form_Mensagem.frx":2FF4
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   915
      Begin VB.Shape Contorno_Sim 
         BorderColor     =   &H00E6964D&
         Height          =   495
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label_Sim 
         Alignment       =   2  'Center
         BackColor       =   &H00272727&
         BackStyle       =   0  'Transparent
         Caption         =   "Sim"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   0
         TabIndex        =   6
         Top             =   135
         Width           =   915
      End
   End
   Begin VB.PictureBox Barra_ControlBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00272727&
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
      ScaleWidth      =   385
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   5775
      Begin VB.Label Label_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00272727&
         BackStyle       =   0  'Transparent
         Caption         =   "Atenção"
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
         TabIndex        =   5
         Top             =   120
         Width           =   795
      End
      Begin VB.Image Fundo_Barra_ControlBox 
         Enabled         =   0   'False
         Height          =   465
         Left            =   0
         Picture         =   "Form_Mensagem.frx":47EE
         Top             =   0
         Width           =   585
      End
   End
   Begin VB.Shape Shape_Contorno 
      BorderColor     =   &H00C0C0C0&
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   615
   End
   Begin VB.Image Pic_Mensagem 
      Enabled         =   0   'False
      Height          =   480
      Left            =   360
      Top             =   720
      Width           =   480
   End
   Begin VB.Label Label_Mensagem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mensagem"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1080
      TabIndex        =   3
      Top             =   840
      Width           =   915
   End
End
Attribute VB_Name = "Form_Mensagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Nikyts tag editor
'   COPYRIGHT © 2011 Nikyts software ™  -  INFORMÁTICA & TECNOLOGIA
'   WWW.NIKYTS.COM.SAPO.PT
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Declaração das variáveis
Dim bMoveFrom As Boolean, LastPoint As POINTAPI

Private Sub Botao_Nao_Click()
    'Fechar o formulário
    Resposta = "Nao"
    Unload Me
End Sub

Private Sub Botao_Nao_GotFocus()
    'Colocar o focus no botao
    Contorno_Nao.Visible = True
End Sub

Private Sub Botao_nao_KeyDown(KeyCode As Integer, Shift As Integer)
    'Atalho de teclas
    If KeyCode = vbKeyReturn Then Botao_Nao_Click
End Sub

Private Sub Botao_nao_LostFocus()
    'Remover o focus no botao
    Contorno_Nao.Visible = False
End Sub

Private Sub Botao_Nao_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Animar o botão
    If Botao_Nao.Picture = Form_Skin.Botao_Over.Picture Then Exit Sub
    Repor_Imagens
    Botao_Nao.Picture = Form_Skin.Botao_Over.Picture
End Sub

Private Sub Botao_Ok_Click()
    'Fechar o formulário
    Unload Me
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

Private Sub Botao_Sim_Click()
    'Fechar o formulário
    Resposta = "Sim"
    Unload Me
End Sub

Private Sub Botao_Sim_GotFocus()
    'Colocar o focus no botao
    Contorno_Sim.Visible = True
End Sub

Private Sub Botao_Sim_KeyDown(KeyCode As Integer, Shift As Integer)
    'Atalho de teclas
    If KeyCode = vbKeyReturn Then Botao_Sim_Click
End Sub

Private Sub Botao_Sim_LostFocus()
    'Remover o focus no botao
    Contorno_Sim.Visible = False
End Sub

Private Sub Botao_Sim_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Animar o botão
    If Botao_Sim.Picture = Form_Skin.Botao_Over.Picture Then Exit Sub
    Repor_Imagens
    Botao_Sim.Picture = Form_Skin.Botao_Over.Picture
End Sub

Private Sub Form_Load()
    'Chamar o procedimento
    Desenhar_Formulario
    Label_Titulo.Caption = App.ProductName
    
    'Propriedades iniciais do formulário
    iTPPX& = Screen.TwipsPerPixelX
    iTPPY& = Screen.TwipsPerPixelY
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Chamar o procedimento
    Repor_Imagens
End Sub

Private Sub Form_Resize()
    'Chamar o procedimento
    Desenhar_Formulario
End Sub

Private Sub Label_Nao_Click()
    'Atalho para
    Botao_Nao_Click
End Sub

Private Sub Label_Nao_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Animar o botão
    If Botao_Nao.Picture = Form_Skin.Botao_Over.Picture Then Exit Sub
    Repor_Imagens
    Botao_Nao.Picture = Form_Skin.Botao_Over.Picture
End Sub

Private Sub Label_Ok_Click()
    'Atalho para
    Botao_Ok_Click
End Sub

Private Sub Label_Ok_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Animar o botão
    If Botao_Ok.Picture = Form_Skin.Botao_Over.Picture Then Exit Sub
    Repor_Imagens
    Botao_Ok.Picture = Form_Skin.Botao_Over.Picture
End Sub

Private Sub Label_Sim_Click()
    'Atalho para
    Botao_Sim_Click
End Sub

Public Sub Desenhar_Formulario()
    'Procedimento para desenhar o formulario, ajustando os objectos
    'Ajustar o formulário
    With Me
        'Descrição dos 3 ultimos valores. Espaço entre: borda do form -> icon, icon -> label messagem, label_mensagem -> borda do form
        .Width = Screen.TwipsPerPixelX * (Label_Mensagem.Width + Pic_Mensagem.Width + 40 + 17 + 80)
        .Height = Screen.TwipsPerPixelX * (Barra_ControlBox.Height + 36 + Label_Mensagem.Height + 50 + 10)
    End With
    
    'Shape_Contorno
    With Shape_Contorno
        .Height = Me.ScaleHeight
        .Top = 0
        .Width = Me.ScaleWidth
        .Left = 0
    End With
    
    'Barra_ControlBox
    With Barra_ControlBox
        .Height = Fundo_Barra_ControlBox.Height
        .Top = 0
        .Width = Me.ScaleWidth
        .Left = 0
    End With
    
    'Fundo_Barra_ControlBox
    With Fundo_Barra_ControlBox
        .Top = 0
        .Stretch = True
        .Width = Me.ScaleWidth
        .Left = 0
    End With
    
    With Label_Titulo
        .Top = (Barra_ControlBox.ScaleHeight - .Height) / 2
        .Left = 10
    End With
    
    'Botao_Ok
    With Botao_Ok
        .Height = Form_Skin.Botao_Normal.Height
        .Left = (Me.ScaleWidth - .Width) / 2
        .Width = Form_Skin.Botao_Normal.Width
        .Top = Me.ScaleHeight - .Height - 10
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
    
    'Botao_Sim
    With Botao_Sim
        .Height = Form_Skin.Botao_Normal.Height
        .Left = (Me.ScaleWidth / 2) - .Width - 3
        .Width = Form_Skin.Botao_Normal.Width
        .Top = Me.ScaleHeight - .Height - 10
    End With
    
    With Contorno_Sim
        .Top = 0
        .Height = Botao_Sim.ScaleHeight
        .Left = 0
        .Width = Botao_Sim.ScaleWidth
    End With
    
    With Label_Sim
        .Top = (Botao_Sim.ScaleHeight - .Height) / 2
        .Width = Botao_Sim.ScaleWidth
    End With
    
    'Botao_Nao
    With Botao_Nao
        .Height = Form_Skin.Botao_Normal.Height
        .Left = (Me.ScaleWidth / 2) + 3
        .Width = Form_Skin.Botao_Normal.Width
        .Top = Me.ScaleHeight - .Height - 10
    End With
    
    With Contorno_Nao
        .Top = 0
        .Height = Botao_Nao.ScaleHeight
        .Left = 0
        .Width = Botao_Nao.ScaleWidth
    End With
    
    With Label_Nao
        .Top = (Botao_Nao.ScaleHeight - .Height) / 2
        .Width = Botao_Nao.ScaleWidth
    End With
End Sub

Private Sub Label_Sim_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Animar o botão
    If Botao_Sim.Picture = Form_Skin.Botao_Over.Picture Then Exit Sub
    Repor_Imagens
    Botao_Sim.Picture = Form_Skin.Botao_Over.Picture
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
    If Me.WindowState = 0 Then
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
    If Me.WindowState = 0 Then
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

Public Sub Repor_Imagens()
    'Procedimento para repor as imagens originais dos botões após o over
    If Botao_Nao.Picture <> Form_Skin.Botao_Normal.Picture Then
        Botao_Nao.Picture = Form_Skin.Botao_Normal.Picture
    End If
    
    If Botao_Ok.Picture <> Form_Skin.Botao_Normal.Picture Then
        Botao_Ok.Picture = Form_Skin.Botao_Normal.Picture
    End If
    
    If Botao_Sim.Picture <> Form_Skin.Botao_Normal.Picture Then
        Botao_Sim.Picture = Form_Skin.Botao_Normal.Picture
    End If
End Sub
