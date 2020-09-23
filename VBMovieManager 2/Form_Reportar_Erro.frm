VERSION 5.00
Begin VB.Form Form_Reportar_Erro 
   Appearance      =   0  'Flat
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   0  'None
   ClientHeight    =   6075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5925
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
   ScaleHeight     =   405
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox Lista_Assunto 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   990
      IntegralHeight  =   0   'False
      ItemData        =   "Form_Reportar_Erro.frx":0000
      Left            =   360
      List            =   "Form_Reportar_Erro.frx":0002
      TabIndex        =   19
      Top             =   2160
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   360
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   345
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1800
      Width           =   5175
      Begin VB.PictureBox Seta_Assunto 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4875
         Picture         =   "Form_Reportar_Erro.frx":0004
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   15
         Width           =   285
      End
      Begin VB.TextBox Text_Assunto 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   4860
      End
      Begin VB.Shape Shape_Assunto 
         BorderColor     =   &H00C0C0C0&
         Height          =   315
         Left            =   0
         Top             =   0
         Width           =   5175
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2115
      Left            =   360
      ScaleHeight     =   141
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   345
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2520
      Width           =   5175
      Begin VB.TextBox Text_Mensagem 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   2085
         Left            =   15
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   15
         Width           =   5145
      End
      Begin VB.Shape Shape_Mensagem 
         BorderColor     =   &H00C0C0C0&
         Height          =   2115
         Left            =   0
         Top             =   0
         Width           =   5175
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   360
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   345
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1200
      Width           =   5175
      Begin VB.TextBox Text_Email 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   15
         TabIndex        =   0
         Top             =   15
         Width           =   5145
      End
      Begin VB.Shape Shape_Email 
         BorderColor     =   &H00C0C0C0&
         Height          =   315
         Left            =   0
         Top             =   0
         Width           =   5175
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
      ScaleWidth      =   393
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   5895
      Begin VB.Label Label_Titulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00272727&
         BackStyle       =   0  'Transparent
         Caption         =   "Suporte técnico"
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
         TabIndex        =   7
         Top             =   120
         Width           =   1530
      End
      Begin VB.Image Botao_Fechar 
         Height          =   225
         Left            =   5280
         Picture         =   "Form_Reportar_Erro.frx":04F6
         ToolTipText     =   "Fechar"
         Top             =   120
         Width           =   240
      End
      Begin VB.Image Fundo_Barra_ControlBox 
         Enabled         =   0   'False
         Height          =   465
         Left            =   0
         Picture         =   "Form_Reportar_Erro.frx":0808
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
      ScaleWidth      =   393
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5160
      Width           =   5895
      Begin VB.PictureBox Botao_Cancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   495
         Left            =   4320
         Picture         =   "Form_Reportar_Erro.frx":0B4D
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   61
         TabIndex        =   13
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
            TabIndex        =   14
            Top             =   135
            Width           =   915
         End
      End
      Begin VB.PictureBox Botao_Ok 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
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
         Height          =   495
         Left            =   3240
         Picture         =   "Form_Reportar_Erro.frx":2347
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   61
         TabIndex        =   3
         Top             =   240
         Width           =   915
         Begin VB.Label Label_Ok 
            Alignment       =   2  'Center
            BackColor       =   &H00272727&
            BackStyle       =   0  'Transparent
            Caption         =   "Ok"
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   0
            TabIndex        =   9
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
         Picture         =   "Form_Reportar_Erro.frx":3B41
         Top             =   0
         Width           =   585
      End
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "(Para possivel contacto caso seja necessário)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   180
      Left            =   900
      TabIndex        =   18
      Top             =   960
      Width           =   3480
   End
   Begin VB.Label Label_Assunto 
      AutoSize        =   -1  'True
      BackColor       =   &H00EEEEEE&
      BackStyle       =   0  'Transparent
      Caption         =   "Assunto"
      Height          =   195
      Left            =   360
      TabIndex        =   16
      Top             =   1560
      Width           =   675
   End
   Begin VB.Label Label_Erro 
      BackColor       =   &H00F5F5F5&
      BackStyle       =   0  'Transparent
      Caption         =   "Campo de preenchimento obrigatório."
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   720
      TabIndex        =   10
      Top             =   630
      Visible         =   0   'False
      Width           =   4725
   End
   Begin VB.Image Image_Erro 
      Enabled         =   0   'False
      Height          =   210
      Left            =   420
      Picture         =   "Form_Reportar_Erro.frx":3EAC
      Top             =   615
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Shape Shape_Contorno 
      BorderColor     =   &H00C0C0C0&
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label_Mensagem 
      AutoSize        =   -1  'True
      BackColor       =   &H00EEEEEE&
      BackStyle       =   0  'Transparent
      Caption         =   "Mensagem"
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   2280
      Width           =   915
   End
   Begin VB.Label Label_De 
      AutoSize        =   -1  'True
      BackColor       =   &H00EEEEEE&
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   960
      Width           =   465
   End
   Begin VB.Shape Shape_Erro 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H008080FF&
      Height          =   255
      Left            =   360
      Top             =   600
      Visible         =   0   'False
      Width           =   5145
   End
End
Attribute VB_Name = "Form_Reportar_Erro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   GADGETS
'   COPYRIGHT © 2010 Nikyts software ™  -  INFORMÁTICA & TECNOLOGIA
'   WWW.NIKYTS.COM.SAPO.PT
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Declaração das variáveis
Dim bMoveFrom As Boolean, LastPoint As POINTAPI

Private Sub Barra_Botoes_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Chamar o procedimento
    Repor_Imagens
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

Private Sub Botao_Fechar_Click()
    'Fechar formulário
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

Private Sub Form_Click()
    'Ocultar lista
    Lista_Assunto.Visible = False
End Sub

Private Sub Form_Load()
    'Iniciar o formulário
    Desenhar_Formulario
    
    'Variáveis para poder mover o formulário
    iTPPX& = Screen.TwipsPerPixelX
    iTPPY& = Screen.TwipsPerPixelY
    
    'Pasta das imagens
    With Lista_Assunto
        .AddItem "Reportar erro"
        .AddItem "Sugestão"
        .AddItem "Questão"
        .AddItem "Outro"
        '.ListIndex = 0
    End With
End Sub

Public Sub Desactivar_Objectos()
    'as textboxs
    Text_Mensagem.Enabled = False
End Sub

Public Sub Activar_Objectos()
    'as textboxs
    Text_Mensagem.Enabled = True
End Sub

Public Sub Limpa_Campos()
    'Limpa o conteudo das caixas de texto
    Text_Email.Text = ""
    Text_Assunto.Text = ""
    Text_Mensagem.Text = ""
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Chamar o procedimento
    Repor_Imagens
End Sub

Private Sub Label_Cancelar_Click()
    'Cancelar operação
    Unload Me
End Sub

Private Sub Label_Cancelar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Animar o botão
    If Botao_Cancelar.Picture = Form_Skin.Botao_Over.Picture Then Exit Sub
    Repor_Imagens
    Botao_Cancelar.Picture = Form_Skin.Botao_Over.Picture
End Sub

Private Sub Label_Ok_Click()
    'Verificar o preencimento das textboxs
    On Error GoTo Corrige_Erro
    Shape_Erro.Visible = False
    Label_Erro.Visible = False
    Image_Erro.Visible = False
    
    'Verifica se o campo email está no formato correcto
    If Not IsEmail(Text_Email.Text) Then
        Label_Erro.Caption = "Indique um endereço de email válido."
        Shape_Erro.Visible = True
        Label_Erro.Visible = True
        Image_Erro.Visible = True
        Text_Email.SetFocus
        Exit Sub
    End If
    
    Label_Titulo.Caption = "Suporte técnico - Aguarde..."
    Me.MousePointer = 11
    Dim servidor As XMLHTTP60: Set servidor = New XMLHTTP60
    servidor.Open "GET", "http://www.nikyts.com/suporte/" & "enviarmensagem.asp?Email=" & Text_Email.Text & "&Assunto=" & App.ProductName & " - " & Text_Assunto.Text & "&Mensagem=" & Text_Mensagem.Text, False
    servidor.send 'envia o pedido para o servidor

    'Verificar os dados acesso
    If Not InStr(servidor.responseText, "HTTP Error") > 0 Then
        If servidor.readyState = 4 And servidor.Status = 200 And servidor.responseText = "sucesso" Then ' 4 - deu resposta e 200 validou
            Mensagem_de_Aviso "Informação", "A sua mensagem foi enviada com sucesso."
            
            'Limpar os campos para se poder enviar uma nova mensagem
            Label_Titulo.Caption = "Suporte técnico"
            Me.MousePointer = 0
            Limpa_Campos
            Botao_Ok.Enabled = False
            Label_Ok.Enabled = False
            Text_Email.SetFocus
        End If
    End If
    
Exit Sub
Corrige_Erro:
Select Case Err.Number
    Case -2146697211
        Mensagem_de_Aviso "Erro", "Ocorreu um erro ao tentar conectar-se ao servidor." & vbNewLine & "Possivel causa: a sua conexão á internet poderá estar desligada."
        
    Case Else
        'Correção de outros erros que poderão surgir
        Mensagem_de_Aviso "Erro", "Ocorreu um erro durante a execução do programa." & vbNewLine & "Erro nº: " & Err.Number & vbNewLine & "Descrição: " & Err.Description
End Select
End Sub

Private Sub Label_Ok_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Animar o botão
    If Botao_Ok.Picture = Form_Skin.Botao_Over.Picture Then Exit Sub
    Repor_Imagens
    Botao_Ok.Picture = Form_Skin.Botao_Over.Picture
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
    
    'Control box
    With Botao_Fechar
        .Top = 8
        .Left = Barra_ControlBox.Width - .Width - 4
    End With
    
    'Barra_Botoes
    With Barra_Botoes
        .Height = Fundo_Barra_Botoes.Height
        .Top = Me.ScaleHeight - .Height - 1
        .Width = Barra_ControlBox.ScaleWidth - 2
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
    
    With Lista_Assunto
        .Top = Picture4.Top + Picture4.ScaleHeight
        .Width = Picture4.ScaleWidth
        .Left = Picture4.Left
    End With
End Sub

Private Sub Lista_Assunto_Click()
    'Ocultar lista
    Text_Assunto.Text = Lista_Assunto.Text
    Lista_Assunto.Visible = False
    Text_Assunto.SetFocus
End Sub

Private Sub Seta_Assunto_Click()
    'Ver/ocultar lista
    If Lista_Assunto.Visible = True Then
        Lista_Assunto.Visible = False
    Else
        Lista_Assunto.Visible = True
    End If
End Sub

Private Sub Text_Assunto_Change()
    'Chamar o procedimento
    Verificar_o_Prenchimento
End Sub

Private Sub Text_Assunto_Click()
    'Ocultar lista
    'Lista_Assunto.Visible = False
End Sub

Private Sub Text_Assunto_KeyDown(KeyCode As Integer, Shift As Integer)
    'Altalho para percorrer as linha da combo de imagens
    On Error Resume Next
    If KeyCode = vbKeyUp Then 'Para cima
        If Lista_Assunto.ListIndex <> 0 Then
            Lista_Assunto.ListIndex = Lista_Assunto.ListIndex - 1
            Text_Assunto.Text = Lista_Assunto.Text
        End If
    End If
    If KeyCode = vbKeyDown Then 'Para baixo
        If Lista_Assunto.ListIndex <> Lista_Assunto.ListCount - 1 Then
            Lista_Assunto.ListIndex = Lista_Assunto.ListIndex + 1
            Text_Assunto.Text = Lista_Assunto.Text
        End If
    End If
End Sub

Private Sub Text_Email_Change()
    'Chamar o procedimento
    Verificar_o_Prenchimento
End Sub

Private Sub Text_Email_Click()
    'Ocultar lista
    Lista_Assunto.Visible = False
End Sub

Private Sub Text_Email_GotFocus()
    'Contorno da text box ao receber o focus
    Shape_Email.BorderColor = &HE6964D    'Azul
End Sub

Private Sub Text_Email_LostFocus()
    'Contorno da text box ao perder o focus
    Shape_Email.BorderColor = &HC0C0C0   'Cinzento
End Sub

Private Sub Text_Mensagem_Change()
    'Chamar o procedimento
    Verificar_o_Prenchimento
End Sub

Private Sub Text_Mensagem_Click()
    'Ocultar lista
    Lista_Assunto.Visible = False
End Sub

Private Sub Text_Mensagem_GotFocus()
    'Contorno da text box ao receber o focus
    Shape_Mensagem.BorderColor = &HE6964D    'Azul
End Sub

Private Sub Text_Mensagem_LostFocus()
    'Contorno da text box ao perder o focus
    Shape_Mensagem.BorderColor = &HC0C0C0   'Cinzento
End Sub

Public Sub Repor_Imagens()
    'Procedimento para repor as imagens originais dos botões após o over
    If Botao_Cancelar.Picture <> Form_Skin.Botao_Normal.Picture Then
        Botao_Cancelar.Picture = Form_Skin.Botao_Normal.Picture
    End If
    
    If Botao_Ok.Picture <> Form_Skin.Botao_Normal.Picture Then
        Botao_Ok.Picture = Form_Skin.Botao_Normal.Picture
    End If
End Sub

Private Sub Text_Assunto_GotFocus()
    'Contorno da text box ao receber o focus
    Shape_Assunto.BorderColor = &HE6964D    'Azul
End Sub

Private Sub Text_Assunto_LostFocus()
    'Contorno da text box ao perder o focus
    Shape_Assunto.BorderColor = &HC0C0C0   'Cinzento
End Sub

Public Sub Verificar_o_Prenchimento()
    'Procedimento para verificar se as caixas de texto estão devidamente bem preenchidas
    If Len(Trim(Text_Email.Text)) = 0 Or Len(Trim(Text_Assunto.Text)) = 0 Or Len(Trim(Text_Mensagem.Text)) = 0 Then
        Botao_Ok.Enabled = False
        Label_Ok.Enabled = False
        
    Else
        Botao_Ok.Enabled = True
        Label_Ok.Enabled = True
    End If
End Sub
