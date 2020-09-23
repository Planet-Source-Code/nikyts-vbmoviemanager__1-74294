VERSION 5.00
Begin VB.Form Form_Instalar 
   Appearance      =   0  'Flat
   BackColor       =   &H00292929&
   BorderStyle     =   0  'None
   Caption         =   "VBMovieManager"
   ClientHeight    =   5970
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   7485
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
   Icon            =   "Form_Instalar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   398
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   499
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      BackColor       =   &H00292929&
      ForeColor       =   &H00FFFFFF&
      Height          =   2115
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox Pic_Pasta_Destino 
      Appearance      =   0  'Flat
      BackColor       =   &H00393939&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2250
      Left            =   1680
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   361
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1980
      Width           =   5415
      Begin VB.FileListBox File1 
         Appearance      =   0  'Flat
         BackColor       =   &H00292929&
         ForeColor       =   &H00FFFFFF&
         Height          =   1590
         Left            =   30
         TabIndex        =   10
         Top             =   30
         Width           =   5355
      End
      Begin VB.Shape Shape_Pasta_Destino 
         BorderColor     =   &H00292929&
         Height          =   2220
         Left            =   15
         Top             =   15
         Width           =   5385
      End
   End
   Begin VB.PictureBox Botao_Cancelar 
      Appearance      =   0  'Flat
      BackColor       =   &H00292929&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   3120
      Picture         =   "Form_Instalar.frx":57E2
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   5
      Top             =   5160
      Width           =   1215
      Begin VB.Label Label_Cancelar 
         Alignment       =   2  'Center
         BackColor       =   &H00272727&
         BackStyle       =   0  'Transparent
         Caption         =   "Cancelar"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   0
         TabIndex        =   6
         Top             =   135
         Width           =   1215
      End
      Begin VB.Shape Contorno_Cancelar 
         BorderColor     =   &H00E6964D&
         Height          =   435
         Left            =   15
         Top             =   15
         Visible         =   0   'False
         Width           =   1185
      End
   End
   Begin VB.PictureBox Botao_Ok 
      Appearance      =   0  'Flat
      BackColor       =   &H00292929&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   1680
      Picture         =   "Form_Instalar.frx":75B0
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   3
      Top             =   5160
      Width           =   1215
      Begin VB.Label Label_Ok 
         Alignment       =   2  'Center
         BackColor       =   &H00272727&
         BackStyle       =   0  'Transparent
         Caption         =   "Registar"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   0
         TabIndex        =   4
         Top             =   135
         Width           =   1215
      End
      Begin VB.Shape Contorno_Ok 
         BorderColor     =   &H00E6964D&
         Height          =   435
         Left            =   15
         Top             =   15
         Visible         =   0   'False
         Width           =   1185
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
      ScaleWidth      =   489
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7335
      Begin VB.Image Botao_Fechar 
         Height          =   225
         Left            =   6960
         Picture         =   "Form_Instalar.frx":937E
         ToolTipText     =   "Fechar"
         Top             =   120
         Width           =   240
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
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   3390
      End
      Begin VB.Image Fundo_Barra_ControlBox 
         Enabled         =   0   'False
         Height          =   465
         Left            =   0
         Picture         =   "Form_Instalar.frx":9690
         Top             =   0
         Width           =   585
      End
   End
   Begin VB.Label Label_Estado 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   192
      Left            =   1680
      TabIndex        =   11
      Top             =   4320
      Width           =   5400
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "que registar algumas componentes."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1680
      TabIndex        =   8
      Top             =   1200
      Width           =   3075
   End
   Begin VB.Label Label_Prefencias 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "O programa para ser executado correctamente poderá ter"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1680
      TabIndex        =   2
      Top             =   960
      Width           =   5025
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   1200
      Left            =   240
      Picture         =   "Form_Instalar.frx":99D5
      Top             =   720
      Width           =   1200
   End
   Begin VB.Shape Shape_Contorno 
      BorderColor     =   &H00000000&
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "Form_Instalar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Nikyts MEDIA PLAYER
'   COPYRIGHT © 2011 Nikyts software ™  -  INFORMÁTICA & TECNOLOGIA
'   WWW.NIKYTS.COM.SAPO.PT
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Variáveis para poder mover o formulário
Dim bMoveFrom As Boolean, LastPoint As POINTAPI

'API's para registar as ocx utilizadas pelo programa
Private Const STATUS_WAIT_0 = &H0
Private Const WAIT_OBJECT_0 = ((STATUS_WAIT_0) + 0)

Private Declare Sub ExitThread Lib "kernel32.dll" (ByVal lngExitCode As Long)
Private Declare Function LoadLibraryRegister Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibraryRegister Lib "kernel32.dll" Alias "FreeLibrary" (ByVal hLibModule As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetProcAddressRegister Lib "kernel32.dll" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CreateThreadForRegister Lib "kernel32.dll" Alias "CreateThread" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lpparameter As Long, ByVal dwCreationFlags As Long, lngThreadID As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32.dll" (ByVal lngThread As Long, lpExitCode As Long) As Long

Public Enum EDllRegister
  DllRegisterServer = 1
  DllUnRegisterServer = 2
End Enum

Public Enum EDllReturnConstants
  edrcNone = 0
  edrcErrorOnLoadInMemory = 1
  edrcInvalidActivex = 2
  edrcFailedRegistration = 3
  edrcRegistered = 4
  edrcUnregistered = 5
End Enum

Private mReturn As EDllReturnConstants

'Variável para identificar as pastas do windows
Dim Pasta_Sistema As String

Public Sub Botao_Cancelar_Click()
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

Private Sub Botao_Fechar_Click()
    'Fechar o formulário
    Unload Me
    End
End Sub

Public Sub Botao_Ok_Click()
    'atalho para
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

Private Sub Form_Load()
    On Error GoTo Corrige_Erro
    'Iniciar o formulário
    Desenhar_Formulario
    
    'Variáveis para poder mover o formulário
    iTPPX& = Screen.TwipsPerPixelX
    iTPPY& = Screen.TwipsPerPixelY
            
    'Pasta dos componentes
    Dir1.path = App.path & "\Components\"
    File1.Pattern = "*.ocx;*.dll;*.tlb" 'Filtar os ficheiros
    File1.path = Dir1.path
    
    'Verificar qual é a pasta do system32
    Pasta_Sistema = Environ$("windir") & IIf(Len(Environ$("OS")), "\SYSTEM32", "\SYSTEM")
    Verificar_Pastas
    If File1.ListCount <> 0 Then
        Verificar_Componentes
    Else
        Form_Principal.Show
        Unload Me
    End If
    
Exit Sub
Corrige_Erro:
Select Case Err.Number
    Case Else
        Mensagem_de_Aviso "Erro", "Ocorreu um erro durante a execução do programa." & vbNewLine & "Erro nº: " & Err.Number & vbNewLine & "Descrição: " & Err.Description
End Select
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
    
    With Botao_Ok
        .Height = Form_Skin.Botao_Instalador.Height
        .Width = Form_Skin.Botao_Instalador.Width
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
    
    With Botao_Cancelar
        .Height = Form_Skin.Botao_Instalador.Height
        .Width = Form_Skin.Botao_Instalador.Width
        .Left = Botao_Ok.Left + Botao_Ok.ScaleWidth + 8
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
    
    With Pic_Pasta_Destino
        .Height = File1.Height + 4
        .Width = File1.Width + 4
    End With
    
    With File1
        .Top = 2
        .Left = 2
    End With
    
    With Shape_Pasta_Destino
        .Height = File1.Height + 2
        .Top = File1.Top - 1
        .Width = File1.Width + 2
        .Left = File1.Left - 1
    End With
End Sub

Private Sub Label_Cancelar_Click()
    'Atalho para
    Botao_Fechar_Click
End Sub

Private Sub Label_Ok_Click()
    On Error GoTo Corrige_Erro
    'Iniciar a instalação
    If Label_Ok.Caption = "Registar" Then
        'Percorrer a lista das compoennets do programa
        Label_Estado.Caption = "Verificando o registo das componentes..."
        Dim Linha As Integer
        For Linha = 0 To File1.ListCount - 1
            File1.ListIndex = Linha
            
            'Verificar se a componente já existe na pasta do windows, caso não exista copia a componente para a pasta e regista
            If Dir(Pasta_Sistema & "\" & File1.List(Linha), vbDirectory) = Empty Then
                'Copia a componente para a pasta do windows
                FileCopy App.path & "\Components\" & File1.List(Linha), Pasta_Sistema & "\" & File1.List(Linha)
                'Regista a componente
                Registar_Ocx_Componente Pasta_Sistema & "\" & File1.List(Linha)
            End If
        Next Linha
        Label_Estado.Caption = "Operação concluida com sucesso"
        Label_Ok.Caption = "Concluir"
    
    ElseIf Label_Ok.Caption = "Concluir" Then
        Form_Principal.Show
        Unload Me
    End If
    
Exit Sub
Corrige_Erro:
Mensagem_de_Aviso "Erro", "Ocorreu um erro ao registar as componetes do programa." & vbNewLine & "Sugestão: execute o programa como administrador."
Unload Me
End
End Sub

Public Function Registar_Componente(ByVal FileName As String, ByVal RegFunction As EDllRegister) As EDllReturnConstants
    'Função para registar as componentes utilizadas pelo programa
    Dim lngLib As Long
    Dim lngProcAddress As Long
    Dim lngThreadID As Long
    Dim lngSucess As Long
    Dim lngExitCode As Long
    Dim lngThread As Long
    
    'Por padrao passa para nada o retorno...
    Registar_Componente = edrcNone
    
    If FileName = "" Then Exit Function
    
    lngLib = LoadLibraryRegister(FileName)
    If lngLib = 0 Then
       Registar_Componente = edrcErrorOnLoadInMemory
       Exit Function
    End If
    
    Select Case RegFunction
    Case EDllRegister.DllRegisterServer
        lngProcAddress = GetProcAddressRegister(lngLib, "DllRegisterServer")
    Case EDllRegister.DllUnRegisterServer
        lngProcAddress = GetProcAddressRegister(lngLib, "DllUnregisterServer")
    Case Else
    End Select
    
    If lngProcAddress = 0 Then
       Registar_Componente = edrcInvalidActivex
       If lngLib Then Call FreeLibraryRegister(lngLib)
       Exit Function
    Else
       lngThread = CreateThreadForRegister(ByVal 0&, 0&, ByVal lngProcAddress, ByVal 0&, 0&, lngThreadID)
       If lngThread Then
            lngSucess = (WaitForSingleObject(lngThread, 10000) = WAIT_OBJECT_0)
            If Not lngSucess Then
               Call GetExitCodeThread(lngThread, lngExitCode)
               Call ExitThread(lngExitCode)
               Registar_Componente = edrcFailedRegistration
               If lngLib Then Call FreeLibraryRegister(lngLib)
               Exit Function
            Else
                If RegFunction = DllRegisterServer Then
                    Registar_Componente = edrcRegistered
                ElseIf RegFunction = DllUnRegisterServer Then
                    Registar_Componente = edrcUnregistered
                End If
            End If
            Call CloseHandle(lngThread)
            If lngLib Then Call FreeLibraryRegister(lngLib)
       End If
    End If
End Function

Public Sub Registar_Ocx_Componente(Directorio_da_Componente As String)
    'Procedimento para registar as ocx do programa
    On Error Resume Next
    'Proceder ao registro das componentes
    mReturn = Registar_Componente(Trim(Directorio_da_Componente), DllRegisterServer)
    Select Case mReturn
        Case edrcErrorOnLoadInMemory
            Label_Estado.Caption = "O componente " & File1.List(Linha) & " não foi carregado na memória"
        Case edrcInvalidActivex
            Label_Estado.Caption = "O componente " & File1.List(Linha) & " é inválido"
        Case edrcFailedRegistration
            Label_Estado.Caption = "Ocorreu um erro ao registar o componente " & File1.List(Linha)
        Case edrcRegistered
            Label_Estado.Caption = "O componente " & File1.List(Linha) & " foi registado com sucesso"
    End Select
End Sub

Public Sub Verificar_Componentes()
    'Procedimento para verificar se existe alguma componente que não esteja registada
    DoEvents
    Dim Linha As Integer
    For Linha = 0 To File1.ListCount - 1
        File1.ListIndex = Linha
        'Verificar se a componente já existe na pasta do windows, caso não exista mostra o formulário install
        If Dir(Pasta_Sistema & "\" & File1.List(Linha), vbDirectory) = Empty Then
'            MsgBox (Pasta_Sistema & "\" & File1.List(Linha) & " Não existe")
            Me.Visible = True
            Exit Sub
        
        Else
'            MsgBox (Pasta_Sistema & "\" & File1.List(Linha) & " existe")
            If Linha = File1.ListCount - 1 Then
'            MsgBox ("Concluido")
                Me.Visible = False
                Form_Principal.Show
                Unload Me
            End If
        End If
    Next Linha
End Sub

Public Sub Verificar_Pastas()
    'Procedimento para verificar se as pastas utilizadas pelo programa existem
    If Not ArquivoExiste(App.path & "\Components", True) Then
        MkDir App.path & "\Components\"
    End If
    
    If Not ArquivoExiste(App.path & "\Covers", True) Then
        MkDir App.path & "\Covers\"
    End If
    
    If Not ArquivoExiste(App.path & "\Data", True) Then
        MkDir App.path & "\Data\"
    End If
    
    If Not ArquivoExiste(App.path & "\Options", True) Then
        MkDir App.path & "\Options\"
    End If
End Sub
