Attribute VB_Name = "Module_Geral"
Option Explicit
'Api's para mover o formulário
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetCapture Lib "user32" () As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Type POINTAPI
        x As Long
        y As Long
End Type
Public Const SRCCOPY = &HCC0020

'Posição x e y
Global iTPPY As Long
Global iTPPX As Long

'API para o procedimento alway's on top
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal CX As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'Para ver o relat´rio através do browser
Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Variável das msgboxs
Public Resposta As String

'Colocar o formulário por cima dos outros
Sub AlwaysOnTop(FrmID As Form, OnTop As Integer)
Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1
Const flags = SWP_NOMOVE Or SWP_NOSIZE
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
    If OnTop = -1 Then
        OnTop = SetWindowPos(FrmID.hWnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
    Else
        OnTop = SetWindowPos(FrmID.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, flags)
    End If
End Sub

Public Sub Mensagem_de_Aviso(Aviso As String, Mensagem As String)
    'Procedimento para mostrar uma mensagem de aviso
    With Form_Mensagem
        If Aviso = "Informação" Then
            .Pic_Mensagem.Picture = Form_Skin.Icon_Info.Picture
            .Botao_Ok.Visible = True
        ElseIf Aviso = "Erro" Then
            .Pic_Mensagem.Picture = Form_Skin.Icon_Error.Picture
            .Botao_Ok.Visible = True
        ElseIf Aviso = "Questão" Then
            .Pic_Mensagem.Picture = Form_Skin.Icon_Error.Picture
            .Botao_Sim.Visible = True
            .Botao_Nao.Visible = True
        End If
        
        .Label_Mensagem.Caption = Mensagem
        .Show vbModal
    End With
End Sub

Public Function ArquivoExiste(ByVal Caminho As String, Optional ByVal SomenteDiretorio As Boolean = False) As Boolean
    'Função para verificar se a pasta existe
    On Error Resume Next
    If SomenteDiretorio Then
        ArquivoExiste = GetAttr(Mid(Caminho, 1, InStrRev(Caminho, ""))) And vbDirectory
    Else
        ArquivoExiste = GetAttr(Caminho)
    End If
    On Error GoTo 0
End Function

