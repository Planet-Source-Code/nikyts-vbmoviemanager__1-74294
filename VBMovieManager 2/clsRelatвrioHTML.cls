VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRelat�rioHTML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'#################################################################################################
'# Nome: !!!!*** clsRelat�rioHTML ***!!!!                                                        #
'# Projeto: Classe para Gera��o de Relat�rios HTML                                               #
'#                                                                                               #
'# Propriedades                                                                                  #
'#   Arquivo                     - string        - Nome do Arquivo a ser Gerado                  #
'#   Data                        - string        - Data de Emiss�o                               #
'#   HoraIn�cio                  - string        - Hor�rio de In�cio da Gera��o                  #
'#   HoraFim                     - string        - Hor�rio de T�rmino da Gera��o                 #
'#   NomeRelat�rio               - string        - Nome do Relat�rio                             #
'#   Empresa                     - string        - Nome da Empresa                               #
'#   LinkEmpresa                 - string        - Email ou Site da Empresa                      #
'#   Desenvolvedor               - string        - Nome do Desenvolvedor                         #
'#   LinkDesenvolvedor           - string        - Link do Desenvolvedor                         #
'#   T�tulosColunas              - array string  - T�tulos das Colunas de Detalhe                #
'#   AlinhamentoT�tulos          - array integer - Alinhamento dos T�tulos                       #
'#   CamposDetalhe               - array string  - Campos de Detalhes das Colunas                #
'#   AlinhamentoDetalhes         - array integer - Alinhamento dos Campos de Detalhes            #
'#   NegritoDetalhes             - array boolean - Utilizar Negrito nos Detalhes                 #
'#   It�licoDetalhes             - array boolean - Utilizar It�lico nos Detalhes                 #
'#   TamanhoColunas              - array integer - Largura das Colunas                           #
'#   Filtros                     - array         - Filtros Utilizados                            #
'#   UsaFonteCourierNoCorpo      - boolean       - Usar Fonte Courier nos Detalhes               #
'#   MostraMensagemDeFinaliza��o - boolean       - Mostrar Mensagem de T�rmino da Gera��o        #
'#                                                                                               #
'# Sub's                                                                                         #
'#   ImprimeCabe�alho            - Imprime o Cabe�alho do Relat�rio - Chamar apenas uma vez      #
'#   ImprimeDetalhe              - Imprime as Colunas de Detalhe do Relat�rio - Chamar "n" vezes #
'#   ImprimeRodap�               - Imprime o Rodap� do Relat�rio - Chamar apenas uma vez         #
'#                                                                                               #
'# Vari�veis de Alinhamento                                                                      #
'#   N�oAlinhar                  - N�o faz Alinamento                                            #
'#   AlinharEsquerda             - For�a Alinhamento pela Esquerda                               #
'#   AlinharCentro               - For�a Alinhamento pelo Centro                                 #
'#   AlinharDireita              - For�a Alinhamento pela Direita                                #
'#                                                                                               #
'# By: Jean Jedson Adriano                                                                       #
'#     jean@mormaii.com.br                                                                       #
'#################################################################################################

Private sFileHTML As String
Private bTemFiltros As Boolean

Private sArquivo As String
Private sData As String
Private sHoraIn�cio As String
Private sHoraFim As String
Private sNomeRelat�rio As String
Private sEmpresa As String
Private sLinkEmpresa As String
Private sDesenvolvedor As String
Private sLinkDesenvolvedor As String
Private aT�tulosColunas() As String
Private aAlinhamentoT�tulos() As Integer
Private aCamposDetalhe() As String
Private aAlinhamentoDetalhes() As Integer
Private aNegritoDetalhes() As Boolean
Private aIt�licoDetalhes() As Boolean
Private aTamanhoColunas() As Integer
Private aFiltros() As String
Private bUsaFonteCourierNoCorpo As Boolean
Private bMostraMensagemDeFinaliza��o As Boolean

Public N�oAlinhar As Integer
Public AlinharEsquerda As Integer
Public AlinharCentro As Integer
Public AlinharDireita As Integer

Private Sub Class_Initialize()
    sFileHTML = FreeFile
    bTemFiltros = False
    N�oAlinhar = 0
    AlinharEsquerda = 1
    AlinharCentro = 2
    AlinharDireita = 3
    bMostraMensagemDeFinaliza��o = False
End Sub

Public Property Let Arquivo(pArquivo As String)
    sArquivo = pArquivo
    On Error Resume Next
    If Dir(sArquivo) <> "" Then
        Kill sArquivo
    End If
End Property

Public Property Let Data(pData As String)
    sData = pData
End Property

Public Property Let HoraIn�cio(pHoraIn�cio As String)
    sHoraIn�cio = pHoraIn�cio
End Property

Public Property Let HoraFim(pHoraFim As String)
    sHoraFim = pHoraFim
End Property

Public Property Let NomeRelat�rio(pNomeRelat�rio As String)
    sNomeRelat�rio = pNomeRelat�rio
End Property

Public Property Let Empresa(pEmpresa As String)
    sEmpresa = pEmpresa
End Property

Public Property Let LinkEmpresa(pLinkEmpresa As String)
    sLinkEmpresa = pLinkEmpresa
End Property

Public Property Let Desenvolvedor(pDesenvolvedor As String)
    sDesenvolvedor = pDesenvolvedor
End Property

Public Property Let LinkDesenvolvedor(pLinkDesenvolvedor As String)
    sLinkDesenvolvedor = pLinkDesenvolvedor
End Property

Public Property Let T�tulosColunas(pT�tulosColunas As Variant)
    aT�tulosColunas = pT�tulosColunas
End Property

Public Property Let AlinhamentoT�tulos(pAlinhamentoT�tulos As Variant)
    aAlinhamentoT�tulos = pAlinhamentoT�tulos
End Property

Public Property Let CamposDetalhe(pCamposDetalhe As Variant)
    aCamposDetalhe = pCamposDetalhe
End Property

Public Property Let AlinhamentoDetalhes(pAlinhamentoDetalhes As Variant)
    aAlinhamentoDetalhes = pAlinhamentoDetalhes
End Property

Public Property Let NegritoDetalhes(pNegritoDetalhes As Variant)
    aNegritoDetalhes = pNegritoDetalhes
End Property

Public Property Let It�licoDetalhes(pIt�licoDetalhes As Variant)
    aIt�licoDetalhes = pIt�licoDetalhes
End Property

Public Property Let TamanhoColunas(pTamanhoColunas As Variant)
    aTamanhoColunas = pTamanhoColunas
End Property

Public Property Let Filtros(pFiltros As Variant)
    aFiltros = pFiltros
    bTemFiltros = True
End Property

Public Property Let UsaFonteCourierNoCorpo(pUsaFonteCourierNoCorpo As Boolean)
    bUsaFonteCourierNoCorpo = pUsaFonteCourierNoCorpo
End Property

Public Property Let MostraMensagemDeFinaliza��o(pMostraMensagemDeFinaliza��o As Boolean)
    bMostraMensagemDeFinaliza��o = pMostraMensagemDeFinaliza��o
End Property

Public Sub ImprimeCabe�alho()
    Open sArquivo For Output As #sFileHTML
    Print #sFileHTML, "<html>" & vbCrLf
    Print #sFileHTML, " <head>" & vbCrLf
    Print #sFileHTML, "  <title>" & sNomeRelat�rio & "</title>" & vbCrLf
    Print #sFileHTML, "  <style>body, td, table {font-family:Tahoma,Verdana,Arial; font-size=10px;}</style>" & vbCrLf
    Print #sFileHTML, " </head>" & vbCrLf
    Print #sFileHTML, " <body leftmargin=15>" & vbCrLf
    Print #sFileHTML, "  <table border=0 width=100% cellpadding=0 cellspacing=0>" & vbCrLf
    Print #sFileHTML, "   <tr>" & vbCrLf
    Print #sFileHTML, "    <td><b><i><font size=5>" & sNomeRelat�rio & "</font></i></b></td>" & vbCrLf
    Print #sFileHTML, "    <td align=right><b><font size=3>" & sEmpresa & "</font></b></td>" & vbCrLf
    Print #sFileHTML, "   </tr>" & vbCrLf
    Print #sFileHTML, "   <tr>" & vbCrLf
    Print #sFileHTML, "    <td><font size=1>&nbsp;</font></td>" & vbCrLf
    Print #sFileHTML, "    <td align=right><font size=1><a href=" & IIf(UCase(Left(sLinkEmpresa, 3)) = "WWW", "http://", "mailto:") & sLinkEmpresa & ">" & sLinkEmpresa & "</a></font></td>" & vbCrLf
    Print #sFileHTML, "   </tr>" & vbCrLf
    ImprimeFiltros
    Print #sFileHTML, "   <tr>" & vbCrLf
    Print #sFileHTML, "    <td colspan=2><hr></td>" & vbCrLf
    Print #sFileHTML, "   </tr>" & vbCrLf
    Print #sFileHTML, "  </table>" & vbCrLf
    Print #sFileHTML, "  <table border=0 width=100% cellpadding=0 cellspacing=0>" & vbCrLf
    Print #sFileHTML, "   <tr>" & vbCrLf
    Print #sFileHTML, "    <td>&nbsp;</td>" & vbCrLf
    Print #sFileHTML, "   </tr>" & vbCrLf
    Print #sFileHTML, "  </table>" & vbCrLf
    IniciaDetalhe
End Sub

Private Sub ImprimeFiltros()
    Dim iCount As Integer
    If bTemFiltros Then
        Print #sFileHTML, "   <tr>" & vbCrLf
        For iCount = 1 To UBound(aFiltros)
            Print #sFileHTML, "    <td colspan=2>" & aFiltros(iCount) & "</td>" & vbCrLf
        Next iCount
        Print #sFileHTML, "   </tr>" & vbCrLf
    End If
End Sub

Private Sub IniciaDetalhe()
    Dim iCount As Integer
    Dim sAlinhamento As String
    Print #sFileHTML, "  <table width=100% border=0 cellpadding=0 cellspacing=0>" & vbCrLf
    Print #sFileHTML, "   <tr>" & vbCrLf
    For iCount = 1 To UBound(aT�tulosColunas)
        Select Case aAlinhamentoT�tulos(iCount)
            Case N�oAlinhar
                sAlinhamento = ""
            Case AlinharEsquerda
                sAlinhamento = "align=left"
            Case AlinharCentro
                sAlinhamento = "align=center"
            Case AlinharDireita
                sAlinhamento = "align=right"
        End Select
        Print #sFileHTML, "    <td width=" & aTamanhoColunas(iCount) & "%" & sAlinhamento & "><font size=1<font size=1" & IIf(bUsaFonteCourierNoCorpo, " face=Courier", "") & "><b>" & "&nbsp;&nbsp;" & aT�tulosColunas(iCount) & "</b></font></td>" & vbCrLf
    Next iCount
    Print #sFileHTML, "   </tr>" & vbCrLf
End Sub

Public Sub ImprimeDetalhe()
    Print #sFileHTML, "   <tr>" & vbCrLf
    For iCount = 1 To UBound(aCamposDetalhe)
        Select Case aAlinhamentoDetalhes(iCount)
            Case N�oAlinhar
                sAlinhamento = ""
            Case AlinharEsquerda
                sAlinhamento = "align=left"
            Case AlinharCentro
                sAlinhamento = "align=center"
            Case AlinharDireita
                sAlinhamento = "align=right"
        End Select
        Print #sFileHTML, "    <td width=" & aTamanhoColunas(iCount) & "% " & sAlinhamento & "><font size=1" & IIf(bUsaFonteCourierNoCorpo, " face=Courier", "") & ">" & IIf(aNegritoDetalhes(iCount), "<b>", "") & IIf(aIt�licoDetalhes(iCount), "<i>", "") & "&nbsp;&nbsp;" & aCamposDetalhe(iCount) & IIf(aIt�licoDetalhes(iCount), "</i>", "") & IIf(aNegritoDetalhes(iCount), "</b>", "") & "</font></td>" & vbCrLf
    Next iCount
    Print #sFileHTML, "   </tr>" & vbCrLf
End Sub

Private Sub FinalizaDetalhe()
    Print #sFileHTML, "  </table>" & vbCrLf
End Sub

Public Sub ImprimeRodap�()
    Dim iCount As Integer
    FinalizaDetalhe
    Print #sFileHTML, "  <table border=0 width=100% cellpadding=0 cellspacing=0>" & vbCrLf
    Print #sFileHTML, "   <tr>" & vbCrLf
    Print #sFileHTML, "    <td>&nbsp;</td>" & vbCrLf
    Print #sFileHTML, "   </tr>" & vbCrLf
    Print #sFileHTML, "  </table>" & vbCrLf
    Print #sFileHTML, "  <table border=0 width=100% cellpout.printing=0 cellspacing=0>" & vbCrLf
    Print #sFileHTML, "   <tr>" & vbCrLf
    Print #sFileHTML, "    <td colspan=2><hr></td>" & vbCrLf
    Print #sFileHTML, "   </tr>" & vbCrLf
    Print #sFileHTML, "   <tr>" & vbCrLf
    Print #sFileHTML, "    <td width=50%><font size=1>" & Format(sData, "dd/mm/yyyy") & "   -   Inicio: " & Format(sHoraIn�cio, "hh:mm:ss") & "   -   Fim: " & Format(sHoraFim, "hh:mm:ss") & "</font></td>" & vbCrLf
    Print #sFileHTML, "    <td width=50% align=right><font size=1><a href=" & IIf(UCase(Left(sLinkDesenvolvedor, 3)) = "WWW", "http://", "mailto:") & sLinkDesenvolvedor & ">Desenvolvido por " & sDesenvolvedor & "</a></font></td>" & vbCrLf
    Print #sFileHTML, "   </tr>" & vbCrLf
    Print #sFileHTML, "  </table>" & vbCrLf
    Print #sFileHTML, " </body>" & vbCrLf
    Print #sFileHTML, "</html>" & vbCrLf
    Close #sFileHTML
    If bMostraMensagemDeFinaliza��o Then
        With Form_Mensagem
            .Pic_Mensagem.Picture = Form_Skin.Icon_Info.Picture
            .Label_Mensagem.Caption = "O relat�rio foi processado com sucesso."
            .Botao_Ok.Visible = True
            .Show
        End With
        Exit Sub
    End If
End Sub

Private Sub Class_Terminate()
    sFileHTML = ""
End Sub
