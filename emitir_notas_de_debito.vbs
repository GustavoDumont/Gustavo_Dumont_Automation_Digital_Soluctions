' Versao publica e anonimizada para portfolio. Valores de negocio foram substituidos.
'=================================================
' 1. DECLARAÇÕES (TODAS AS VARIÁVEIS AQUI)
'=================================================
Option Explicit
Dim SapGuiAuto, application, connection, session
Dim objExcel, objSheet, linhaExcel
Dim ContaSAP, DataFaturamento, DataVencimento, Valor, PEP, Texto
Dim mensagemSAP, codigoGerado, tentativas, i, ch
Dim JanelaPopUp, Elemento, Partes, NumLinha, UltimaLinhaPopUp, AlvoIndex
Dim shell, windowReady, caminhoArquivo, tentativa, nomeSimples

'=================================================
' 2. CONEXÃO COM SAP
'=================================================
On Error Resume Next
Set SapGuiAuto = GetObject("SAPGUI")
If Err.Number <> 0 Then
    MsgBox "SAP GUI não está aberto.", vbCritical
    WScript.Quit
End If
On Error GoTo 0

Set application = SapGuiAuto.GetScriptingEngine
If application.Children.Count = 0 Then
    MsgBox "Nenhuma conexão SAP encontrada.", vbCritical
    WScript.Quit
End If

Set connection = application.Children(0)
Set session = connection.Children(0)

'=================================================
' 3. CONEXÃO COM EXCEL E FUNÇÃO DE ESPERA
'=================================================
Set objExcel = GetObject(, "Excel.Application")
Set objSheet = objExcel.ActiveWorkbook.ActiveSheet

Sub EsperarSAP()
    Do While session.Busy
        WScript.Sleep 300
    Loop
    WScript.Sleep 500
End Sub

'=================================================
' 4. LOOP PRINCIPAL
'=================================================
For linhaExcel = 2 To objSheet.UsedRange.Rows.Count

    ContaSAP        = Trim(CStr(objSheet.Cells(linhaExcel, 1).Value))
    DataVencimento  = Trim(CStr(objSheet.Cells(linhaExcel, 3).Value))
    Valor           = Trim(CStr(objSheet.Cells(linhaExcel, 4).Value))
    PEP             = Trim(CStr(objSheet.Cells(linhaExcel, 5).Value))
    Texto           = Trim(CStr(objSheet.Cells(linhaExcel, 6).Value))

    If ContaSAP = "" Then Exit For

    EsperarSAP
    session.findById("wnd[0]").maximize

    ' Fechar popup residual antes de começar
    If session.Children.Count > 1 Then
        On Error Resume Next
        session.findById("wnd[1]").close
        On Error GoTo 0
        EsperarSAP
    End If

    '=================================================
    ' VA01 - TELA INICIAL
    '=================================================
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nVA01"
    session.findById("wnd[0]").sendVKey 0
    EsperarSAP

    session.findById("wnd[0]/usr/ctxtVBAK-AUART").Text = "TIPO_DEMO"
    session.findById("wnd[0]/usr/ctxtVBAK-VKORG").Text = "ORG_DEMO"
    session.findById("wnd[0]/usr/ctxtVBAK-VTWEG").Text = "CANAL_DEMO"
    session.findById("wnd[0]/usr/ctxtVBAK-SPART").Text = "SETOR_DEMO"
    session.findById("wnd[0]").sendVKey 0
    EsperarSAP

    ' Inserir Cliente
    session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR").Text = ContaSAP
    session.findById("wnd[0]").sendVKey 0
    
    '=================================================
    ' 5. MÉTODO: SELEÇÃO DA PENÚLTIMA ÁREA (COM SCROLL)
    '=================================================
    WScript.Sleep 1500 ' Tempo vital para o pop-up carregar
    
    If session.Children.Count > 1 Then
        Set JanelaPopUp = session.findById("wnd[1]/usr")
        
        ' --- PASSO 1: ROLAR ATÉ O FINAL ---
        ' Verifica se existe barra de rolagem e se ela é maior que 0
        If JanelaPopUp.verticalScrollbar.maximum > 0 Then
            ' Joga a barra de rolagem para a última posição possível
            JanelaPopUp.verticalScrollbar.position = JanelaPopUp.verticalScrollbar.maximum
            ' Espera o SAP carregar os dados do fundo da lista
            WScript.Sleep 800 
        End If

        ' --- PASSO 2: ENCONTRAR O ÍNDICE VISUAL MAIS ALTO ---
        UltimaLinhaPopUp = -1
        
        ' Varre apenas os elementos visíveis agora (que estamos no fundo)
        For Each Elemento In JanelaPopUp.Children
            ' Procura pelos labels da coluna 15 (onde fica o texto da área de vendas)
            If InStr(Elemento.Id, "lbl[15,") > 0 Then
                Partes = Split(Elemento.Id, ",")
                ' O ID vem no formato lbl[15,9]. Pegamos o 9.
                NumLinha = CInt(Replace(Partes(1), "]", ""))
                
                If NumLinha > UltimaLinhaPopUp Then UltimaLinhaPopUp = NumLinha
            End If
        Next

        ' --- PASSO 3: SELECIONAR A PENÚLTIMA ---
        ' Se a última linha visual é 9, a penúltima é 8.
        AlvoIndex = UltimaLinhaPopUp - 1
        
        ' Proteção: Se houver apenas 1 item (AlvoIndex seria -1), seleciona o único item (0)
        If AlvoIndex < 0 Then AlvoIndex = UltimaLinhaPopUp

        If AlvoIndex >= 0 Then
            On Error Resume Next
            session.findById("wnd[1]/usr/lbl[15," & AlvoIndex & "]").setFocus
            ' Se der erro no setFocus, tenta clicar direto
            If Err.Number <> 0 Then Err.Clear
            
            session.findById("wnd[1]").sendVKey 2 ' Tecla F2 (Selecionar)
            On Error GoTo 0
        Else
            ' Fallback: Se não achou índices, apenas dá Enter
            session.findById("wnd[1]").sendVKey 0 
        End If
        
        ' Limpa janelas residuais (confirmações extras)
        WScript.Sleep 500
        Do While session.Children.Count > 1
            session.ActiveWindow.sendVKey 0
            WScript.Sleep 500
        Loop
    End If
    
    '=================================================
    ' 6. ITEM OVERVIEW (FORMATO ORIGINAL ADAPTADO)
    '=================================================
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02").select
    EsperarSAP

    ' Inserção do Material
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/" & _
        "ssubSUBSCREEN_BODY:SAPMV45A:4415/" & _
        "subSUBSCREEN_TC:SAPMV45A:4902/" & _
        "tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").Text = "MATERIAL_DEMO"

    ' Inserção da Quantidade
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/" & _
        "ssubSUBSCREEN_BODY:SAPMV45A:4415/" & _
        "subSUBSCREEN_TC:SAPMV45A:4902/" & _
        "tblSAPMV45ATCTRL_U_ERF_GUTLAST/txtVBAP-ZMENG[2,0]").Text = "1"

    ' Inserção do Centro
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/" & _
        "ssubSUBSCREEN_BODY:SAPMV45A:4415/" & _
        "subSUBSCREEN_TC:SAPMV45A:4902/" & _
        "tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtVBAP-WERKS[44,0]").Text = "CENTRO_DEMO"

    session.findById("wnd[0]").sendVKey 0
    EsperarSAP

    '=================================================
    ' DETALHE DO ITEM
    '=================================================
    ' ABRIR DADOS DO ITEM (DUPLO CLIQUE NO MATERIAL)
    '=================================================
    EsperarSAP
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02").select
    EsperarSAP

    session.findById( _
    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/" & _
    "ssubSUBSCREEN_BODY:SAPMV45A:4415/" & _
    "subSUBSCREEN_TC:SAPMV45A:4902/" & _
    "tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]" _
    ).setFocus

    session.findById( _
    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/" & _
    "ssubSUBSCREEN_BODY:SAPMV45A:4415/" & _
    "subSUBSCREEN_TC:SAPMV45A:4902/" & _
    "tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]" _
    ).caretPosition = 4

    session.findById("wnd[0]").sendVKey 2   ' duplo clique real no item
    EsperarSAP

    '=================================================
    ' DOCUMENTOS DE FATURAMENTO
    '=================================================
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\04").select
    EsperarSAP
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\04/ssubSUBSCREEN_BODY:SAPMV45A:4453/ctxtVBKD-VALDT").Text = DataVencimento
    EsperarSAP

    '========================
    ' CONDIÇÕES – SAPLV69A (FORMA COMPATÍVEL)
    '========================
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06").select
    EsperarSAP

    ' Rolar até a última linha (como no script gravado)
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN").verticalScrollbar.position = 139
    EsperarSAP

    ' Tipo de condição
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/ctxtKOMV-KSCHL[1,0]").Text = "CONDICAO_DEMO"
    EsperarSAP

    ' Montante
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]").Text = Valor
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,0]").setFocus
    session.findById("wnd[0]").sendVKey 0
    EsperarSAP

    '=================================================
    ' CLASSIFICAÇÃO CONTÁBIL
    '=================================================
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\07").select
    EsperarSAP
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4457/subCOBL:SAPLKACB:1006/ctxtCOBL-PS_POSID").Text = PEP
    session.findById("wnd[0]").sendVKey 0
    EsperarSAP

    '=================================================
    ' TEXTOS
    '=================================================
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\09").select
    EsperarSAP
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").Text = Texto
    session.findById("wnd[0]").sendVKey 0
    EsperarSAP

    '=================================================
    ' SALVAR 
    '=================================================
    EsperarSAP
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/btn[11]").press
    EsperarSAP


'=================================================
    ' --- VF01 (FATURAMENTO) ---
    '=================================================
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nVF01"
    session.findById("wnd[0]").sendVKey 0 
    EsperarSAP

    ' Garante que o foco está no número da ordem e processa
    session.findById("wnd[0]/usr/tblSAPMV60ATCTRL_ERF_FAKT/ctxtKOMFK-VBELN[0,0]").setFocus
    session.findById("wnd[0]").sendVKey 0
    EsperarSAP

    ' VERIFICAÇÃO DE SEGURANÇA (Se a barra de status indicar erro "E", pula a linha)
    If session.findById("wnd[0]/sbar").MessageType = "E" Then
        ' Opcional: Registar no Excel que esta linha deu erro
    Else
        ' --- ETAPA 2: CONFIGURAR MENSAGENS ---
        session.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[3]").Select
        EsperarSAP

        ' Botão de novos detalhes / criar
        session.findById("wnd[0]/tbar[1]/btn[5]").press 
        EsperarSAP

        ' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        ' ESPAÇO RESERVADO: INSIRA AQUI A GRAVAÇÃO DO CÓDIGO 9000 E RE
        ' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        
        ' Define Enviar Imediatamente (4)
        session.findById("wnd[0]/usr/cmbNAST-VSZTP").key = "4"
        session.findById("wnd[0]").sendVKey 0 ' Enter para confirmar
        
        ' Voltar para a tela principal da VF01
        session.findById("wnd[0]/tbar[0]/btn[3]").press 
        ' Se aparecer popup de confirmação ao voltar, dá um Enter extra
        If session.Children.Count > 1 Then session.ActiveWindow.sendVKey 0
        EsperarSAP

        ' --- ETAPA 3: GRAVAÇÃO AUTOMÁTICA E PDF ---
        
        ' 1. Clica no botão Gravar no SAP
        session.findById("wnd[0]/tbar[0]/btn[11]").press 
        
        ' Inicializa o objeto shell e limpa verificadores
        Set shell = CreateObject("WScript.Shell")
        windowReady = False
        tentativa = 0
        
        ' --- LOOP DE ESPERA PELA JANELA DO WINDOWS ---
        ' Às vezes a janela demora. Vamos esperar até 15 segundos antes de desistir.
        Do While windowReady = False And tentativa < 15
            ' Tenta ativar a janela de Impressão ou a de Salvar Como
            If shell.AppActivate("Imprimir") Then
                WScript.Sleep 500
                shell.SendKeys "{ENTER}"
                windowReady = True 
            ElseIf shell.AppActivate("Salvar Saída de Impressão como") Then
                windowReady = True
            End If
            
            WScript.Sleep 1000
            tentativa = tentativa + 1
        Loop

        ' --- AÇÃO DE DIGITAÇÃO ---
        If windowReady Then
            ' Dá um tempo extra para a janela "estabilizar"
            WScript.Sleep 1500 
            
            ' Traz a janela para a frente com foco total
            shell.AppActivate "Salvar Saída de Impressão como"
            WScript.Sleep 500
            
            ' Define o nome: ND 1, ND 2...
            nomeSimples = "ND " & (linhaExcel - 1)
            
            ' Garante que o campo de nome está limpo (Seleciona tudo e apaga)
            shell.SendKeys "^a" 
            WScript.Sleep 200
            shell.SendKeys "{BACKSPACE}"
            WScript.Sleep 300
            
            ' Digita o nome do arquivo
            shell.SendKeys nomeSimples
            WScript.Sleep 500
            
            ' Pressiona ENTER para salvar
            shell.SendKeys "{ENTER}"
            
            ' Aguarda o Windows fechar a janela antes de ir para o próximo loop
            WScript.Sleep 2000 
        Else
            ' Caso a janela não abra de jeito nenhum, avisa qual linha falhou
            MsgBox "A janela de impressão não abriu para a linha " & linhaExcel, vbExclamation
        End If
    End If

Next

MsgBox "Automacao finalizada!", 64

