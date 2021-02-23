Attribute VB_Name = "NF3eAPI"
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'activate Microsoft XML, v6.0 in references

'Atributo privado da classe
Private Const tempoResposta = 500
'Private Const impressaoParam = """impressao"":{" & """tipo"":""pdf""," & """ecologica"":false," & """itemLinhas"":""1""," & """itemDesconto"":false," & """larguraPapel"":""80mm""}"
Private Const token = "SEU_TOKEN_AQUI"

'Esta funcao envia um conteudo para uma URL, em requisicoes do tipo POST
Function enviaConteudoParaAPI(conteudo As String, url As String, tpConteudo As String) As String
On Error GoTo SAI
    Dim contentType As String
    
    If (tpConteudo = "txt") Then
        contentType = "text/plain;charset=utf-8"
    ElseIf (tpConteudo = "xml") Then
        contentType = "application/xml;charset=utf-8"
    Else
        contentType = "application/json;charset=utf-8"
    End If
    
    Dim obj As MSXML2.ServerXMLHTTP60
    Set obj = New MSXML2.ServerXMLHTTP60
    obj.Open "POST", url
    obj.setRequestHeader "Content-Type", contentType
    If Trim(token) <> "" Then
        obj.setRequestHeader "X-AUTH-TOKEN", token
    End If
    obj.send conteudo
    Dim resposta As String
    resposta = obj.responseText
    
    Select Case obj.status
        Case 401
            MsgBox ("Token nao enviado ou invalido")
        Case 403
            MsgBox ("Token sem permissao")
    End Select
    
    enviaConteudoParaAPI = resposta
    Exit Function
SAI:
  enviaConteudoParaAPI = "{" & """status"":""" & Err.Number & """," & """motivo"":""" & Err.Description & """" & "}"
End Function

'Esta funcao realiza o processo completo de emissao: envio e download do documento
Public Function emitirN3eSincrono(conteudo As String, tpConteudo As String, cnpj As String, tpDown As String, tpAmb As String, caminho As String, exibeNaTela As Boolean) As String
    Dim retorno As String
    Dim resposta As String
    Dim statusEnvio As String
    Dim statusConsulta As String
    Dim statusDownload As String
    Dim motivo As String
    Dim erros As String
    Dim nsNRec As String
    Dim chNFe As String
    Dim cStat As String
    Dim nProt As String

    statusEnvio = ""
    statusConsulta = ""
    statusDownload = ""
    motivo = ""
    erros = ""
    nsNRec = ""
    chNFe = ""
    cStat = ""
    nProt = ""
    
    gravaLinhaLog ("[EMISSAO_SINCRONA_INICIO]")
    
    resposta = emitirNF3e(conteudo, tpConteudo)
    statusEnvio = LerDadosJSON(resposta, "status", "", "")

    If (statusEnvio = "200") Or (statusEnvio = "-6") Then
    
        nsNRec = LerDadosJSON(resposta, "nsNRec", "", "")

        Sleep (tempoResposta)

        resposta = consultarStatusProcessamento(cnpj, nsNRec, tpAmb)
        statusConsulta = LerDadosJSON(resposta, "status", "", "")

        If (statusConsulta = "200") Then
            
            cStat = LerDadosJSON(resposta, "cStat", "", "")

            If (cStat = "100") Or (cStat = "150") Then
            
                chNFe = LerDadosJSON(resposta, "chNFe", "", "")
                nProt = LerDadosJSON(resposta, "nProt", "", "")
                motivo = LerDadosJSON(resposta, "xMotivo", "", "")

                resposta = downloadNF3eESalvar(chNFe, tpAmb, tpDown, caminho, exibeNaTela)
                statusDownload = LerDadosJSON(resposta, "status", "", "")
                
                If (statusDownload <> "200") Then
                
                    motivo = LerDadosJSON(resposta, "motivo", "", "")
                    
                End If
            Else
            
                motivo = LerDadosJSON(resposta, "xMotivo", "", "")
                
            End If
        ElseIf (statusConsulta = "-2") Then
                
            erros = Split(resposta, """erro"":""")
            erros = LerDadosJSON(resposta, "erro", "", "")
            motivo = LerDadosJSON(erros, "xMotivo", "", "")
            cStat = LerDadosJSON(erros, "cStat", "", "")
                        
        Else
        
            motivo = LerDadosJSON(resposta, "motivo", "", "")
            
        End If
        
    ElseIf (status = "-7") Then
    
        motivo = LerDadosJSON(resposta, "motivo", "", "")
        nsNRec = LerDadosJSON(resposta, "nsNRec", "", "")
   
    ElseIf (statusEnvio = "-4") Or (statusEnvio = "-2") Then

        motivo = LerDadosJSON(resposta, "motivo", "", "")
        erros = LerDadosJSON(resposta, "erros", "", "")

    ElseIf (statusEnvio = "-999") Or (statusEnvio = "-5") Then
    
        erros = Split(resposta, """erro"":""")
        erros = LerDadosJSON(resposta, "erro", "", "")
        erros = LerDadosJSON(erros, "xMotivo", "", "")
        
    Else
    
        motivo = LerDadosJSON(resposta, "motivo", "", "")
        
    End If
    
    'Monta o JSON de retorno
    retorno = "{"
    retorno = retorno & """statusEnvio"":""" & statusEnvio & ""","
    retorno = retorno & """statusConsulta"":""" & statusConsulta & ""","
    retorno = retorno & """statusDownload"":""" & statusDownload & ""","
    retorno = retorno & """cStat"":""" & cStat & ""","
    retorno = retorno & """chNFe"":""" & chNFe & ""","
    retorno = retorno & """nProt"":""" & nProt & ""","
    retorno = retorno & """motivo"":""" & motivo & ""","
    retorno = retorno & """nsNRec"":""" & nsNRec & ""","
    retorno = retorno & """erros"":""" & erros & """"
    retorno = retorno & "}"
    
    gravaLinhaLog ("[JSON_RETORNO]")
    gravaLinhaLog (retorno)
    gravaLinhaLog ("[EMISSAO_SINCRONA_FIM]")

    emitirNFeSincrono = retorno
End Function

'Esta funcao realiza o envio de uma NF3e
Public Function emitirNF3e(conteudo As String, tpConteudo As String) As String

    Dim url As String
    Dim resposta As String

    url = "https://nf3e.ns.eti.br/v1/nf3e/issue"

    gravaLinhaLog ("[ENVIO_DADOS]")
    gravaLinhaLog (conteudo)
        
    resposta = enviaConteudoParaAPI(conteudo, url, tpConteudo)
    
    gravaLinhaLog ("[ENVIO_RESPOSTA]")
    gravaLinhaLog (resposta)

    emitirNF3e = resposta
End Function

Public Function consultarStatusProcessamento(cnpj As String, nsNRec As String, tpAmb As String) As String
    Dim json As String
    Dim url As String
    Dim resposta As String

    'Monta o JSON
    json = "{"
    json = json & """CNPJ"":""" & cnpj & ""","
    json = json & """nsNRec"":""" & nsNRec & ""","
    json = json & """tpAmb"":""" & tpAmb & """"
    json = json & "}"

    url = "https://nf3e.ns.eti.br/v1/nf3e/issue/status"
    
    gravaLinhaLog ("[CONSULTA_DADOS]")
    gravaLinhaLog (json)
        
    resposta = enviaConteudoParaAPI(json, url, "json")
    
    gravaLinhaLog ("[CONSULTA_RESPOSTA]")
    gravaLinhaLog (resposta)

    consultarStatusProcessamento = resposta
End Function


'Esta funcao realiza o download de documentos de uma NF3-e
Public Function downloadNF3e(chNFe As String, tpAmb As String) As String

    Dim json As String
    Dim url As String
    Dim resposta As String
    Dim status As String

    'Monta o JSON
    json = "{"
    json = json & """chNFe"":""" & chNFe & ""","
    json = json & """tpAmb"":""" & tpAmb & ""","
    'json = json & impressaoParam
    json = json & "}"

    url = "https://nf3e.ns.eti.br/v1/nf3e/get"

    gravaLinhaLog ("[DOWNLOAD_NF3E_DADOS]")
    gravaLinhaLog (json)
        
    resposta = enviaConteudoParaAPI(json, url, "json")
    
    status = LerDadosJSON(resposta, "status", "", "")
        
    'O retorno da API sera gravado somente em caso de erro,
    'para nao gerar um log extenso com o PDF e XML
    If (status <> "100") Then
    
        gravaLinhaLog ("[DOWNLOAD_NF3E_RESPOSTA]")
        gravaLinhaLog (resposta)
        
    Else

        gravaLinhaLog ("[DOWNLOAD_NF3E_STATUS]")
        gravaLinhaLog (status)
        
    End If

    downloadNF3e = resposta
End Function

'Esta funcao realiza o download de documentos de uma NFC-e e salva-os
Public Function downloadNF3eESalvar(chNFe As String, tpAmb As String, caminho As String, exibeNaTela As Boolean, imprimePDF As Boolean) As String

    Dim xml As String
    Dim pdf As String
    Dim status As String
    Dim resposta As String
    
    resposta = downloadNF3e(chNFe, tpAmb)
    status = LerDadosJSON(resposta, "status", "", "")

    If status = "100" Then
        
        'Cria o diretorio, caso nao exista
        If Dir(caminho, vbDirectory) = "" Then
            MkDir (caminho)
        End If
    
        xml = LerDadosJSON(resposta, "nfeProc", "xml", "")
        Call salvarXML(xml, caminho, chNFe)
        
        If InStr(1, impressaoParam, "pdf") Then
        
            pdf = LerDadosJSON(resposta, "pdf", "", "")
            Call salvarPDF(pdf, caminho, chNFe)
            
            If exibeNaTela Then
            
                ShellExecute 0, "open", caminho & chNFe & "-procNFe.pdf", "", "", vbNormalFocus
            
            End If

            If imprimePDF Then

                Call imprimirNF3e(caminho, chNFe)
            
            End If

        End If
    Else
        MsgBox ("Ocorreu um erro, veja o Retorno da API para mais informacoes")
        gravaLinhaLog ("[Ocorreu um erro, veja o Retorno da API para mais informacoes  - Metodo: downloadNF3eESalvar]")
    End If

    downloadNF3eESalvar = resposta
End Function


'Esta funcao realiza o download de eventos de uma NF3e e salva-os
Public Function downloadEventoNF3eESalvar(chNFe As String, tpAmb As String, caminho As String, exibeNaTela As Boolean) As String
    Dim xml As String
    Dim chNFeCanc As String
    Dim pdf As String
    Dim status As String
    Dim resposta As String
    
    resposta = downloadNF3e(chNFe, tpAmb)
    status = LerDadosJSON(resposta, "status", "", "")

    If status = "100" Then
    
        'Cria o diretorio, caso nao exista
        If Dir(caminho, vbDirectory) = "" Then
            MkDir (caminho)
        End If
        
        xml = LerDadosJSON(resposta, "retEvento", "xml", "")
        chNFeCanc = LerDadosJSON(resposta, "retEvento", "chNFeCanc", "")
        Call salvarXML(xml, caminho, chNFeCanc, "CANC")

        If InStr(1, impressaoParam, "pdf") Then
        
            pdf = LerDadosJSON(resposta, "pdfCancelamento", "", "")
            Call salvarPDF(pdf, caminho, chNFeCanc, "CANC")
            
            If exibeNaTela Then
    
                ShellExecute 0, "open", caminho & chNFeCanc & "-procEvenNFe.pdf", "", "", vbNormalFocus
            
            End If
        End If
    Else
        MsgBox ("Ocorreu um erro, veja o Retorno da API para mais informacoes")
         gravaLinhaLog ("[Ocorreu um erro, veja o Retorno da API para mais informacoes  - Metodo: downloadEventoNF3eESalvar]")
    End If

    downloadEventoNF3eESalvar = resposta
End Function

'Esta fun��o realiza o cancelamento de uma NFC-e
Public Function cancelarNF3e(chNFe As String, tpAmb As String, dhEvento As String, nProt As String, xJust As String, caminho As String, exibeNaTela As Boolean) As String
    Dim json As String
    Dim url As String
    Dim resposta As String
    Dim status As String
    Dim respostaDownload As String

    'Monta o JSON
    json = "{"
    json = json & """chNFe"":""" & chNFe & ""","
    json = json & """tpAmb"":""" & tpAmb & ""","
    json = json & """dhEvento"":""" & dhEvento & ""","
    json = json & """nProt"":""" & nProt & ""","
    json = json & """xJust"":""" & xJust & """"
    json = json & "}"
    
    url = "https://nf3e.ns.eti.br/v1/nf3e/cancel"
    
    gravaLinhaLog ("[CANCELAMENTO_DADOS]")
    gravaLinhaLog (json)
    
    resposta = enviaConteudoParaAPI(json, url, "json")
        
    gravaLinhaLog ("[CANCELAMENTO_RESPOSTA]")
    gravaLinhaLog (resposta)
    
    status = LerDadosJSON(resposta, "status", "", "")
    
    'Se houve sucesso no evento, realiza o download
    If (status = "135") Then
    
        respostaDownload = downloadEventoNF3eESalvar(chNFe, tpAmb, caminho, exibeNaTela)
        status = LerDadosJSON(respostaDownload, "status", "", "")
        
        If (status <> "100") Then
            MsgBox ("Ocorreu um erro ao fazer o download. Verifique os logs.")
        End If
    End If
    
    cancelarNF3e = resposta
End Function

'Esta fun��o realiza a consulta de situa��o de uma NFC-e
Public Function consultarSituacao(chNFe As String, tpAmb As String) As String
    Dim json As String
    Dim url As String
    Dim resposta As String

    'Monta o JSON
    json = "{"
    json = json & """chNFe"":""" & chNFe & ""","
    json = json & """tpAmb"":""" & tpAmb & """"
    json = json & "}"

    url = "https://nf3e.ns.eti.br/v1/nf3e/status"
    
    gravaLinhaLog ("[CONSULTA_SITUACAO_DADOS]")
    gravaLinhaLog (json)

    resposta = enviaConteudoParaAPI(json, url, "json")
    
    gravaLinhaLog ("[CONSULTA_SITUACAO_RESPOSTA]")
    gravaLinhaLog (resposta)
    
    consultarSituacao = resposta
End Function


'Esta fun��o salva um XML
Public Sub salvarXML(xml As String, caminho As String, chNFe As String, Optional tipo As String = "")
    Dim fsT As Object
    Set fsT = CreateObject("ADODB.Stream")
    Dim conteudoSalvar  As String
    Dim localParaSalvar As String
    Dim extensao As String
    
    If (tipo = "CANC") Then
        extensao = "-procEvenNFe.xml"
    Else
        extensao = "-procNFe.xml"
    End If
    'Seta o caminho para o arquivo XML
    localParaSalvar = caminho & chNFe & extensao

    'Remove as contrabarras
    conteudoSalvar = Replace(xml, "\""", "")

    fsT.Type = 2
    fsT.Charset = "utf-8"
    fsT.Open
    fsT.WriteText conteudoSalvar
    fsT.SaveToFile localParaSalvar
End Sub

'Esta fun��o salva um PDF
Public Function salvarPDF(pdf As String, caminho As String, chNFe As String, Optional tipo As String = "") As Boolean
On Error GoTo SAI
    Dim conteudoSalvar  As String
    Dim localParaSalvar As String
    Dim extensao As String
    If (tipo = "CANC") Then
        extensao = "-procEvenNFe.pdf"
    Else
        extensao = "-procNFe.pdf"
    End If
    'Seta o caminho para o arquivo PDF
    localParaSalvar = caminho & chNFe & extensao

    Dim fnum
    fnum = FreeFile
    Open localParaSalvar For Binary As #fnum
    Put #fnum, 1, Base64Decode(pdf)
    Close fnum
    Exit Function
SAI:
    MsgBox (Err.Number & " - " & Err.Description), vbCritical
End Function

'Esta fun��o l� os dados de um JSON
Public Function LerDadosJSON(sJsonString As String, key1 As String, key2 As String, key3 As String, Optional key4 As String, Optional key5 As String) As String
On Error GoTo err_handler
    Dim oScriptEngine As ScriptControl
    Set oScriptEngine = New ScriptControl
    oScriptEngine.Language = "JScript"
    Dim objJSON As Object
    Set objJSON = oScriptEngine.Eval("(" + sJsonString + ")")
    If key1 <> "" And key2 <> "" And key3 <> "" And key4 <> "" And key5 <> "" Then
        LerDadosJSON = VBA.CallByName(VBA.CallByName(VBA.CallByName(VBA.CallByName(VBA.CallByName(objJSON, key1, VbGet), key2, VbGet), key3, VbGet), key4, VbGet), key5, VbGet)
    ElseIf key1 <> "" And key2 <> "" And key3 <> "" And key4 <> "" Then
        LerDadosJSON = VBA.CallByName(VBA.CallByName(VBA.CallByName(VBA.CallByName(objJSON, key1, VbGet), key2, VbGet), key3, VbGet), key4, VbGet)
    ElseIf key1 <> "" And key2 <> "" And key3 <> "" Then
        LerDadosJSON = VBA.CallByName(VBA.CallByName(VBA.CallByName(objJSON, key1, VbGet), key2, VbGet), key3, VbGet)
    ElseIf key1 <> "" And key2 <> "" Then
        LerDadosJSON = VBA.CallByName(VBA.CallByName(objJSON, key1, VbGet), key2, VbGet)
    ElseIf key1 <> "" Then
        LerDadosJSON = VBA.CallByName(objJSON, key1, VbGet)
    End If
Err_Exit:
    Exit Function
err_handler:
    LerDadosJSON = "Error: " & Err.Description
    Resume Err_Exit
End Function

'Esta funcao le os dados de um XML
Public Function LerDadosXML(sXml As String, key1 As String, key2 As String) As String
    On Error Resume Next
    LerDadosXML = ""
    
    Set xml = New DOMDocument60
    xml.async = False
    
    If xml.loadXML(sXml) Then
        'Tentar pegar o strCampoXML
        Set objNodeList = xml.getElementsByTagName(key1 & "//" & key2)
        Set objNode = objNodeList.nextNode
        
        Dim valor As String
        valor = objNode.Text
        
        If Len(Trim(valor)) > 0 Then 'CONSEGUI LER O XML NODE
            LerDadosXML = valor
        End If
        Else
        MsgBox "Nao foi possivel ler o conteudo do XML da NF3e especificado para leitura.", vbCritical, "ERRO"
    End If
End Function

'Esta fun��o grava uma linha de texto em um arquivo de log
Public Sub gravaLinhaLog(conteudoSalvar As String)
    Dim fsT As Object
    Set fsT = CreateObject("ADODB.Stream")
    Dim localParaSalvar As String
    Dim data As String
    
    'Diret�rio para salvar os logs
    localParaSalvar = App.Path & "\log\"
    
    'Checa se existe o caminho passado para salvar os arquivos
    If Dir(localParaSalvar, vbDirectory) = "" Then
        MkDir (localParaSalvar)
    End If
    
    'Pega data atual
    data = Format(Date, "yyyyMMdd")
    
    'Diret�rio + nome do arquivo para salvar os logs
    localParaSalvar = App.Path & "\log\" & data & ".txt"
    
    'Pega data e hora atual
    data = DateTime.Now
    
    Dim fnum
    fnum = FreeFile
    Open localParaSalvar For Append Shared As #fnum
    Print #fnum, data & " - " & conteudoSalvar
    Close fnum
End Sub

'Esta função imprime o cupom fiscal diretamente na impressora padrão do windows
Public Function imprimirNF3e(caminho As String, chNFe As String) As Long
    Dim impressao_aux As Long

    impressao_aux = ShellExecute(0, "print", caminho & chNFe & "-procNFe.pdf", vbNullString, vbNullString, vbNormalFocus)
    'recomendamos utilizar o leitor de PDF PDF-XChange por motivos de compatibilidade com o código em questão
End Function
