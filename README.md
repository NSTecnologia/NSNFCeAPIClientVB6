# NSNF3eAPIClientVB6

Esta página apresenta trechos de códigos de um módulo em VB6 que foi desenvolvido para consumir as funcionalidades da NS NF3e API.

-------

## Primeiros passos:

### Integrando ao sistema:

Para utilizar as funções de comunicação com a API, você precisa realizar os seguintes passos:

1. Extraia o conteúdo da pasta compactada que você baixou;
2. Copie para a pasta da sua aplicação os módulos **NF3eAPI.bas** e **base64Convert.bas**, que estão na pasta raiz;
3. Abra o seu projeto e importe a pasta copiada.
4.A aplicação utiliza as bibliotecas **Microsoft Script Control 1.0** e **Active Microsoft XML, v6.0** para realizar a comunicação com a API e fazer a manipulação de dados JSON, respectivamente. Ative as duas referencias em: **Project > References**. 

**Pronto!** Agora, você já pode consumir a NS NF3e API através do seu sistema. Todas as funcionalidades de comunicação foram implementadas no módulo NF3eAPI.bas. Confira abaixo sobre realizar uma emissão completa.

------

## Emissão Sincrona:

### Realizando uma Emissão:

Para realizar uma emissão completa, você poderá utilizar a função emitirNF3eSincrono do módulo NF3eAPI. Veja abaixo sobre os parâmetros necessários, e um exemplo de chamada do método.

##### Parâmetros:

**ATENÇÃO:** o **token** também é um parâmetro necessário e você deve primeiramente defini-lo no módulo NF3eAPI.bas. Ele é uma constante do módulo. 

Parametros     | Descrição
:-------------:|:-----------
conteudo       | Conteúdo de emissão do documento.
tpConteudo     | Tipo de conteúdo que está sendo enviado. Valores possíveis: json, xml, txt
tpAmb          | Ambiente onde foi autorizado o documento.Valores possíveis:<ul> <li>1 - produção</li> <li>2 - homologação</li> </ul>
caminho        | Caminho onde devem ser salvos os documentos baixados.
exibeNaTela    | Se for baixado, exibir o PDF na tela após a autorização.Valores possíveis: <ul> <li>**True** - será exibido</li> <li>**False** - não será exibido</li> </ul> 

##### Exemplo de chamada:

Após ter todos os parâmetros listados acima, você deverá fazer a chamada da função. Veja o código de exemplo abaixo:
           
    Dim retorno As String
    retorno = emitirNF3eSincrono(conteudoEnviar, "json", "2", "C:\Documentos", True)
    MessageBox(retorno)

A função **emitirNF3eSincrono** fará o envio, a consulta e download do documento, utilizando as funções emitirNFe, consultarStatusProcessamento e downloadNFeAndSave, presentes no módulo NFeAPI.bas. Por isso, o retorno será um JSON com os principais campos retornados pelos métodos citados anteriormente. No exemplo abaixo, veja como tratar o retorno da função emitirNFeSincrono:

##### Exemplo de tratamento de retorno:

O JSON retornado pelo método terá os seguintes campos: statusEnvio, statusDownload, cStat, chNFe, nProt, motivo, erros. Veja o exemplo abaixo:

    {
        "statusEnvio": "200",
        "statusDownload": "200",
        "cStat": "100",
        "chNFe": "43181007364617000135550000000119741004621864",
        "nProt": "143180007036833",
        "motivo": "Autorizado o uso da NF-e",
        "erros": ""
    }
      
Confira um código para tratamento do retorno, no qual pegará as informações dispostas no JSON de Retorno disponibilizado:

    Dim retorno As String
    retorno = emitirNF3eSincrono(conteudoEnviar, "json", "2", "C:\Documentos", True)

    Dim statusEnvio, statusDownload, cStat, chNFe, nProt, motivo, erros As String

    statusEnvio = LerDadosJSON(retorno, "statusEnvio", "", "")
    statusDownload = LerDadosJSON(retorno, "statusDownload", "", "")
    cStat = LerDadosJSON(retorno, "cStat", "", "")
    chNFe = LerDadosJSON(retorno, "chNFe", "", "")
    nProt = LerDadosJSON(retorno, "nProt", "", "")
    motivo = LerDadosJSON(retorno, "motivo", "", "")
    erros = LerDadosJSON(retorno, "erros", "", "")
    
    'Testa se houve sucesso na emissão
    If (statusEnvio = 100) Or (statusEnvio = -100) Then
        'Testa se a nota foi autorizada
        If (cStat = 100) Then

            'Aqui dentro você pode realizar procedimentos como desabilitar o botão de emitir, etc
            MsgBox (motivo)

            'Testa se o download teve problemas
            If (statusDownload <> 100) Then

            MsgBox ("Erro no Download")

            End If
        'Caso tenha dado erro na consulta
        Else
            MsgBox (motivo)

        End If
    Else
        'Aqui você pode exibir para o usuário o erro que ocorreu no envio
        MsgBox (motivo + Chr(13) + erros)
    End If

-----

## Demais Funcionalidades:

No módulo NFeAPI, você pode encontrar também as seguintes funcionalidades:

NOME                     | FINALIDADE             | DOCUMENTAÇÂO CONFLUENCE
:-----------------------:|:----------------------:|:-----------------------
**enviaConteudoParaAPI** |Função genérica que envia um conteúdo para API. Requisições do tipo POST.|
**emitirNF3e** | Envia uma NF3-e para processamento.|[Emissão de NF3e]().
**downloadNF3e** | Baixa documentos de emissão de uma NF3e autorizada. | [Download da NF3e]()
**downloadNF3eESalvar** | Baixa documentos de emissão de uma NF3e autorizada e salva-os em um diretório. | Por utilizar o método downloadNF3e, a documentação é a mesma. 
**downloadEventoNF3eESalvar** | Baixa documentos de evento de uma NF3e autorizada e salva-os em um diretório. | Por utilizar o método downloadNF3e, a documentação é a mesma. 
**cancelarNF3e** | Realiza o cancelamento de uma NF3-e. | [Cancelamento de NF3e]().
**consultarSituacao** | Consulta a situação de uma NFC-e na Sefaz. | [Consulta Situação da NF3e]().
**salvarXML** | Salva um XML em um diretório. | 
**salvarJSON** | Salva um JSON em um diretório. |
**salvarPDF** |	Salva um PDF em um diretório. | 
**LerDadosJSON** | 	Lê o valor de um campo de um JSON. |
**LerDadosXML** | Lê o valor de um campo de um XML. | 
**gravaLinhaLog** | Grava uma linha de texto no arquivo de log. | 



![Ns](https://nstecnologia.com.br/blog/wp-content/uploads/2018/11/ns%C2%B4tecnologia.png) | Obrigado pela atenção!
