VERSION 5.00
Begin VB.Form frmNF3eAPI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NF3e API"
   ClientHeight    =   9300
   ClientLeft      =   6810
   ClientTop       =   990
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9300
   ScaleWidth      =   10500
   Begin VB.CheckBox checkImprime 
      Caption         =   "Imprimir PDF"
      Height          =   195
      Left            =   1530
      TabIndex        =   11
      Top             =   5550
      Value           =   1  'Checked
      Width           =   1200
   End
   Begin VB.TextBox txtCaminho 
      Height          =   315
      Left            =   120
      TabIndex        =   9
      Text            =   "C:\Notas\"
      Top             =   360
      Width           =   8055
   End
   Begin VB.ComboBox cbTpConteudo 
      Height          =   315
      ItemData        =   "frmNF3eAPI.frx":0000
      Left            =   8400
      List            =   "frmNF3eAPI.frx":000D
      TabIndex        =   8
      Text            =   "txt"
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox txtTpAmb 
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Text            =   "2"
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CheckBox checkExibir 
      Caption         =   "Exibir PDF"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   5520
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.TextBox txtResult 
      Height          =   3015
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   6120
      Width           =   10215
   End
   Begin VB.TextBox txtConteudo 
      Height          =   3615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   1080
      Width           =   10215
   End
   Begin VB.CommandButton cmdEnviar 
      Caption         =   "Enviar Documento para Processamento"
      Height          =   615
      Left            =   4440
      TabIndex        =   0
      Top             =   5040
      Width           =   5775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Salvar em:"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   750
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Ambiente:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   4800
      Width           =   1290
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Resposta do Servidor"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   5880
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Conteudo"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   690
   End
End
Attribute VB_Name = "frmNF3eAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnviar_Click()
    On Error GoTo SAI
    Dim retorno As String
    Dim token As String
    
    If (txtCaminho.Text <> "") And (txtConteudo.Text <> "") And (cbTpConteudo.Text <> "") And (txtTpAmb.Text <> "") Then
        
        'Faz a emissao sincrona
        retorno = emitirNF3eSincrono(txtConteudo.Text, cbTpConteudo.Text, txtTpAmb.Text, txtCaminho.Text, checkExibir.Value, checkImprime.Value)
        txtResult.Text = retorno
        
        'Abaixo, confira um exemplo de tratamento de retorno da funcao emitirNF3eSincrono
        
        Dim statusEnvio, statusDownload, cStat, chNFe, nProt, motivo, erros As String
        
        'Le o statusEnvio
        statusEnvio = LerDadosJSON(retorno, "statusEnvio", "", "")
        'Le o statusDownload
        statusDownload = LerDadosJSON(retorno, "statusDownload", "", "")
        'Le o cStat
        cStat = LerDadosJSON(retorno, "cStat", "", "")
        'Le a chNFe
        chNFe = LerDadosJSON(retorno, "chNFe", "", "")
        'Le o nProt
        nProt = LerDadosJSON(retorno, "nProt", "", "")
        'Le o motivo
        motivo = LerDadosJSON(retorno, "motivo", "", "")
        'Le os erros
        erros = LerDadosJSON(retorno, "erros", "", "")
        
        'Agora que voce ja leu os dados, e aconselhavel que faca o salvamento de todos
        'eles no seu banco de dados antes de prosseguir para o teste abaixo
                 
        'Testa se houve sucesso na emissao
        If (statusEnvio = 100) Or (statusEnvio = -100) Then

                'Testa se a nota foi autorizada
                If (cStat = 100) Then
                
                    'Aqui dentro voce pode realizar procedimentos como desabilitar o botao de emitir, etc
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
            'Aqui voce pode exibir para o usuario o erro que ocorreu no envio
            MsgBox (motivo + Chr(13) + erros)
        End If
    Else
        MsgBox ("Todos os campos devem ser preenchidos")
    End If
    
    Exit Sub
SAI:
    MsgBox ("Problemas ao Requisitar emissao ao servidor" & vbNewLine & Err.Description), vbInformation, titleCTeAPI

End Sub
