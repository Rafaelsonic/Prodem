VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form im_webxml 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   Begin InetCtlsObjects.Inet Inet 
      Left            =   390
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   9435
   End
End
Attribute VB_Name = "im_webxml"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private m_objDOMPessoa As DOMDocument
Private m_blnItemClicked As Boolean
Private m_strXmlPath As String
Dim flagpreenche As Boolean

Sub GERA()
  Dim objPessoaRoot As IXMLDOMElement
  Dim objPessoaElement As IXMLDOMElement
  Dim tvwRoot As Node
  Dim X As IXMLDOMNodeList
  
  flagpreenche = True
  Set m_objDOMPessoa = New DOMDocument
  m_objDOMPessoa.resolveExternals = True
  m_objDOMPessoa.validateOnParse = True
  'carrega o XML no documento DOM
  m_objDOMPessoa.async = False
  Call m_objDOMPessoa.Load(m_strXmlPath)
  'verifica se a carga do XML foi feita com sucesso
  If m_objDOMPessoa.parseError.reason <> "" Then
    MsgBox m_objDOMPessoa.parseError.reason
    Exit Sub
  End If
  'obtem o elemento raiz do XML
   Set objPessoaRoot = m_objDOMPessoa.documentElement
  
  
  
  ' iteracao atraves de cada elemento para encher a arvore
  ' que por sua vez interaagem atraves de cada childNode
  ' do element(objPessoaElement)
  For Each objPessoaElement In objPessoaRoot.childNodes
    PreencherTreeWithChildren objPessoaElement
  Next objPessoaElement
  

End Sub

Private Sub Form_Load()
m_strXmlPath = App.Path & "\GERAXML.xml"
'm_strXmlPath = App.Path & "\agenda.xml"
BaixaXML
flagpreenche = False
GERA
End Sub

Private Sub PreencherTreeWithChildren(objDOMNode As IXMLDOMElement)

  Dim objDataNode As IXMLDOMNode
  Dim objAttributes As IXMLDOMNamedNodeMap
  Dim objAttributeNode As IXMLDOMNode
  Dim objPessoaElement As IXMLDOMElement
  Dim intIndex As Integer
  
  
  Dim data As String
  Dim cliente As String
  Dim junta As String
  'obtem o nome do elemento selecionado
  Set objDataNode = objDOMNode.selectSingleNode("cliente")

  'inclui os elementos aos nós
  
  
  Set objAttributes = objDOMNode.Attributes
  
  'verifica os atributos
  If objAttributes.length > 0 Then
   
    ' obtendo o item para a referencia  ''PERSONID',
    ' com NameNodeListMap para o Nó atual
    Set objAttributeNode = objAttributes.getNamedItem("cli")
    cliente = objAttributeNode.nodeName
    cliente = cliente & objAttributeNode.nodeTypedValue
  
    Set objAttributeNode = objAttributes.getNamedItem("dt")
    data = objAttributeNode.nodeName
    data = cliente & objAttributeNode.nodeTypedValue
  
  End If
  


  
  'interagem através dos Nós filhos(childNodes) do objeto DOMNode
  ' para preencher o TreeView os seus valores
  For Each objPessoaElement In objDOMNode.childNodes
    'Set tvwChildElement = tvwPessoa.Nodes.Add(intIndex, tvwChild)
            junta = junta & objPessoaElement.nodeName & " " & objPessoaElement.nodeTypedValue
            If UCase$(objPessoaElement.nodeName) = "RG" Then
                List1.AddItem data & cliente & junta
                junta = ""
            End If
  Next
       
       junta = ""
End Sub





Sub BaixaXML()
'colocar funcao para excluir o arquivo local atual
Dim Retorno As Boolean
Dim Endereco As String, Destino As String, Registro As String
    Endereco = "http://colwt06004/avg/avg6info.ctf"
    Endereco = "http://colsv03/avg/avg6info.ctf"
    Endereco = "http://comp/pages/wf/geraxml.asp"
    Destino = App.Path & "\GERAXML.xml"
    Retorno = download("http://comp/pages/wf/geraxml.asp", App.Path & "\GERAXML.xml")
    
End Sub

