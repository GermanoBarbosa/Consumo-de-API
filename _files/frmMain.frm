VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   9915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   4935
      Left            =   960
      TabIndex        =   3
      Top             =   2280
      Width           =   8175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Text            =   "06990590000123"
      Top             =   780
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Pesquisa"
      Height          =   855
      Left            =   600
      TabIndex        =   0
      Top             =   540
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "CNPJ"
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   360
      Width           =   3315
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim m_Curl As New hCurl
Dim m_saida As String
Dim m_Headers As String
Dim m_meu_token As String
Dim m_jbag As New hJsonBag
Dim cc As Long
    
    'Cadastre aqui o seu token
    'http://portal.easycodar.com.br:2020/open.do?action=open&sys=EAS
    
    m_meu_token = "cole aqui"
    m_Headers = "Authorization: " & m_meu_token
    m_saida = m_Curl.CurlGet("https://comunidade.easycodar.com.br/cnpj?cnpj=" & Text1.Text, m_Headers)
    m_jbag.JSON = m_saida
    
    List1.AddItem "UF: " & m_jbag("uf")
    List1.AddItem "cep: " & m_jbag("uf")
    List1.AddItem "cnpj: " & m_jbag("cnpj")
    List1.AddItem "pais: " & m_jbag("pais")
    List1.AddItem "email: " & m_jbag("email")
    List1.AddItem "porte: " & m_jbag("porte")
    List1.AddItem "bairro: " & m_jbag("bairro")
    List1.AddItem "numero: " & m_jbag("numero")
    List1.AddItem "ddd_fax: " & m_jbag("ddd_fax")
    List1.AddItem "municipio: " & m_jbag("municipio")
    List1.AddItem "logradouro: " & m_jbag("logradouro")
    List1.AddItem "cnae_fiscal: " & m_jbag("cnae_fiscal")
    List1.AddItem "codigo_pais: " & m_jbag("codigo_pais")
    List1.AddItem "complemento: " & m_jbag("complemento")
    List1.AddItem "codigo_porte: " & m_jbag("codigo_porte")
    List1.AddItem "razao_social: " & m_jbag("razao_social")
    List1.AddItem "nome_fantasia: " & m_jbag("nome_fantasia")
    List1.AddItem "capital_social: " & m_jbag("capital_social")
    List1.AddItem "ddd_telefone_1: " & m_jbag("ddd_telefone_1")
    List1.AddItem "ddd_telefone_2: " & m_jbag("ddd_telefone_2")
    List1.AddItem "opcao_pelo_mei: " & m_jbag("opcao_pelo_mei")
    List1.AddItem "descricao_porte: " & m_jbag("descricao_porte")
    
    For cc = 1 To m_jbag("qsa").Count - 1
        List1.AddItem "pais: " & m_jbag("qsa")(cc)("pais")
        List1.AddItem "nome_socio: " & m_jbag("qsa")(cc)("nome_socio")
    Next
End Sub

