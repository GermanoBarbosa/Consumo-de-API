VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Pesquisa"
      Height          =   855
      Left            =   960
      TabIndex        =   0
      Top             =   2820
      Width           =   1935
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
    
    'Cadastre aqui o seu token
    'http://portal.easycodar.com.br:2020/open.do?action=open&sys=EAS
    
    m_meu_token = "cole Aqui"
    m_Headers = "Authorization: " & m_meu_token
    m_saida = m_Curl.CurlGet("https://comunidade.easycodar.com.br/cnpj?cnpj=06990590000123", m_Headers)
    MsgBox m_saida
End Sub

