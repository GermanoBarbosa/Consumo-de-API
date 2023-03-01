VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "hCurl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Progress(ByVal m_por As Double)
'Public m_CProgress As Object

Dim FormItems() As CurlFormItem
Dim MaxFI As Long

Sub AddFormFieldText(Name As String, Text As String)
    MaxFI = MaxFI + 1
    ReDim Preserve FormItems(MaxFI)
    FormItems(MaxFI).Name = Name
    FormItems(MaxFI).Text = Text
    FormItems(MaxFI).IsFile = False
End Sub

Sub AddFormFieldFile(Name As String, Filename As String)
    MaxFI = MaxFI + 1
    ReDim Preserve FormItems(MaxFI)
    FormItems(MaxFI).Name = Name
    FormItems(MaxFI).Text = Filename
    FormItems(MaxFI).IsFile = True
End Sub

Function CurlGet(Url As String, Optional Headers As String) As String
    CurlGet = mCurl.CurlGet(Url, Headers)
End Function

Function CurlSftp(Url As String, Command As String) As String
    CurlSftp = mCurl.CurlSftp(Url, Command)
End Function

Function CurlSftpUpload(Url As String, Dest As String, Data As String) As String
    CurlSftpUpload = mCurl.CurlSftpUpload(Url, Dest, Data)
End Function

Friend Sub setProgress(ByVal m_por As Double)
    RaiseEvent Progress(m_por)
End Sub

Private Sub Class_Initialize()
    Set m_P = Me
    MaxFI = -1
End Sub

Function CurlPost(Url As String, Data As String, Optional Headers As String) As String
    CurlPost = mCurl.CurlPost(Url, Data, Headers)
End Function

Function CurlForm(Url As String, Optional Headers As String) As String
    CurlForm = mCurl.CurlForm(Url, FormItems, Headers)
    Erase FormItems
    MaxFI = -1
End Function

Function CurlPut(Url As String, Data As String, Optional Headers As String) As String
    CurlPut = mCurl.CurlPut(Url, Data, Headers)
End Function
