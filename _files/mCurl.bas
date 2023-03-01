Attribute VB_Name = "mCurl"
Option Explicit

Public m_P As hCurl

Private Declare Function curl_easy_init CDecl Lib "libcurl.dll" () As Long
Private Declare Function curl_easy_setopt CDecl Lib "libcurl.dll" (ByVal curl_easy_init As Long, ByVal OpTS As Long, Data As Any) As Long
Private Declare Function curl_easy_perform CDecl Lib "libcurl.dll" (ByVal curl_easy_init As Long) As Long
Private Declare Function curl_easy_cleanup CDecl Lib "libcurl.dll" (ByVal curl_easy_init As Long) As Long
Private Declare Function curl_slist_append CDecl Lib "libcurl.dll" (ByVal curl_list_handle As Long, ByVal lpStr As String) As Long
Private Declare Function curl_slist_free_all CDecl Lib "libcurl.dll" (ByVal curl_list_handle As Long) As Long
Private Declare Function curl_easy_escape CDecl Lib "libcurl.dll" (ByVal curl_easy_init As Long, ByVal lpStr As String, ByVal strsize As Long) As Long

Private Declare Function curl_mime_init CDecl Lib "libcurl.dll" (ByVal curl_easy_init As Long) As Long
Private Declare Function curl_mime_addpart CDecl Lib "libcurl.dll" (ByVal curl_mime_init As Long) As Long
Private Declare Function curl_mime_name CDecl Lib "libcurl.dll" (ByVal curl_mime_addpart As Long, ByVal Name As String) As Long
Private Declare Function curl_mime_filedata CDecl Lib "libcurl.dll" (ByVal curl_mime_addpart As Long, ByVal Path As String) As Long
Private Declare Function curl_mime_data CDecl Lib "libcurl.dll" (ByVal curl_mime_addpart As Long, ByVal Data As String, ByVal DataLength As Long) As Long
Private Declare Function curl_mime_free CDecl Lib "libcurl.dll" (ByVal curl_mime_init As Long) As Long

Private Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (desc As Any, src As Any, ByVal Length As Long) As Long

Private Const CURL_ZERO_TERMINATED As Long = -1

Private Const CURLOPT_URL As Long = 10002
Private Const CURLOPT_QUOTE As Long = 10028
Private Const CURLOPT_POSTQUOTE As Long = 10039
Private Const CURLOPT_UPLOAD As Long = 46
Private Const CURLOPT_WRITEFUNCTION As Long = 20011
Private Const CURLOPT_READFUNCTION As Long = 20012
Private Const CURLOPT_PROGRESSFUNCTION As Long = 20056
Private Const CURLOPT_HTTPHEADER As Long = 10023
Private Const CURLOPT_SSL_VERIFYPEER As Long = 64
Private Const CURLOPT_SSL_VERIFYHOST As Long = 81
Private Const CURLOPT_NOPROGRESS As Long = 43
Private Const CURLOPT_MIMEPOST As Long = 10269

Private Const CURLOPT_POST As Long = 47
Private Const CURLOPT_PUT As Long = 54

Private Const HeaderSeparator As String = vbCrLf
Private Const CommandSeparator As String = vbCrLf

Public m_file_output As String
Public rBuffer() As Byte, rBuffPtr As Long, rBuffEnd As Long
Public wBuffer() As Byte, MaxW As Long
 
Public m_total As Long

Public Type CurlFormItem
    Name As String
    Text As String
    IsFile As Boolean
End Type

Private EmptyForms() As CurlFormItem

Public Function WriteCallback CDecl(ByVal Data As Long, ByVal Size As Long, ByVal count As Long, ByVal custom As Long) As Long
Static m_last_timer As Double, m_last_total As Long
Dim m_dif As Double
Dim l_Buffer() As Byte
    count = count * Size
    If Len(m_file_output) Then
        Open m_file_output For Binary As #1
        ReDim l_Buffer(MaxW + count)
        CopyMemory l_Buffer(MaxW + 1), ByVal Data, count
        Put #1, , l_Buffer
        Close #1
    Else
        ReDim Preserve wBuffer(MaxW + count)
        CopyMemory wBuffer(MaxW + 1), ByVal Data, count
        MaxW = MaxW + count
    End If
    WriteCallback = count
    
    m_total = m_total + count
    m_dif = (Timer - m_last_timer)
    If m_dif > 0.5 Then
        'Debug.Print Time, Data, custom, size, count, FormatarEm(m_total, FoLong), FormatarEm((m_total - m_last_total) / m_dif, FoLong)
        m_last_timer = Timer
        m_last_total = m_total
    End If
End Function

'void *ptr, size_t size, size_t nmemb, void *stream
Public Function ReadCallback CDecl(ByVal Data As Long, ByVal Size As Long, ByVal count As Long, ByVal custom As Long) As Long
    If rBuffPtr = 0 Then Exit Function
    
    count = count * Size
    If count > rBuffEnd - rBuffPtr Then count = rBuffEnd - rBuffPtr
    If count = 0 Then Exit Function
    CopyMemory ByVal Data, ByVal rBuffPtr, count
    rBuffPtr = rBuffPtr + count
    ReadCallback = count / Size
End Function

Function ProgressCallback CDecl(ByVal clientp As Long, ByVal dlTotal As Double, ByVal dlNow As Double, ByVal ulTotal As Double, ByVal ulNow As Double) As Long
'Debug.Print "ProgressCallback(" & dlTotal & "," & dlNow & "," & ulTotal & "," & ulNow & ")"
    If dlTotal <> 0 And Not m_P Is Nothing Then
        'm_CProgress.Value = (dlNow / dlTotal) * 100
        m_P.setProgress (dlNow / dlTotal) * 100
    End If
'DoEvents
End Function

Function CurlForm(Url As String, FormItems() As CurlFormItem, Optional Headers As String) As String
Dim Size As Long
On Error GoTo err
Size = -1
Size = UBound(FormItems)
err: On Error GoTo 0
CurlForm = CurlGeneric(0, Url, Headers, FormItems, Size)
End Function

Function CurlGet(Url As String, Optional Headers As String) As String
    CurlGet = CurlGeneric(0, Url, Headers, EmptyForms, -1)
End Function

Function CurlPost(Url As String, Data As String, Optional Headers As String) As String
    rBuffer = StrConv(Data, vbFromUnicode)
    
    If UBound(rBuffer) <> -1 Then
        rBuffPtr = VarPtr(rBuffer(0))
        rBuffEnd = UBound(rBuffer) + rBuffPtr + 1
    Else
        rBuffPtr = 0
    End If
    
    CurlPost = CurlGeneric(CURLOPT_POST, Url, Headers, EmptyForms, -1)
    Erase rBuffer
End Function

Function CurlPut(Url As String, Data As String, Optional Headers As String) As String
    rBuffer = StrConv(Data, vbFromUnicode)
    
    If UBound(rBuffer) <> -1 Then
        rBuffPtr = VarPtr(rBuffer(0))
        rBuffEnd = UBound(rBuffer) + rBuffPtr + 1
    Else
        rBuffPtr = 0
    End If
    
    CurlPut = CurlGeneric(CURLOPT_PUT, Url, Headers, EmptyForms, -1)
    Erase rBuffer
End Function

Function CurlSftp(Url As String, Command As String) As String
    CurlSftp = CurlSftpCommand(Url, Command)
End Function

Function CurlSftpUpload(Url As String, Dest As String, Data As String) As String
    rBuffer = StrConv(Data, vbFromUnicode)
    
    If UBound(rBuffer) <> -1 Then
        rBuffPtr = VarPtr(rBuffer(0))
        rBuffEnd = UBound(rBuffer) + rBuffPtr + 1
    Else
        rBuffPtr = 0
    End If
    
    CurlSftpUpload = CurlSftpUpload2(Url)
    Erase rBuffer
End Function

Private Function CurlGeneric(ByVal OptToSet As Long, Url As String, Headers As String, Forms() As CurlFormItem, ByVal Max As Long) As String
      Dim H() As String
      Dim Curl As Long
      Dim list As Long
      Dim Z As Long
      Dim Ret As Long
      
      Dim Form As Long, Field As Long
10       On Error GoTo CurlGeneric_Error

20        H = Split(Headers, HeaderSeparator)
          
30        For Z = 0 To UBound(H)
40            list = curl_slist_append(list, H(Z))
50        Next

60        Erase wBuffer
70        MaxW = -1

80        Curl = curl_easy_init
90        Debug.Print curl_easy_setopt(Curl, CURLOPT_URL, ByVal Url & Chr$(0))
100       Debug.Print curl_easy_setopt(Curl, CURLOPT_WRITEFUNCTION, ByVal l(AddressOf WriteCallback))
110       Debug.Print curl_easy_setopt(Curl, CURLOPT_READFUNCTION, ByVal l(AddressOf ReadCallback))
'120       curl_easy_setopt Curl, CURLOPT_PROGRESSFUNCTION, ByVal l(AddressOf ProgressCallback))
'130       Debug.Print curl_easy_setopt(Curl, CURLOPT_NOPROGRESS, ByVal 0&)
140       Debug.Print curl_easy_setopt(Curl, CURLOPT_SSL_VERIFYPEER, ByVal 0&)
150       Debug.Print curl_easy_setopt(Curl, CURLOPT_SSL_VERIFYHOST, ByVal 0&)

            
    If Max > -1 Then
        Form = curl_mime_init(Curl)
        For Z = 0 To Max
            Field = curl_mime_addpart(Form)
            
            Debug.Assert Ret = 0
            If Forms(Z).IsFile Then
                Ret = curl_mime_filedata(Field, Forms(Z).Text)
                'Ret = curl_mime_name(Field, "p")
                Debug.Assert Ret = 0
            Else
                Ret = curl_mime_data(Field, Forms(Z).Text, CURL_ZERO_TERMINATED)
                Ret = curl_mime_name(Field, Forms(Z).Name)
                Debug.Assert Ret = 0
            End If
        Next
        curl_easy_setopt Curl, CURLOPT_MIMEPOST, ByVal Form
    End If
            
165       If OptToSet Then
    Debug.Print curl_easy_setopt(Curl, OptToSet, ByVal 1&)
End If
160       If list Then
    Debug.Print curl_easy_setopt(Curl, CURLOPT_HTTPHEADER, ByVal list)
    End If

170       Ret = curl_easy_perform(Curl)

180       CurlGeneric = StrConv(wBuffer, vbUnicode)
190       Erase wBuffer

200       If list Then curl_slist_free_all list

210       curl_easy_cleanup Curl

    If Form Then curl_mime_free Form

220      On Error GoTo 0
230      Exit Function

CurlGeneric_Error:

240       MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure CurlGeneric of Módulo mod_Curl, " & Erl
         
End Function

Private Function CurlSftpCommand(Url As String, Command As String) As String
Dim H() As String
Dim Curl As Long
Dim list As Long
Dim Z As Long
Dim Ret As Long
    H = Split(Command, CommandSeparator)
    
    For Z = 0 To UBound(H)
        list = curl_slist_append(list, H(Z))
    Next

    Erase wBuffer
    MaxW = -1
    
    Curl = curl_easy_init
    curl_easy_setopt Curl, CURLOPT_URL, ByVal Url
    curl_easy_setopt Curl, CURLOPT_WRITEFUNCTION, ByVal l(AddressOf WriteCallback)
    curl_easy_setopt Curl, CURLOPT_READFUNCTION, ByVal l(AddressOf ReadCallback)
    curl_easy_setopt Curl, CURLOPT_PROGRESSFUNCTION, ByVal l(AddressOf ProgressCallback)
    curl_easy_setopt Curl, CURLOPT_NOPROGRESS, ByVal 0&
    curl_easy_setopt Curl, CURLOPT_SSL_VERIFYPEER, ByVal 0&

    If list Then curl_easy_setopt Curl, CURLOPT_QUOTE, ByVal list

    Ret = curl_easy_perform(Curl)
    'Debug.Print "ret:" & Ret
    CurlSftpCommand = StrConv(wBuffer, vbUnicode)
    Erase wBuffer
    
    If list Then curl_slist_free_all list

    curl_easy_cleanup Curl

End Function

Private Function l(ByVal v As Long) As Long
    l = v
End Function

Function CurlSftpUpload2(Url As String) As String
Dim H() As String
Dim Curl As Long
Dim list As Long
Dim Z As Long
Dim Ret As Long

    'list = curl_slist_append(list, "RNFR " & "TEMPORARY_FILE_" & Time)
    'list = curl_slist_append(list, "RNTO " & Filename)
    
    Erase wBuffer
    MaxW = -1
        
    Curl = curl_easy_init
    curl_easy_setopt Curl, CURLOPT_URL, ByVal Url
    curl_easy_setopt Curl, CURLOPT_WRITEFUNCTION, ByVal l(AddressOf WriteCallback)
    curl_easy_setopt Curl, CURLOPT_READFUNCTION, ByVal l(AddressOf ReadCallback)
    curl_easy_setopt Curl, CURLOPT_PROGRESSFUNCTION, ByVal l(AddressOf ProgressCallback)
    curl_easy_setopt Curl, CURLOPT_NOPROGRESS, ByVal 0&
    curl_easy_setopt Curl, CURLOPT_UPLOAD, ByVal 1&
    curl_easy_setopt Curl, CURLOPT_SSL_VERIFYPEER, ByVal 0&

    Ret = curl_easy_perform(Curl)
    
    CurlSftpUpload2 = StrConv(wBuffer, vbUnicode)
    Erase wBuffer
    
    If list Then curl_slist_free_all list

    curl_easy_cleanup Curl
    'Debug.Print StrConv(wBuffer, vbUnicode)
End Function
