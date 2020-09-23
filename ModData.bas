Attribute VB_Name = "Module"
''''''''''''''''''''''''''''' Hotmail Check Message ''''''''''''''''''''''''''''
'                                                                              '
'    This code uses the http/1.1 protocol to connect to the hotmail server     '
'    and retrieve the mail box (note: when i use the term mailbox 'data'       '
'    I am actually referring to the SOURCE CODE of the mailbox, which of       '
'    course is sent in html format). This program does not use any special     '
'    mail features, nor does it implement POP mail, it simply uses http        '
'    commands to get the mailbox. Because it is so confusing, I tried the      '
'    best i could to comment anywhere that there may be confusion, but         '
'    if you are not familiar with socket programming or the http protocol,     '
'    you will most likely have a difficult time understanding it.              '
'    And although the only piece of data you see as a result of this program   '
'    is how many new messages you have, once you understand how the program    '
'    works, retrieving any other information about your hotmail account is     '
'    a piece of cake. If you have any questions or comments, you can contact   '
'    me at:  nmjblue@hotmail.com                                               '
'                                                                              '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public StrLogin As String, StrPass As String ' holds login and password
Public NewHost As String, NewUrl As String ' new server and url after redirection (see below)
Public BatchNumber As Integer ' holds the current batch number we need to send
Public Cookies(6) As String ' stores cookies received, required for receiving mailbox (contains encrypted information read by server)
Public CurrentCookie As Integer ' stores current cookie number, as there are numerous different ones
Public MailData As String ' once we begin to receive data about mailbox, this is the string that stores it so we can retrieve the information
Public ReadBox As Boolean, BoxBatch As Integer ' boolean for whether or not we are receiving the mailbox data, and batch number of the data we are receiving

' Socket Values
Public Const AF_INET = 2
Public Const SOCK_STREAM = 1
Public Const IPPROTO_IP = 0
Public Const SOCKET_CONNECT = 2
Public Const SOCKET_CANCEL = 5
Public Const SOCKET_FLUSH = 6
Public Const SOCKET_DISCONNECT = 7

Public Function MakeString(Connection As Integer) As String
Dim strdata As String ' for temporary storage of data to send
Dim feed As String
feed = (Chr(13) & Chr(10)) ' carriage return & linefeed

Select Case Connection
Case 0 'first batch of data sent, contains login information
    Dim content As String
    content$ = "login=" & StrLogin$ & "&domain=hotmail.com&passwd=" & StrPass$ & "&enter=Sign+in&sec=no&curmbox=ACTIVE&js=yes&_lang=&beta=&ishotmail=1&id=2&ct=963865176"
    strdata = "POST /cgi-bin/dologin HTTP/1.1" & feed & "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, */*" & feed
    strdata = strdata & "Accept -Language: en -us" & feed & "Content-Type: application/x-www-form-urlencoded" & feed
    strdata = strdata & "Accept -Encoding: gzip , deflate" & feed & "User-Agent: Mozilla/4.0 (compatible; MSIE 5.0; Windows 98; DigExt)" & feed
    strdata = strdata & "Host: lc5.law5.hotmail.passport.com" & feed
    strdata = strdata & "Content-Length: " & Len(content$) & feed & "Connection: Keep -Alive" & feed & feed
    strdata = strdata & content$ & feed & feed
    MakeString = strdata
Case 1 'we get relocated to a new hotmail server (NewHost) containing the mailbox. here we request a new page, because contained in the url of the page (NewUrl) is our encrypted login and password
    strdata = "GET /" & NewUrl$ & " HTTP/1.1" & feed
    strdata = strdata & "User-Agent: Mozilla/4.0 (compatible; MSIE 5.0; Windows 98; DigExt)" & feed
    strdata = strdata & "Host: " & NewHost$ & feed
    strdata = strdata & "Cookie: MC1=V=2&GUID=B8E9C518070C49B18A9884F543033C33; mh=ENCA; MSPDom=; MSPAuth=; MSPProf=; MSPVis=; LO=; HMSC0899=; HMP1=1; HMSC0899="
    strdata = strdata & feed & feed
     '& feed
    MakeString = strdata
Case 2 'finally, we request the mailbox on the new server, by sending the cookies we received with all the encrypted information needed
    strdata = "GET " & NewUrl$ & " HTTP/1.1" & feed
    strdata = strdata & "User-Agent: Mozilla/4.0 (compatible; MSIE 5.0; Windows 98; DigExt)" & feed
    strdata = strdata & "Host: " & NewHost$ & feed
    strdata = strdata & "Connection: Keep-Alive" & feed
    strdata = strdata & "Cookie: HMP1=1; " & Cookies(4) & "; MSPDom=; " & Cookies(1) & "; " & Cookies(2) & "; MSPVis=1; LO=;" & feed & feed
    MakeString = strdata
End Select
End Function
