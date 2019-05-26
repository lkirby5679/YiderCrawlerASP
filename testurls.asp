<%Option Explicit%>
<%Server.ScriptTimeout = 2500%>
<html>
<head>
<meta name="GENERATOR" Content="Microsoft Visual Studio.NET 7.0">
</head>
<body>

<%

Dim html, XMLHttp, url

url = "http://www.google.com"
  
'on error resume next

set XMLHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
  

if err.number <> 0 then
  err.number = 0
  set XMLHttp = Server.CreateObject("WinHttp.WinHttpRequest.5.1")
  
  if err.number <> 0 then
    set XMLHttp = Server.CreateObject("WinHttp.WinHttpRequest.5")
    Response.Write "<br>WinHttp.WinHttpRequest.5"
  else
    Response.Write "<br>WinHttp.WinHttpRequest.5.1"
  end if
  
else
 
  Response.Write "<br>MSXML2.ServerXMLHTTP"
  
end if


XMLHttp.open "GET", url, false, "", ""
XMLHttp.send()
html = XMLHttp.ResponseText

Response.Write "<br><br>html is " & XMLHttp.ResponseText



%>



</body>
</html>
