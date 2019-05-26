<%
'Copyright (c) Tom Kirby 2005
'This program and its associated source code is distributed under the terms of the GNU General Public 
'License.
'See the attached file COPYING.txt for more information
%>

<%Option Explicit%>
<%Server.ScriptTimeout = 2500%>
<html>
<head></head>
<body>

  
<%
  Dim xmlHttp
  
  on error resume next
  err.Clear
  set xmlHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
  
  
  if err.number <> 0 then
    err.number = 0
    set xmlHttp = Server.CreateObject("WinHttp.WinHttpRequest.5.1")

    if err.number <> 0 then
    
      err.number = 0
      set xmlHttp = Server.CreateObject("WinHttp.WinHttpRequest.5")
      
      if err.number <> 0 then
        response.Write "Oh o - xmlHttp is not installed on your server :-("
      else
        response.Write "Hooray - WinHTTP 5.0 is installed on your server.<br>However, XMLHTTP is not installed. If you use Unicode character sets, you will need to download and install XMLHTTP.<br>Please refer to the Yider's help file."
      end if
      
    else
      response.Write "Hooray - WinHTTP 5.1 is installed on your server.<br>However, XMLHTTP is not installed. If you use Unicode character sets, you will need to download and install XMLHTTP.<br>Please refer to the Yider's help file."
    end if

  else
    response.Write "Hooray - XMLHTTP is installed on your server"
  end if  
      

  g_set = g_set + 1
  set xmlHttp = Nothing
  g_set = g_set - 1
%>


</body>
</html>