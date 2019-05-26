<%@CodePage=1250 Language=VBScript%>
<%
'Copyright (c) Tom Kirby 2005-2007
'This program and its associated source code is distributed under the terms of the GNU General Public 
'License.
'See the attached file COPYING.txt for more information
%>

<%Option Explicit%>

<!-- #include file="configuration.asp" -->
<%
response.charset = g_charset
'Copyright (c) Tom Kirby 2005-2007
'All rights reserved
'This source code is subject to the licensing conditions at http://www.transworldinteractive.net
%>


<html>
<head>
<!--If your site is in a foreign language, uncomment and modify the link below to your charset-->
<meta http-equiv="Content-Type" content="text/html; charset=<%=g_charset%>">
</head>
<body>

<div> Welcome to Transworld Interactive Search Engine<br>
  &nbsp;
    
  <!-- Abyss Logo At Top of page. -->
    
  <br><center>
  <img src="images/abysslogo.jpg" alt="Abyss Search" hspace="250" longdesc="http://search.transworldinteractive.net"><br>
  
  <!-- #include file="CYiderSearch.asp" -->
  <!-- #include file="search_button_input.asp" --></center>
    
  <!-- #include file="search_include.asp" -->
    
  <br>
  <br>
  Copyright 2012 Transworld Interactive</div>
</body>
</html>

<%

if g_set <> 0 and InStr(g_url_to_spider, "127.0.0.1") <> 0 then
  response.Write vbcrlf & "<script language=""JavaScript"">"
  response.write "alert('The search has found an error. g_set is " & g_set & "');"
  response.write vbcrlf & "</script>"
  SendErrorReport "g_set"
end if

if g_open <> 0 and InStr(g_url_to_spider, "127.0.0.1") <> 0 then
  response.Write vbcrlf & "<script language=""JavaScript"">"
  response.write "alert('The search has found an error. g_open is " & g_open & "');"
  response.write vbcrlf & "</script>"
  SendErrorReport "g_open"
end if

%>