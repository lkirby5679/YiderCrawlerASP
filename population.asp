<%@CodePage=1250 Language=VBScript%>

<%
'Copyright (c) Tom Kirby 2005
'This program and its associated source code is distributed under the terms of the GNU General Public 
'License.
'See the attached file COPYING.txt for more information
%>

<%Option Explicit%>
<%Server.ScriptTimeout = 2500%>
<%
'Copyright (c) Tom Kirby 2005
'All rights reserved
'This source code is subject to the licensing conditions at http://www.transworld.dyndns.ws

  Response.charset = g_charset
  Response.Buffer = True
  Response.Expires = 0
  Response.Expiresabsolute = Now() -1
  Response.AddHeader "pragma","no-cache"
  Response.AddHeader "cache-control", "private" 'stops proxy server cache
  Response.CacheControl = "no-cache"
%> 

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
	<head>
	</head>
	<body>
		<!--start of code you must include-->
		<!-- #include file="CHTMLTextExtractor.asp" -->
		<!-- #include file="configuration.asp" -->
		<form name="population" method="post" action="<%Response.Write(Filename(Request.ServerVariables("SCRIPT_NAME")))%>">
			<table cellpadding="0" cellspacing="0" border="0" align="center" valign="middle">
				<tr>
					<td style="font-family:arial; font-size:10pt;"><span style="font-weight:bold">Transworld Interactive  Options<br></span>
						<br>
						<input type="submit" name="populate" value="Populate"> - Click 
						here to populate the Yider table
						<br>
						<input type="submit" name="clear" value="Clear"> - Click here to 
						remove all data from the Yider table
						<%

  Dim extractor
  set extractor = new CHTMLTextExtractor
  g_set = g_set + 1
  
  'on error resume next

  extractor.Constructor g_database_connection
  extractor.m_username = g_username
  extractor.m_password = g_password
  extractor.m_english = g_english
  extractor.m_charset = g_charset
  extractor.m_local_ID = g_local_ID
  extractor.m_compact = g_compact
  extractor.m_strip_url_parameters = g_strip_url_parameters

  if request("clear") = "Clear" or request.QueryString("clear") = "Clear" then
    extractor.Clear
    if err.number = 0 then
      response.write ("<script scr=""JavaScript1.2"">alert('The Yider table has been cleared!');</script>")
    end if

  elseif request("populate") = "Populate" or request.QueryString=("populate") = "Populate" or request.QueryString=("auto_populate") = 1 or request("auto_populate") = "1" then
    extractor.m_wait = g_pause
    extractor.m_urls_per_iteration = g_urls_per_iteration
    
    extractor.StoreTextThroughoutDomain g_url_to_spider, g_valid_file_extensions, g_valid_url_strings, g_bad_page_strings, g_urls_not_to_view, g_urls_to_view_not_store, g_default_documents, g_delete_between_tags, g_delete_between_tags_complete, g_max_pages

  end if

  extractor.Destructor
  set extractor = Nothing
  g_set = g_set - 1
  
  
  if InStr(g_database_connection, "JET") <> 0 then
  
    if err.number = -2147467259 or err.number = -2147217911 then
    'write permissions disabled
      response.Write vbcrlf & "<script language=""JavaScript"">"
      response.Write vbcrlf & "alert('Error - The Yider cannot access the Access database.\n\nTry giving the web server write permissions to use the Yider\nMake sure you don\'t have the Access database file open as well.');"
      response.Write vbcrlf & "</script>"
    elseif err.number = 3704 or err.number = 424 or err.number = 70 then
    'update permissions disabled
      response.Write vbcrlf & "<script language=""JavaScript"">"
      response.Write vbcrlf & "alert('Error - The Yider cannot access the Access database.\n\nTry giving the web server modify permissions to use the Yider\nMake sure you don\'t have the Access database file open as well.');"
      response.Write vbcrlf & "</script>"
    elseif err.number <> 0 then
      response.write vbcrlf & "<script language=""JavaScript"">"
      response.write vbcrlf & "alert('The Yider has just encountered an error.\nError number: " & RemoveCarriageReturn(err.number) & "\nError Description: " & RemoveCarriageReturn(err.description) & "\nAn error report has been sent to peter.surna@yart.com.au but it would be nice if you could contact him direct to sort this issue out.');"
      response.write vbcrlf & "</script>"
      SendErrorReport "population"
    end if
  
  else
  
    if err.number = 3704 then
      response.Write vbcrlf & "<script language=""JavaScript"">"
      response.Write vbcrlf & "alert('Error - cannot connect to the SQL Server database.\nPlease check the database name, username and password are correct.');"
      response.Write vbcrlf & "</script>"

      if g_your_email_address = "yider.user@somewhere.com" then
        response.Write vbcrlf & "<script language=""JavaScript"">"
        response.write "alert('Please complete the variable g_your_email_address in configuration.asp so this error can be sent to Yart and addressed.\n\nThis is my only way to test the Yider over many platforms and in many situations so I\'d appreciate it if you help me out.\n\nI can\'t address this error if I can\'t contact you.');"
        response.Write vbcrlf & "</script>"
      end if
      
    elseif err.number <> 0 then
      response.write vbcrlf & "<script language=""JavaScript"">"
      response.write vbcrlf & "alert('The Yider has just encountered an error.\nError number: " & err.number & "\nError Description: " & RemoveCarriageReturn(err.description) & ".\nAn error report has been sent to administrator@transworld.dyndns.ws but it would be nice if you could contact him direct to sort this issue out.');"
      response.write vbcrlf & "</script>"
      SendErrorReport "population"

      if g_your_email_address = "yider.user@somewhere.com" then
        response.Write vbcrlf & "<script language=""JavaScript"">"
        response.write "alert('Please complete the variable g_your_email_address in configuration.asp so this error can be sent to Yart and addressed.\n\nThis is my only way to test the Yider over many platforms and in many situations so I\'d appreciate it if you help me out.\n\nI can\'t address this error if I can\'t contact you.');"
        response.write vbcrlf & "</script>"
      end if

    end if

  end if
    
  'response.Write "<br>" & err.Description
  

%>
		</form>
		</td></tr> </table> 
		<!--end of code you must include-->
	</body>
</html>

<%

if g_set <> 0 and InStr(g_url_to_spider, "127.0.0.1") <> 0 then
  response.Write vbcrlf & "<script language=""JavaScript"">"
  response.write "alert('The Yider has found an error. g_set is " & g_set & "');"
  response.write vbcrlf & "</script>"
  SendErrorReport "g_set"
end if

if g_open <> 0 and InStr(g_url_to_spider, "127.0.0.1") <> 0 then
  response.Write vbcrlf & "<script language=""JavaScript"">"
  response.write "alert('The Yider has found an error. g_open is " & g_open & "');"
  response.write vbcrlf & "</script>"
  SendErrorReport "g_open"
end if

%>