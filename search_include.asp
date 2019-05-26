<%
'Copyright (c) Tom Kirby 2005
'This program and its associated source code is distributed under the terms of the GNU General Public 
'License.
'See the attached file COPYING.txt for more information
%>
<%Server.ScriptTimeout = 2500%>
<%

  Dim yider_search
  
  on error resume next
  
  set yider_search = new CYiderSearch
  g_set = g_set + 1

  yider_search.m_charset = g_charset
  yider_search.Constructor g_database_connection

  yider_search.m_style_title = g_style_title
  yider_search.m_style_text = g_style_text
  yider_search.m_style_you_searched_for = g_style_you_searched_for
  yider_search.m_style_url = g_style_url
  yider_search.m_style_br = g_style_br
  yider_search.m_style_more_results_text = g_style_more_results_text
  yider_search.m_style_more_results_link = g_style_more_results_link
  yider_search.m_style_more_results_link_Next = g_style_more_results_link_Next
  yider_search.m_trailing_words = g_trailing_words

  yider_search.DisplayResults
  yider_search.Destructor
  set yider_search = Nothing
  g_set = g_set - 1
  
  if err.number <> 0 then
    response.write "<script language=""JavaScript"">"
    
    response.write vbcrlf & "alert('The Yider has just encountered an error.\nError number: " & RemoveCarriageReturn(err.number) & "\nError Description: " & RemoveCarriageReturn(err.description) & "\nAn error report has been sent to peter.surna@yart.com.au but it would be nice if you could contact him direct to sort this issue out.');"

    if g_your_email_address = "yider.user@somewhere.com" then
      response.write "alert('Please complete the variable g_your_email_address in configuration.asp so this error can be sent to Yart and addressed.\n\nThis is my only way to test the Yider over many platforms and in many situations so I\'d appreciate it if you help me out.\n\nI can\'t address this error if I can\'t contact you.');"
    end if
    
    response.write "</script>"
    SendErrorReport "search"
  end if

  showTime()

%>