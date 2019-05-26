<%Server.ScriptTimeout = 2500%>
<%
'Copyright (c) Tom Kirby 2005
'This program and its associated source code is distributed under the terms of the GNU General Public 
'License.
'See the attached file COPYING.txt for more information
%>

<!-- #include file="functions.asp" -->

<%

class CURLExtractor

  Public m_charset, m_database, m_english, m_update, m_XMLHttp, m_RegExpObject
  Public m_urls_not_to_view

  'username and password for authentication
  Public m_username, m_password
  
  Public m_strip_url_parameters

  'add the url to the database
  'set [parsed] to parsed
  'set [URLsize] to size_html if size_html is equal to -1
  private sub AddToDatabase(url, parsed, firstLocated)
  
    'response.write "<br>66url is " & url & " parsed is " & parsed
  
    url = AdjustName(url)
    firstLocated = replace(firstLocated, "'", "''")
    
    if DataBaseType = 0 then
      AddToDatabaseAccess url, parsed, firstLocated
    else
      AddToDatabaseSQL url, parsed, firstLocated
    end if
    
  end sub
  

  private sub AddToDatabaseAccess(url, parsed, firstLocated)
  
    Dim query, recordset
    
    CreateRecordset recordset

    recordset.Open "select [parsed], [url] from [Yider] where [url]='" & replace(url, "'", "''") & "'", m_database
    g_open = g_open + 1
    
    if recordset.eof then
      query = "insert into Yider ([url], [title], [text], [parsed], [URLSize], [firstLocated]) values ('" & replace(url, "'", "''") & "', '', '', 0, -1, '" & firstLocated & "')"
      m_database.Execute = query
      
    elseif recordset(1) <> url then
      query = "insert into Yider ([url], [title], [text], [parsed], [URLSize], [firstLocated]) values ('" & replace(url, "'", "''") & "', '', '', 0, -1, '" & firstLocated & "')"
      m_database.Execute = query
    else
    
      if recordset(0) = 0 or recordset(0) = 2 then
        query = "update [Yider] set [parsed]= " & parsed & " where [url]='" & replace(url, "'", "''") & "'"
        'response.Write "<br>" & query
        
        m_database.Execute query
      end if
    
    end if

    recordset.Close
    g_open = g_open - 1

    set recordset = Nothing
    g_set = g_set - 1
    
  end sub


  private sub AddToDatabaseSQL(byval url, parsed, firstLocated)
  
    Dim query, recordset
    
    CreateRecordset recordset
    url = replace(url, "'", "''")

    recordset.Open "select [parsed] from [Yider] where [url]='" & url & "'", m_database    
    g_open = g_open + 1
    
    'on error resume next
    
    if recordset.eof then
    
      query = "begin tran insert into [Yider] values ( N'" & url & "', '', '', " & parsed & ", -1, '" & firstLocated & "', DEFAULT) commit tran"
      m_database.Execute query

    else
      if recordset(0) = 0 or recordset(0) = 2 then
      
        query = "begin tran update [Yider] set [parsed]=" & parsed & " where [url]='" & url & "' commit tran"
        'response.Write "<br>" & query
        m_database.Execute query
        
      end if
    end if

    recordset.Close
    g_open = g_open - 1

    set recordset = Nothing
    g_set = g_set - 1
    
  end sub


  'if url is a directory, make sure it has a / at the end
  'see URLIsDirectory for more comments
  public function AdjustName(url)
      
    if URLIsDirectory(url) then
    
      if Right(url, 1) <> "/" then
        url = url & "/"
      end if
      
    end if
    
    AdjustName = url
    
  end function
  
  
  private function AdjustURL(url)
    Dim str
  
    if IsArray(m_strip_url_parameters) then
      for each str in m_strip_url_parameters

        'response.Write "<br><br>url is " & url
        url = StripParameterFromURL(url, str)
        
      next
    end if
  
    'response.Write " now it's " & url

    AdjustURL = replace(url, "'", "''")
    
  end function   
  
  
  'remove mailto:
  'remove javascript:
  'replace \ with /
  'remove / if it's the last character
  'remove everything after the # because it's an internal tag
  'if href is of the form href="javascript:newwin('xyz.asp')", return xyz.asp
  'if href is of the form javascript:newwin("xyz.asp"), return xyz.asp
  private function CleanUpHref(byval href)
    Dim pos, pos_question
    
    pos = InStr(href, "mailto:")
    
    if pos <> 0 then
      href = Left(href, pos - 1)
    end if
    
    href = replace(href, "\", "/")

    if GetFileExtension(href) = "" then
      pos_question =  InStr(href, "?")

      if Right(href, 1) <> "/" and pos_question = 0 then
      'required for url's like http://www.vailmountaineers.org/skate?doc&ID=index
        href = href & "/"
      end if
    end if
   
    pos = InStr(href, "#")
    
    if pos <> 0 then
      href = Left(href, pos - 1)
    end if
    
    pos = InStr(href, "javascript:")
    
    if pos <> 0 then
    
      Dim pos1, pos2, pos3
      pos1 = InStr(href, "newwin")
      pos2 = 0
      pos3 = 0
            
      if pos1 <> 0 then

        pos2 = InStr(pos1, href, "'")
        
        if pos2 = 0 then
          pos2 = InStr(pos1, href, """")
          
          if pos2 <> 0 then
            pos3 = InStr(pos2 + 1, href, """")
          else
            href = Left(href, pos - 1)
          end if
          
        else
          pos3 = InStr(pos2 + 1, href, "'")
        end if
        
        if pos2 <> 0 and pos3 <> 0 then
          href = Mid(href, pos2 + 1, pos3 - pos2 - 1)
        else
          href = Left(href, pos - 1)
        end if
        
      else
        href = Left(href, pos - 1)
      end if
    
    end if
    
    CleanUpHref = href
    
  end function


  public sub Constructor(update, byref database)
  
    on error resume next
    err.Clear
  
    set m_XMLHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
    
    if err.number <> 0 then
      err.number = 0
      set m_XMLHttp = Server.CreateObject("WinHttp.WinHttpRequest.5.1")

      if err.number <> 0 then
        set m_XMLHttp = Server.CreateObject("WinHttp.WinHttpRequest.5")
      else
        g_set = g_set + 1
      end if

    else
      g_set = g_set + 1
    end if
    
    
    set m_RegExpObject = New RegExp
    g_set = g_set + 1

    m_english = true
    m_charset = ""
    m_update = update
    
    set m_database = database
    g_set = g_set + 1
        
  end sub
  

  public sub Destructor
  
    set m_RegExpObject = Nothing
    g_set = g_set - 1

    set m_XMLHttp = Nothing
    g_set = g_set - 1

    set m_database = Nothing
    g_set = g_set - 1
  end sub
  

  private function DomainValid(href, valid_url_strings, default_documents)
    Dim valid
    
    valid = DomainValid1(href, valid_url_strings)
    
    if valid then
      valid = DomainValid2(href, default_documents)
    end if
      
    DomainValid = valid
     
  end function
  
  
  private function DomainValid1(href, valid_url_strings)
    Dim valid
    
    valid = false
        
    if IsArray(valid_url_strings) then
    
      if InStr(href, "http://") <> 0 or InStr(href, "https://") <> 0 then 
      
        valid = InArrayRegExp(href, valid_url_strings(0), valid_url_strings(1), m_RegExpObject)
        
      end if
      
    else
        
      valid =true
      
    end if
    
    'response.write "<br>valid is " & valid & " href is " & href & " valid_url_strings(0) is " & valid_url_strings(0)
    
    DomainValid1 = valid
    
  end function
  
  
  private function DomainValid2(byval href, valid_url_strings)
    Dim base_url, page, query, query_url, recordset, valid, url
    
    valid = true
    base_url = GetBaseURL(href) & "/"
  
    'response.Write "<br>href is " & href
    
    if IsArray(valid_url_strings) then
      if not URLIsDirectory(href) then
      
        page = FileName(href)
      
        if InArrayStr(page, valid_url_strings) then
          query  = "select [key] from [Yider] where [url]='" & replace(base_url, "'", "''") & "'"
          
          set recordset = m_database.Execute(query)
          g_set = g_set + 1
          g_open = g_open + 1
          
          if not recordset.eof then
            valid = false
          end if
          
          recordset.Close
          g_open = g_open - 1
          set recordset = Nothing
          g_set = g_set - 1
          
        end if
        
      else
        query = "select [key] from [Yider] where "
      
        for each url in valid_url_strings
          if Len(query_url) > 0 then
            query_url = query_url & " or "
          end if
          
          query_url = query_url & " url = '" & replace(base_url & url, "'", "''") & "'"
        next
          
        query = "select [key] from [Yider] where " & query_url
        
        set recordset = m_database.Execute(query)
        g_set = g_set + 1
        g_open = g_open + 1
          
        if not recordset.eof then
          valid = false
        end if
        
        recordset.Close
        g_open = g_open - 1
        set recordset = Nothing
        g_set = g_set - 1
        
      end if
    end if
    
    'if not valid then 
    '  response.Write "<br>not valid " & href
    'end if
    
    DomainValid2 = valid
  end function

  
  function EscapeURL(url)
    
    EscapeURL = replace(Trim(url), " ", "%20")

  end function


  'arr(0) - the html at start_url
  'arr(1) - false if the URL needs to be parsed, true if it doesn't
  public function ExtractHREFsFromURL(start_url, valid_file_extensions, valid_url_strings, default_documents, delete_between_tags_complete)
    Dim arr, url_array, html
    
    start_url = AdjustName(start_url)
        
    'response.Write "<br><br>arr(0) is " & start_url

    if InStr(start_url, "http://") = 0 and InStr(start_url, "https://") = 0 then
      Response.Write "<br><br>The url you supplied must contain a ""http://"" e.g. http://www.yart.com.au"
      Response.End
    else
      arr = GetURLsDirect(start_url, valid_file_extensions, valid_url_strings, default_documents, delete_between_tags_complete)      
    end if
    
    
    ExtractHREFsFromURL = arr
 
  end function
  
  
  'http://www.yart.com.au returns http://www.yart.com.au
  'http://www.yart.com.au/a/b/c returns http://www.yart.com.au
  'url must contain http or https
  'url must not contain \
  private function GetBaseDomain(url)
    Dim base_domain, pos
    
    pos = InStr(9, url, "/")
    
    if pos = 0 then
      base_domain = url
    else
      base_domain = Left(url, pos - 1)
    end if
    
    GetBaseDomain = base_domain
    
  end function


  'eg GetBaseURL("http://www.yart.com.au") returns http://www.yart.com.au
  'eg GetBaseURL("http://www.yart.com.au/") returns http://www.yart.com.au
  'eg GetBaseURL("http://www.yart.com.au/page10.htm") returns http://www.yart.com.au
  'eg GetBaseURL("http://www.yart.com.au/page10") returns http://www.yart.com.au/page10
  'eg GetBaseURL("http://www.yart.com.au/page10/") returns http://www.yart.com.au/page10
  'url must contain "http://" at least
  private function GetBaseURL(url)
  
    Dim base_url
        
    url = replace(url, "\", "/")
   
    if InStr(8, url, "/") <> 0 then
      base_url = GetBaseURLWithSlashes(url)
    else
      base_url = url
    end if

    GetBaseURL = base_url
    
  end function

  
  'eg GetBaseURLWithSlashes("http://www.yart.com.au/") returns http://www.yart.com.au
  'eg GetBaseURLWithSlashes("http://www.yart.com.au/page10.htm") returns http://www.yart.com.au
  'eg GetBaseURLWithSlashes("http://www.yart.com.au/page10") returns http://www.yart.com.au/page10
  'eg GetBaseURLWithSlashes("http://www.yart.com.au/page10/") returns http://www.yart.com.au/page10
  'url must contain a '/' after 'http://'
  private function GetBaseURLWithSlashes(url)
  
    Dim base_url, position_dot, position_last_slash
            
    if Mid(url, Len(url)) = "/" then
      base_url = Mid(url, 1, Len(url) - 1)
      
    else
    'if there is a dot after the last slash, this is a specific file otherwise it's a base url
      position_last_slash = InStrRev(url, "/")
      position_dot = InStr(position_last_slash, url, ".")
      
      if position_dot > position_last_slash then
        base_url = Mid(url, 1, position_last_slash - 1)
      else
        base_url = url
      end if
    
    end if
    
    GetBaseURLWithSlashes = base_url
    
  end function
  
  
  'url - http://www.yart.com.au/
  'url - http://www.yart.com.au
  'url - http://www.yart.com.au/index
  'url - http://www.yart.com.au/index.asp
  'the above should return http://www.yart.com.au/
  private function GetDirectory(url)
    Dim pos_slash, pos_dot
        
    url = replace(url, "\", "/")
    
    pos_slash = InStrRev(url, "/")
    pos_dot = InStrRev(url, ".")
    
    if pos_dot > pos_slash then
    'this could be a filename
    'check the slashes aren't the // after http://
      if InStrRev(url, "//") <> pos_slash - 1 then
        url = Left(url, pos_slash)
      else
        if Right(url, 1) <> "/" then
          url = url & "/"
        end if
      end if
    else
    'this is a directory
      if Right(url, 1) <> "/" then
        url = url & "/"
      end if
    end if
    
    GetDirectory = url
    
  end function
  
  
  'href = www.yart.com.au/index.html returns html
  'href = www.yart.com.au/articles returns ""
  'href = www.yart.com.au/articles/ returns ""
  private function GetFileExtension(href)
    Dim extension, pos_dot, pos_last_slash, pos_question, pos_slash
    
    extension = ""
    pos_question = InStr(href, "?")

    if pos_question <> 0 then
    'the url contains a ?
    
      pos_last_slash =  InStrRev(href, "/")
      pos_dot =  InStrRev(href, ".", pos_question)
      
      if pos_dot > pos_last_slash then
      'this will not occur for url's like http://www.vailmountaineers.org/skate?doc&ID=index
        extension = Mid(href, pos_dot + 1, pos_question - pos_dot - 1)
      end if
      
    else

      pos_dot = InStrRev(href, ".")
      href = replace(href, "\", "/")
      pos_slash = InStrRev(href, "/")
      
      if pos_dot <> 0 and (pos_dot > pos_slash) then
        extension = Mid(href, pos_dot + 1)
      end if
        
    end if
    
    'response.Write "<br>href is " & href & " extension is " & extension
    
    GetFileExtension = extension
  end function
  
  
  'URL is a www address
  'href is href found at that address
  'this function returns the new fully qualified URL the href refers to
  'response.Write "<br><br>1 it is " & url.GetFullyQualifiedURL("http://127.0.0.1/yidertest", "index.asp") & " should be http://127.0.0.1/yidertest/index.asp"
  'response.Write "<br><br>2 it is " & url.GetFullyQualifiedURL("http://127.0.0.1/yidertest/articles/jack.htm", "/index.asp") & " should be http://127.0.0.1/index.asp"
  'response.Write "<br><br>3 it is " & url.GetFullyQualifiedURL("http://127.0.0.1/yidertest/articles", "../index.asp") & " should be http://127.0.0.1/yiderTest/index.asp"
  'response.Write "<br><br>4 it is " & url.GetFullyQualifiedURL("https://127.0.0.1/yidertest", "../index.asp") & " should be https://127.0.0.1/index.asp"
  'response.Write "<br><br>5 it is " & url.GetFullyQualifiedURL("https://127.0.0.1/yidertest/", "../index.asp") & " should be https://127.0.0.1/index.asp"
  'response.Write "<br><br>6 it is " & url.GetFullyQualifiedURL("http://127.0.0.1/yidertest/articles/", "index.asp") & " should be http://127.0.0.1/yiderTest/articles/index.asp"
  'response.Write "<br><br>7 it is " & url.GetFullyQualifiedURL("http://127.0.0.1/yidertest/articles/", "/index.asp") & " should be http://127.0.0.1/index.asp"
  'response.Write "<br><br>8 it is " & url.GetFullyQualifiedURL("https://127.0.0.1/yidertest/articles/test.asp", "/stories/") & " should be https://127.0.0.1/stories/"
  'response.Write "<br><br>9 it is " & url.GetFullyQualifiedURL("http://127.0.0.1/yidertest/articles/test.asp", "#comment") & " should be http://127.0.0.1/yidertest/articles/test.asp"
  'response.Write "<br><br>10 it is " & url.GetFullyQualifiedURL("http://127.0.0.1/yidertest/articles/test.asp", "abc.html#comment") & " should be http://127.0.0.1/yidertest/articles/abc.html"
  'response.Write "<br><br>11 it is " & url.GetFullyQualifiedURL("http://127.0.0.1/yidertest", "/stories/") & " should be http://127.0.0.1/stories/"
  'response.Write "<br><br>12 it is " & url.GetFullyQualifiedURL("https://127.0.0.1/yidertest/", "/articles/stories/test.asp") & " should be https://127.0.0.1/articles/stories/test.asp"
  'response.Write "<br><br>13 it is " & url.GetFullyQualifiedURL("https://127.0.0.1/yidertest/", "/stories/test.asp") & " should be https://127.0.0.1/stories/test.asp"
  'response.Write "<br><br>14 it is " & url.GetFullyQualifiedURL("https://127.0.0.1/yidertest", "/articles/stories/test.asp") & " should be https://127.0.0.1/articles/stories/test.asp"
  'response.Write "<br><br>15 it is " & url.GetFullyQualifiedURL("https://127.0.0.1/yidertest", "javascript:newWIn('xyz.asp', 20, 30)") & " should be https://127.0.0.1/yidertest/xyz.asp"
  'response.Write "<br><br>16 it is " & url.GetFullyQualifiedURL("https://127.0.0.1/yidertest", "javascript:newWIn(""xyz.asp"", 20, 30)") & " should be https://127.0.0.1/yidertest/xyz.asp"
  'response.Write "<br><br>17 it is " & url.GetFullyQualifiedURL("https://127.0.0.1/yidertest", "javascript:test();") & " should be https://127.0.0.1/yidertest"
  'response.Write "<br><br>18 it is " & url.GetFullyQualifiedURL("https://127.0.0.1/yidertest", "harry.asp#javascript:test();") & " should be https://127.0.0.1/yidertest/harry.asp"
  'response.Write "<br><br>19 it is " & url.GetFullyQualifiedURL("https://127.0.0.1/yidertest", "../../../../index") & " should be https://127.0.0.1/index/"
  'response.Write "<br><br>20 it is " & url.GetFullyQualifiedURL("https://127.0.0.1", "../index") & " should be https://127.0.0.1/index/"
  'response.Write "<br><br>21 it is " & url.GetFullyQualifiedURL("https://127.0.0.1/a/b/c/d", "../../index.asp") & " should be https://127.0.0.1/a/b/index.asp"
  'response.Write "<br><br>22 it is " & url.GetFullyQualifiedURL("https://127.0.0.1/a/b/c/d", "../../../../../index/index.asp") & " should be https://127.0.0.1/index/index.asp"
  'response.Write "<br><br>23 it is " & url.GetFullyQualifiedURL("https://127.0.0.1/a/b/harry.htm", "../index/index.asp") & " should be https://127.0.0.1/a/index/index.asp"
  public function GetFullyQualifiedURL(url, href)
    Dim char_left, fq_url, new_url, pos, pos_dot, pos_second_slash, pos_hash
    
    'response.Write "<br>url is " & url & " href is " & href
        
    href = CleanUpHref(href)
    url = replace(url, "\", "/")
    
    if InStr(href, "http://") = 0 and InStr(href, "https://") = 0 then
    'the href doesn't contain http:// or https://
    
      if Left(href, 1) = "/" then
        new_url = GetBaseDomain(url)
        fq_url = new_url & href        
      else
      
        if Len(href) > 0 then
          new_url = GetDirectory(url)
          fq_url = new_url & href        
        else
          fq_url = url
        end if
        
      end if
        
    else
      fq_url = href
    end if
        
    'response.Write "<br>fq_url is " & fq_url 
    'response.write "<br>1 url is " & url & " fq_url is " & fq_url & " href is " & href
    
    fq_url = RemoveDotReferencing(fq_url)
    'response.Write "<br>1 fq_url is " & fq_url 
    fq_url = RemoveDotDotReferencing(fq_url)
    'response.Write "<br>2 fq_url is " & fq_url & g_set 
        
    GetFullyQualifiedURL = fq_url
      
  end function

  
  'html' is a text string of a plain old html page
  ''valid_file_extensions' is an array of acceptable file extensions in hrefs in this page that we will
  'consider extracting eg Array("htm", "html", "asp", "php", "php3")
  ''url' is the fully qualified name of the page where 'html' has beed extracted
  'from eg http://www.somewhere.com/page.htm
  'this function builds pairs of values in 'urls_examined'
  
  'the first value in the pair will be 'url'
  'the second value in the pair is a fully qualified url that was extracted from within a href tag
  'within the 'html' provided it contains at least one of the strings in the string array 'valid_url_strings'
  'tag_attribute - must be lower case
  private sub GetHREFs(url, html, html_lcase, valid_file_extensions, valid_url_strings, default_documents)
    Dim expressionmatch, expressionmatched, file_extension, fq_url, href, in_str, original_url, recordset, valid
    
    CreateRecordset recordset
    
    'response.Write server.HTMLEncode(html)
   
    'required for bug in vbscript
    'see below
    original_url = url
    
    With m_RegExpObject
      .Pattern = "href\s*=\s*""[^""]*""|href\s*=\s*'[^']*'|src\s*=\s*""[^""]*""|src\s*=\s*'[^']*'"
      .IgnoreCase = true
      .Global = True
    End With

    set expressionmatch = m_RegExpObject.Execute(html)
                     
    for each expressionmatched in expressionmatch
    
      href = GetURLFromHref(expressionmatched.value)
            
      file_extension = GetFileExtension(href)
  
      'response.write "<br><br>1 fq_url is " & fq_url & " url is '" & url & "' href is '" & href & "' file_extension is " & file_extension
      
      if ValidFileExtension(file_extension, valid_file_extensions) then
            
      'response.write "<br>wow"

        'bug in vbscript?
        'url get's reassigned in this statement for no reason I can tell!   
        fq_url = GetFullyQualifiedURL(url, href)
        url = original_url
        
        fq_url = AdjustURL(fq_url)

        valid = DomainValid(fq_url, valid_url_strings, default_documents)
        'in_str = InStrArray(fq_url, m_urls_not_to_view)
        'response.Write "<br>valid is " & valid
        
        if IsArray(m_urls_not_to_view) then
          in_str = InArrayRegExp(fq_url, m_urls_not_to_view(0), m_urls_not_to_view(1), m_RegExpObject)
        else
          in_str = false
        end if


        if valid and not in_str then
          set recordset = m_database.Execute("select [key], [url] from [Yider] where [url]='" & replace(fq_url, "'", "''") & "'")
          g_open = g_open + 1
          g_set = g_set + 1
                              
          'response.write "<br> fq_url is " & fq_url & " url is " & url & " recordset.eof is " & recordset.eof

          if recordset.eof then
            AddToDatabase fq_url, 0, url
          elseif recordset(1) <> fq_url then
            AddToDatabase fq_url, 0, url
          end if
          
          recordset.Close
          g_open = g_open - 1

          set recordset = Nothing
          g_set = g_set - 1

        end if
        
      end if
      
    next
    
    set recordset = Nothing
    g_set = g_set - 1
            
  end sub
  
  
  sub GetRedirects(url, html, valid_file_extensions, valid_url_strings, default_documents)
    Dim fq_url, in_str, recordset, valid, x
    
    CreateRecordset recordset
    
    fq_url = GetRedirect(html)
    
    'response.write "<br>1 it is " & fq_url
    fq_url = GetFullyQualifiedURL(url, fq_url)
    'response.write "<br>2 it is " & x
    
    fq_url = AdjustURL(fq_url)

    valid = DomainValid(fq_url, valid_url_strings, default_documents)

    if IsArray(m_urls_not_to_view) then
      in_str = InArrayRegExp(fq_url, m_urls_not_to_view(0), m_urls_not_to_view(1), m_RegExpObject)
    else
      in_str = false
    end if

    if valid and not in_str then
      set recordset = m_database.Execute("select [key] from [Yider] where [url]='" & replace(fq_url, "'", "''") & "'")
      g_open = g_open + 1
      g_set = g_set + 1
                          
      if recordset.eof then
        AddToDatabase fq_url, 0, url
      end if
      
      recordset.Close
      g_open = g_open - 1

      set recordset = Nothing
      g_set = g_set - 1

    end if
    
    set recordset = Nothing
    g_set = g_set - 1

  end sub
  
  
  'eg href = 'href="abc.asp"'
  'eg href = "href='abc.asp'"
  function GetURLFromHref(href)
    Dim pos, url
    
    pos = InStr(href, "=")
    
    if pos <> 0 then
      url = Mid(href, pos + 1)
      
      url = Trim(url)
      url = Mid(url, 2, Len(url) - 2)
      
    end if
    
    GetURLFromHref = url
    
  end function

  
  ''url' is the fully qualified name of the page where hrefs are to be extracted
  'from eg http://www.somewhere.com/page.htm
  ''valid_file_extensions' is an array of acceptable file extensions in hrefs in this page that we will
  'consider extracting eg Array("htm", "html", "asp", "php", "php3"), Array("esolutions")
  
  'urls_examined is an Array of fully qualified url pairs
  'the first value in the pair is the value of an url that has to be openend and examined for hrefs
  'the second value is true if it has been examined and false if it hasn't  
  public function GetURLsDirect(url, valid_file_extensions, valid_url_strings, default_documents, delete_between_tags_complete)
   
    Dim count, html, html_lcase, requires_parsing, size_html
     
    Err.Clear              
    on error resume next
    'this is necessary because URLs like 
    'http://www.yart.com.au?x=a&#39;e or
    'http://www.yart.com.au?x=a'e
    'will throw an exception here
    m_XMLHttp.open "GET", EscapeURL(url), false, m_username, m_password
    m_XMLHttp.send()
        
    if err.number = -2147012744 then
    'this is caused by URL's with no DNS entry
    'err.Description = The server returned an invalid or unrecognised response 
      err.Clear
    end if

    if err.number <> 0 then
    'this is to catch url's with &20 in them
    'one user said mixed case URLs like NTShutdown.htm cause a crash here
      query = "update [Yider] set [parsed]= 1, [URLsize] = -2 where [url]='" & replace(url, "'", "''") & "'"
      m_database.Execute query
    end if

    'response.Write "<br>m_XMLHttp.Status is " & m_XMLHttp.Status
    'response.Write "<br>ValidContentType(m_XMLHttp) is " & ValidContentType(m_XMLHttp)
    
    if m_XMLHttp.Status = 200 and ValidContentType(m_XMLHttp) then
    
      if not m_english then
        html = m_XMLHttp.ResponseBody
        
        html = BinaryToString1(html, m_charset)
        
        'response.Write "m_charset is " & m_charset
        'response.Write "<br><br>html is " & server.HTMLEncode(html)
      else
        'response.Write "<br>url is " & url
        on error resume next
        html = m_XMLHttp.ResponseText
        
        if err.number = -1072896658 then
        'a xmlhttp error caused by getting http://www.dpreview.com/shop/product.asp?affiliate=1&id=banner2
        'why? - I don't know
          err.Clear
          html = ""
        end if
      end if
      
      'response.Write html
      
      size_html = Len(html)
      
      'has the URL been parsed before and its size changed?
      requires_parsing = URLRequiresParsing(url, size_html)
      UpdateURLSize url, size_html
      
      'delete text between the following tags for both spidering and searching
      if IsArray(delete_between_tags_complete) then
        for count = 0 to UBound(delete_between_tags_complete) step 2
          html = ReplaceBetweenStrings(html, delete_between_tags_complete(count), delete_between_tags_complete(count + 1), "")
        next
      end if
      
      'remove text between JavaScript tags
      html = ReplaceBetweenStrings(html, "<script", "/script>", "")
      
      if requires_parsing then
        html_lcase = lcase(html)
        
        GetHREFs url, html, html_lcase, valid_file_extensions, valid_url_strings, default_documents
        
      end if
      
      GetRedirects url, html, valid_file_extensions, valid_url_strings, default_documents

    else
      html = ""
    end if
    
    'response.Write "<br>url is " & url
            
    AddToDatabase url, 1, url
    
    GetURLsDirect = Array(html, requires_parsing)
          
  end function
  
  
  'http://www.yart.com.au - true
  'http://www.yart.com.au/ - true
  'http://www.yart.com.au/articles - false
  'url must not contain \
  private function IsBaseURL(url)
  
    Dim count, i, is_base
    count = 0
  
    is_base = true

    url  = replace(url, "http://", "")
    url  = replace(url, "https://", "")
    
    for i = 1 to Len(url)
     
     if Mid(url, i, 1) = "/" then
       count = count + 1
     end if
     
     if count > 1 then
       is_base = false
       Exit For
     end if
     
    next
    
    IsBaseURL = is_base
    
  end function
  
  
  private function GetRedirect(html)

    Dim expressionmatch, expressionmatched, pos, redirect, RegExpObject
    
    redirect = ""

    set RegExpObject = New RegExp

    With RegExpObject
      
      .Pattern = "<\s*meta\s*http-equiv=""refresh""\s*content\s*=\s*""3\s*;\s*url=\s*"
      .IgnoreCase = true
      .Global = True
      
    End With

    set expressionmatch = RegExpObject.Execute(html)
          
    for each expressionmatched in expressionmatch
    'there will only be one match

      pos = InStr(expressionmatched.length + expressionmatched.firstIndex + 1, html, """")
      redirect = Trim(Mid(html, expressionmatched.length + expressionmatched.firstIndex + 1, pos - (expressionmatched.length + expressionmatched.firstIndex + 1)))

    next
    
    GetRedirect = redirect

  end function

    
  
  'some urls are of the form http://www.yart.com.au/./index.asp
  'this function returns http://www.yart.com.au/index.asp
  'url must contain 'http' or 'https'
  'url must never contain a \
  private function RemoveDotReferencing(url)
  
		Dim keep_going, removed, removed_old
		
		keep_going = true
		removed_old = url
		removed = url

		while(keep_going)
		  
			removed = replace(removed, "/./", "/")
			
			if Len(removed_old) = Len(removed) then
			  keep_going = false
			else
			  removed_old = removed
			end if
			
			'response.Write "<br><br>1 removed_old is " & removed_old
			'response.Write "<br>2 removed is " & removed
			'response.Write "<br>3 keep_going is " & keep_going
		
		wend

    RemoveDotReferencing = removed
  end function


  'some urls are of the form http://www.yart.com.au/articles/../index.asp
  'or http://www.yart.com.au/stuff/../articles/../index.asp
  'the first url is exactly equivalent to http://www.yart.com.au/index.asp
  'the second url is exactly equivalent to http://www.yart.com.au/index.asp
  'this function returns the equivalent url
  'url must contain 'http' or 'https'
  'url must never contain a \
  private function RemoveDotDotReferencing(url)
    Dim pos_dotdot, pos_first_slash, pos_replace, pos_slash, start, str_left, str_right, url_fixed
    
    if InStr(url, "https:") <> 0 then
      pos_replace = 9
    else
      pos_replace = 8
    end if
    
    start = 1
    
    url = replace(url, "//", "/", pos_replace)
    
    if pos_replace = 8 then
      url = "http://" & url
    else
      url = "https://" & url
    end if
    
    pos_dotdot = InStr(start, url, "/..")
    pos_first_slash = InStr(9, url, "/")
        
    while pos_dotdot <> 0
    
      if pos_dotdot = pos_first_slash then
        url = Left(url, pos_first_slash) & Mid(url, InStrRev(url, "../") + 3)
        pos_dotdot = 0
        
      else
      
        pos_slash = InStrRev(url, "/", pos_dotdot - 1)

        str_left = Left(url, pos_slash - 1)
        str_right = Right(url, Len(url) - pos_dotdot - 2)
        url = str_left + str_right
        
        start = pos_slash
        pos_dotdot = InStr(start, url, "/..")

      end if

      'response.Write "<br>url is " & url & " pos_dotdot is  " & InStr(start, url, "/..")

    wend
      
    RemoveDotDotReferencing = url
    
  end function
  

  'StripParameterFromURL("http://www.website.com", "wow") returns http://www.website.com
  'StripParameterFromURL("http://www.website.com?wow=1", "wow") returns http://www.website.com
  'StripParameterFromURL("http://www.website.com?x=1&wow=1", "wow") returns http://www.website.com?x=1
  'StripParameterFromURL("http://www.website.com?x=1&wow=1&z=2", "wow") returns http://www.website.com?x=1&z=2
  'StripParameterFromURL("/tree2/tree3/tree4/tree5/tree5.htm?y=2&z=55&x=1", "x") returns /tree2/tree3/tree4/tree5/tree5.htm?y=2&z=55
  'StripParameterFromURL("/tree2/tree3/tree4/tree5/tree5.htm?y=2&z=55&x=1", "y") returns /tree2/tree3/tree4/tree5/tree5.htm?z=55
  private function StripParameterFromURL(url, str)
    Dim pos, pos1, pos2, left_url, right_url
    
    pos = InStr(url, "?")
    
    if pos <> 0 then
    
      pos1 = InStr(pos + 1, url, str)
      
      if pos1 <> 0 then
      
        pos2 = InStr(pos1 + 1, url, "&")
        
        if pos2 = 0  then
          url = Left(url, pos1 - 2)
        else
        
        'response.write "<br>pos1 is " & pos1
        'response.write "<br>pos2 is " & pos2
        
          left_url = Left(url, pos1 - 1)
          right_url = Mid(url, pos2)
          url = left_url & right_url
        end if
        
      end if
    
    end if
    
    url = replace(url, "?&", "?")
    url = replace(url, "&&", "&")
    
    StripParameterFromURL = url
      
  end function


  private sub UpdateURLSize(byval url, size_html)
    Dim recordset
      
    CreateRecordset recordset
    
    'response.Write "<br>it is "
    'response.Write URLIsDirectory(url)    
    
    url = replace(AdjustName(url), "'", "''")
    
    recordset.Open "select [key] from [Yider] where [url]='" & url & "'", m_database
    g_open = g_open + 1
    
    if not recordset.eof then
      m_database.Execute("update [Yider] set [URLsize] = " & size_html & " where [key] = " & recordset(0))
    else
    
      if DataBaseType = 0 then
      'Access
        m_database.Execute "insert into Yider ([url], [title], [text], [parsed], [URLSize], [firstLocated]) values ('" & url & "', '', '', 0, " & size_html & ", '')"
        'response.write "<br>1 it is " & url

      else
      'SQL Server
        m_database.Execute "begin tran insert into [Yider] values ( N'" & url & "', '', '', 0, " & size_html & ", '', DEFAULT) commit tran"
      end if

    end if
      
    recordset.Close
    g_open = g_open - 1

    set recordset = Nothing
    g_set = g_set - 1

  end sub

  
  
  'if the URL has been parsed and it's [URLSize] is not equal to 'size', return true
  'otherwise return false
  private function URLRequiresParsing(url, size)
    Dim requires_parsing, query, recordset
    
    CreateRecordset recordset
    
    query = "select [URLsize], [parsed] from [Yider] where [url]='" & replace(url, "'", "''") & "'"
    recordset.Open query, m_database
    g_open = g_open + 1
    
    if not recordset.eof then
    
      if recordset(1) = 0 then
        requires_parsing = true
      elseif recordset(1) = 1 then
        requires_parsing = false
      elseif recordset(1) = 2 then
      
        if recordset(0) <> size then
          requires_parsing = true
        else
          requires_parsing = false
        end if
      
      else
        
      end if
        
      'response.write "<br><br>url is " & url & " requires_parsing is <b>" & requires_parsing & "</b> recordset(0) is " & recordset(0) & " size is " & size
    
    else    
      requires_parsing = true
    end if
    
    recordset.Close
    g_open = g_open - 1

    
    set recordset = Nothing
    g_set = g_set - 1
    
    URLRequiresParsing = requires_parsing
    
  end function
  
  
  'response.Write "<br>1 is " & URLIsDirectory("http://www.yart.com.au/a?") & " true"
  'response.Write "<br>2 is " & URLIsDirectory("http://www.yart.com.au/") & " true"
  'response.Write "<br>2 is " & URLIsDirectory("http://www.yart.com.au/?") & " true"
  'response.Write "<br>2 is " & URLIsDirectory("http://www.yart.com.au?") & " true"
  'response.Write "<br>2 is " & URLIsDirectory("http://www.yart.com.au?a.asp") & " false"
  'response.Write "<br>3 is " & URLIsDirectory("http://www.yart.com.au/a.asp") & " false"
  'response.Write "<br>4 is " & URLIsDirectory("http://www.yart.com.au/a.asp?z=6") & " false"
  'response.Write "<br>5 is " & URLIsDirectory("http://www.vailmountaineers.org/skate?ann&ID=index") & " true"
  'response.Write "<br>6 is " & URLIsDirectory("http://www.vailmountaineers.org/skate.asp?ann&ID=index") & " false"
  'response.Write "<br>7 is " & URLIsDirectory("http://www.esoluions.com.au/articles/") & " true"
  'response.Write "<br>8 is " & URLIsDirectory("http://www.esoluions.com.au/articles/more") & " true"
  private function URLIsDirectory(byval url)

    Dim is_directory, pos_dot, pos_question, pos_slash
    
    url = replace(url, "http://", "")
    url = replace(url, "https://", "")
    
    pos_dot = InStrRev(url, ".")
    pos_slash = InStrRev(url, "/")
    pos_question = InStrRev(url, "?")
    
    if pos_slash <> 0 and pos_question <> 0 then
      if pos_dot < pos_slash and pos_slash < pos_question then
        is_directory = true
      elseif pos_slash < pos_dot and pos_dot < pos_question then
        is_directory = false
      end if
      
    elseif pos_slash <> 0 and pos_question = 0 then
      if pos_dot < pos_slash then 
        is_directory = true
      else
        is_directory = false
      end if

    elseif pos_slash = 0 and pos_question <> 0 then
      if pos_question < Len(url) then
        is_directory = false
      else
        is_directory = true
      end if

    elseif pos_slash = 0 and pos_question = 0 then
      is_directory = true

    end if
    
    URLIsDirectory = is_directory

  end function

  
  'true if the content in winHttp is of a valid form to parse for href tags
  'the Yider can not analyse binary files
  'text is acceptable
  private function ValidContentType(winHttp)
    Dim content_type, valid
    
    valid = false
    content_type  = winHttp.GetResponseHeader("content-type")
    
    'response.Write "<br>it is " & content_type
    
    'm_database.Execute "insert into test ([url]) values ('" & content_type & "')"

    if InStr(content_type, "text") <> 0 then
    'see http://www.asahi-net.or.jp/en/guide/cgi/mimetype.html for a list of MIME content types
      valid = true      
    end if
        
    ValidContentType = valid
    
  end function
  
  
  'returns true if the file_extension could be a valid file extension
  'by could, we mean that file_extension could be empty because it is a directory
  'this might be valid so we can't exclude it
  'alternatively, it must be in the array valid_file_extensions
  private function ValidFileExtension(byval file_extension, valid_file_extensions)
  
    Dim pos, valid
      
    pos = InStr(file_extension, "#")
    
    if pos <> 0 then
      file_extension = Left(file_extension, pos - 1)
    end if
    
    if Len(file_extension) = 0 then
      valid = true
    else
      'valid = InArrayStrExact(file_extension, valid_file_extensions)
      valid = InArrayRegExp(file_extension, valid_file_extensions(0), valid_file_extensions(1), m_RegExpObject)
      
      'response.write "<br><br>file_extension is " & file_extension
      'response.write "<br>valid_file_extensions(0) is " &  valid_file_extensions(0)
      'response.write "<br>valid_file_extensions(1) is " &  valid_file_extensions(1)
      'response.write "<br><br>valid is " & valid
    end if
    
    ValidFileExtension = valid
    
  end function
  

end class


%>
