<%Server.ScriptTimeout = 2500%>
<%
'Copyright (c) Tom Kirby 2005
'All rights reserved
'This source code is subject to the licensing conditions at http://www.transworld.dyndns.ws
%>

<!-- #include file="CURLExtractor.asp" -->


<%

class CHTMLTextExtractor

  Private m_count_invalid_text, m_count, m_database, m_heading_required, m_proceed, m_RegExpObject
  Private m_max_urls 'the maximum number of urls to ever parse
  
  'the connection string to the database
  Private m_database_connection
  
  'the maximum number of urls to parse per iteration
  Public m_urls_per_iteration 
  
  Public m_too_busy_text, m_wait
  
  'if true, compact the database when clearing it
  Public m_compact
  
  'true if you want existing URL's to be updated
  Public m_update
  
  'default is true, set to false for foreign languages
  Public m_english
  
  'required for full text searching
  Public m_local_ID
  
  'default is ""
  Public m_charset

  'username and password for authentication
  Public m_username, m_password
  
  Public m_strip_url_parameters
    
  public function AccessFileName()
    
    Dim filename, pos1, pos2, temp_connection
    
    pos2 = 0
    
    temp_connection = replace(lcase(m_database_connection), " ", "")
    
    pos1 = InStr(temp_connection, "source=")
    
    if pos1 <> 0 then
      pos2 =  InStr(pos1 + 7, temp_connection, ".mdb")
    end if
    
    if pos2 <> 0 then
      filename = Mid(temp_connection, pos1 + 7, (pos2 + 4) - (pos1 + 7))
    end if
    
    AccessFileName = filename
    
  end function
  

  'remove all data
  public sub Clear
  
    if DatabaseTableExists("Yider") then
      DeleteData "Yider"
    end if
    
    if DatabaseTableExists("YiderResult") then
      DeleteData "YiderResult"
    end if
      
    if m_compact then
      CompactDatabase
    end if
    
    
  end sub
  
  
  
  private sub CompactDatabase
  
    if DataBaseType = 0 then
      CompactDatabaseAccess
    else
      CompactDatabaseSQL
    end if
    
  end sub
  
  
  private sub CompactDatabaseAccess
    Dim access_file, engine, file_object, temp_file
    
    set file_object = Server.CreateObject("Scripting.FileSystemObject")
    g_set = g_set + 1
    
    if file_object.FileExists(GetPath() & "yider_temp_yider.mdb") then
      file_object.DeleteFile GetPath() & "yider_temp_yider.mdb"
    end if
    
    m_database.Close
    g_open = g_open - 1
    
    set engine = CreateObject("JRO.JetEngine")
    g_set = g_set + 1
    
    access_file = AccessFileName()
    temp_file = Left(access_file, Len(access_file) - 4) & "_temp_yider.mdb"
    
    engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & access_file, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & temp_file
    

    file_object.DeleteFile access_file
    file_object.MoveFile temp_file, access_file
    
    set file_object = Nothing
    g_set = g_set - 1

    set engine = Nothing
    g_set = g_set - 1
    
    m_database.Open m_database_connection
    g_open = g_open + 1

  end sub
  
  
  private sub CompactDatabaseSQL
    
    if DatabaseTableExists("Yider") then
      m_database.Execute("DBCC DBREINDEX('Yider')")
    end if

    if DatabaseTableExists("YiderResult") then
      m_database.Execute("DBCC DBREINDEX('YiderResult')")
    end if
    
  end sub

  
  public sub Constructor(database_connection)
  
    m_update = true
    m_proceed = true
    m_count_invalid_text = 0
    m_heading_required = true
    m_count = 0
    m_max_urls = 0
    m_wait = 30
    m_urls_per_iteration = 0
    m_compact = false
    m_database_connection = database_connection
    m_english = true
    m_charset = ""
    set m_RegExpObject = New RegExp
    g_set = g_set + 1
        
    set m_database = Server.CreateObject("ADODB.Connection")
    g_set = g_set + 1
    m_database.Open database_connection
    g_open = g_open + 1
    
  end sub

  
  private function DatabaseTableExists(table_name)
    Dim exists
    
    if DataBaseType = 0 then
      exists = DatabaseTableExistsAccess(table_name)
    else
      exists = DatabaseTableExistsSQL(table_name)
    end if
      
    DatabaseTableExists = exists
    
  end function
  
  
  private function DatabaseTableExistsAccess(table_name)
    Dim exists, recordset

    on error resume next '- this is now present in population.asp
    set recordset = m_database.Execute("select top 1 * from [" & table_name & "]")
    g_set = g_set + 1
    g_open = g_open + 1
    
    if err.number = -2147217865 then
      exists = false
    else
      exists = true
    end if
    
    recordset.Close
    set recordset = Nothing
    g_set = g_set - 1
    g_open = g_open - 1
    
    err.Clear
    DatabaseTableExistsAccess = exists
    
  end function


  private function DatabaseTableExistsSQL(table_name)
    Dim exists, recordset
    
    exists = false
    set recordset = m_database.Execute("select [id] from [sysobjects] where [name]='" & table_name & "'")
    g_set = g_set + 1
    g_open = g_open + 1
      
    if not recordset.eof then
      exists = true
    end if
      
    recordset.Close
    g_open = g_open - 1

    set recordset = Nothing
    g_set = g_set - 1
    
    DatabaseTableExistsSQL = exists
    
  end function


  sub DeleteData(table_name)
  
    if DataBaseType = 0 then
      m_database.Execute "delete * from [" & table_name & "]"
    else
      m_database.Execute "begin tran truncate table [" & table_name & "] commit tran"
    end if

  end sub


  public sub Destructor   

    set m_RegExpObject = Nothing
    g_set = g_set - 1
    
    m_database.Close
    g_open = g_open - 1

    set m_database = Nothing
    g_set = g_set - 1

  end sub
  
  
  private sub ExtractSearchableText(url, text, bad_page_strings, delete_between_tags)
  
    url = replace(url, Chr(0), "")
    text = replace(text, Chr(0), "")
      
    if DataBaseType = 0 then
      ExtractSearchableTextAccess url, text, bad_page_strings, delete_between_tags
    else
      ExtractSearchableTextSQL url, text, bad_page_strings, delete_between_tags
    end if
    
  end sub


  private sub ExtractSearchableTextAccess(url, text, bad_page_strings, delete_between_tags)

    Dim count, query, ret, title
    
    ret = GetTitleAndText(text, delete_between_tags)
    title = ret(0)
    text = ret(1)
    
    'response.Write "m_english is " & m_english 
    'on error resume next

   'response.Write "<br><br>url is " & url & " text is " & server.HTMLEncode(text)
    
    if IsArray(m_too_busy_text) then
    
      if InArrayStr(text, m_too_busy_text) then
        m_proceed = false
      end if
      
    end if
          
    
    if m_proceed then
    
      if IsArray(bad_page_strings) then
      
        'response.write "<br>text is " & text
      
        if not InArrayRegExp(text, bad_page_strings(0), bad_page_strings(1), m_RegExpObject) then
                  
          query = "update [Yider] set [title]=' " & Left(replace(title, "'", "''"), 255) & " ', [text]='" & replace(text, "'", "''") & "', [parsed]=1 where [url]='" & replace(url, "'", "''") & "'"
          'response.Write "<br>query is " & query
          m_database.Execute query
          
          PrintResults vbcrlf & "<br>" & url & " was parsed"
          'Response.Write "<br>text is " & text
          
        else
          m_database.Execute "update [Yider] set [parsed]=1 where [url]='" & replace(url, "'", "''") & "'"

          PrintResults "<br>*Invalid text: " & url & " was not added to the database because it contains invalid text"
          m_count_invalid_text = m_count_invalid_text + 1
        end if
        
      else
        m_database.Execute "update [Yider] set [title]='" & replace(title, "'", "''") & "', [text]='" & replace(text, "'", "''") & "', [parsed]=1 where [url]='" & replace(url, "'", "''") & "'"
        PrintResults "<br>" & url & " was parsed"
        
        'response.Write "<br><br>text is " & replace(text, "'", "''")
        'response.Write "<br><br>title is " & ASC(replace(title, "'", "''"))
      end if      
      
    else
      
      PrintResults "<br><br>Spidering was abandoned because " & url & " contained one of the following words/phrases:"
        
      for count = 0 to UBound(m_too_busy_text)
        response.write " '" & m_too_busy_text(count) & "'"
          
        if count <  UBound(m_too_busy_text) then
          response.write ","
        end if
      next
        
    end if
      
  end sub


  private sub ExtractSearchableTextSQL(url, text, bad_page_strings, delete_between_tags)

    Dim count, query, ret, title
        
    ret = GetTitleAndText(text, delete_between_tags)
    title = ret(0)
    text = ret(1)
    'on error resume next
    
    if IsArray(m_too_busy_text) then
    
      if InArrayStr(text, m_too_busy_text) then
        m_proceed = false
      end if
      
    end if
          
    
    if m_proceed then
    
      if IsArray(bad_page_strings) then
      
        if not InArrayStr(text, bad_page_strings) then
        
          m_database.Execute "update [Yider] set [title]=N' " & replace(title, "'", "''") & " ', [text]=N'" & replace(text, "'", "''") & "', [parsed]=1 where [url]='" & replace(url, "'", "''") & "'"
          
          PrintResults "<br>" & url & " was parsed"
          
        else
          m_database.Execute "update [Yider] set [parsed]=1 where [url]=N'" & replace(url, "'", "''") & "'"

          PrintResults "<br>*Invalid text: " & url & " was not added to the database because it contains invalid text"
          m_count_invalid_text = m_count_invalid_text + 1
        end if
        
      else
        m_database.Execute "update [Yider] set [title]=N'" & replace(title, "'", "''") & "', [text]=N'" & replace(text, "'", "''") & "', [parsed]=1 where [url]='" & replace(url, "'", "''") & "'"
        
        
        PrintResults "<br>" & url & " was parsed"
        
        'response.Write "<br><br>text is " & replace(text, "'", "''")
        'response.Write "<br><br>title is " & ASC(replace(title, "'", "''"))
      end if      
      
    else
      
      PrintResults "<br><br>Spidering was abandoned because " & url & " contained one of the following words/phrases:"
        
      for count = 0 to UBound(m_too_busy_text)
        response.write " '" & m_too_busy_text(count) & "'"
          
        if count <  UBound(m_too_busy_text) then
          response.write ","
        end if
      next
        
    end if
      
  end sub


  'a not very good way to determine whether a database has full text enabled but I'm not sure where this
  'information is stored
  private function FullTextEabled
    Dim enabled, recordset
    
    enabled = false
    
    CreateRecordset recordset
    
    recordset.Open "select ftcatid from sysfulltextcatalogs", m_database
    g_open = g_open + 1
    
    if not recordset.eof then
      enabled = true
    end if
    
    recordset.Close
    g_open = g_open - 1

    set recordset = Nothing
    g_set = g_set - 1
    
    FullTextEabled = enabled
  
  end function
  
  
  'e.g. GetAttribute("<META name=""keywords"" content=""ebay, electronics, cars"">", "content")
  'returns ebay, electronics, cars
  private function GetAttribute(tag, the_attribute)

    Dim expressionmatch, expressionmatched, RegExpObject, value

    set RegExpObject = New RegExp
    
    With RegExpObject
      .Pattern = the_attribute & "=""[^""]*"""
      .IgnoreCase = true
      .Global = True
    End With
             
    set expressionmatch = RegExpObject.Execute(tag)
         
    For Each expressionmatched in expressionmatch
      value = expressionmatched.value
      Exit For
    Next
    
    GetAttribute = GetValBetweenQuotes(value)
    
  end function  
  
  
  private function GetTagValue(html, tag)

    Dim expressionmatch, expressionmatched, RegExpObject, value

    set RegExpObject = New RegExp
      
    With RegExpObject
      .Pattern = "<[\s]*" & tag & "[\s]*[^>]*>[^>]*</" & tag & ">"
      .IgnoreCase = true
      .Global = True
    End With
             
    set expressionmatch = RegExpObject.Execute(html)
         
    For Each expressionmatched in expressionmatch
      value = GetInnerText(expressionmatched.value)
      Exit For
    Next

    GetTagValue = value
    
  end function


  private function GetInnerText(tag)
    Dim expressionmatch, expressionmatched, RegExpObject, value

    set RegExpObject = New RegExp
      
    With RegExpObject
      .Pattern = ">[^<]*</"
      .IgnoreCase = true
      .Global = True
    End With
             
    set expressionmatch = RegExpObject.Execute(tag)
         
    For Each expressionmatched in expressionmatch
      value = expressionmatched.value
      value = replace(value, ">", "")
      value = replace(value, "</", "")
      Exit For
    Next

    GetInnerText = value
    
  end function

  
  public sub GetTextThroughoutDatabase(valid_file_extensions, valid_url_strings, bad_page_strings, urls_not_to_view, urls_to_view_not_store, default_documents, delete_between_tags, delete_between_tags_complete, urlextractor)
  
    Dim arr, finished, html, query, recordset

    CreateRecordset recordset
    finished = false
    
    if m_urls_per_iteration = 0 then
      m_urls_per_iteration = m_max_urls
    end if
    
    while not finished and m_proceed and m_count < m_urls_per_iteration and m_count < m_max_urls

      recordset.Open "select [url] from [Yider] where [parsed]=0 or [parsed]=2", m_database
      g_open = g_open + 1
      
      if recordset.eof then
        finished = true
      end if
      
      while not recordset.eof and m_proceed and m_count < m_urls_per_iteration and m_count < m_max_urls

        arr = urlextractor.ExtractHREFsFromURL(recordset(0), valid_file_extensions, valid_url_strings, default_documents, delete_between_tags_complete)
        html = arr(0)
        
        'response.Write "<br><br>html is " & html
        
        if Len(html) > 0 and arr(1) then 
        
          if IsArray(urls_to_view_not_store) then
            if not InArrayRegExp(recordset(0), urls_to_view_not_store(0), urls_to_view_not_store(1), m_RegExpObject) then
              ExtractSearchableText recordset(0), html, bad_page_strings, AddArrays(delete_between_tags, delete_between_tags_complete)
            else
              query = "update [Yider] set [parsed]=1 where [url]='" & replace(recordset(0), "'", "''") & "'"
              m_database.Execute(query)
              'response.write query
            end if
          else
            ExtractSearchableText recordset(0), html, bad_page_strings, AddArrays(delete_between_tags, delete_between_tags_complete)
          end if
          
        end if
                
        if arr(1) then
          m_count = m_count + 1
        end if
        
        'response.Write "<br>recordset(0) is " & recordset(0)

        recordset.MoveNext
        
      wend
    
      recordset.Close
      g_open = g_open - 1
    wend
    

    if FullTextEnabled then
      m_database.Execute("exec sp_fulltext_table @tabname='Yider', @action='start_full'")
    end if


    set recordset = Nothing
    g_set = g_set - 1
    
  end sub

  
  public sub GetTextThroughoutDomain(start_url, valid_file_extensions, valid_url_strings, max_urls, bad_page_strings, urls_not_to_view, urls_to_view_not_store, default_documents, delete_between_tags, delete_between_tags_complete)
    Dim arr, html, recordset, urlextractor
    
    m_max_urls = max_urls
    
    set urlextractor = new CURLExtractor
    g_set = g_set + 1

    urlextractor.m_username = m_username
    urlextractor.m_password = m_password

    urlextractor.Constructor m_update, m_database
    urlextractor.m_english = m_english
    urlextractor.m_charset = m_charset
         
    urlextractor.m_urls_not_to_view = urls_not_to_view
    urlextractor.m_strip_url_parameters = m_strip_url_parameters
        
    arr = urlextractor.ExtractHREFsFromURL(start_url, valid_file_extensions, valid_url_strings, default_documents, delete_between_tags_complete)
    html = arr(0)
    
    'response.Write "<br><br>9 html is " & html
    
    if arr(1) then
      ExtractSearchableText start_url, html, bad_page_strings, AddArrays(delete_between_tags, delete_between_tags_complete)
      m_count = 1
    end if
    
    'response.write "<br>m_count is " & m_count
    
    GetTextThroughoutDatabase valid_file_extensions, valid_url_strings, bad_page_strings, urls_not_to_view, urls_to_view_not_store, default_documents, delete_between_tags, delete_between_tags_complete, urlextractor
    

    if m_count <> 0 then
      if m_count = 1 then
        PrintResults "<br><br>" & m_count & " URL was spidered in this pass"
      else
        PrintResults "<br><br>" & m_count & " URL's were spidered in this pass"
      end if
    else
      m_heading_required = false
      PrintResults "<br><br>No URL's have been spidered."
    end if
      
    if IsArray(bad_page_strings) and m_count_invalid_text <> 0 then
      
      if m_count_invalid_text = 1 then
        response.write "<br>" & m_count_invalid_text & " URL contained some of the invalid text in <b>" & bad_page_strings(0)
      else
        response.write "<br>" & m_count_invalid_text & " URLs contained some of the invalid text in <b>" & bad_page_strings(0)
      end if
      
      response.write "</b>"
    end if
    
    
    set recordset = m_database.Execute("select * from [YiderConstants]")
    g_set = g_set + 1
    g_open = g_open + 1
    
    if recordset.eof and InStr(start_url, "yart") = 0 and InStr(start_url, "localhost") = 0 and InStr(start_url, "127.0.0.1") = 0 then
      RegisterYider urlextractor, start_url
      m_database.Execute("insert into [YiderConstants] ([usage]) values (1)")
    end if
    
    recordset.Close
    g_open = g_open - 1
    set recordset = Nothing
    g_set = g_set - 1
    
            
    urlextractor.Destructor
    set urlextractor = Nothing
    g_set = g_set - 1

  end sub
  
  
  'html is the raw html from an url
  'returns an array
  '(0) is the value of the title tag
  '(1) is the searchable text within html
  public function GetTitleAndText(html, delete_between_tags)
  
    Dim count, tag, meta, title, position
    
    'response.Write "<br><br>html is " & server.HTMLEncode(html)
    
    tag = GetTagValueWithAttribute(html, "meta", "name", "keywords")
    
    if g_use_keywords then
      meta = GetAttribute(tag, "content") & " "
      meta = replace(meta, ",", " ")
      meta = replace(meta, "  ", " ")
    else
      meta = ""
    end if
    
    'response.Write "<br><br>99 meta is " & meta
        
    title = " " & GetTagValue(html, "title") & " "
        
    html = replace(html, vbCr, " ") 'get rid of all carriage returns 
    html = replace(html, vbLf, " ") 'get rid of all line feeds
    'html = replace(html, "<br>", vbCrLf) 'change <br> to carriage returns so they don't get stripped
    html = ReplaceTag(html, "head", "") 'remove all content between <script></script> tags
    html = ReplaceTag(html, "script", "") 'remove all content between <script></script> tags
    html = ReplaceTag(html, "title", "") 'remove all content between <title></title> tags
          
    if IsArray(delete_between_tags) then
      for count = 0 to UBound(delete_between_tags) step 2
        html = ReplaceBetweenStrings(html, delete_between_tags(count), delete_between_tags(count + 1), "")
      next
    end if
                
    position = 1
    while position <> 0 
      position = StripBetween(position, "<!--", "-->", true, html)
    wend

    position = 1
    while position <> 0
      position = StripBetween(position, "<select", "</select>", true, html)
    wend

    position = 1
    while position <> 0
      position = StripBetween(position, "<%", Chr(37) & ">", true, html)
    wend
    
    'response.Write server.HTMLEncode(html)

    html = replace(html, "><", "> <")
    position = 1
    while position <> 0
      position = StripBetween(position, "<", ">", true, html)
    wend

    html = replace(html, "&nbsp;", " ") 'must convert &nbsp; to spaces before removing duplicate spaces
    html = replace(html, Chr(9), " ") 'convert tabs to spaces before removing duplicate spaces
        
    html = RemoveDuplicateStrings(html, " ") 'remove duplicate spaces
    html = replace(html, " " & vbCrLf, vbCrLf) 'remove spaces immediately before line breaks
    html = RemoveDuplicateStrings(html, vbCrLf) 'remove duplicate <br>'s (remember that these were converted to vbCrLf's)
    html = replace(html, vbCrLf, vbCrLf & "<br>")
    html = meta & html & " "
    
    GetTitleAndText = Array(title, html)

  end function
  
  
  'GetValBetweenQuotes("content=""ebay, electronics, cars"")
  'returns ebay, electronics, cars
  private function GetValBetweenQuotes(str)
  
    Dim expressionmatch, expressionmatched, RegExpObject, value

    set RegExpObject = New RegExp
    
    With RegExpObject
      .Pattern = """[^""]*"""
      .IgnoreCase = true
      .Global = True
    End With
             
    set expressionmatch = RegExpObject.Execute(str)
         
    For Each expressionmatched in expressionmatch
      value = expressionmatched.value
      value = replace(value, """", "")
      Exit For
    Next

    GetValBetweenQuotes = value
  
  end function
  
  
  'html = "etc<META name=""keywords"" content=""online shopping, auction, online auction""><title>eBay - New & used electronics, cars, apparel, collectibles, sporting goods & more at low prices</title>etc"
  'tag = "meta"
  'the_attribute = "name"
  'value = "keywords"
  'GetTagValueWithAttribute(html, tag, the_attribute, value)
  'returns <META name=""keywords"" content=""online shopping, auction, online auction"">
  private function GetTagValueWithAttribute(html, tag, the_attribute, value)

    Dim expressionmatch, expressionmatched, RegExpObject, val

    set RegExpObject = New RegExp
      
    With RegExpObject
      .Pattern = "<[\s]*" & tag & "[^>]*" & the_attribute & "=""" & value & """[^>]*>"
      .IgnoreCase = true
      .Global = True
    End With
             
    set expressionmatch = RegExpObject.Execute(html)
         
    For Each expressionmatched in expressionmatch
      val = expressionmatched.value
      'response.Write "<br><br>it is " & Server.HTMLEncode(val)
      Exit For
    Next

    GetTagValueWithAttribute = val
    
  end function  


  private function MakeTable(table_name)
    Dim ok
    
    ok = true
  
    Select Case table_name
     
      Case "Yider" 
          
        if DataBaseType = 0 then
          MakeTableAccessYider
        else
          ok = MakeTableSQLYider
        end if
        
      Case "YiderResult"  
      
        if DataBaseType = 0 then
          MakeTableAccessYiderResult
        else
          MakeTableSQLYiderResult
        end if

      Case Else
        ResponseWrite "error MakeTable"
        
    End Select
   
    MakeTable = ok
    
  end function
    
  
  '[parsed] = 0 never parsed
  '[parsed] = 1 parsed
  '[parsed] = 2 parsed a previous time the Tider was run
  private sub MakeTableAccessYider
  
    m_database.Execute "CREATE TABLE Yider([key] IDENTITY PRIMARY KEY, url MEMO, title MEMO, [text] MEMO, parsed INT, URLsize INT NOT NULL, firstLocated VARCHAR(255) NOT NULL)"
    m_database.Execute "CREATE INDEX urlIndex ON Yider ([url])"
    
  end sub
  
  
  private sub MakeTableAccessYiderResult
  
    m_database.Execute "CREATE TABLE YiderResult([key] IDENTITY PRIMARY KEY, [keyYider] INT, [pageRank] DOUBLE)"
    m_database.Execute "CREATE INDEX keyYiderIndex ON YiderResult ([keyYider])"
    m_database.Execute "CREATE INDEX keyMatchTye ON YiderResult ([pageRank])"
    
    m_database.Execute "CREATE TABLE YiderConstants([key] IDENTITY PRIMARY KEY, [usage] INT)"

  end sub
  
  
  private function MakeTableSQLYider
    Dim ok
    
    ok = true

    m_database.Execute "begin tran CREATE TABLE [Yider] ([key] [int] IDENTITY (1, 1) NOT NULL, [url] [nvarchar] (4000) NOT NULL, [title] [ntext] NOT NULL, [text] [ntext] NOT NULL, [parsed] [int] NOT NULL, [URLSize] int NOT NULL, [firstLocated] [nvarchar] (900) NOT NULL, [fulltext_timestamp] [timestamp] NOT NULL) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY] commit tran"
    m_database.Execute "begin tran ALTER TABLE [Yider] WITH NOCHECK ADD CONSTRAINT [PK_Yider] PRIMARY KEY  CLUSTERED ([key]) ON [PRIMARY] commit tran"
    

    if FullTextEnabled then
    
    
      on error resume next
      err.Clear
      m_database.Execute("if not exists (select * from dbo.sysfulltextcatalogs) exec sp_fulltext_database @action='Enable'")
      m_database.Execute("if not exists (select * from dbo.sysfulltextcatalogs where name='YiderFullText') exec sp_fulltext_catalog @ftcat='YiderFullText', @action='CREATE'")
      
      if err.number = -2147217900 then
        JavaScriptAlert("Full-text has not been installed in your SQL Server database.\nFull-text is not, in general, installed by default so you must reload SQL server with full-text features enabled.\nTo check whether full-text is installed, run the command:\n\nexec sp_fulltext_database @action=\'Enable\'\n\nin Query Analyzer on any database in the system.\nIf full-text is not installed, you will see an error.")
        Response.End
      end if
      
      m_database.Execute("exec sp_fulltext_table @tabname='Yider', @action='create', @ftcat='YiderFullText', @keyname='PK_Yider'")
      
      if Len(m_local_ID) = 0 then
      
        response.Write NewLine(0) & "<script language=""JavaScript1.2"">"
        response.Write NewLine(0) & "alert('You have not set g_local_ID in configuration.asp.\nPopulation cannot proceed!');"
        response.Write NewLine(0) & "</script>"
        ok = false
        m_database.Execute("drop table Yider")
      
      else
        m_database.Execute("exec sp_fulltext_column @tabname='Yider', @colname='title', @action='add', @language=" & m_local_ID)
        m_database.Execute("exec sp_fulltext_column @tabname='Yider', @colname='text', @action='add', @language=" & m_local_ID)
      end if
      
    end if
    
    MakeTableSQLYider = ok

  end function
  

  private sub MakeTableSQLYiderResult

    '[pageRank] 0 - complete phrase match
    '[pageRank] 1 - match of all words in the phrase
    m_database.Execute "CREATE TABLE [YiderResult] ([key] [int] IDENTITY (1, 1) NOT NULL, [keyYider] [int] NOT NULL, [pageRank] [float] NOT NULL) ON [PRIMARY]"
    m_database.Execute "ALTER TABLE [YiderResult] WITH NOCHECK ADD CONSTRAINT [PK_YiderResult] PRIMARY KEY  CLUSTERED ([key]) ON [PRIMARY]"
    m_database.Execute "CREATE  INDEX [IX_YiderResult] ON [YiderResult]([keyYider]) ON [PRIMARY]"
    
    m_database.Execute "CREATE TABLE [YiderConstants] ([key] [int] IDENTITY (1, 1) NOT NULL, [usage] [int] NOT NULL)"

  end sub


  private sub PrintResults(str)
  
    if m_heading_required then      
      response.write "<br><br>The Yider crawled through the following URL's:<br>"
      m_heading_required = false
    end if
      
    response.write str
        
  end sub
    
  
  private function RegisterYider(urlextractor, start_url)
    
    'on error resume next
    urlextractor.m_XMLHttp.open "GET", "http://www.yart.com.au/register_yider.asp?start_url=" & Server.HTMLEncode(start_url), false
    urlextractor.m_XMLHttp.send()

  end function
 

  'remove duplicates of the string 'duplicate_string' in 'str'
  'eg RemoveDuplicateStrings("aaa11bbb11ccc", "1") returns "aaa1bbb1ccc"
  'tested for
  'str = RemoveDuplicateStrings("", "b")
  'str = RemoveDuplicateStrings("b", "b")
  'str = RemoveDuplicateStrings("bb", "b")
  'str = RemoveDuplicateStrings("bbb", "b")
  'str = RemoveDuplicateStrings("1bbb", "b")
  'str = RemoveDuplicateStrings("bbb3", "b")
  'str = RemoveDuplicateStrings("1bbb3", "b")
  'str = RemoveDuplicateStrings("1bb3bb5", "b")
  'str = RemoveDuplicateStrings("1bb3bb5b", "b")
  'str = RemoveDuplicateStrings("1bb3bb5bb", "b")
  'str = RemoveDuplicateStrings("b1bb3bb5bb", "b")
  'str = RemoveDuplicateStrings("bb1bb3bb5bb", "b")
  'str = RemoveDuplicateStrings("bb1bb3bb5bb", "bb")
  'str = RemoveDuplicateStrings("bbb1bb3bb5bb", "bb")
  'str = RemoveDuplicateStrings("bbbb1bb3bb5bb", "bb")
  private function RemoveDuplicateStrings(str, duplicate_string)
    Dim position, search
    
    position = InStr(str, duplicate_string & duplicate_string)
    
    while position <> 0

      str = replace(str, duplicate_string & duplicate_string, duplicate_string)
      position = InStr(str, duplicate_string & duplicate_string)
      
    wend
    
    RemoveDuplicateStrings = str
    
  end function 


  'replaces the string between the start and end of a tag with 'replacement'
  'eg ReplaceTag("here comes a <a href="""">tag</a>", "a", "poo")
  'returns "here comes a poo"
  'this function checks for all cases of tag_name and will handle
  'ReplaceTag("<SCRIPT>contents...</script>", "script", "poo")
  private function ReplaceTag(str, tag_name, replacement)
    Dim tag_start, tag_end, str_lcase, str_left, str_right
    
    str_lcase = lcase(str)
    tag_name = lcase(tag_name)

    tag_start = InStr(str_lcase, "<" & tag_name)
    
    while tag_start <> 0
    
      if tag_start <> 0 then
        tag_end = InStr(tag_start + Len("<" & tag_name), str_lcase, "</" & tag_name)
        
        if tag_end <> 0 then
          tag_end = InStr(tag_end, str_lcase, ">")
        end if
      end if
      
      if tag_start <> 0 and tag_end <> 0 then
        str_left = Left(str, tag_start - 1)
        str_right = Right(str, Len(str) - tag_end)
        str = str_left + replacement + str_right
        str_lcase = lcase(str)
        tag_start = InStr(tag_start, str_lcase, "<" & tag_name)
      else
        tag_start = 0
      end if
      
    wend
         
    ReplaceTag = str
    
  end function  
  
  
  private function SpideringComplete
    Dim complete, recordset
    
    CreateRecordset recordset
    
    recordset.Open "select top 1 [key] from [Yider] where [parsed]=0 or [parsed]=2", m_database
    g_open = g_open + 1
    
    if recordset.eof then
      complete = true
    else
      complete = false
    end if
    
    recordset.Close
    g_open = g_open - 1

    set recordset = Nothing
    g_set = g_set - 1
    
    SpideringComplete = complete
    
  end function

  
  public sub StoreTextThroughoutDomain(url, valid_file_extensions, valid_url_strings, bad_page_strings, urls_not_to_view, urls_to_view_not_store, default_documents, delete_between_tags, g_delete_between_tags_complete, max_urls)
    
  
    Dim count, no, ok, recordset, text_array
    
    ok = true
                
    if not DatabaseTableExists("Yider") then
      ok = MakeTable("Yider")
    end if
    
    if not DatabaseTableExists("YiderResult") and ok then
      ok = MakeTable("YiderResult")
    end if
    
    if ok then
    
      GetTextThroughoutDomain url, valid_file_extensions, valid_url_strings, max_urls, bad_page_strings, urls_not_to_view, urls_to_view_not_store, default_documents, delete_between_tags, g_delete_between_tags_complete
      CreateRecordset recordset
      
      recordset.Open "select count(*) from [Yider] where [parsed]=1", m_database
      g_open = g_open + 1
      
      if not SpideringComplete and recordset(0) < max_urls then
      
        response.Write NewLine(0) & "<script language=""JavaScript1.2"">"
        response.Write NewLine(0) & "window.setTimeout('document.population.submit()', " & m_wait & "000);"
        response.Write NewLine(0) & "</script>"

        response.Write NewLine(0) & "<input type=""hidden"" name=""auto_populate"" value=""1"">"
        response.Write NewLine(0) & "<br><br><font color=""red"">PLEASE WAIT - Spidering is incomplete for " & url & " and will recommence in " & m_wait & " seconds</font>"
        response.Write NewLine(0) & "<br>" & recordset(0) & " URL's parsed so far..."
      
      else
      
        response.Write NewLine(0) & "<br><br>FINISHED - Spidering is complete for " & url
        response.Write NewLine(0) & "<br>" & recordset(0) & " URL's parsed in the database"
        
        m_database.Execute("update [Yider] set [parsed] = 2")

      end if
      
      recordset.Close
      g_open = g_open - 1

      set recordset = Nothing
      g_set = g_set - 1
    
    end if
        
  end sub
  

  'removes the text between the first occurrence of 'first_str' and the next occurrence of 'last_str'
  'returns the position of the character after 'last_str' in the new string or 0 if there isn't one
  public function StripBetween(position, first_str, last_str, add_space, byref str)
    Dim length, first_str_pos, last_str_pos, str_left, str_right
    
    length = Len(str)
    first_str_pos = 0
    last_str_pos = 0

    first_str_pos = InStr(position, str, first_str)
    
    if first_str_pos <> 0 then
    'we found the first character
      last_str_pos = InStr(first_str_pos + 1, str, last_str)
      
    'response.write "<br><br>first_str is " & first_str & " last_str is " & last_str

      if last_str_pos = 0 then
        position = 0
      
      elseif last_str_pos <> 0 then
        str_left = Left(str, first_str_pos - 1)
        str_right = Right(str, Len(str) - (last_str_pos + Len(last_str) - 1))
        
        if add_space then
          str = str_left + " " + str_right
        else
          str = str_left + str_right
        end if
                
        if last_str_pos = length then
        'the last character is at the end of the string
          position = 0
        else
          position = Len(str_left)
                    
          if position = 0 then
            position = 1
          end if
        end if
        
        'response.write "<br><br>first_str is " & first_str & "<br>first_str_next is " & first_str_next & "<br>position_ascii_62 is " & position_ascii_62 & "<br>start is " & start & " strlen is " & len(str)
      end if
      
    else
    
      position = 0
      
    end if
    
    'response.write "<br>str is '" & str & "'"

    StripBetween = position
      
  end function
  
  
  private function UpdateRequired(url)
    Dim query, recordset, required
    
    set recordset = Server.CreateObject("ADODB.Recordset")
    g_set = g_set - 1

    recordset.CursorType = 0
    recordset.CacheSize = 1

    required = true
            
    if not m_update then
        
      query =  "select [key] from [Yider] where [url] = '" & replace(url, "'", "''") & "'"
      recordset.Open query, m_database
      g_open = g_open + 1

      if not recordset.eof then
       required = false
      end if
      
      recordset.Close
      g_open = g_open - 1
    end if
    
    set recordset = Nothing
    g_set = g_set - 1
    
    UpdateRequired = required
    
  end function

end class

%>
