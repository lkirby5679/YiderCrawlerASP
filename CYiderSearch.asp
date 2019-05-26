<%Server.ScriptTimeout = 2500%><%
'Copyright (c) Tom Kirby 2005
'This program and its associated source code is distributed under the terms of the GNU General Public 
'License.
'See the attached file COPYING.txt for more information
%>

<!-- #include file="functions.asp" -->


<script language="JScript" runat="server">

function Ceil(number)
{
  return Math.ceil(number)
}

function Floor(number)
{
  return Math.floor(number)
}

function TheMaximum(number1, number2)
{
  return Math.max(number1, number2)
}

function TheMinimum(number1, number2)
{
  return Math.min(number1, number2)
}

</script>


<%

class CYiderSearch

  private m_database, m_isolation_vbscript, m_isolation_sql_access, m_isolation_sql_sqlserver
  public m_no_results, m_style_br, m_style_title, m_style_you_searched_for, m_style_text, m_style_url, m_style_more_results_text, m_style_more_results_link, m_style_more_results_link_Next, m_trailing_words
  public m_charset, m_rank_exact_match, m_rank_partial_match

  public sub Constructor(database_connection)
  
    set m_database = Server.CreateObject("ADODB.Connection")
    g_set = g_set + 1

'response.Write "<br>database connection is " & database_connection 

    m_database.Open database_connection
    g_open = g_open + 1
            
    m_rank_exact_match = 1
    m_rank_partial_match = 0.5
    
    m_no_results = 10
    
    m_isolation_vbscript = "($|\b|\40|\41|\42|\43|\44|\45|\46|\47|\50|\51|\52|\53|\54|\55|\56|\57|\72|\73|\74|\75|\76|\77|\100|\133|\134|\135|\136|\137|\140|\173|\174|\175|\176|\222|\224|\226|\227)"
    m_isolation_sql_access = " #@()-:;'"""",./<>$|!_%&" & Chr(37)
    m_isolation_sql_sqlserver = " #@()-:;''"",./<>$|!_%&" & Chr(37)

    if not MultiByteCharSet(g_charset) then
      m_isolation_sql_access = m_isolation_sql_access & Chr(146) & Chr(148) & Chr(150) & Chr(151)
      m_isolation_sql_sqlserver = m_isolation_sql_sqlserver & Chr(146) & Chr(148) & Chr(150) & Chr(151)
    end if
    
  end sub


  public sub Destructor()
  
    m_database.Close
    g_open = g_open - 1

    set m_database = Nothing
    g_set = g_set - 1
  
  end sub
  
  
  sub DisplayNoiseWords(noise_words)
    Dim count
  
    if IsArray(noise_words) and FullTextEnabled then
      response.write NewLine(0) & "<span style=""" & g_style_noise_word & """>"

      for count = 0 to UBound(noise_words)
        Response.Write "<b>""" & noise_words(count) & """</b>"
        
        if count = UBound(noise_words) - 1 then
          Response.Write " and "
        elseif count < UBound(noise_words) then
          Response.Write ", "
        end if
        
      next
      
      if count = 1 then
        Response.Write " is a very common word and was not included in your search</span><br>"
      else
        Response.Write " are very common words and were not included in your search</span><br>"
      end if
    end if
        
  end sub      
          

  public sub DisplayResults
    Dim recordset, records_displayed
            
    if Len(Trim(YiderSearchTerm())) > 0 then
        
      set recordset = Server.CreateObject("ADODB.Recordset")
      g_set = g_set + 1
  
      recordset.CursorType = 3

      records_displayed = DisplayResults_search(recordset)
      
      if records_displayed > 0 then
        DisplayResults_more recordset, records_displayed
      end if
      
      if records_displayed <> -1 then
        recordset.Close
      end if
      g_open = g_open - 1
  
      set recordset = Nothing
      g_set = g_set - 1
    
    end if
    
  end sub

  
  'returns the number of results displayed
  '-1 if there are none because a full text search on ignored words did not return any results
  public function DisplayResults_search(recordset)
  
    Dim arr, count, extracted, noise_word, noise_words, search_array, search_maintain_search_term, search_phrase, success, title, title_highlighted, text, total, url
        
    if Len(g_search_maintain_search_term) > 0 then
      search_maintain_search_term = g_search_maintain_search_term & "=" & StripExteriorQuotes(YiderSearchTerm())
    else
      search_maintain_search_term = ""
    end if
    
    search_phrase = StripExteriorQuotes(YiderSearchTerm())
    'response.write "<br>search_phrase is  " & search_phrase

    if PhraseSearch then
      search_array = Array(StripExteriorQuotes(search_phrase))
    else
      search_array = GetArrayString(search_phrase, " ")
      search_array = RemoveString(search_array, "")
    end if

    arr = GetResults(recordset, search_phrase, search_array)
    success = arr(0)
    
    if FullTextEnabled then
      noise_words = arr(1)
      search_array = arr(2)
    end if
    
    'response.write "<br>search_array is " & UBound(search_array)
    
    if success then
    
      if not recordset.eof then
    
        recordset.Move(Start() - 1)
        
        if recordset.RecordCount = 1 then
          response.write NewLine(0) & NewLine(0) & "<span style=""" & m_style_you_searched_for & """>" & g_you_searched & " " & SearchTextDescription & " <b>" & StripExteriorQuotes(YiderSearchTerm()) & "</b> " & g_and & " " & recordset.RecordCount & " " & g_page_was_found & "</a><br></span>"
        else
          response.write NewLine(0) & NewLine(0) & "<span style=""" & m_style_you_searched_for & """>" & g_you_searched & " " & SearchTextDescription & " <b>" & StripExteriorQuotes(YiderSearchTerm()) & "</b> " & g_and & " " & recordset.RecordCount & " " & g_pages_were_found & "</a><br></span>"
        end if
        
        DisplayNoiseWords noise_words
        
        count = 0

        while not recordset.eof and count < m_no_results
        
          count = count + 1
            
          title = Trim(recordset(0))
          
          if Len(title) = 0 then
            title_highlighted = "(Untitled)"
          else
          
            title_highlighted = Highlight(title, search_array)

            if Len(title_highlighted) = 0 then
              title_highlighted = title
            end if 
                     
          end if

          text = recordset(1)
          url = recordset(2)

          if count = 1 and Len(search_maintain_search_term) > 0 then
            if InStr(url, "?") = 0 then
              search_maintain_search_term = "?" & search_maintain_search_term
            else
              search_maintain_search_term = "&" & search_maintain_search_term
            end if
          end if
          
          response.write NewLine(0) & NewLine(0) & "<br><a href=""" & url & search_maintain_search_term & """ style=""" & m_style_title & """>" & title_highlighted & "</a>"
          
          extracted = ExtractAndHighlight(text, search_phrase, search_array, m_trailing_words)
          
          if Len(Trim(extracted)) > 0 then
            response.write NewLine(0) & "<br><span style=""" & m_style_text & """>" & extracted & "</span>"
          end if
          
          response.write NewLine(0) & "<br><span style=""" & m_style_url & """>" & url & "</span>"
          response.write NewLine(0) & "<span style=""" & m_style_br & """><br></span>"
          
          recordset.MoveNext
          
        wend
        
      else
      
        response.write NewLine(0) & "<br><span style=""" & m_style_text & """>" & g_no_pages & " " & SearchTextDescription & " <b>" & Trim(YiderSearchTerm()) & "</b></span>"
      
      end if
      
    else
      DisplayNoiseWords noise_words
      count = -1
    end if
    
    DisplayResults_search = count
        
  end function
  
  
  public sub DisplayResults_more(recordset, records_displayed)
    Dim i, i_for_next, last_page_no, first_page_no, current_record, search_maintain_url_params, total
    
    current_record = Start()
    total = recordset.RecordCount
    
    last_page_no = TheMinimum(  (total - 1)/m_no_results + 1, 10 * Floor(current_record/(10*m_no_results)) + 10)
    first_page_no = 10 * Floor(current_record/(10*m_no_results)) + 1
    
    if IsArray(g_search_maintain_url_params) then
      
      search_maintain_url_params = ""
    
      for i = 0 to UBound(g_search_maintain_url_params)
      
        if i <> 0 then
          search_maintain_url_params = search_maintain_url_params & "&" & g_search_maintain_url_params(i) & "=" & Server.URLEncode(Request(g_search_maintain_url_params(i)))
        else
          search_maintain_url_params = search_maintain_url_params & g_search_maintain_url_params(i) & "=" & Server.URLEncode(Request(g_search_maintain_url_params(i)))
        end if
        
      next
      
    end if
    
    'response.write "<br>current_record is " & current_record
    'response.write "<br>m_no_results is " & m_no_results
    'response.write "<br>first_page_no is " & first_page_no
    'response.write "<br>last_page_no is " & last_page_no
    'response.write "<br>total is " & total
    'response.write "<br>recordset.eof is " & recordset.eof
    'response.write "<br>recordset.bof is " & recordset.bof
    'response.write "<br><br><br>"
    
    if not recordset.bof then
      response.write "<br><span style=""" & m_style_more_results_text & """>" & g_result_pages & ":&nbsp;</span>"
    end if

    'if first_page_no > 10 then
    '  response.write "<a href=""" & Filename(Request.ServerVariables("SCRIPT_NAME")) & "?yider=" & YiderSearchTerm() & "&start=" & (first_page_no - 10) & """ style=""" & m_style_more_results_link_Next_100 & """>&lt;&lt;&nbsp;Previous " & m_no_results*10 & "</a>&nbsp;&nbsp;&nbsp;"
    'end if
    
    if current_record <> 1 then
      response.Write "<a href=""" & Filename(Request.ServerVariables("SCRIPT_NAME")) & "?yider=" & replace(YiderSearchTerm(), """", "&quot;") & "&start=" & current_record - m_no_results & "&" & search_maintain_url_params & """ style=""" & m_style_more_results_link_Next & """>" & g_previous & "</a>&nbsp;&nbsp;&nbsp;"
    end if


    if not recordset.bof then
    'there are records
    
      last_page_no = TheMinimum((current_record - 1)/m_no_results + m_no_results, Ceil(total / m_no_results))
        
      'for i = first_page_no to last_page_no
      for i = 1 to last_page_no

        if current_record = (i * m_no_results) - m_no_results + 1 then
          i_for_next = i + 1
          response.write "<span style=""" & m_style_more_results_link & """>" & i & "</span>&nbsp;&nbsp;"
        else
          response.write "<a href=""" & Filename(Request.ServerVariables("SCRIPT_NAME")) & "?yider=" & replace(YiderSearchTerm(), """", "&quot;") & "&start=" & m_no_results*(i - 1) + 1 & "&" & search_maintain_url_params & """ style=""" & m_style_more_results_link & """>" & i & "</a>&nbsp;&nbsp;"
        end if
          
      next
      
      if m_no_results*(i_for_next - 1) + 1 <= total then
        response.Write "<a href=""" & Filename(Request.ServerVariables("SCRIPT_NAME")) & "?yider=" & replace(YiderSearchTerm(), """", "&quot;") & "&start=" & m_no_results*(i_for_next - 1) + 1  & "&" & search_maintain_url_params & """ style=""" & m_style_more_results_link_Next & """>" & g_next & "</a>&nbsp;&nbsp;&nbsp;"
      end if
                  
      if total >= m_no_results*last_page_no + 1 then
      
        'if total - (m_no_results*(first_page_no - 1) + m_no_results*10) < m_no_results*10 then
        '  response.write "<a href=""" & Filename(Request.ServerVariables("SCRIPT_NAME")) & "?yider=" & YiderSearchTerm() & "&start=" & m_no_results*last_page_no + 1 & """ style=""" & m_style_more_results_link_Next_100 & """>Next " & total - (m_no_results*(first_page_no - 1) + m_no_results*10) & "&nbsp;&gt;&gt;</a>&nbsp;&nbsp;&nbsp;"
        'else
        '  response.write "<a href=""" & Filename(Request.ServerVariables("SCRIPT_NAME")) & "?yider=" & YiderSearchTerm() & "&start=" & m_no_results*last_page_no + 1 & """ style=""" & m_style_more_results_link_Next_100 & """>Next " & m_no_results*10 & "&nbsp;&gt;&gt;</a>&nbsp;&nbsp;&nbsp;"
        'end if
        
      end if
      
    end if
    
  end sub
  
  
  private function EscapeRegExpChars(byval text_to_find)
  
    text_to_find = replace(text_to_find, "\", "\\")
    text_to_find = replace(text_to_find, ".", "\.")
    'text_to_find = replace(text_to_find, "*", "\*") ' * is a wildcard and should not be escaped
    text_to_find = replace(text_to_find, "?", "\?")
    text_to_find = replace(text_to_find, "(", "\(")
    text_to_find = replace(text_to_find, ")", "\)")
    text_to_find = replace(text_to_find, "{", "\{")
    text_to_find = replace(text_to_find, "}", "\}")
    text_to_find = replace(text_to_find, "[", "\[")
    text_to_find = replace(text_to_find, "]", "\]")
    text_to_find = replace(text_to_find, "$", "\$")
    text_to_find = replace(text_to_find, "|", "\|")
    text_to_find = replace(text_to_find, "+", "\+")
    
    'response.Write "<br>text_to_find is " & text_to_find
    
    EscapeRegExpChars = text_to_find

  end function


  private function Extract(byval text_to_search, byval text_to_find, word_array, max_words)
    Dim count, extracted, matched_text
        
    extracted = ExtractMatchedText(text_to_search, text_to_find, max_words)
    
    'response.Write "<br>extracted is '" & extracted & "'"
    'response.Write "<br>text_to_find is '" & text_to_find & "'"
    'response.Write "<br>text_to_search is '" & text_to_search & "'"
    'response.Write "<br>word_array is " & word_array(0)
    'response.Write "<br>word_array is " & word_array(1)
    
    if Len(extracted) = 0 then
      for count = 0 to UBound(word_array)

        'response.Write "<br><br>count is " & count
      
        if InStr(lcase(matched_text), lcase(word_array(count))) = 0 then
        'check the word is not already in the string
        
          matched_text = ExtractMatchedText(text_to_search, word_array(count), max_words)
                  
          if SearchComplete(matched_text, extracted, word_array) then
            extracted = extracted & matched_text
            Exit For
          elseif Len(extracted) > 0 and Len(matched_text) > 0 then
            extracted = extracted & "<br>"
            extracted = extracted & matched_text
          else
            extracted = extracted & matched_text
          end if

          'response.Write "<br>extracted is " & extracted & "<br><br>"
      
        end if
        
      next
    end if

    Extract = extracted
  end function

  
  'finds the first occurrence of the 'text_to_find' in 'text_to_search'
  'places <b></b> tags around the 'text_to_find' and returns 'max_words' words before and after the text_to_find
  'if 'text_to_find' can't be found, the function returns the first occurrence of each word in 'text_to_find' and
  'places <b></b> tags around each word with 'max_words' words before and after each word
  'no_match_str if true, return a default comment
  private function ExtractAndHighlight(byval text_to_search, byval text_to_find, word_array, max_words)
    Dim extracted
                
    extracted = Extract(text_to_search, text_to_find, word_array, max_words)
    
    'response.Write "<br>text_to_search is " & text_to_search
    'response.Write "<br>extracted is '" & extracted & "'"
    
    if Len(extracted) = 0 and max_words <> -1 then
      extracted = GetWords(text_to_search, 2*max_words)
    else
      extracted = Highlight(extracted, word_array)
    end if
    
    'response.Write "<br>extracted is '" & Len(Trim(extracted)) & "'"

    ExtractAndHighlight = extracted
    
  end function
   
  
  private function ExtractMatchedText(text_to_search, byval text_to_find, byval max_words)
    Dim ret
  
    if MultiByteSearch(YiderSearchTerm()) then
      ret = ExtractMatchedTextMultiByte(text_to_search, text_to_find, max_words)
    else
      ret = ExtractMatchedTextSingleByte(text_to_search, text_to_find, max_words)
    end if
    
    ExtractMatchedText = ret
    
  end function
  

  private function ExtractMatchedTextMultiByte(text_to_search, byval text_to_find, byval max_words)
    Dim count, extracted, pos, str, str_left, str_right, size
    
    size = 40
         
    pos = InStr(text_to_search, text_to_find)

    if pos <> 0 then
    
      if max_words <> -1 then
      
        if pos <= size then
          str_left = Left(text_to_search, pos - 1)
        else
          str_left = Mid(text_to_search, pos - size, size)

        end if
        
        str_right = Mid(text_to_search, pos + Len(text_to_find), size)
        max_words = -1

      else
        str_left = Left(text_to_search, pos - 1)
        str_right = Mid(text_to_search, pos + Len(text_to_find))
      end if
      
      extracted = str_left & Mid(text_to_search, pos, Len(text_to_find)) & str_right

    end if
    
    'response.Write "<br>extracted is " & extracted
    
    ExtractMatchedTextMultiByte = extracted

  end function
  

  'this function returns a string with a number of words before and after text_to_find in text_to_search
  'it retains max_words before the first occurrence of text_to_search
  'it retains max_words after the end of the first occurrence of text_to_search
  'if the retained string is not the beginning/end of text_to_search, three dots are appended
  'if max_words is -1, it returns all of text_to_search
  private function ExtractMatchedTextSingleByte(text_to_search, byval text_to_find, byval max_words)

    Dim count, expressionmatch, expressionmatched, extracted, RegExpObject, str, str_left, str_right
         
    text_to_find = EscapeRegExpChars(text_to_find)
    
    text_to_find = WordIsolationPattern(text_to_find)

    set RegExpObject = New RegExp
    g_set = g_set + 1
    
    With RegExpObject
      .Pattern = text_to_find
      'response.Write "<br>pattern is " & text_to_find
      .IgnoreCase = true
      .Global = True
    End With
           
    set expressionmatch = RegExpObject.Execute(text_to_search)
    g_set = g_set + 1
       
    'response.Write "<br><br>text_to_search is " & text_to_search
    'response.Write "<br>Pattern is " & RegExpObject.Pattern
    'response.Write "<br>count is " & expressionmatch.Count
    
    For Each expressionmatched in expressionmatch

      'response.Write "<br>expressionmatched.FirstIndex is " & expressionmatched.FirstIndex
    
      if max_words <> -1 then
        str_left = WordsBefore(text_to_search, expressionmatched.FirstIndex + 1, max_words)
        str_right = WordsAfter(text_to_search, expressionmatched.FirstIndex + 1 + expressionmatched.length, max_words)
        max_words = -1
      else
        str_left = Left(text_to_search, expressionmatched.FirstIndex)
        str_right = Mid(text_to_search, expressionmatched.FirstIndex + 1 + expressionmatched.length)
      end if
      
      'response.Write "<br><br>str_left is '" & str_left & "'"
      'response.Write "<br>str_right is '" & str_right & "'"
      
      str = Mid(text_to_search, expressionmatched.FirstIndex + 1, expressionmatched.length)
            
      extracted = str_left & str & str_right

      'response.Write "<br><br>extracted is " & extracted

      Exit For
      
    Next
    
    'response.Write "<br>extracted is '" & extracted & "'"

    set expressionmatch = Nothing
    g_set = g_set - 1

    set RegExpObject = Nothing
    g_set = g_set - 1

    ExtractMatchedTextSingleByte = extracted
  end function
  

  'returns the filename only of a path if its path and filename is in a string
  private function FileName(complete_file_name)
    Dim file_name, pos

    complete_file_name = Replace(complete_file_name, "/", "\")

    pos = InStrRev(complete_file_name, "\")

    if pos <> 0 then
      file_name = Right(complete_file_name, Len(complete_file_name) - pos)
    else
      file_name = complete_file_name
    end if

    FileName = file_name

  end function

  
  private function GetWords(text_to_search, max_words)
    Dim expressionmatch, RegExpObject, words
    
    set RegExpObject = New RegExp
    g_set = g_set + 1

    With RegExpObject
      .Pattern = "\s"
      .IgnoreCase = true
      .Global = True
    End With
    
    set expressionmatch = RegExpObject.Execute(text_to_search)
    g_set = g_set + 1
    
    'response.write "<br>expressionmatch.count is " & expressionmatch.count
    'response.write "<br>max_words is " & max_words
    
    if expressionmatch.count < max_words then
      words = text_to_search
    else
      words = Left(text_to_search, expressionmatch.item(max_words - 1).FirstIndex - 1) & "..."
    end if
      
    set expressionmatch = Nothing
    g_set = g_set - 1

    set RegExpObject = Nothing
    g_set = g_set - 1
    
    GetWords = words

  end function
    

  private function GetResults(byref recordset, search_phrase, search_array)
    Dim arr
    
    if FullTextEnabled then
      arr = GetResultsLikeFullText(recordset, search_phrase)
    else
      arr = GetResultsLike(recordset, search_phrase, search_array)
    end if
    
    GetResults = arr
    
  end function
  
  
  private function GetLikePhrase(search_phrase)
    Dim count, like_phrase
    
    'if this is a multibyte web site and multibyte characters are being searched for, detect the search string 
    'no matter what characters preceed/follow it
    
    'response.Write "it is "
    'response.Write m_charset

    if MultiByteSearch(search_phrase) then
      m_isolation_sql_access = ""
      m_isolation_sql_sqlserver = ""
      m_isolation_vbscript = ""
    end if
    
    if DataBaseType = 0 then
      GetLikePhrase = GetLikePhraseAccess(search_phrase)
    else
      GetLikePhrase = GetLikePhraseSQLServer(search_phrase)
    end if
    
    like_phrase = GetLikePhrase
    
  end function
  
  
  private function GetLikePhraseAccess(byval search_phrase)
    Dim left_prefix, like_phrase, RegExpObject, right_prefix
    
    Set RegExpObject = New RegExp
    g_set = g_set + 1
    
    search_phrase = Trim(search_phrase)
    search_phrase = StripExteriorQuotes(search_phrase)
    search_phrase = replace(search_phrase, "[", "[[]")
    search_phrase = replace(search_phrase, "_", "[_]")
    search_phrase = replace(search_phrase, "%", "[%]")
    search_phrase = replace(search_phrase, """", """""")
   
    if Left(search_phrase, 1) <> "*" and Right(search_phrase, 1) = "*" then
      like_phrase = Left(search_phrase, Len(search_phrase) - 1) & "[!" & m_isolation_sql_access & "]%"
    
    elseif Left(search_phrase, 1) = "*" and Right(search_phrase, 1) <> "*" then
      like_phrase = "%[!" & m_isolation_sql_access & "]" & Mid(search_phrase, 2)

    elseif Left(search_phrase, 1) = "*" and Right(search_phrase, 1) = "*" then
      like_phrase = "%[!" & m_isolation_sql_access & "]" & Mid(search_phrase, 2, Len(search_phrase) - 2) & "[!" & m_isolation_sql_access & "]%"

    else
      like_phrase = search_phrase
    
    end if

    if Len(m_isolation_sql_access) > 0 then
      like_phrase = """%[" & m_isolation_sql_access & "]" & like_phrase & "[" & m_isolation_sql_access & "]%"""
    else
      like_phrase = """%" & m_isolation_sql_access & like_phrase & m_isolation_sql_access & "%"""
    end if

    
    Set RegExpObject = New RegExp
    g_set = g_set - 1
    
    GetLikePhraseAccess = like_phrase
  end function

  

  private function GetLikePhraseSQLServer(byval search_phrase)
    Dim left_prefix, like_phrase, RegExpObject, right_prefix
    
    Set RegExpObject = New RegExp
    g_set = g_set + 1
    
    search_phrase = Trim(search_phrase)
    search_phrase = StripExteriorQuotes(search_phrase)
    search_phrase = replace(search_phrase, "[", "[[]")
    search_phrase = replace(search_phrase, "^", "[^]")
    search_phrase = replace(search_phrase, "%", "[%]")
    search_phrase = replace(search_phrase, "_", "[_]")

    if Left(search_phrase, 1) <> "*" and Right(search_phrase, 1) = "*" then
      like_phrase = Left(search_phrase, Len(search_phrase) - 1) & "%"
    
    elseif Left(search_phrase, 1) = "*" and Right(search_phrase, 1) <> "*" then
      like_phrase = "%" & Mid(search_phrase, 2)

    elseif Left(search_phrase, 1) = "*" and Right(search_phrase, 1) = "*" then
      like_phrase = "%" & Mid(search_phrase, 2, Len(search_phrase) - 2) & "%"

    else
      like_phrase = search_phrase
    
    end if

    if Len(m_isolation_sql_sqlserver) > 0 then
      like_phrase = "N'%[" & m_isolation_sql_sqlserver & "]" & replace(like_phrase, "'", "''") & "[" & m_isolation_sql_sqlserver & "]%'"
    else
      like_phrase = "N'%" & replace(like_phrase, "'", "''") & "" & m_isolation_sql_sqlserver & "%'"
    end if
    
    Set RegExpObject = New RegExp
    g_set = g_set - 1
    
    GetLikePhraseSQLServer = like_phrase
  end function
  
    
  private function GetResultsLike(byref recordset, search_phrase, search_array)
    Dim count, like_phrase, query, query1, query2, query3, value
    
    like_phrase = GetLikePhrase(search_phrase)
    query1 = "select [key], 0 from [Yider] where [text] like " & like_phrase & " or [title] like " & like_phrase

    'response.Write vbcrlf & "<br>it is  " & query1
        
    if UBound(search_array) > 0 then
      query2 = "select [key], 1 from [Yider] where "

      for count = 0 to UBound(search_array)
      
        if count > 0 then
          query2 =  query2 & " and"
        end if
        
        like_phrase = GetLikePhrase(search_array(count))
        query2 =  query2 & " [text] like " & like_phrase
        
      next
      
      query2 = query2 & " and [key] not in (select [keyYider] from [YiderResult])"

    else
      query2 = ""
    end if
    
    'response.Write "<br>it is " & query2
   
    if DataBaseType = 0 then

      m_database.Execute "delete * from [YiderResult]"
      m_database.Execute "insert into [YiderResult] (keyYider, pageRank) " & query1
      
      'response.write "<br>insert into [YiderResult] (keyYider, pageRank) " & query1

      if Len(query2) > 0 then
        query2 = "insert into [YiderResult] (keyYider, pageRank) " & query2
        m_database.Execute query2
      end if

    else
    
      m_database.Execute "begin tran truncate table [YiderResult] commit tran"
      m_database.Execute "begin tran insert into [YiderResult] " & query1 & " commit tran"
      'response.Write "begin tran insert into [YiderResult] " & query1 & " commit tran"

      if Len(query2) > 0 then
        query2 = "begin tran insert into [YiderResult] " & query2 & " commit tran"
        m_database.Execute query2
      end if
    
    end if
    
    RankResults search_phrase, search_array
    query = "select [Yider].[title], [Yider].[text], [Yider].[url] from [Yider], [YiderResult] where [YiderResult].[keyYider]=[Yider].[key] order by [YiderResult].[pageRank] desc"

    recordset.open query, m_database
    g_open = g_open + 1
    'this recordset is closed in DisplayResults
    'response.write recordset(1)
    
    GetResultsLike = Array(true, 0)
        
  end function
  
  
  private function GetResultsLikeFullText(byref recordset, search_phrase)
    Dim ret

    if PhraseSearch then
      ret = GetResultsLikeFullTextPhrase(recordset, search_phrase)
    else
      ret = GetResultsLikeFullTextNotPhrase(recordset, search_phrase)
    end if
    
    GetResultsLikeFullText = ret
    
  end function

  
  private function GetResultsLikeFullTextNotPhrase(byref recordset, search_phrase)
    Dim count, noise_words, query, search_array, search_string, search_strings, success
    
    search_strings = Split(search_phrase, " ", -1, 1)
    
    query = "select [title], [text], [url], [Rank] from containstable(Yider, *, '"
    
    '"brown dog" and "lazy"') as ct join Yider as y on ct.[key]=y.[key] order by Rank desc
    
    count = 0
    for each search_string in search_strings
      
      if not IsNoiseWord(search_string) then
      
        if count <> 0 then
          query = query & " and """ & search_string & """"
        else
          query = query & """" & search_string & """"
        end if
        
        count = count + 1
        
        AddToArray search_string, search_array
        
      else
      
        AddToArray search_string, noise_words
      
      end if
      
    next
    
    query = query & "') as ct join Yider as y on ct.[key]=y.[key] order by Rank desc"
    
    'Response.Write "it is " & query
    
    on error resume next
    'the line below will fail if the query only contains noise words
    recordset.open query, m_database
    g_open = g_open + 1
    
    'response.Write "it is " & err.number
    
    if err.number = -2147217900 then
      success = false
    else
      success = true
    end if
    
    err.Clear
    
    GetResultsLikeFullTextNotPhrase = Array(success, noise_words, search_array)
        
  end function


  private function GetResultsLikeFullTextPhrase(byref recordset, search_phrase)
    Dim noise_words, query, search_strings, success, word
    
    query = "select [title], [text], [url], [Rank] from containstable(Yider, *, '""" & search_phrase & """') as ct join Yider as y on ct.[key]=y.[key] order by Rank desc"
    'response.Write "it is " & query
    
    search_strings = Split(search_phrase, " ", -1, 1)
    
    for each word in search_strings
      
      if IsNoiseWord(word) then
        AddToArray word, noise_words
      end if
      
    next    

    on error resume next
    'the line below will fail if the query only contains noise words
    recordset.open query, m_database
    g_open = g_open + 1

    if err.number = -2147217900 then
      success = false
    else
      success = true
    end if
    
    err.Clear
    
    
    GetResultsLikeFullTextPhrase = Array(success, noise_words, Array(search_phrase))
        
  end function


  private function Highlight(byval text_to_search, byval word_array_to_find)
    Dim ret
    
    if MultiByteSearch(YiderSearchTerm()) then
      ret = HighlightMultiByte(text_to_search, word_array_to_find)
    else
      ret = HighlightSingleByte(text_to_search, word_array_to_find)
    end if
    
    Highlight = ret
    
  end function
  

 'highlight all instances of the words in word_array_to_find within text_to_search with a <b> tag
  private function HighlightMultiByte(byval text_to_search, byval word_array_to_find)
    
    Dim bold_tag, count, find, pos, start
    
    bold_tag = "<b>"
    
    for count = 0 to UBound(word_array_to_find)
   
      pos = -1
      start = 1
      find = word_array_to_find(count)
      pos = InStr(start, text_to_search, find)
 
      while pos <> 0
      
        text_to_search = Left(text_to_search, pos - 1) & bold_tag & find & "</b>" & Mid(text_to_search, pos + Len(find))
        start = pos + Len(bold_tag) + 1
        pos = InStr(start, text_to_search, find)
        
      wend
      
    next
    
    HighlightMultiByte = text_to_search
    
  end function

    
  'highlight all instances of the words in word_array_to_find within text_to_search with a <b> tag
  private function HighlightSingleByte(byval text_to_search, byval word_array_to_find)
    Dim count, expressionmatch, expressionmatched, highlighted, matches, pos, RegExpObject, str, str_left, str_right, text_to_find
  
    set RegExpObject = New RegExp
    g_set = g_set + 1
    
    if not PhraseSearch then
      highlighted = replace(text_to_search, " ", "  ")
      'if adjacent matches are present, they must be separated by double spaces or the regular expression won't find them
    else
      highlighted = text_to_search
    end if
    
    for count = 0 to UBound(word_array_to_find)
    
      matches = 0
      'response.Write "<br><br>1 it is " & text_to_search
      'response.Write "<br>it is " & word_array_to_find(count)
      'response.Write "<br>it is " & EscapeRegExpChars(word_array_to_find(count))

      'response.Write "<br>it is " & word_array_to_find(count)
      text_to_find = EscapeRegExpChars(word_array_to_find(count))
      'response.Write "<br>it is " & text_to_find
      
      'response.Write "<br>highlighted is " & Mid(highlighted, 1, 1)

      With RegExpObject
        .Pattern = WordIsolationPattern(text_to_find)
        .IgnoreCase = true
        .Global = True
      End With

      set expressionmatch = RegExpObject.Execute(highlighted)
      g_set = g_set + 1
      'response.Write "<br>expressionmatch.count is " & expressionmatch.count
      'response.Write "<br>highlighted is " & highlighted
      
      For Each expressionmatched in expressionmatch
             
        pos = expressionmatched.FirstIndex + 7*matches
  
        'response.Write "<br>length is " & expressionmatched.length
        'response.Write "<br>pos is " & pos
        'response.Write "<br>expressionmatched is " & expressionmatched
                
        str_left = Left(highlighted, pos)
        str = Mid(highlighted, pos + 1, expressionmatched.length)
        str_right = Mid(highlighted, pos + 1 + expressionmatched.length)
        
        'response.Write "<br>str_left is '" & str_left & "'"
        'response.Write "<br>str is '" & str & "'"
        'response.Write "<br>str_right is '" & str_right & "'"
        
        if Left(Trim(YiderSearchTerm()), 1) <> "*" and Right(Trim(YiderSearchTerm()), 1) = "*" then
          str = Left(str, 1) & "<b>" & Mid(str, 2, Len(str) - 1) & "</b>"

        elseif Left(Trim(YiderSearchTerm()), 1) = "*" and Right(Trim(YiderSearchTerm()), 1) <> "*" then
          str = "<b>" & Mid(str, 1, Len(str)) & "</b>"
        
        elseif Left(Trim(YiderSearchTerm()), 1) = "*" and Right(Trim(YiderSearchTerm()), 1) = "*" then
          str = "<b>" & str & "</b>"
        
        else

          if MultiByteSearch(YiderSearchTerm()) then
            str = "<b>" & str & "</b>"
          else
        
            if pos = 0 then
            'match at beginning of string
              str = "<b>" & Mid(str, 1, Len(str)) & "</b>"
            
            else
              str = Left(str, 1) & "<b>" & Mid(str, 2, Len(str) - 1) & "</b>"
            end if
            
          end if
          
        end if
        
         highlighted = str_left & str & str_right

        'response.Write "<br>highlighted is '" & highlighted & "'"

        matches = matches + 1
        
      Next

      set expressionmatch = Nothing
      g_set = g_set - 1
    
    next
    
        'response.Write "<br>highlighted is '" & highlighted & "'"

    set RegExpObject = Nothing
    g_set = g_set - 1
    
    HighlightSingleByte = highlighted
  end function
  
  
  private function IsNoiseWord(noise_word)
    Dim is_noise_word, query
    
    on error resume next
    err.Clear
    
    query = "select top 1 [y].[key] from containstable(Yider, *, '""" & noise_word & """') as ct join Yider as y on ct.[key]=y.[key] order by Rank desc"
    
    'response.Write "<br>query is " & query
    m_database.Execute(query)
    
    if err.number = -2147217900 then
      err.Clear
      is_noise_word = true
    else
      is_noise_word = false
    end if
    
    IsNoiseWord = is_noise_word
    
  end function
  
  
  private function MultiByteSearch(search_phrase)
    Dim count, is_multi, the_asc
    
    is_multi = false
  
    if MultiByteCharSet(m_charset) then
    
      for count = 1 to Len(search_phrase)
        the_asc = Asc(Mid(search_phrase, count))
        
        'response.Write "<br>count is " & count & " the_asc is " & the_asc
        
        if (the_asc < 1 or the_asc > 127) then
          is_multi = true
          Exit For
        end if
      next
    end if
    
    'response.write "<br>it is " & is_multi
    
    MultiByteSearch = is_multi
    
  end function

  
  private function NewLine(spaces)
    NewLine = vbCrLf & Space(spaces)
  end function
  
  
  private function NextNonSpace(str, pos)
  
    Dim ch, pos_non_space

	  pos_non_space = pos
    ch = Mid(str, pos, 1)

	  while(ch = " " and pos_non_space < Len(str))

	    pos_non_space = pos_non_space + 1
	    ch = Mid(str, pos_non_space, 1)

	  wend

    NextNonSpace = pos_non_space

  end function


  'returns the previous charcater that is not a space from pos in str
  private function PreviousNonSpace(str, pos)
  
    Dim ch, pos_non_space

	  pos_non_space = pos
    ch = " "
    
    'response.write "<br>-pos is " & pos & " ch is " & ch

	  while(ch = " " and pos_non_space > 1)

	    pos_non_space = pos_non_space - 1
	    ch = Mid(str, pos_non_space, 1)

      'response.write "<br>--pos is " & pos & " pos_non_space is " & pos_non_space & " ch is " & ch
	  wend

	  if pos_non_space < 1 then
	    pos_non_space = 1
	  end if

    PreviousNonSpace = pos_non_space

  end function
 

  private sub RankResults(search_phrase, search_array)

    Dim count, match_count, query, rank, recordset, RegExpObject, search_array_size, text, text_to_find
    
    'response.Write "<br>search_phrase is " & search_phrase

    match_count = Array()

    query = "select [YiderResult].[key], [Yider].[text] from [Yider] inner join [YiderResult] on [Yider].[key] = [YiderResult].[keyYider]"
    'response.Write query
    set recordset = m_database.Execute(query)
    g_set = g_set + 1
    g_open = g_open + 1
    
    Set RegExpObject = New RegExp
    g_set = g_set + 1

    search_array_size = UBound(search_array)

    'response.Write "<br>search_array_size  is " & search_array_size 
    
    while not recordset.eof
    
      text = recordset(1)
      'response.Write "<br><br>[Yider].[url] is " & recordset(0)

      rank = RankResultsPage(text, Array(search_phrase), RegExpObject, m_rank_exact_match)

      if not PhraseSearch then
        rank = rank + RankResultsPage(text, search_array, RegExpObject, m_rank_partial_match * 1 / (search_array_size + 1))
      end if

      query = "update [YiderResult] set [YiderResult].[pageRank]=[YiderResult].[pageRank] + " & replace(CStr(FormatNumber(rank, 2, 0, 0, 0)), ",", ".") & " where [YiderResult].[key]=" & recordset(0)
      'response.write query
      
      m_database.Execute(query)
      recordset.MoveNext

    wend
    
    recordset.Close
    g_open = g_open - 1
    
    set recordset = Nothing
    g_set = g_set - 1
    
    set RegExpObject = Nothing
    g_set = g_set - 1
    
  end sub
  
  
  private function RankResultsPage(text_to_search, byval text_to_find, byref RegExpObject, points)
    Dim i, j, expressionmatch, expressionmatched, match_positions, min_matches, next_rank, next_str, no_strings, text_to_search_length, value
    
    value = 0
    min_matches = "n"
    text_to_search_length = Len(text_to_search)
    no_strings = UBound(text_to_find) + 1
    match_positions = Array()
    ReDim match_positions(no_strings - 1, 1)
    
    'response.Write "<br>text_to_find is  " & text_to_find(0)
    'response.Write "<br><br>no_strings is  " & no_strings

    for i = 0 to no_strings - 1
      next_str = text_to_find(i)

      With RegExpObject
        .Pattern = WordIsolationPattern(EscapeRegExpChars(next_str))
        .IgnoreCase = true
        .Global = True
      End With
      
      'response.Write "<br>text_to_search is " & text_to_search
      'response.Write "<br>pattern is " & EscapeRegExpChars(next_str)

      set expressionmatch = RegExpObject.Execute(text_to_search)
      g_set = g_set + 1
      
      if not IsNumeric(min_matches) then
        min_matches = expressionmatch.count
        ReDim Preserve match_positions(no_strings - 1, expressionmatch.count - 1)
      elseif min_matches < expressionmatch.count then
        min_matches = expressionmatch.count
        ReDim Preserve match_positions(no_strings - 1, expressionmatch.count - 1)
      end if
      
      'response.Write "<br><br>text_to_find(i) is " & text_to_find(i)
      'response.Write "<br>min_matches is " & min_matches
      'response.Write "<br>expressionmatch.count is " & expressionmatch.count
      'response.Write "<br>expressionmatch.count is " & expressionmatch.count

      if expressionmatch.count > 0 then

        'response.Write "<br>UBound zzz is " & UBound(match_positions, 1)
        
        j = 0
        for each expressionmatched in expressionmatch
          match_positions(i, j) = expressionmatched.FirstIndex + 1

          'response.Write "<br>i is " & i
          'response.Write "<br>j is " & j
          'response.Write "<br>expressionmatched.FirstIndex is " & expressionmatched.FirstIndex + 1

          j = j + 1
        next
      
      end if
      
      set expressionmatch = Nothing
      g_set = g_set - 1

      'response.Write "<br>ozzzz is " & UBound(match_positions, i + 1)

    next

    value = PageRankArray(match_positions, min_matches, no_strings, points, text_to_search_length)
    'response.Write "<br>value is " & value
    
    RankResultsPage = value
    
  end function
  
  
  private function PageRankArray(match_positions, min_matches, no_strings, points, text_to_search_length)
    Dim i, j, minimum_matches, value
    
    value = 0
        
    'response.Write "<br>j is " & min_matches - 1

    for i = 0 to UBound(match_positions)
      for j = 0 to min_matches - 1
      
        'response.Write "<br>i is " & i
        'response.Write "<br>j is " & j
        'response.Write "<br>points is " & points
        'response.Write "<br>match_positions(i, j) is "
        'response.write match_positions(i, j) <> 0
        'response.Write "<br>text_to_search_length is " & text_to_search_length
      
        if match_positions(i, j) <> 0 then 'note that if the array is empty, match_positions(i, j) would print a blank but IsNill, IsNumeric still return false!
          'response.Write "<br>next is " & (text_to_search_length - match_positions(i, j)) / text_to_search_length
          'response.write "<br>value  = " & value & " + (" & text_to_search_length & " - " & match_positions(i, j) & ") / " & text_to_search_length & "*" & points
          value = value + (text_to_search_length - match_positions(i, j)) / text_to_search_length
        end if

      next
    next
    
    'response.Write "<br>text_to_search_length is " & text_to_search_length
    'response.Write "<br>value is " & points * value
    
    PageRankArray = points * value
    
  end function
  
  
  private function PhraseSearch
    Dim is_phrase_search, search_term 
    
    is_phrase_search = false
    search_term = Trim(YiderSearchTerm())
    
    if Left(search_term, 1) = """" and Right(search_term, 1) = """" then
      is_phrase_search = true
    end if
    
    PhraseSearch = is_phrase_search
  
  end function
  
  
  function SearchComplete(text_found, text_to_search, word_array)
    Dim complete, count, expressionmatch, RegExpObject, text_to_find, total
    
    'response.Write "<br>UBound(word_array) is " & UBound(word_array)
  
    set RegExpObject = New RegExp
    g_set = g_set + 1
    total = 0

    for count = 0 to UBound(word_array) 
    
      text_to_find = EscapeRegExpChars(word_array(count))
      text_to_find = WordIsolationPattern(text_to_find)

      'response.Write "<br><br>text_to_find is " & text_to_find
      'response.Write "<br>text_to_search is " & text_to_search

      With RegExpObject
        .Pattern = text_to_find
        .IgnoreCase = true
        .Global = True
      End With
             
      set expressionmatch = RegExpObject.Execute(text_to_search)
      
      'response.Write "<br>expressionmatch.count is " & expressionmatch.count

      if expressionmatch.count <> 0 then
        total = total + 1
      else
        Exit For
      end if
    
    next
    
    if UBound(word_array) + 1 = total then
      complete = true
    else
      complete = false
    end if
    
    set RegExpObject = Nothing
    g_set = g_set - 1
    
    SearchComplete = complete
    
  end function
  
  
  private function SearchTextDescription
    Dim description
  
    if PhraseSearch then
      description = g_phrase
    else
      description = g_text
    end if
    
    SearchTextDescription = description
    
  end function

  
  private function Start()
  
    Dim the_start
    
    the_start = Request.QueryString("start")
        
    if Len(the_start) > 0 then
      the_start = CInt(the_start)
    else
      the_start = 1
    end if
    
    Start = the_start
  
  end function
  
  
  private function StripExteriorQuotes(byval str)
    
    if Left(str, 1) = """" and Right(str, 1) = """" and Len(str) > 1 then
      str = Mid(str, 2, Len(str) - 2)
    end if
    
    StripExteriorQuotes = str

  end function
  
  
  'this function places a 'tag' around every instance of the strings in 'word_array' that occur in 'str'
  'eg str = "the cat in the hat is very lazy indeed"
  'word_array = Array("cat", "dog", "cow", "hat")
  'tag = "b"
  'this function returns the string
  ' "the <b>cat</b> in the <b>hat</b> is very lazy indeed" with matches set to 2
  private function TagWordsInArray(str, word_array, tag, byref matches)
    Dim match, word
    
    matches = 0
        
    for each word in word_array
    
      str = TagWordsInStr(str, word, tag, match)

      if match then
        matches = matches + 1      
      end if
            
    next
    
    TagWordsInArray = str
  end function
  
  
  'this function places a 'tag' around every instance of the string 'word' that occur in 'str'
  'eg str = "the cat in the hat is a very lazy cat indeed"
  'word = "cat"
  'tag = "b"
  'this function returns the string
  ' "the <b>cat</b> in the hat is a very lazy <b>cat</b> indeed"
  'match is set to true if there is at least one match
  public function TagWordsInStr(str, word, tag, byref match)
    
    Dim RegExpObject, expressionmatch, expressionmatched, lcase_str, new_str, pos_previous, pos_next
    
    pos_previous = 1
    lcase_str = str
     
    Set RegExpObject = New RegExp
    g_set = g_set + 1
    
    'str = "that we all know what a Functional Specification is and why you need one, it's time to "
    'response.Write "<br><br>str is '" & str & "'"
    'response.Write "<br>word '" & word & "'"
        
    With RegExpObject
      .Pattern = "\b" & word & "\b"
      .IgnoreCase = true
      .Global = True
    End With

    Set expressionmatch = RegExpObject.Execute(lcase_str)
    g_set = g_set + 1

    For Each expressionmatched in expressionmatch
      'Response.Write "<br>" & expressionmatched.Value & " was matched at position " & expressionmatched.FirstIndex
      
      'check if the first watch was with a non word
      if Mid(str, expressionmatched.FirstIndex + 1, 1) <> Left(word, 1) then
        pos_next = expressionmatched.FirstIndex + 1
      else
        pos_next = expressionmatched.FirstIndex
      end if
      
      new_str = new_str & Mid(str, pos_previous, pos_next - pos_previous + 1) & "<" & tag & ">" & Mid(str, pos_next + 1, Len(word)) & "</" & tag & ">"
      pos_previous = pos_next + Len(word) + 1
      
      'response.write "<br>pos_previous is " & pos_previous & " new_str is '" & new_str & "'"
      'response.write "<br>str is " & str

      match = true
    Next
    
    'response.write "<br>2 new_str is '" & new_str & "'"

    if match then
      new_str = new_str & Right(str, Len(str) - pos_previous + 1)
    else
      new_str = str
    end if

    'response.write "<br>3 new_str is '" & new_str & "'"

    Set RegExpObject = Nothing
    g_set = g_set - 1

    Set expressionmatch = Nothing
    g_set = g_set - 1
    
    TagWordsInStr = new_str
    
  end function


  'returns 'words_after' words after 'start' bytes from the beginning of the string 'str'
  private function WordsAfter(str, start, words_after)
	  Dim count, pos, ret

	  'Response.Write("<br> str is '" & str & "' start is " & start & " words_after is " & words_after)

    count = 0
    
    if Mid(start, start, 1) <> " " and Right(YiderSearchTerm(), 1) = "*" then
      pos = InStr(start, str, " ")
    else
      pos = start
    end if 
    
    if pos <> 0 then   

	    pos = InStr(NextNonSpace(str, pos), str, " ")

	    if pos <> 0 then
		    count = count + 1
	    end if

	    while(pos <> 0 and count < words_after)

        pos = InStr(NextNonSpace(str, pos + 1), str, " ")
      
		    if pos <> 0 then
		      count = count + 1
		    end if

	    wend
	  
	  end if

	  if pos = 0 then
	    pos = Len(str)
	  else
	    ret = " ..."
	  end if
	  
	  'response.Write "<br><br>start is " & start
	  'response.Write "<br>pos is " & pos
	  'response.Write "<br>str is " & str & " Len(str) is " & Len(str)

	  ret = RTrim(Mid(str, start, pos - start + 1)) & ret

	  'response.Write "<br>words after is " & ret

	  WordsAfter = ret

  end function
  

  'returns 'words_before' words after 'start' bytes from the beginning of the string 'str'
  private function WordsBefore(str, start, words_before)
    Dim count, pos, prev_non_space, ret

    count = 0

	  'Response.Write("<br><br>str is " & str)
	  'Response.Write("<br>start is " & start)
	  'Response.Write("<br>words_before is " & words_before)

    pos = InStrRev(str, " ", start)
    'InStrRev(str, " ", PreviousNonSpace(str, start))

	  'Response.Write("<br>pos is " & pos)

	  while(pos <> 0 and count < words_before)

	    prev_non_space = PreviousNonSpace(str, pos)

	    'Response.Write("<br>pos is " & pos & " prev_non_space is " & prev_non_space)
	    pos = InStrRev(str, " ", prev_non_space)

		  if pos <> 0 then
		    count = count + 1
		  end if

	  wend

	  if pos = 0 then
	    pos = 1
	  end if

	  ret = Mid(str, pos, start - pos)

	  if pos <> 1 then
	    ret = "... " & ret
	  end if

	  'Response.Write("<br>pos is " & pos & " start - pos is " & start - pos)
	  'Response.Write("<br>ret is '" & ret & "'<br><br>")
	  'Response.Write("<br>start is " & start)

	  WordsBefore  = ret

  end function

  
  private function WordIsolationPattern(byval word)
  
    Dim isolation, original_length
    
    isolation = m_isolation_vbscript
    
    'response.Write "<br>1 word is " & word

    word = replace(word, "*", "[^ \t\r\n\v\f\/-]+")
    
    'response.Write "<br>2 word is " & word
        
    if Left(word, 1) = "*" and Right(word, 1) <> "*" then
      word = "[^" & m_isolation_vbscript & "]+" & Right(word, Len(word) - 1) & m_isolation_vbscript
    
    elseif Left(word, 1) <> "*" and Right(word, 1) = "*" then
      word = m_isolation_vbscript & Left(word, Len(word) - 1) & "[^" & m_isolation_vbscript & "]+"
      
    elseif Left(word, 1) = "*" and Right(word, 1) = "*" then
      word = "[^" & m_isolation_vbscript & "]+" & Mid(word, 2, Len(word) - 2) & "[^" & m_isolation_vbscript & "]+"
      
    else
      word = m_isolation_vbscript & Left(word, Len(word)) & m_isolation_vbscript
    
    end if
        
    'multibyte stuff
    if Len(m_isolation_vbscript) = 0 then
      m_isolation_vbscript = replace(m_isolation_vbscript, "[]", "")
    end if
    
    WordIsolationPattern = word
    
  end function
  
  
  function YiderSearchTerm()
    Dim yider
    
    yider = Request("yider")
    
    if Len(yider) = 0 then
      yider = Request.QueryString("yider")
    end if
    
    YiderSearchTerm = yider
    
  end function
  


end class


%>