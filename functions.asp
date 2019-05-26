<%
'Copyright (c) Tom Kirby 2005
'All rights reserved
'This source code is subject to the licensing conditions at http://www.transworld.dyndns.ws
%>
<%Server.ScriptTimeout = 2500%>
<script language="JScript" runat="server">

function escape(str)
{
  return escape(str);
}

startTime = new Date();
function showTime()
{
  thisTime = new Date();
  Response.Write("<br><hr><font face=arial size=2>This form took ")
  Response.Write((thisTime.getTime() - startTime.getTime()) / 1000)
  Response.Write(" seconds to load from the server</font><hr>")
}

function showElapsedTime(str)
{
  var diff;
  
  thisTime = new Date();
  Response.Write(str)
  diff = thisTime.getTime() - startTime.getTime();
  
  Response.Write(diff + " milli seconds to process</font><hr>")
  startTime = new Date();
  
}

</script>

<%


function AddArrays(array1, array2)
  Dim array_size, count, new_array
  
  if not IsArray(array1) and not IsArray(array2) then
    new_array = 0
  else
  
    new_array = Array()
    
    if IsArray(array1) then
      for count = 0 to UBound(array1)
        array_size = UBound(new_array) + 1
        ReDim Preserve new_array(array_size)
        new_array(array_size) = array1(count)
      next
    end if
    
    if IsArray(array2) then
      for count = 0 to UBound(array2)
        array_size = UBound(new_array) + 1
        ReDim Preserve new_array(array_size)
        new_array(array_size) = array2(count)
      next
    end if
  
  end if
  
  AddArrays = new_array
  
end function


sub AddToArray(item, byref the_array)

  Dim array_size

  if not IsArray(the_array) then
    the_array = Array()
  end if
  
  array_size = UBound(the_array) + 1
  ReDim Preserve the_array(array_size)
  the_array(array_size) = item
  
end sub


'returns the minimum number in the array
function ArrayMinimum(arr)
  Dim count, min
  min = "n"
  
  if IsArray(arr) then
    
    for count = 0 to UBound(arr) - 1
    
      if not IsNumeric(min) then
        min = arr(count)
      elseif arr(count) < min then
        min = arr(count)
      end if
    
    next
  
  end if
  
  ArrayMinimum = min
  
end function


Function BinaryToString1(Binary, CharSet)
'http://www.pstruh.cz/tips/detpg_binarytostring.htm

  Const adTypeText = 2
  Const adTypeBinary = 1
  
  'Create Stream object
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")
  
  'Specify stream type - we want To save text/string data.
  BinaryStream.Type = adTypeBinary
  
  'Open the stream And write text/string data To the object
  BinaryStream.Open
  BinaryStream.Write Binary
  
  'Change stream type To binary
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeText
  
  'Specify charset For the source text (unicode) data.
  BinaryStream.CharSet = CharSet
  'response.Write Charset
    
  'Open the stream And get binary data from the object
  BinaryToString1 = BinaryStream.ReadText
  
End Function


Function BinaryToString(Binary)
  'Antonin Foller, http://www.pstruh.cz
  'Optimized version of a simple BinaryToString algorithm.
  
  Dim cl1, cl2, cl3, pl1, pl2, pl3
  Dim L
  cl1 = 1
  cl2 = 1
  cl3 = 1
  L = LenB(Binary)
  
  Do While cl1<=L
    pl3 = pl3 & Chr(AscB(MidB(Binary,cl1,1)))
    cl1 = cl1 + 1
    cl3 = cl3 + 1
    If cl3>300 Then
      pl2 = pl2 & pl3
      pl3 = ""
      cl3 = 1
      cl2 = cl2 + 1
      If cl2>200 Then
        pl1 = pl1 & pl2
        pl2 = ""
        cl2 = 1
      End If
    End If
  Loop
  BinaryToString = pl1 & pl2 & pl3
End Function


sub CreateRecordset(byref recordset)

  set recordset = Server.CreateObject("ADODB.Recordset")
  g_set = g_set + 1

  recordset.CursorType = 0
  recordset.CursorLocation = 1
  recordset.CacheSize = 1
  recordset.LockType = 1
  
end sub


'0 - Access
'1 - SQL Server
public function DataBaseType()
  Dim db_type

  if InStr(lcase(g_database_connection), ".mdb") <> 0 then
    db_type = 0
  else
    db_type = 1
  end if
   
  DataBaseType = db_type
  
end function
  
  
'returns the filename only of a path if its path and filename is in a string
function FileName(complete_file_name)
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


function FullTextEnabled
  Dim enabled

  if g_full_text and InStr(LCase(g_database_connection), "uid") then
    enabled = true
  else
    enabled = false
  end if
  
  FullTextEnabled = enabled

end function


'returns an Array from a delimited string
'eg "a,ddd,dd" would return GetArrayString(0) = "a", GetArrayString(1) = "ddd", GetArrayString(2)="dd"
'all strings are trimed
function GetArrayString(array_def, delimiter)
  Dim array_size, next_pos, old_pos, str, the_array

  the_array = Array()
  next_pos = InStr(1, array_def, delimiter)
  old_pos = 1
  
  'response.Write "<br>array_def is " & array_def
  'response.Write "<br>next_pos is " & next_pos
  
  while old_pos <> 0

    if next_pos = 0 then
      str = Mid(array_def, old_pos)
      old_pos = 0
    else
      str = Mid(array_def, old_pos, next_pos - old_pos)
    end if

    if Len(str) > 0 then
      if UBound(the_array) = -1 then
        the_array = Array(Trim(str))
      else
        array_size = UBound(the_array) + 1
        ReDim Preserve the_array(array_size)
        the_array(array_size) = Trim(str)
      end if
      
      'response.Write "<br>trimed it is " & Trim(str)
      'response.Write "<br>array is " & the_array(0)
    end if

    if old_pos <> 0 then
      old_pos = next_pos + 1
      next_pos = InStr(old_pos + 1, array_def, delimiter)
    end if
    
    'response.write "<br>old_pos is " & old_pos & " next_pos is " & next_pos & " str is '" & str & "'" & " UBound(the_array) is " & UBound(the_array)

  wend
  
  'response.Write "<br>the_array is " & the_array(0)
  'response.Write "<br>the_array is " & the_array(1)
 

  GetArrayString = the_array

end function


'eg c:\yider\
function GetPath
  Dim path, pos
  
  if Len(g_path) > 0 then
    path = g_path
    
    if Right(path, 1) <> "\" then
      path = path & "\"
    end if
    
    'Response.Write("<br>1 path is " & path)

  else
    'Response.Write("<br>2 path is " & path)
    path = Request.ServerVariables("PATH_TRANSLATED")
    pos = InStrRev(path, "\")
    path = Left(path, pos)
  end if


  GetPath = path 'replace(path, "\", "/")

end function


'returns true if regexp_to_find match any of the strings in array_to_search
function InArrayRegExp(text_to_search, regexp_to_find, ignore_case, RegExpObject)
  Dim expressionmatch, found, no

  found = false
  no = 0

  if Len(regexp_to_find) > 0 then
  
    With RegExpObject
      .Pattern = regexp_to_find
      .IgnoreCase = ignore_case
      .Global = true
    End With
    
    set expressionmatch = RegExpObject.Execute(text_to_search)
    
    if expressionmatch.Count <> 0 then
      found = true
    end if

    'response.Write "<br><br>text_to_search is " & text_to_search
    'response.Write "<br>regexp_to_find is " & regexp_to_find
    'response.Write "<br>found is " & found

  end if

  InArrayRegExp = found
end function


'this function returns true if 'element_to_find' is within 'the_array' but examines
'only every 'nth' element starting at element 'start'
'0 is the first element
'if 'the_array' is not an array it returns 0
'if 'the_array' is in the array it returns its position
function InArraySkip(element_to_find, the_array, nth, start)
  Dim count, found
  
  found = false
    
  if IsArray(the_array) then

    for count = start to UBound(the_array) step nth 
      if element_to_find = the_array(count) then
        found = true
        Exit For
      end if
    next
    
    if not found then
    'not found
      count = -1
    end if
  
  end if  
  
  InArraySkip = count
  
end function


'returns true if the string 'search_for' is in contained in any of the strings in the array 'array_to_search'
function InArrayStr(search_for, array_to_search)
  Dim found, no, str

  found = false
  no = 0

  if IsArray(array_to_search) then

    while not found and no <= UBound(array_to_search)
    
      str = array_to_search(no)
    
      if Len(str) > 0 then
           
        if InStr(CStr(search_for), CStr(str)) <> 0 then
          found = true
        end if
      
      end if

      no = no + 1
    wend

  end if

  InArrayStr = found
end function


'returns true if str_to_find is equal to any of the strings in array_to_search
function InArrayStrExact(str_to_find, array_to_search)
  Dim found, no

  found = false
  no = 0

  if IsArray(array_to_search) then

    while not found and no <= UBound(array_to_search)
          
      if CStr(str_to_find) = CStr(array_to_search(no)) then
        found = true
      end if

      no = no + 1
    wend

  end if

  InArrayStrExact = found
end function


'returns true if any of the elements in the_array can be found in string_to_search
function InStrArray(string_to_search, the_array)
  Dim in_str, str
  
  in_str = false

  if IsArray(the_array) then
  
    for each str in the_array
    
      if Len(str) > 0 then
    
        if InStr(string_to_search, str) <> 0 then
          in_str = true
          exit for
        end if
      
      end if
      
    next
    
  end if
  
  InStrArray = in_str
  
end function


sub JavaScriptAlert(message)

  response.Write NewLine(0) & "<script language=""JavaScript"">"
  response.Write NewLine(0) & "alert('" & message & "')"
  response.Write NewLine(0) & "</script>"
        
end sub


function MultiByteCharSet(charset)

  Dim mulitbyte

  Select Case charset
  
    Case "ASMO-708" mulitbyte = false
    Case "DOS-720" mulitbyte = false
    Case "iso-8859-6" mulitbyte = false
    Case "x-mac-arabic" mulitbyte = false
    Case "windows-1256" mulitbyte = false
    Case "ibm775" mulitbyte = false
    Case "iso-8859-4" mulitbyte = false
    Case "windows-1257" mulitbyte = false
    Case "ibm852" mulitbyte = false
    Case "iso-8859-2" mulitbyte = false
    Case "x-mac-ce" mulitbyte = false
    Case "windows-1250" mulitbyte = false
    Case "x-Europa" mulitbyte = false
    Case "x-IA5-German" mulitbyte = false
    Case "ibm737" mulitbyte = false
    Case "iso-8859-7" mulitbyte = false
    Case "x-mac-greek" mulitbyte = false
    Case "windows-1253" mulitbyte = false
    Case "ibm869" mulitbyte = false
    Case "DOS-862" mulitbyte = false
    Case "iso-8859-8-i" mulitbyte = false
    Case "iso-8859-8" mulitbyte = false
    Case "x-mac-hebrew" mulitbyte = false
    Case "windows-1255" mulitbyte = false
    Case "x-EBCDIC-Arabic" mulitbyte = false
    Case "x-EBCDIC-CyrillicRussian" mulitbyte = false
    Case "x-EBCDIC-CyrillicSerbianBulgarian" mulitbyte = false
    Case "x-EBCDIC-DenmarkNorway" mulitbyte = false
    Case "x-ebcdic-denmarknorway-euro" mulitbyte = false
    Case "x-EBCDIC-FinlandSweden" mulitbyte = false
    Case "x-ebcdic-finlandsweden-euro" mulitbyte = false
    Case "x-ebcdic-finlandsweden-euro" mulitbyte = false
    Case "x-ebcdic-france-euro" mulitbyte = false
    Case "x-EBCDIC-Germany" mulitbyte = false
    Case "x-ebcdic-germany-euro" mulitbyte = false
    Case "x-EBCDIC-GreekModern" mulitbyte = false
    Case "x-EBCDIC-Greek" mulitbyte = false
    Case "x-EBCDIC-Hebrew" mulitbyte = false
    Case "x-EBCDIC-Icelandic" mulitbyte = false
    Case "x-ebcdic-icelandic-euro" mulitbyte = false
    Case "x-ebcdic-international-euro" mulitbyte = false
    Case "x-EBCDIC-Italy" mulitbyte = false
    Case "x-ebcdic-italy-euro" mulitbyte = false
    Case "X-EBCDIC-Spain" mulitbyte = false
    Case "x-ebcdic-spain-euro" mulitbyte = false
    Case "x-EBCDIC-Thai" mulitbyte = false
    Case "CP1026" mulitbyte = false
    Case "x-EBCDIC-Turkish" mulitbyte = false
    Case "x-EBCDIC-UK" mulitbyte = false
    Case "x-ebcdic-uk-euro" mulitbyte = false
    Case "ebcdic-cp-us" mulitbyte = false
    Case "x-ebcdic-cp-us-euro" mulitbyte = false
    Case "ibm861" mulitbyte = false
    Case "x-mac-icelandic" mulitbyte = false
    Case "iso-8859-3" mulitbyte = false
    Case "iso-8859-15" mulitbyte = false
    Case "x-IA5-Norwegian" mulitbyte = false
    Case "IBM437" mulitbyte = false
    Case "x-IA5-Swedish" mulitbyte = false
    Case "ibm857" mulitbyte = false
    Case "iso-8859-9" mulitbyte = false
    Case "x-mac-turkish" mulitbyte = false
    Case "windows-1254" mulitbyte = false
    Case "us-ascii" mulitbyte = false
    Case "ibm850" mulitbyte = false
    Case "x-IA5" mulitbyte = false
    Case "iso-8859-1" mulitbyte = false
    Case "macintosh" mulitbyte = false
    Case "Windows-1252" mulitbyte = false


    Case "EUC-CN" mulitbyte = true
    Case "hz-gb-2312" mulitbyte = true
    Case "x-mac-chinesesimp" mulitbyte = true
    Case "big5" mulitbyte = true
    Case "x-Chinese-CNS" mulitbyte = true
    Case "x-Chinese-Eten" mulitbyte = true
    Case "x-mac-chinesetrad" mulitbyte = true
    Case "cp866" mulitbyte = true
    Case "iso-8859-5" mulitbyte = true
    Case "koi8-r" mulitbyte = true
    Case "koi8-u" mulitbyte = true
    Case "x-mac-cyrillic" mulitbyte = true
    Case "windows-1251" mulitbyte = true
    Case "x-EBCDIC-JapaneseAndKana" mulitbyte = true
    Case "x-EBCDIC-JapaneseAndJapaneseLatin" mulitbyte = true
    Case "x-EBCDIC-JapaneseAndUSCanada" mulitbyte = true
    Case "x-EBCDIC-JapaneseKatakana" mulitbyte = true
    Case "x-EBCDIC-KoreanAndKoreanExtended" mulitbyte = true
    Case "x-EBCDIC-KoreanExtended" mulitbyte = true
    Case "CP870" mulitbyte = true
    Case "x-EBCDIC-SimplifiedChinese" mulitbyte = true
    Case "x-EBCDIC-Thai" mulitbyte = true
    Case "x-EBCDIC-TraditionalChinese" mulitbyte = true
    Case "x-iscii-as" mulitbyte = true
    Case "x-iscii-be" mulitbyte = true
    Case "x-iscii-de" mulitbyte = true
    Case "x-iscii-gu" mulitbyte = true
    Case "x-iscii-ka" mulitbyte = true
    Case "x-iscii-ma" mulitbyte = true
    Case "x-iscii-or" mulitbyte = true
    Case "x-iscii-pa" mulitbyte = true
    Case "x-iscii-ta" mulitbyte = true
    Case "x-iscii-te" mulitbyte = true
    Case "euc-jp" mulitbyte = true
    Case "iso-2022-jp" mulitbyte = true
    Case "iso-2022-jp" mulitbyte = true
    Case "csISO2022JP" mulitbyte = true
    Case "x-mac-japanese" mulitbyte = true
    Case "shift_jis" mulitbyte = true
    Case "ks_c_5601-1987" mulitbyte = true
    Case "euc-kr" mulitbyte = true
    Case "iso-2022-kr" mulitbyte = true
    Case "Johab" mulitbyte = true
    Case "x-mac-korean" mulitbyte = true
    Case "windows-874" mulitbyte = true
    Case "unicode" mulitbyte = true
    Case "unicodeFFFE" mulitbyte = true
    Case "utf-7" mulitbyte = true
    Case "utf-8" mulitbyte = true
    Case "gb2312" mulitbyte = true
    Case "windows-1258" mulitbyte = true

	  'return true to be conservative, multi byte charsets have more restrictions than non-multibyte charsets
	  Case else mulitbyte = true 
	  
  End Select
  
  'response.Write "MultiByteCharSet is " & charset & " "
  'response.Write mulitbyte
  
  MultiByteCharSet = mulitbyte
    
end function


'returns the number of strings in word_array that occur in str_to_search
'the strings in word_array are prefix by 'prefix' and postfixed by 'postfix'
'this is a case insensitive comparison
function NumberMatchesInStr(byval str_to_search, word_array, prefix, postfix)

  Dim no, word

  no = 0
  str_to_search = lcase(str_to_search)

  if IsArray(word_array) then

    for each word in word_array
      
      if InStr(str_to_search, prefix & lcase(word) & postfix) <> 0 then
        no = no + 1
      end if
      
      'response.Write "<br><br>str_to_search is '" & str_to_search & "' word is '" & word & "' " & Asc(Mid(str_to_search, 22, 1)) & " " & Mid(str_to_search, 22, 5)
      
    next
    
  end if

  NumberMatchesInStr = no
  
end function


function NewLine(spaces)
  NewLine = vbCrLf & Space(spaces)
end function


'removes all occurences of 'the_string' in the array 'the_array'
function RemoveString(the_array, the_string)
  Dim new_array, size, value

  new_array = Array()
  
  for each value in the_array
  
    if value <> the_string then
        
      if UBound(new_array) = -1 then
        new_array = Array(value)
      else
        size = UBound(new_array) + 1
        ReDim Preserve new_array(size)
        new_array(size) = value
      end if
      
    end if
    
  next

  RemoveString = new_array

end function


function RemoveDuplicates(the_array)
  Dim count, replacement_array
  
  replacement_array = Array()
      
  if IsArray(the_array) then

    for count = 0 to UBound(the_array)
          
      if not InArrayStrExact(the_array(count), replacement_array) then

        ReDim Preserve replacement_array(UBound(replacement_array) + 1)
        replacement_array(UBound(replacement_array)) = the_array(count) 
  
      end if
      
    next
  
  end if
  
  RemoveDuplicates = replacement_array
  
end function


function RemoveCarriageReturn(str)

  str = replace(str, vbcrlf, "\n")
  str = replace(str, vbcr, "\n")
  str = replace(str, vblf, "\n")
  str = replace(str, "'", "\'")
  
  RemoveCarriageReturn = str

end function


'replaces the string between begin_rep and end_rep with 'replacement'
'this function checks for all cases of begin_rep/end_rep
'ReplaceBetweenStrings("here is a string 111 to replace 222 man", "111", "222", "333") returns "here is a string 333 man"
'ReplaceBetweenStrings("here is a string 111 to replace 222", "111", "222", "333") returns "here is a string 333"
'ReplaceBetweenStrings("here is a string 111 to replace 222", "222", "111", "333") returns "here is a string 111 to replace 222"
private function ReplaceBetweenStrings(str, begin_rep, end_rep, replacement)
  Dim tag_start, tag_end, str_lcase, str_left, str_right
  
  'response.Write "<br>begin_rep is " & server.htmlencode(begin_rep) & " end_rep is " & server.htmlencode(end_rep)
  'response.Write "<br><br>1 str is " & server.HTMLEncode(str)
  
  str_lcase = lcase(str)
  begin_rep = lcase(begin_rep)
  end_rep = lcase(end_rep)

  tag_start = InStr(str_lcase, begin_rep)
  
  while tag_start <> 0
  
    if tag_start <> 0 then
      tag_end = InStr(tag_start + Len(begin_rep), str_lcase, end_rep)
    end if
    
    'response.Write "<br>tag_start is " & tag_start
    'response.Write "<br>tag_end is " & tag_end
    
    if tag_start <> 0 and tag_end <> 0 then
      str_left = Left(str, tag_start - 1)
      str_right = Mid(str, tag_end + Len(end_rep))
      str = str_left + replacement + str_right
      str_lcase = lcase(str)
      tag_start = InStr(tag_start, str_lcase, begin_rep)
    else
      tag_start = 0
    end if
    
  wend
       
  'response.Write "<br>2 str is " & server.HTMLEncode(str)

  ReplaceBetweenStrings = str
  
end function 


sub ResponseWriteArray(the_array)
  Dim count

  if IsArray(the_array) then
    for count = 0 to UBound(the_array)
      response.write "<br>count is " & count & " value is " & the_array(count)
    next
  end if
        
end sub


function MakeQueryStringEmailable(str)
  
  str = replace(str, "http:", "http")
  str = replace(str, "/", "slash")
  str = replace(str, ".", "dot")
  str = replace(str, " ", "chr32")
  str = replace(str, vbcrlf, "vbcrlf")
  
  MakeQueryStringEmailable = str
  
end function


sub SendErrorReport(error_type)
  Dim body, XMLHttp, url, url_to_spider
  set XMLHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
  
  body = "Error type is " & error_type & vbcrlf & "g_url_to_spider is " & g_url_to_spider
  body = body & vbcrlf & "Error number is " & err.number
  body = body & vbcrlf & "Error Description is " & err.Description

  if error_type = "search" then
    body = body & vbcrlf & "Search term is " & Request("yider")
  end if

  body = body & vbcrlf & "Email address is " & g_your_email_address
  
  
  'response.Write body
  body = MakeQueryStringEmailable(body)
  
  url = "http://www.yart.com.au/admin/general_asp_functions/senderror.asp?number=" & MakeQueryStringEmailable(err.number) & "&email=" & MakeQueryStringEmailable(g_your_email_address) & "&the_error=" & body
    
  on error resume next 'we might not be on the Internet
  
  XMLHttp.open "GET", url, false
  XMLHttp.send()
  
  set XMLHttp = Nothing
end sub


function TextStringToBinaryString(str)
  Dim bin_str, count, next_char, str_len

  str_len = Len(str)

  'turn str into a binary string
  for count = 1 to str_len
    next_char = Mid(str, count, 1)
	  bin_str = bin_str & ChrB(AscB(next_char))
  next

  TextStringToBinaryString = bin_str

end function

  
%>