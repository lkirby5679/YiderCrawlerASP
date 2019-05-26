<%
'Copyright (c) Tom Kirby 2005
'This program and its associated source code is distributed under the terms of the GNU General Public 
'License.
'See the attached file COPYING.txt for more information
%>
<%Server.ScriptTimeout = 25000%>
<%

Dim g_and, g_bad_page_strings, g_button_filename, g_charset, g_compact, g_database_connection, g_default_documents, g_delete_between_tags, g_delete_between_tags_complete, g_english, g_full_text, g_local_ID, g_max_pages, g_next, g_no_pages, g_open, g_page_was_found, g_pages_were_found, g_password, g_path, g_pause, g_phrase, g_previous, g_result_pages, g_results, g_search_box_length, g_search_maintain_search_term, g_search_maintain_url_params, g_search_style, g_set, g_strip_url_parameters, g_style_br, g_style_more_res, g_style_more_results_link, g_style_more_results_link_Next, g_style_more_results_text, g_style_noise_word, g_style_title, g_style_text, g_style_url, g_style_you_searched_for, g_text, g_text_go, g_text_help, g_text_search, g_trailing_words, g_urls_not_to_view, g_urls_per_iteration, g_url_to_spider, g_urls_to_view_not_store, g_use_keywords, g_username, g_valid_file_extensions, g_valid_url_strings, g_you_searched, g_your_email_address

'if you installed the Yider files in one directory, leave the setting below
'if you installed the Access databsase in a different directory, you need to modify g_path to point to this directory
'eg g_path = "c:\a_directory_somewhere\access_databases\"
g_path = GetPath

'database connection string
'SQL Server 2000 connection string
'g_database_connection = "PROVIDER=SQLOLEDB;SERVER=10.0.0.7;Database=yider;UID=sa;PWD=227oct36"

'Microsoft Access connection string
g_database_connection = "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & g_path & "yider.mdb"

'start spidering at this url
g_url_to_spider = "http://database.transworldinteractive.net"

'every url the Yider parses must contain this string in the url or it will be ignored
'g_valid_url_strings = Array("(www.twiportal.com)|(www.transworldinteractive.net)", true) 
'g_valid_url_strings = Array("www.twiportal.com", true) 
g_valid_url_strings = 0

'every url extension the Yider parses must be a page which satisfies a regular expression match
g_valid_file_extensions = Array("(htm)|(html)|(asp)|(php)|(jsp)|(cfm)", true)

'maximum pages for the Yider to spider
g_max_pages = 500000

'your email address
'please complete this so we can contact you to debug the Yider if it crashes during use
g_your_email_address = "administrator@transworldinteractive.net"

'set compact to true if you want your database compacted when it is cleared
'g_compact = false
g_compact = true

'the Yider will not duplicate storage of pages that can be obtained by URLs that require the web server to
'issue default documents e.g:
g_default_documents = Array("index.htm", "index.asp", "default.htm", "default.asp", "default.aspx")
'g_default_documents = 0

'url's that contain these strings will not be spidered e.g:
'g_urls_not_to_view = Array("http://www.yart.com.au/finances", true) or
'g_urls_not_to_view = Array("(http://www.yart.com.au/finances)|(http://www.yart.com.au/pays_due)", true) for multiple urls
'g_urls_not_to_view = 0
g_urls_not_to_view = Array("(http://click)|(http://ad)", true)

'url's that contain these strings will be spidered for further urls but not themselves stored for searching e.g.:
'g_urls_to_view_not_store = Array("http://www.yart.com.au/finances", true) or
'g_urls_to_view_not_store = Array("(http://www.yart.com.au/finances)|(http://www.yart.com.au/pays_due)", true) for multiple urls
g_urls_to_view_not_store = 0

'web pages (be careful - that's web pages, not URLs) that contain these strings will be spidered for further urls 
'but not themselves stored for searching e.g:
'g_bad_page_strings = Array("bad string", true) or
'g_bad_page_strings = Array("(bad string)|(another bad string)", true) for multiple strings
g_bad_page_strings = 0

'pages linking to search results will have these parameters appended if they can be obtained from the 
'ASP Request("var_name") function, see the Yider's help e.g:
'Array("foo", "param")
g_search_maintain_url_params = 0

'if given a string value, the search term is appended to the document hyperlinks in the search results pages
'g_search_maintain_search_term = "yider_search_term"
g_search_maintain_search_term = ""

'delete text between the following tags for searching but not spidering e.g:
'g_delete_between_tags = Array("<noindex>", "</noindex>", "<!--", "-->") 
g_delete_between_tags = 0

'delete text between the following tags for both spidering and searching e.g:
'g_delete_between_tags_complete = Array("<noindex>", "</noindex>", "<!--", "-->") 
g_delete_between_tags_complete = 0

'strip parameters from an URL when they are found 
'e.g g_strip_url_parameters = Array("offset")
'This would result in url's of the form http://www.website.com?offset=1&page=2 being stored in the database as
'http://www.website.com?page=2 i.e. the value of "offset" is stripped
g_strip_url_parameters = 0
'g_strip_url_parameters = Array("cmd")

'if true, the Yider will add the keywords from the meta tag to the beginning of the text in each web page
'g_use_keywords = false
g_use_keywords = true

'Optional Yider variables you can complete to run population.asp
g_urls_per_iteration = 10
g_pause = 10

'username for web servers that require authentication
g_username = ""

'password for web servers that require authentication
g_password = ""

'Optional Yider variables you probably won't want to change
g_search_style = "font-family:Arial; font-size:10pt; color:#000000"
g_style_you_searched_for = "font-family:Arial; font-size:10pt; color:#000000"
g_style_title = "font-size:12pt; font-family:Arial"
g_style_text = "font-size:10pt; font-family:Arial"
g_style_noise_word = "font-size:10pt; font-family:Arial; color:#6f6f6f"
g_style_url = "font-size:10pt; font-family:Arial; color:green"
g_style_br = "font-size:10pt; font-family:Arial"
g_style_more_results_text = "font-size:10pt; font-family:Arial"
g_style_more_results_link = "font-size:10pt; font-family:Arial"
g_style_more_results_link_Next = "font-size:12pt; font-family:Arial; font-weight:bold; color:#00c"

'variables governing searching
g_results = "search.asp"
g_button_filename = "button_go.jpg"
g_trailing_words = 10
g_search_box_length = 50

'languages other than English
g_english = true
g_charset = "windows-1250"

'text on the page search.asp
g_text_search = "Search:"
g_text_go = "Go"
g_text_help = "Help"

'text on the page search.asp when results are displayed
g_result_pages = "Result pages"
g_next = "Next"
g_previous = "Previous"
'the next 4 variables are used to construct the sentence:
'You searched for the text 'xml' and 1 page was found
'or
'You searched for the text 'xml' and 5 pages were found
g_you_searched = "You searched for the "
g_and= "and"
g_page_was_found = "page was found"
g_pages_were_found = "pages were found"
'the next 3 variables are used to construct a sentence when no match is found:
'There are no pages that contain the search phrase 'xml xslt'
'or
'There are no pages that contain the search text 'xml'
g_no_pages = "There are no pages that contain the search"
g_phrase = "phrase"
g_text = "text"

'if you use a SQL Server database, setting this variable to true will enable full text searching
g_full_text = true

g_local_ID = "0"         
'This is the default for g_local_ID and you should stick with it if:
'a) None of the below apply or
'b) More than one of the below
'Otherwise, comment out the line above and uncomment one of the lines below

'g_local_ID = "0x0804"    'Chinese_Simplified 
'g_local_ID = "0x0404"    'Chinese_Traditional
'g_local_ID = "0x0413"    'Dutch
'g_local_ID = "0x0809"    'English_UK
'g_local_ID = "0x0409"    'English_US
'g_local_ID = "0x040c"    'French
'g_local_ID = "0x0407"    'German
'g_local_ID = "0x0410"    'Italian
'g_local_ID = "0x0411"    'Japanese
'g_local_ID = "0x0412"    'Korean
'g_local_ID = "0x0c0a"    'Spanish_Modern
'g_local_ID = "0x041d"    'Swedish_Default

%>


<%

'forget about these variables, they are used for debugging and won't do a thing unless the 
'string 127.0.0.1 appears in URL's you Yider
g_set = 0
g_open = 0

%>