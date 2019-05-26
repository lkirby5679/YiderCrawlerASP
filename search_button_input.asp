<%
'Copyright (c) Tom Kirby 2005
'This program and its associated source code is distributed under the terms of the GNU General Public 
'License.
'See the attached file COPYING.txt for more information
%>
<%Server.ScriptTimeout = 2500%>
<form name="yider_form" onsubmit="OnSearch(this);">
<table cellpadding="0" cellspacing="0" border="0">
  <tr>

<%    
    Response.Write "<td style=""" & g_search_style & """> " & g_text_search & "&nbsp;<input type=""text"" name=""yider"" size=""" & g_search_box_length & """ maxlength=""100"" value=""" & replace(Request("yider"), """", "&quot;") & """ style=""" & g_search_style & """></td>"
%>

    <!--HTML button start-->
    <td>
<%  
    Response.Write "&nbsp;<input type=""button"" value=""" & g_text_go  & """ onClick=""OnSearch(document.yider_form); return false;"" style=""" & g_search_style & """>&nbsp;"
%>
    </td>
    <!--HTML button end-->

    <!--Alternative 3 - Help-->
    <td>
<% 
    if not MultiByteCharSet(g_charset) then
      Response.Write "&nbsp;<a href="""" border=""0"" style=""font-family:Arial;font-size=10pt"" onclick=""OnHelp(); return false;"">" & g_text_help & "</a>&nbsp;"
    end if
%>
    </td>
    <!--Alternative 3 ends here-->

  </tr>
  
</table>

<%

  if FullTextEnabled then
    Response.Write "<script language=""JavaScript"">function OnHelp() {var width, height, left, top; width=400; height=400; left=(screen.width-width)/2; top=((screen.height-50)-height)/2; window.open('search_help_ft.htm', 'win', 'width=' + width + ',height=' + height + ',alwaysLowered=1,alwaysRaised=0,channelmode=0,dependent=0,directories=0,fullscreen=0,hotkeys=0,location=0,menubar=0,resizable=0,scrollbars=1,status=0,titlebar=0,toolbar=0,left=' + left + ',top=' + top + ',z-lock=0'); return false;} function OnSearch(form) { var ok;  ok = true; if(form.yider.value.length == 0) {alert('You must enter some text to search for!'); form.yider.focus(); ok = false;} else if(form.yider.value.search(/\*/) != -1 && form.yider.value.search(/\s/) != -1) {alert('If you are using the * to generalise your search, you can only search for one word at a time!'); form.yider.focus(); ok = false;} else if(form.yider.value.search(/^\S+\*\S+$/) != -1) { alert('You can only use the * at the beginning or end of your search term'); form.yider.focus(); ok = false; } else if(form.yider.value == '*') {alert('You cannot search for a general * character!\nSee the search Help'); form.yider.focus(); ok = false;} if(ok){form.action = '" & g_results & "'; form.submit();} }</script>"
  else
    Response.Write "<script language=""JavaScript"">function OnHelp() {var width, height, left, top; width=400; height=400; left=(screen.width-width)/2; top=((screen.height-50)-height)/2; window.open('search_help.htm', 'win', 'width=' + width + ',height=' + height + ',alwaysLowered=1,alwaysRaised=0,channelmode=0,dependent=0,directories=0,fullscreen=0,hotkeys=0,location=0,menubar=0,resizable=0,scrollbars=1,status=0,titlebar=0,toolbar=0,left=' + left + ',top=' + top + ',z-lock=0'); return false;} function OnSearch(form) { var ok;  ok = true; if(form.yider.value.length == 0) {alert('You must enter some text to search for!'); form.yider.focus(); ok = false;} else if(form.yider.value.search(/\*/) != -1 && form.yider.value.search(/\s/) != -1) {alert('If you are using the * to generalise your search, you can only search for one word at a time!'); form.yider.focus(); ok = false;} else if(form.yider.value.search(/^\S+\*\S+$/) != -1) { alert('You can only use the * at the beginning or end of your search term'); form.yider.focus(); ok = false; } else if(form.yider.value == '*') {alert('You cannot search for a general * character!\nSee the search Help'); form.yider.focus(); ok = false;} if(ok){form.action = '" & g_results & "'; form.submit();} }</script>"
  end if

%>

</form>