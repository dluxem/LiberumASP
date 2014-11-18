<%@ LANGUAGE="VBScript" %>
<% 
  Option Explicit
  'Buffer the response, so Response.Expires can be used
  Response.Buffer = TRUE
%>


<?xml version="1.0"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

  <!--
  Liberum Help Desk, Copyright (C) 2000-2001 Doug Luxem
  Liberum Help Desk comes with ABSOLUTELY NO WARRANTY
  Please view the license.html file for the full GNU General Public License.

  Filename: status.asp
  Date:     $Date: 2002/06/15 23:49:20 $
  Version:  $Revision: 1.50.4.1 $
  Purpose:  Page to update the status information for the In/Out Board.
  -->

  <!-- 	#include file = "../public.asp" -->
  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid

    ' Look for the problem id in the query string, if it's not
    ' there display an error.
    
    If Len(Request.QueryString("id")) = 0 Then
    	Response.Write("<title>" & lang(cnnDB, "ERROR") & "</title></head><body>")
    	cnnDB.Close
    	Call DisplayError(3, lang(cnnDB, "NovalidIDgiven"))
    End If
  %>
  <head>
    <title>
      <% = Cfg(cnnDB, "SiteName") %>&nbsp;<%=lang(cnnDB, "InOutBoard")%>
    </title>
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>
  <%
    ' See if authenticated
    Call CheckUser(cnnDB, sid)

  	'Check if user is In/Out Board admin
    dim inoutadmin
    inoutadmin = Usr(cnnDB, sid, "inoutadmin")

    'Call DisplayHeader(cnnDB, sid)
  %>
  <div align="center">
  <table class="normal">
      <tr class="Head1">
    <td>
	    <% = Cfg(cnnDB, "SiteName") %>
      <br />
      <%=lang(cnnDB, "StatusInformation")%>
    </td>
  </tr>
  </table>

  <%
  	Dim strSql, rstUser, rstUpdate 
  	Dim usid
  	usid = request.querystring("id")
  
  	If request.form("save") = "1" then
  		Dim sStatus, Statustext, StatusDate
  		Dim frm_status, frm_statustext
  		if request.form("frm_status") = "on" then
  			sStatus = 1
  			Statustext = request.form("frm_statustext")
        statustext = Replace(statustext , "'", "''")
  		else
  			sStatus = 0
  			Statustext = ""
  		end if
  		StatusDate = SQLDate(Now, lhdAddSQLDelim)
  		strSql = "UPDATE tblUsers SET " &_
  			"statuscode = " & sStatus & ", " &_
  			"statustext = '" & statustext & "', " &_
  			"statusdate = " & statusdate & " " &_
  			"WHERE sid = " & usid
  		set rstUpdate = SQLQuery(cnnDB, strSql)
  	end if
  	
  	strSQL = "SELECT * FROM tblUsers WHERE sid=" & usid
  	set rstUser = SQLQuery(cnnDB, strSql)
  
  	if rstUser.EOF then
  		Call DisplayError(3, lang(cnnDB, "Noresultsfound"))
  	end if
  
  	dim uname, check
  	uname = rstUser("fname")
  	if rstUser("statuscode") > 0 then
  		check = "checked"
  	else
  		check = ""
  	end if
  	
  	if rstUser("statuscode") = 2 and inoutadmin <> 1 then
  	  Call DisplayError(3, lang(cnnDB, "OnlyAdministratorscanchangethisstatus"))
  	else
    %>
      <table class="inout">
      <form method="POST" action="status.asp?id=<% = usid %>">
      <input type="hidden" name="save" value="1">
      <% 
        If Request.Form("save") = "1" Then
        	response.write "<tr class=""head2""><td colspan=2><div align=""center"">" & lang(cnnDB, "Status") & " " & lang(cnnDB, "isupdated") & "</div></td></tr>"
        end if
      %>
      <tr class="body1"><td><%=lang(cnnDB, "Fullname")%><br></td><td><b><%= uname %></b></td></tr>
      <tr class="body1"><td><%=lang(cnnDB, "Out")%></td><td><input type="checkbox" name="frm_status" <% = check %>></td></tr>
      <tr class="body1"><td valign="top"><%=lang(cnnDB, "Text")%></td><td><TEXTAREA NAME="frm_statustext" ROWS=4 COLS=47><% =rstUser("statustext") %></TEXTAREA></td></tr>
      <tr class="body1"><td align="center" colspan="2"><input type="submit" value="<%=lang(cnnDB, "Submit")%>"></td></tr>
      <tr class="body1"><td align="right" colspan="2"><a href="details.asp?id=<% = usid %>"><%=lang(cnnDB, "Details")%></a></td></tr>
      </table></form>	
  <%
    end if
  	rstUser.close
  	Call DisplayFooter(cnnDB, sid)
  	cnnDB.Close
  %>
  </div>
</body>
</html>
		