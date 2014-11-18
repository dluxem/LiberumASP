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

  Filename: details.asp
  Date:     $Date: 2001/12/13 21:47:38 $
  Version:  $Revision: 1.51 $
  Purpose:  This is the details form for the In/Out Board.
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
      ' See is user is validated
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
        <%=lang(cnnDB, "UserDetails")%>
      </td>
    </tr>
    </table>
    <%	
    	dim usid
    	usid = request.QueryString("id")
    
    	Dim strSql, rstUser
    	strSql = "SELECT tblUsers.*, departments.dname" &_
    		" FROM tblUsers, departments " &_
    		" WHERE (sid = " & usid & ")" &_
    		" AND (tblUsers.department = departments.department_id) "
    
    	Set rstUser = SQLQuery(cnnDB, strSql)

    	' If no results are returned, display an error
    	If rstUser.EOF Then
    		cnnDB.Close
    		'Call DisplayError(3, lang(cnnDB, "Noresultsfound"))
    	end if
    
    	dim uname, uid
    	uname = rstUser("fname")
    	uid = rstUser("uid")
    	
    %>
    <table class="Inout">
    <tr class="body1"><td><%=lang(cnnDB, "Fullname")%><br></td>
    <td><b><%=uname%></b><br></td>
    <%
    Dim fsFileSys, strFullFileName
    Set fsFileSys = Server.CreateObject("Scripting.FileSystemObject")
    strFullFileName = Server.MapPath("../image/" & uid & ".jpg")
    If fsFileSys.FileExists(strFullFileName) Then 
      Response.Write "<td rowspan=""8"" valign=""top"" align=""right"">" &_ 
        "<img src=""../image/" & uid & ".jpg"" width=""140"" height=""170""" & _
        " alt=""" & uname & """ border=""0""></td></tr>"
    Else
      Response.Write "<td rowspan=""8"" valign=""top"" align=""right"">" &_ 
        "<img src=""../image/nopicture.gif"" width=""140"" height=""170""" & _
        " alt=""" & uname & """ border=""0""></td></tr>"
    End if
    %>
    <tr class="body1"><td><%=lang(cnnDB, "Department")%></td><td><b><%=rstUser("dname")%></b></td></tr>
    <tr class="body1"><td><%=lang(cnnDB, "UserName")%></td><td><b><%=uid%></b></td></tr>
    <tr class="body1"><td><%=lang(cnnDB, "Email")%></td><td><b><%=rstUser("email1")%></b></td></tr>
    <tr class="body1"><td><%=lang(cnnDB, "Phone")%></td><td><b><%=rstUser("phone")%></b></td></tr>
    <tr class="body1"><td><%=lang(cnnDB, "HomePhone")%></td>
    <td><b><%
      dim str
    	str = rstUser("phone_home")
    	if len (str) = 8 then 
    		Str = Mid(Str, 1, 2) & " " & Mid(Str, 3, 2) & " " & Mid(Str, 5, 2) & " " & Mid(Str, 7, 2)
    	end if
    	response.write str
    	%>
    	</b></td></tr>
    <tr class="body1"><td><%=lang(cnnDB, "MobilePhone")%></td>
    <td><b>
    <%
    	str =rstUser("phone_mobile")
    	if len (str) = 8 then 
    		Str = Mid(Str, 1, 3) & " " & Mid(Str, 4, 2) & " " & Mid(Str, 6, 3)
    	end if
    	response.write str
    %>
    </b></td></tr>
    <tr class="body1"><td valign="top"><%=lang(cnnDB, "Status")%></td>
    <%
    select case rstUser("statuscode")
    	case 2
    		response.write "<td><img src=""../image/yellow_pin.gif"" border=0>&nbsp;" & rstUser("statustext") & "</td>"
    	case 1
    		response.write "<td><img src=""../image/red_pin.gif"" border=0>&nbsp;" & rstUser("statustext") & "</td>"
    	case 0
    		response.write "<td><img src=""../image/green_pin.gif"" border=0></td>"
    end select
    %>
    </tr>
    <tr class="body1"><td valign="top"><%=lang(cnnDB, "Function")%></td><td colspan=2><%=rstUser("jobfunction")%></td></tr>
    <tr class="body1"><td valign="top"><%=lang(cnnDB, "Resume")%></td><td colspan=2><%=rstUser("userresume")%></td></tr>
    <tr class="body1"><td align="right" colspan=3>
    <%
      if trim(usid) = trim(sid) or inoutadmin = 1 then
        response.write "<a href=""update.asp?id=" & usid & """>" & lang(cnnDB, "Edit") & "</a> | "
      end if
      if rstUser("statuscode") <> 2 or inoutadmin = 1 then
        response.write "<a href=""status.asp?id=" & usid & """>" & lang(cnnDB, "ChangeStatus") & "</a>"
      end if
    %>
    </td>&nbsp;</tr>
    </table>
    </div>
    <%
    	rstUser.close
    	Set fsFileSys = Nothing
    	Call DisplayFooter(cnnDB, sid)
    	cnnDB.Close
    %>
</body>
</html>
