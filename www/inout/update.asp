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

  Filename: update.asp
  Date:     $Date: 2001/12/13 09:38:21 $
  Version:  $Revision: 1.51 $
  Purpose:  Update user information in the In/Out Board.
  -->

  <!-- 	#include file = "../public.asp" -->

  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
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

    ' Look for the problem id in the query string, if it's not
    ' there display an error.
    Dim mId
    mId = request.querystring("id")

    If (trim(mId) <> trim(sid)) and (Usr(cnnDB, sid, "inoutadmin") <> 1) or (len(mId) = 0) then
    	cnnDB.Close
    	Call DisplayError(3, lang(cnnDB, "AccessDenied"))
    Else

    	Dim strSql, rstUser, rstUpdate
    	Dim phone_home, phone_mobile, jobfunction, userresume
      Dim frm_phone_home, frm_phone_mobile, frm_jobfunction, frm_userresume
    	If request.form("save") = "1" then
    		phone_home = request.form("frm_phone_home")
    		phone_mobile = request.form("frm_phone_mobile")
    		jobfunction = request.form("frm_jobfunction")
        jobfunction = Replace(jobfunction, "'", "''")
    		userresume = request.form("frm_userresume")
        userresume = Replace(userresume, "'", "''")
    
    		strSql = "UPDATE tblUsers SET " &_
    			"phone_home = '" & phone_home & "', " &_
    			"phone_mobile = '" & phone_mobile & "', " &_
    			"jobfunction = '" & jobfunction & "', " &_
    			"userresume = '" & userresume & "' " &_
    			"WHERE sid = " & mId
    		set rstUpdate = SQLQuery(cnnDB, strSql)
    	end if
    	' Fetch userinformation from the database
    	strSql = "SELECT tblUsers.*, departments.dname" &_
    		" FROM tblUsers, departments " &_
    		" WHERE (sid = " & mId & ") AND (department = department_id)"
    
    	set rstUser = SQLQuery(cnnDB, strSql)
    
    	if rstUser.EOF then
    		Call DisplayError(3, lang(cnnDB, "Noresultsfound"))
    	end if
    
    	dim uname, uid, check
    	uname = rstUser("fname")
    	uid = rstUser("uid")
    
    	if rstUser("statuscode") > 0 then
    		check = "checked"
    	else
    		check = ""
    	end if
    %>
  
    <div align="center">
    <table class="normal">
      <tr class="Head1"><td><% = Cfg(cnnDB, "SiteName") %></td></tr>
      <tr class="Head2" align="center"><td><%=lang(cnnDB, "UserDetails")%></td></tr>
    </table>
  
  <table class="inout">
  <form method="POST" action="update.asp?id=<% = mId %>">
  <input type="hidden" name="save" value="1">
  <% If Request.Form("save") = "1" Then
  	response.write "<tr class=""head2""><td colspan=3><div align=""center"">" & lang(cnnDB, "Information") & "&nbsp;" & lang(cnnDB, "isUpdated") & "</div></td><tr>"
  end if %>
  
  <tr class="body1"><td><%=lang(cnnDB, "Fullname")%><br></td><td><b><% = uname %></b><br></td>
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
  <tr class="body1"><td><%=lang(cnnDB, "Department")%><br></td><td><b><% = rstUser("dname") %></b><br></td>
  <tr class="body1"><td><%=lang(cnnDB, "Username")%></td><td><b><%=uid%></b></td></tr>
  <tr class="body1"><td><%=lang(cnnDB, "Email")%></td><td><b><%=rstUser("email1")%></b></td></tr>
  <tr class="body1"><td><%=lang(cnnDB, "Phone")%></td><td><b><%=rstUser("phone")%></b></td></tr>
  <tr class="body1"><td><%=lang(cnnDB, "HomePhone")%></td><td><input type="text" name="frm_phone_home" value="<% = rstUser("phone_home")%>" size= "10"></td></tr>
  <tr class="body1"><td><%=lang(cnnDB, "MobilePhone")%></td><td><input type="text" name="frm_phone_mobile" value="<% = rstUser("phone_mobile")%>" size= "10"></td></tr>
  <tr class="body1"><td>&nbsp;</td><td>&nbsp;</td></tr>
  <tr class="body1"><td valign="top"><%=lang(cnnDB, "Function")%></td><td colspan=2><textarea name="frm_jobfunction" ROWS=4 COLS=45><% = rstUser("jobfunction") %></textarea></td></tr>
  <tr class="body1"><td valign="top"><%=lang(cnnDB, "Resume")%></td><td colspan=2><textarea name="frm_userresume" ROWS=4 COLS=45><% = rstUser("userresume") %></textarea></td></tr>
  <tr class="body1"><td align="center" colspan="3"><input type="submit" value="<%=lang(cnnDB, "Submit")%>"></td></tr>
  </table></form>	

	<form method="post" action="savefile.asp?uid=<%=mId%>" enctype="multipart/form-data">
  	<table class="inout">
  	<tr class="Head2" align="center">
  	  <td colspan="2"><%=lang(cnnDB, "Uploadimage")%></td>
  	</tr>
  	<tr class="body1">
  	  <td colspan="2">&nbsp;</td>
  	</tr>
  	<tr class="body1">
  	  <td><%=lang(cnnDB, "Image")%></td>
  		<td><input type="file" name="blob" size="40"></td>
  	</tr>
  	<tr class="body1" align="center">
  	  <td colspan="2"><i>(<%=lang(cnnDB, "MaxImageSize")%>:&nbsp;<%=cfg(cnnDB,"MaxImageSize")%>&nbsp;<%=lang(cnnDB, "Bytes")%>)</i></td>
  	</tr>
  	<tr class="body1" align="center">
  	  <td colspan="2"><input type="submit" value="<%=lang(cnnDB, "Upload")%>"></td>
  	</tr>
    <tr class="body1">
      <td align="right" colspan="3"><a href="details.asp?id=<% = mId %>"><%=lang(cnnDB, "Details")%></a></td>
    </tr>
    </table>
  </form>
  
<%
  	rstUser.close
  	Call DisplayFooter(cnnDB, sid)
  	cnnDB.Close
	end if
%>
</div>
</body>
</html>
		