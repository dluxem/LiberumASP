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

  Filename: new.asp
  Date:     $Date: 2002/08/28 15:30:54 $
  Version:  $Revision: 1.51.2.1.2.2 $
  Purpose:  This page displays the form used for entering new problems.
  -->

  <!-- 	#include file = "../public.asp" -->
  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid

    Dim strSearchName, blnPost
    strSearchName = Trim(Request.Form("searchname"))
    blnPost = False    
    If (Cint(Request.Form("postform")) = 1) And (Len(strSearchName) >= 1) Then
      blnPost = True
      Dim rstUsers
      Set rstUsers = SQLQuery(cnnDB, "SELECT sid, uid, firstname, lastname, location1 FROM tblUsers WHERE " & _
        "uid LIKE '" & strSearchName & "%' OR " & _
        "firstname LIKE '" & strSearchName & "%' OR " & _
        "lastname LIKE '" & strSearchName & "%' " & _
        "ORDER BY lastname ASC")
    End If
  %>

<head>
<title><%=lang(cnnDB, "HelpDesk")%> - <%=lang(cnnDB, "SelectUser")%></title>
<link rel="stylesheet" type="text/css" href="../default.css">
<script language="JavaScript">
  function updateParent(formval, uid) {
     opener.document.newProbForm.uselectid.value = formval;
     var element = opener.document.getElementById("selectUserText");
     if (element.hasChildNodes())
       element.firstChild.nodeValue = uid;
     else
       element.appendChild(opener.document.createTextNode(uid));
     opener.document.newProbForm.uid.disabled = true;
     opener.document.newProbForm.uid.value = '';
     opener.document.newProbForm.uemail.disabled = true;
     opener.document.newProbForm.uemail.value = '';
     opener.document.newProbForm.department.disabled = true;
     opener.document.newProbForm.department.value = 0;
     opener.document.newProbForm.ulocation.disabled = true;
     opener.document.newProbForm.ulocation.value = '';
     opener.document.newProbForm.uphone.disabled = true;
     opener.document.newProbForm.uphone.value = '';
     self.close();
     return false;
  }
</script>
</head>
<body>

<%
	' Check if user has permissions for this page
	Call CheckRep(cnnDB, sid)
%>
<form method="post" action="selectuser.asp" name="selectUserForm">
  <input type="hidden" name="postform" value="1">
  <div align="center">
    <table class="Normal">
      <tr class="Head1">
        <td colspan="5">
          <%=lang(cnnDB, "SelectUser")%>
        </td>
      </tr>
      <tr class="Body1">
        <td colspan="5">
          <input type="text" size="15" name="searchname" value="<% = strSearchName %>"> <input type="submit" value="<%=lang(cnnDB, "Search")%>">
        </td>
      </tr>
      <% If blnPost Then 
           If rstUsers.EOF Then
      %>
              <tr align="center" class="Body1">
                <td colspan="5">
                  <%=lang(cnnDB, "Noresultsfound")%>.
                </td>
              </tr>
      <%    Else  %>
              <tr align="center" Class="Head2">
                <td nowrap><%=lang(cnnDB, "LastName")%></td>
                <td><%=lang(cnnDB, "FirstName")%></td>
                <td nowrap><%=lang(cnnDB, "UserName")%></td>
                <td nowrap><%=lang(cnnDB, "Location")%></td>
                <td nowrap>&nbsp;</td>
              </tr>

      <%      Do While Not rstUsers.EOF %>
                <tr valign="center" class="Body1">
                  <td nowrap><% = rstUsers("lastname") %></td>
                  <td nowrap><% = rstUsers("firstname") %></td>
                  <td nowrap><% = rstUsers("uid") %></td>
                  <td nowrap><% = rstUsers("location1") %></td>
                  <td nowrap align="center" ><a href="javascript:updateParent('<% = rstUsers("sid") %>', '<% = rstUsers("uid") %>')">Select</a></td>
                
      <%        rstUsers.MoveNext
              Loop
            End If
         rstUsers.Close
         End If %>
    </table>
  </div>
</form>
<%
	cnnDB.Close
%>

</body>

</html>