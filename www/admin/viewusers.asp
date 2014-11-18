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

  Filename: viewusers.asp
  Date:     $Date: 2001/12/09 02:01:24 $
  Version:  $Revision: 1.50 $
  Purpose:  Displays the list of users.
  -->
  
  <!-- 	#include file = "../public.asp" -->

  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title><%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "ManageUsers")%>
    </title>
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>

    <%
      ' Check for perms to view this page
      Call CheckAdmin

      Dim rstUserList
      Set rstUserList = SQLQuery(cnnDB, "SELECT sid, uid, fname FROM tblUsers WHERE sid>0 ORDER BY uid ASC")

      Dim intCounter
      intCounter = 1
    %>
    <div align="center">
      <table class="Normal">
        <tr class="Head1">
          <td colspan="4">
            <%=lang(cnnDB, "ManageUsers")%>
          </td>
        </tr>
        <tr class="Body1">
          <td>
            <div align="center">
            <% If Not rstUserList.EOF Then %>
                <form method="POST" action="moduser.asp">
                  <select name="usersid" size="5">
                    <% Do While Not rstUserList.EOF 
                        If intcounter = 1 Then
                          Response.Write("<option value=""" & rstUserList("sid") & """ selected>")
                          intCounter = intCounter + 1
                        Else
                          Response.Write("<option value=""" & rstUserList("sid") & """>")
                        End If
                    %>
                        <% = rstUserList("uid") %> (<% = rstUserList("fname") %>)
                      </option>
                    <% 	rstUserList.MoveNext
                      Loop
                    %>
                  </select>
                  <p><input type="submit" value="<%=lang(cnnDB, "EditUserAccount")%>"></p>
                </form>
              <% End If %>
              <form name="addusr" method="POST" action="adduser.asp">
                <input type="submit" value="<%=lang(cnnDB, "AddNewUsers")%>">
              </form>
            </div>
          </td>
        </tr>
      </table>
      <p>
      <a href="default.asp"><%=lang(cnnDB, "AdministrativeMenu")%></a>
    </div>

    <%
      rstUserList.Close

      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
