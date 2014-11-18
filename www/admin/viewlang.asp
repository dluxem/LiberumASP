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

  Filename: viewlang.asp
  Date:     $Date: 2001/12/09 02:01:24 $
  Version:  $Revision: 1.50 $
  Purpose:  Manage the list of availbable Languages
  -->
  
  <!-- 	#include file = "../public.asp" -->

  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title>
      <%=lang(cnnDB, "HelpDesk")%>&nbsp;<%=lang(cnnDB, "Manage")%>&nbsp;<%=lang(cnnDB, "Languages")%>
    </title>
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>

    <%
      ' Check for perms to view this page
      Call CheckAdmin

      Dim rstLanguage
      Set rstLanguage = SQLQuery(cnnDB, "SELECT * FROM tblLanguage ORDER BY id ASC")

    %>
    <div align="center">
      <table class="Normal">
        <tr class="Head1">
          <td colspan="3">
            <%=lang(cnnDB, "Languages")%>
          </td>
        </tr>
        <tr class="Head2">
          <td>
            <div align="center">
              <%=lang(cnnDB, "Language")%>
            </div>
          </td>
          <td>
            <div align="center">
              <%=lang(cnnDB, "Localized")%>
            </div>
          </td>
          <td>
            <div align="center">
              <%=lang(cnnDB, "Modify")%>
            </div>
          </td>
        </tr>
        <%
          Do While Not rstLanguage.EOF
        %>
        <tr class="Body1">
          <td align="center">
            <% = rstLanguage("LangName") %>
          </td>
          <td align="center">
             <% = rstLanguage("Localized") %>
          </td>
          <td align="center">
            <% If rstLanguage("id") <> 1 Then %>
              <a href="modify.asp?mtype=6&id=<% = rstLanguage("id") %>"><%=lang(cnnDB, "Edit")%></a>
              &nbsp;|&nbsp;
              <a href="confdelete.asp?mtype=6&id=<% = rstLanguage("id") %>"><%=lang(cnnDB, "Delete")%></a>
              &nbsp;|&nbsp;
            <% End If %>
            <a href="viewlangstring.asp?lang_id=<% = rstLanguage("id") %>"><%=lang(cnnDB, "Strings")%></a>
          </td>
        </tr>
        <%
          rstLanguage.MoveNext
          Loop
        %>
      </table>
      <p><form method="post" action="modify.asp?mtype=6"><input type="submit" value="<%=lang(cnnDB, "AddNew")%>&nbsp;<%=lang(cnnDB, "Language")%>"></form></p>
      <p><a href="default.asp"><%=lang(cnnDB, "AdministrativeMenu")%></a></p>
    </div>

    <%
      rstLanguage.Close

      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
