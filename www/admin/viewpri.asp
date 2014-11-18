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

  Filename: viewpri.asp
  Date:     $Date: 2001/12/09 02:01:24 $
  Version:  $Revision: 1.50 $
  Purpose:  View the list of priorities.
  -->
  
  <!-- 	#include file = "../public.asp" -->

  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title>
      <%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "Manage")%>&nbsp;<%=lang(cnnDB, "Priorities")%>
    </title>
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>

    <%
      ' Check for perms to view this page
      Call CheckAdmin

      Dim rstPriList
      Set rstPriList = SQLQuery(cnnDB, "SELECT priority_id, pname FROM priority WHERE priority_id > 0 ORDER BY priority_id ASC")

    %>
    <div align="center">
      <table class="Normal">
        <tr class="Head1">
          <td colspan="3">
            <%=lang(cnnDB, "Priorities")%>
          </td>
        </tr>
        <tr class="Head2">
          <td>
            <div align="center">
              <%=lang(cnnDB, "ID")%>
            </div>
          </td>
          <td>
            <div align="center">
              <%=lang(cnnDB, "Priority")%>
            </div>
          </td>
          <td>
            <div align="center">
              <%=lang(cnnDB, "Modify")%>
            </div>
          </td>
        </tr>
        <%
          Do While Not rstPriList.EOF
        %>
        <tr class="Body1">
          <td align="center">
            <% = rstPriList("priority_id") %>
          </td>
          <td align="center">
            <% = rstPriList("pname") %>
          </td>
          <td align="center">
            <a href="modify.asp?mtype=4&id=<% = rstPriList("priority_id") %>"><%=lang(cnnDB, "Edit")%></a>
            &nbsp|&nbsp <a href="confdelete.asp?mtype=4&id=<% = rstPriList("priority_id") %>"><%=lang(cnnDB, "Delete")%></a>
          </td>
        </tr>
        <%
          rstPriList.MoveNext
          Loop
        %>
      </table>
      <p>
        <form method="post" action="modify.asp?mtype=4">
          <input type="submit" value="<%=lang(cnnDB, "AddNew")%>&nbsp;<%=lang(cnnDB, "Priority")%>">
        </form><br />
        <a href="default.asp"><%=lang(cnnDB, "AdministrativeMenu")%></a>
      </p>
    </div>

    <%
      rstPriList.Close

      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
