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

  Filename: viewstatus.asp
  Date:     $Date: 2001/12/09 02:01:24 $
  Version:  $Revision: 1.50 $
  Purpose:  View the list of statuses.
  -->
  
  <!-- 	#include file = "../public.asp" -->

  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title>
      <%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "Manage")%>&nbsp;<%=lang(cnnDB, "Reports")%>
    </title>
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>

    <%
      ' Check for perms to view this page
      Call CheckAdmin

      Dim rstStat
      Set rstStat = SQLQuery(cnnDB, "SELECT status_id, sname FROM status WHERE status_id > 0 ORDER BY status_id ASC")

    %>
    <div align="center">
      <table class="Normal">
        <tr class="Head1">
          <td colspan="3">
            <%=lang(cnnDB, "Statuses")%>
          </td>
        </tr>
        <tr class="Head2">
          <td>
            <div align="center">
              <%=lang(cnnDB, "ID")%>
          </td>
          <td>
            <div align="center">
              <%=lang(cnnDB, "Status")%>
          </td>
          <td>
            <div align="center">
              <%=lang(cnnDB, "Modify")%>
            </div>
          </td>
        </tr>
        <%
          Do While Not rstStat.EOF
        %>
        <tr class="Body1">
          <td align="center">
            <% = rstStat("status_id") %>
            <% If Cint(rstStat("status_id")) = Cfg(cnnDB, "CloseStatus") Then
              Response.Write("<em>*</em>")
              End If
            %>
          </td>
          <td align="center">
            <% = rstStat("sname") %>
          </td>
          <td align="center">
            <a href="modify.asp?mtype=5&id=<% = rstStat("status_id") %>"><%=lang(cnnDB, "Edit")%></a>
            &nbsp;|&nbsp; <a href="confdelete.asp?mtype=5&id=<% = rstStat("status_id") %>"><%=lang(cnnDB, "Delete")%></a>
          </td>
        </tr>
        <%
          rstStat.MoveNext
          Loop
        %>
      </table>
      <i><em>*</em><%=lang(cnnDB, "ClosedStatusDonotdelete")%>.</i>
      <p><form method="post" action="modify.asp?mtype=5"><input type="submit" value="<%=lang(cnnDB, "AddNew")%>&nbsp;<%=lang(cnnDB, "Status")%>"></form></p>
      <p><a href="default.asp"><%=lang(cnnDB, "AdministrativeMenu")%></a></p>
    </div>

    <%
      rstStat.Close

      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
