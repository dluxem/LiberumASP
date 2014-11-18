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

  Filename: viewdep.asp
  Date:     $Date: 2001/12/09 02:01:24 $
  Version:  $Revision: 1.50 $
  Purpose:  The list of departments.
  -->
  
  <!-- 	#include file = "../public.asp" -->

  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title>
      <%=lang(cnnDB, "HelpDesk")%>&nbsp;<%=lang(cnnDB, "Manage")%>&nbsp;<%=lang(cnnDB, "Departments")%>
    </title>
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>

    <%
      ' Check for perms to view this page
      Call CheckAdmin

      Dim rstDepts
      Set rstDepts = SQLQuery(cnnDB, "SELECT department_id, dname FROM departments WHERE department_id > 0 ORDER BY dname ASC")

    %>
    <div align="center">
      <table class="Normal">
        <tr class="Head1">
          <td colspan="2">
            <%=lang(cnnDB, "Departments")%>
          </td>
        </tr>
        <tr class="Head2">
          <td>
            <div align="center">
              <%=lang(cnnDB, "Department")%>
            </div>
          </td>
          <td>
            <div align="center">
              <%=lang(cnnDB, "Modify")%>
            </div>
          </td>
        </tr>
        <%
          Do While Not rstDepts.EOF
        %>
        <tr class="Body1">
          <td align="center">
            <% = rstDepts("dname") %>
          </td>
          <td align="center">
            <a href="modify.asp?mtype=3&id=<% = rstDepts("department_id") %>"><%=lang(cnnDB, "Edit")%></a>
            &nbsp;|&nbsp; <a href="confdelete.asp?mtype=3&id=<% = rstDepts("department_id") %>"><%=lang(cnnDB, "Delete")%></a>
          </td>
        </tr>
        <%
          rstDepts.MoveNext
          Loop
        %>
      </table>
      <p><form method="post" action="modify.asp?mtype=3"><input type="submit" value="<%=lang(cnnDB, "AddNew")%>&nbsp;<%=lang(cnnDB, "Department")%>"></form></p>
      <p><a href="default.asp"><%=lang(cnnDB, "AdministrativeMenu")%></a></p>
    </div>

    <%
      rstDepts.Close

      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
