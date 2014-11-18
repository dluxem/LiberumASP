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

  Filename: viewcat.asp
  Date:     $Date: 2001/12/09 02:01:24 $
  Version:  $Revision: 1.50 $
  Purpose:  Displays a list of categories
  -->
  
  <!-- 	#include file = "../public.asp" -->

  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title>
      <%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "Manage")%>&nbsp;<%=lang(cnnDB, "Categories")%>
    </title>
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>

    <%
      ' Check for perms to view this page
      Call CheckAdmin

      Dim rstCategories
      Set rstCategories = SQLQuery(cnnDB, "SELECT category_id, cname, rep_id FROM categories WHERE category_id > 0 ORDER BY cname ASC")
    %>
    <div align="center">
      <table class="Normal">
        <tr class="Head1">
          <td colspan="3">
            <%=lang(cnnDB, "Categories")%>
          </td>
        </tr>
        <tr class="Head2">
          <td>
            <div align="center">
              <%=lang(cnnDB, "Category")%>
            </div>
          </td>
          <td>
            <div align="center">
              <%=lang(cnnDB, "Rep")%>
            </div>
          </td>
          <td>
            <div align="center">
              <%=lang(cnnDB, "Modify")%>
            </div>
          </td>
        </tr>
        <%
          Do While Not rstCategories.EOF
        %>
        <tr class="Body1">
          <td align="center">
            <% = rstCategories("cname") %>
          </td>
          <td align="center">
            <%
              Dim repRes
              Set repRes = SQLQuery(cnnDB, "SELECT uid FROM tblUsers WHERE sid=" & rstCategories("rep_id"))
              Response.Write(repRes("uid"))
              repRes.Close
             %>
          </td>
          <td align="center">
            <a href="modify.asp?mtype=2&id=<% = rstCategories("category_id") %>"><%=lang(cnnDB, "Edit")%></a>
            &nbsp;|&nbsp; <a href="confdelete.asp?mtype=2&id=<% = rstCategories("category_id") %>"><%=lang(cnnDB, "Delete")%></a>
          </td>
        </tr>
        <%
          rstCategories.MoveNext
          Loop
        %>
      </table>
      <p><form method="post" action="modify.asp?mtype=2"><input type="submit" value="<%=lang(cnnDB, "AddNew")%>&nbsp;<%=lang(cnnDB, "Category")%>"></form></p>
      <p><a href="default.asp"><%=lang(cnnDB, "AdministrativeMenu")%></a></p>
    </div>

    <%
      rstCategories.Close

      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
