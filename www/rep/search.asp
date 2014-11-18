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

  Filename: search.asp
  Date:     $Date: 2001/12/09 02:01:24 $
  Version:  $Revision: 1.50 $
  Purpose:  This page displays a form of search paramaters which are posted
    to results.asp.

  -->

  <!-- 	#include file = "../public.asp" -->
  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

<head>
  <title><%=lang(cnnDB, "HelpDesk")%> - <%=lang(cnnDB, "ProblemSearch")%></title>
  <link rel="stylesheet" type="text/css" href="../default.css">
</head>
<body>

<%

	' Check if user has permissions for this page
	Call CheckRep(cnnDB, sid)

%>
<div align="center">
  <table class="Normal">
  <tr class="Head1">
    <td>
      <%=lang(cnnDB, "ProblemSearch")%>
    </td>
  </tr>
  <tr class="Head2">
    <td>
      <%=lang(cnnDB, "ProblemSpecifications")%>:
    </td>
  </tr>
  <tr class="Body1">
    <td>
      <form method="post" action="results.asp">
      <table class="Normal">
        <tr class="body1">
          <td>
            <b><%=lang(cnnDB, "ReportedBy")%>:</b>
          </td>
          <td>
            <input type="text" name="uid" size="30">
          </td>
        </tr>
        <tr class="body1">
          <td>
            <b><%=lang(cnnDB, "ProblemID")%>:</b>
          </td>
          <td>
            <input type="text" name="id" size="30">
          </td>
        </tr>
        <tr class="body1">
          <td>
            <b><%=lang(cnnDB, "AssignedTo")%>:</b>
          </td>
          <td>
            <select name="rep">
            <option value="0"><%=lang(cnnDB, "Any")%></option>
            <%
              Dim repRes
              Set repRes = SQLQuery(cnnDB, "SELECT sid, uid From tblUsers WHERE IsRep = 1 AND RepAccess <> 2 AND sid > 0 ORDER BY uid ASC")
              If Not repRes.EOF Then
              Do While Not repRes.EOF
            %>
              <option value="<% = repRes("sid")%>">
              <% = repRes("uid") %></option>
      
            <%	repRes.MoveNext
              Loop
              End If
              repRes.Close
            %>
            </select>
          </td>
        </tr>
        <tr class="body1">
          <td>
            <b><%=lang(cnnDB, "Category")%>:</b>
          </td>
          <td>
            <select name="category">
            <option value="0"><%=lang(cnnDB, "Any")%></option>
            <%
              Dim catRes
              Set catRes = SQLQuery(cnnDB, "SELECT category_id, cname From categories WHERE category_id > 0 ORDER BY cname ASC")
              If Not catRes.EOF Then
              Do While Not catRes.EOF
            %>
              <option value="<% = catRes("category_id")%>">
              <% = catRes("cname") %></option>
      
            <%	catRes.MoveNext
              Loop
              End If
              catRes.Close
            %>
            </select>
          </td>
        </tr>
        <tr class="body1">
          <td>
            <b><%=lang(cnnDB, "Department")%>:</b>
          </td>
          <td>
            <select name="department">
            <option value="0"><%=lang(cnnDB, "Any")%></option>
            <%
              Dim depRes
              Set depRes = SQLQuery(cnnDB, "SELECT * From departments WHERE department_id > 0 ORDER BY dname ASC")
              If Not depRes.EOF Then
              Do While Not depRes.EOF
            %>
              <option value="<% = depRes("department_id")%>">
              <% = depRes("dname") %></option>
      
            <%	depRes.MoveNext
              Loop
              End If
              depRes.Close
            %>
            </select>
          </td>
        </tr>
        <tr class="body1">
          <td>
            <b><%=lang(cnnDB, "Status")%>:</b>
          </td>
          <td>
            <select name="status" width="30">
            <option value="0"><%=lang(cnnDB, "Any")%></option>
            <option value="-1"><%=lang(cnnDB, "AllNonClosed")%></option>
            <%
              Dim statRes
              Set statRes = SQLQuery(cnnDB, "SELECT * From status WHERE status_id > 0 ORDER BY status_id ASC")
              If Not statRes.EOF Then
              Do While Not statRes.EOF
            %>
              <option value="<% = statRes("status_id")%>">
              <% = statRes("sname") %></option>
      
            <%	statRes.MoveNext
              Loop
              End If
              statRes.Close
            %>
            </select>
          </td>
        </tr>
        <tr class="body1">
          <td>
            <b><%=lang(cnnDB, "Priority")%>:</b>
          </td>
          <td>
            <select name="priority">
            <option value="0"><%=lang(cnnDB, "Any")%></option>
            <%
              Dim priRes
              Set priRes = SQLQuery(cnnDB, "SELECT * From priority WHERE priority_id > 0 ORDER BY priority_id ASC")
              If Not priRes.EOF Then
              Do While Not priRes.EOF
            %>
              <option value="<% = priRes("priority_id")%>">
              <% = priRes("pname") %></option>
      
            <%	priRes.MoveNext
              Loop
              End If
              priRes.Close
            %>
            </select>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr class="Head2">
    <td>
      <%=lang(cnnDB, "Contains")%>:
    </td>
  </tr>
  <tr class="Body1">
    <td>
      <table  class="Normal">
        <tr>
          <td width="100">
            <b><%=lang(cnnDB, "Keywords")%>:</b>
          </td>
          <td>
            <input type="text" size="50" name="keywords">
          </td>
        </tr>
        <tr>
          <td valign="top">
            <b><%=lang(cnnDB, "SearchFields")%>:</b>
          </td>
          <td>
            <input type="checkbox" name="title" checked> <%=lang(cnnDB, "Title")%><br />
            <input type="checkbox" name="description" checked> <%=lang(cnnDB, "Description")%><br />
            <input type="checkbox" name="solution" checked> <%=lang(cnnDB, "Solution")%><br />
          </td>
        </tr>
      </table>
    </td>
  <tr class="Head2">
    <td>
      <%=lang(cnnDB, "OrderedBy")%>:
    </td>
  </tr>
  <tr class="Body1">
    <td>
      <input type="radio" name="order" value="1" checked> <%=lang(cnnDB, "ProblemID")%>
      <input type="radio" name="order" value="2"> <%=lang(cnnDB, "UserName")%>
      <input type="radio" name="order" value="3"> <%=lang(cnnDB, "RepUserName")%>
      <input type="radio" name="order" value="4"> <%=lang(cnnDB, "Status")%>
  <tr class="Head2">
    <td>
      <%=lang(cnnDB, "Dates")%>:
    </td>
  </tr>
  <tr class="Body1">
    <td>
      <table class="normal">
        <tr class="body1">
          <td>
            <%=lang(cnnDB, "From")%>&nbsp;
            <select name="s_month" size="1">
            <%
              Dim temp_date, year_adj, count
              year_adj = 0
      
              For count = 1 to 12
                temp_date = Month(now) - 1
                If temp_date = 0 Then
                  temp_date = 12
                  year_adj = -1
                End If
                If count = temp_date Then
                  Response.Write("<option value=""" & count & """ selected>" & count & "</option>")
                Else
                  Response.Write("<option value=""" & count & """>" & count & "</option>")
                End If
              Next
            %>
            </select> / <select name="s_day" size="1">
            <%
              For count = 1 to 31
                If count = Day(now) Then
                  Response.Write("<option value=""" & count & """ selected>" & count & "</option>")
                Else
                  Response.Write("<option value=""" & count & """>" & count & "</option>")
                End If
              Next
            %>
            </select> / <select name="s_year" size="1">
            <%
              year_adj = Year(now) + year_adj
              For count = 2000 to 2010
                If count = year_adj Then
                  Response.Write("<option value=""" & count & """ selected>" & count & "</option>")
                Else
                  Response.Write("<option value=""" & count & """>" & count & "</option>")
                End If
              Next
            %>
            </select>
          </td>
          <td>
            <%=lang(cnnDB, "through")%>&nbsp;
            <select name="e_month" size="1">
            <%
              For count = 1 to 12
                If count = Month(now) Then
                  Response.Write("<option value=""" & count & """ selected>" & count & "</option>")
                Else
                  Response.Write("<option value=""" & count & """>" & count & "</option>")
                End If
              Next
            %>
            </select> / <select name="e_day" size="1">
            <%
              For count = 1 to 31
                If count = Day(now) Then
                  Response.Write("<option value=""" & count & """ selected>" & count & "</option>")
                Else
                  Response.Write("<option value=""" & count & """>" & count & "</option>")
                End If
              Next
            %>
            </select> / <select name="e_year" size="1">
            <%
              For count = 2000 to 2010
                If count = Year(now) Then
                  Response.Write("<option value=""" & count & """ selected>" & count & "</option>")
                Else
                  Response.Write("<option value=""" & count & """>" & count & "</option>")
                End If
              Next
            %>
            </select>
          </td>
        </tr>
      </table>
      <center><br /><input type="submit" value="<%=lang(cnnDB, "Search")%>"></center>
      </form>
    </td>
  </tr>
  </table>
</div>

<%
	Call DisplayFooter(cnnDB, sid)
	cnnDB.Close
%>
</body>
</html>
