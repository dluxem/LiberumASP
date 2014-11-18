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

  Filename: viewreports.asp
  Date:     $Date: 2002/06/15 23:49:20 $
  Version:  $Revision: 1.50.4.1 $
  Purpose:  Displays the report results with parameters from reports.asp.
  -->
  
  <!-- 	#include file = "../public.asp" -->

  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title>
      <%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "Reports")%>
    </title>
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>

    <%
      ' Check for perms to view this page
      Call CheckAdmin

      Dim queryStr, start_date, end_date, intStartDay, intEndDay
      intStartDay = FixDay(Request.Form("s_month"), Request.Form("s_day"), Request.Form("s_year"))
      intEndDay = FixDay(Request.Form("e_month"), Request.Form("e_day"),  Request.Form("e_year"))
      start_date = Request.Form("s_month") & "/" & intStartDay & "/" & Request.Form("s_year")
      end_date = Request.Form("e_month") & "/" & intEndDay & "/" & Request.Form("e_year")
      start_date = start_date & " 00:00:00"
      end_date = end_date & " 23:59:59"

      ' Date criteria used in query.
      Dim dateCriteria
      dateCriteria = "start_date > " & SQLDate(start_date, lhdAddSQLDelim) & " AND start_date < " & SQLDate(end_date, lhdAddSQLDelim)

      ' Generate the sql queries
      Select Case Cint(Request.Form("type"))
        Case 0	' Departments
            queryStr = "SELECT d.dname AS name, Count(*) AS total, sum(p.time_spent) AS total_time " & _
              "FROM problems AS p INNER JOIN departments AS d ON p.department = d.department_id " & _
              "WHERE " & dateCriteria & " " & _
              "GROUP BY dname ORDER BY dname ASC"
        Case 1	' Categories
            queryStr = "SELECT c.cname AS name, Count(*) AS total, sum(p.time_spent) AS total_time " & _
              "FROM problems AS p INNER JOIN categories AS c ON p.category = c.category_id " & _
              "WHERE " & dateCriteria & " " & _
              "GROUP BY cname ORDER BY cname ASC"
        Case 2	' Reps
            queryStr = "SELECT r.uid AS name, Count(*) AS total, sum(p.time_spent) AS total_time " & _
              "FROM problems AS p INNER JOIN tblUsers AS r ON p.rep = r.sid " & _
              "WHERE " & dateCriteria & " AND sid>0 " & _
              "GROUP BY r.uid ORDER BY r.uid ASC"

      End Select ' end select
      
      Dim rstResults
      Set rstResults = SQLQuery(cnnDB, queryStr)

      Dim total, total_time
      total = 0
      total_time = 0
      If Not rstResults.EOF Then
        Do While NOT rstResults.EOF
          total = total + Cint(rstResults("total"))
          total_time = total_time + Cint(rstResults("total_time"))
          rstResults.MoveNext
        Loop
        rstResults.MoveFirst
      End If

    %>
    <div align="center">
      <table class="Normal">
        <tr class="Head1">
          <td colspan="6">
            <%=lang(cnnDB, "Reports")%>
          </td>
        </tr>
        <tr class="Head2" align="center">
          <td>
            <%
              Select Case Cint(Request.Form("type"))
                Case 0
                    Response.Write(lang(cnnDB, "Department"))
                Case 1
                    Response.Write(lang(cnnDB, "Category"))
                Case 2
                    Response.Write(lang(cnnDB, "Rep"))
              End Select
            %>
          </td>
          <td><%=lang(cnnDB, "Total")%></td>
          <td><%=lang(cnnDB, "Time")%></td>
          <td><%=lang(cnnDB, "AvgTime")%></td>
          <td><%=lang(cnnDB, "PofProblems")%></td>
          <td><%=lang(cnnDB, "PofTime")%></td>
        </tr>
        <% If Not rstResults.EOF Then
          Do While Not rstResults.EOF %>
            <tr class="Body1" align="center">
              <td>
                <% = rstResults("name") %>
              </td>
              <td>
                <% = rstResults("total") %>
              </td>
              <td>
                <% = rstResults("total_time") %>
              </td>
              <td>
                <% = FormatNumber((Cint(rstResults("total_time")) / Cint(rstResults("total"))), 1) %>
              </td>
              <td>
                <% = FormatNumber((Cint(rstResults("total")) / total * 100), 1) %>%
              </td>
              <td>
                <% If total_time <> 0 Then
                  Response.Write(FormatNumber((Cint(rstResults("total_time")) / total_time * 100), 1))
                   Else
                    Response.Write("0")
                   End If
                   Response.Write("%")
                 %>
              </td>
            </tr>
        <%
          rstResults.MoveNext
          Loop
          rstResults.Close
        %>
        <tr class="Head2" align="center">
          <td>
            <%=lang(cnnDB, "Totals")%>
          </td>
          <td>
            <% = total %>
          </td>
          <td>
            <% = total_time %>
          </td>
          <td>
            <% = FormatNumber((total_time / total), 1) %>
          </td>
          <td>
            100%
          </td>
          <td>
            100%
          </td>
        </tr>
        <% Else ' no data %>
        <tr class="Body1">
          <td colspan="6" align="center">
            <b><%=lang(cnnDB, "Noresultsfound")%></b>
          </td>
        </tr>
        <% End If %>
      </table>
      <p>
      <a href="default.asp"><%=lang(cnnDB, "AdministrativeMenu")%></a><br>
      <a href="reports.asp"><%=lang(cnnDB, "ReportsMenu")%></a>
    </div>
    <%
      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
