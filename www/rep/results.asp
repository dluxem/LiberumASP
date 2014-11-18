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

  Filename: results.asp
  Date:     $Date: 2002/06/15 23:49:20 $
  Version:  $Revision: 1.51.4.1 $
  Purpose:  This page display a list of search results posted from search.asp.

  -->

  <!-- 	#include file = "../public.asp" -->
  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>
  <head>
    <title><%=lang(cnnDB, "HelpDesk")%> - <%=lang(cnnDB, "SearchResults")%></title>
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>

  <%
    ' Check if user has permissions for this page
    Call CheckRep(cnnDB, sid)

    ' Get the search fields from the form.
    Dim uid, id, rep, category, department, status, priority, order

    uid = Trim(Request.Form("uid"))
    rep = Cint(Request.Form("rep"))
    category = Cint(Request.Form("category"))
    department = Cint(Request.Form("department"))
    status = Cint(Request.Form("status"))
    priority = Cint(Request.Form("priority"))
    order = Cint(Request.Form("order"))

    id = Trim(Request.Form("id"))
    If IsNumeric(id) Then
      id = Cint(id)
    Else
      id = 0
    End If

    ' Convert to valid characters
    uid = Replace(uid,"'","''")

    ' Create the query string.  Only get 100 results. String is in two
    ' parts, listStr (here) and whereStr (below).
    Dim listStr, whereStr
    listStr = "SELECT TOP 100 p.id, p.title, p.start_date, p.uid, p.uemail, r.uid AS ruid, s.sname " & _
    "FROM (problems AS p " & _
    "INNER JOIN tblUsers AS r ON p.rep = r.sid) " & _
    "INNER JOIN status AS s ON p.status = s.status_id " & _
    "WHERE "

    ' Convert the start and end dates to valid ODBC dates
    Dim start_date, end_date
    If Len(Request.Form("start_date")) > 0 Then
      start_date = Request.Form("start_date")
    Else
      Dim intStartDay
      intStartDay = FixDay(Request.Form("s_month"), Request.Form("s_day"), Request.Form("s_year"))
      start_date = Request.Form("s_month") & "/" & intStartDay & "/" & Request.Form("s_year")
      start_date = start_date & " 00:00:00"
    End If

    If Len(Request.Form("end_date")) > 0 Then
      end_date = Request.Form("end_date")
    Else
      Dim intEndDay
      intEndDay = FixDay(Request.Form("e_month"), Request.Form("e_day"), Request.Form("e_year"))
      end_date = Request.Form("e_month") & "/" & intEndDay & "/" & Request.Form("e_year")
      end_date = end_date & " 23:59:59"
    End If

    ' If a field was filled in on the search form, add it to the where string.

    ' Insert the dates to search between.
    whereStr = "p.start_date >" & SQLDate(start_date, lhdAddSQLDelim) & " AND p.start_date<" & SQLDate(end_date, lhdAddSQLDelim)

    If Len(uid) > 0 Then
      whereStr = whereStr & " AND p.uid='" & uid & "'"
    End If

    If id <> 0 Then
      whereStr = whereStr & " AND p.id=" & id
    End If

    If rep <> 0 Then
      whereStr = whereStr & " AND p.rep=" & rep
    End If

    If category <> 0 Then
      whereStr = whereStr & " AND p.category=" & category
    End If

    If department <> 0 Then
      whereStr = whereStr & " AND p.department=" & department
    End If

    If status <> 0 Then
      If status > 0 Then
        whereStr = whereStr & " AND p.status=" & status
      Else
        whereStr = whereStr & " AND p.status<>" & Cfg(cnnDB, "CloseStatus")
      End If
    End If

    If priority <> 0 Then
      whereStr = whereStr & " AND p.priority=" & priority
    End If

    ' Add the kewords to the list
    If Len(Request.Form("keywords")) > 0 Then
      Dim strWhere2, keywords, srchRes, varWordList, blAllOff
      keywords = Trim(Request.Form("keywords"))
      varWordList = Split(keywords, " ")

      If Request.Form("title") <> "on" And Request.Form("description") <> "on" And Request.Form("solution") <> "on" Then
        blAllOff = True
      Else
        blAllOff = False
      End If

      If Cfg(cnnDB, "KBFreeText") = 1 Then
        If Request.Form("title") = "on" Or blAllOff Then
          If Len(strWhere2) < 1 Then
            strWhere2 = " AND (FREETEXT(title, '" & keywords & "')"
          Else
            strWhere2 = strWhere2 & " OR FREETEXT(title, '" & keywords & "')"
          End If
        End If
        If Request.Form("description") = "on" Or blAllOff Then
          If Len(strWhere2) < 1 Then
            strWhere2 = " AND (FREETEXT(description, '" & keywords & "')"
          Else
            strWhere2 = strWhere2 & " OR FREETEXT(description, '" & keywords & "')"
          End If
        End If
        If Request.Form("solution") = "on" Or blAllOff Then
          If Len(strWhere2) < 1 Then
            strWhere2 = " AND (FREETEXT(solution, '" & keywords & "')"
          Else
            strWhere2 = strWhere2 & " OR FREETEXT(solution, '" & keywords & "')"
          End If
        End If
        strWhere2 = strWhere2 & ")"
      Else
        If Request.Form("title") = "on" Or blAllOff Then
          Dim strTitleKW, strWhereTitle
          For Each strTitleKW in varWordList
            If Len(strWhereTitle) < 1 Then
              strWhereTitle = "("
            Else
              strWhereTitle = strWhereTitle & " AND "
            End If
            strWhereTitle = strWhereTitle & "title LIKE '%" & strTitleKW & "%'"
          Next
          strWhereTitle = strWhereTitle & ")"
          If Len(strWhere2) < 1 Then
            strWhere2 = " AND (" & strWhereTitle
          End IF
        End If

        If Request.Form("description") = "on" Or blAllOff Then
          Dim strDescKW, strWhereDesc
          For Each strDescKW in varWordList
            If Len(strWhereDesc) < 1 Then
              strWhereDesc = "("
            Else
              strWhereDesc = strWhereDesc & " AND "
            End If
            strWhereDesc = strWhereDesc & "description LIKE '%" & strDescKW & "%'"
          Next
          strWhereDesc = strWhereDesc & ")"
          If Len(strWhere2) < 1 Then
            strWhere2 = " AND (" & strWhereDesc
          Else
            strWhere2 = strWhere2 & " OR " & strWhereDesc
          End IF
        End If

        If Request.Form("solution") = "on" Or blAllOff Then
          Dim strSolKW, strWhereSol
          For Each strSolKW in varWordList
            If Len(strWhereSol) < 1 Then
              strWhereSol = "("
            Else
              strWhereSol = strWhereSol & " AND "
            End If
            strWhereSol = strWhereSol & "solution LIKE '%" & strSolKW & "%'"
          Next
          strWhereSol = strWhereSol & ")"
          If Len(strWhere2) < 1 Then
            strWhere2 = " AND (" & strWhereSol
          Else
            strWhere2 = strWhere2 & " OR " & strWhereSol
          End IF
        End If
        strWhere2 = strWhere2 & ")"
      End If
      whereStr = whereStr & strWhere2
    End If

    Select Case order
      Case 1 whereStr = whereStr & " ORDER BY p.id ASC"
      Case 2 whereStr = whereStr & " ORDER BY p.uid ASC"
      Case 3 whereStr = whereStr & " ORDER BY r.uid ASC"
      Case 4 whereStr = whereStr & " ORDER BY p.status ASC"
    End Select


    ' Concatenate the two strings together.
    listStr = listStr & whereStr

    ' Query the database
    Dim rstProbList, start
    Set rstProbList = SQLQuery(cnnDB, listStr)

    ' If results are retuned, display them.  Only 10 results
    ' per page.
    If Not rstProbList.EOF Then
      Dim Counter, numToDisplay, startNum
      Counter = 1
      If Len(Request.Form("num")) > 0 Then
        numToDisplay = CInt(Request.Form("num"))
      Else
        numToDisplay = 10
      End if
      If Len(Request.Form("start")) > 0 Then
        start = CInt(Request.Form("start"))
      Else
        start = 1
      End if
    %>
    <div align="center">
      <table class="Wide">
      <tr class="Head1">
        <td colspan="6">
          <%=lang(cnnDB, "SearchResults")%>
        </td>
      </tr>
      <tr class="Head2" align="center">
        <td nowrap><%=lang(cnnDB, "ID")%></td>
        <td><%=lang(cnnDB, "Title")%></td>
        <td nowrap><%=lang(cnnDB, "UserName")%></td>
        <td nowrap><%=lang(cnnDB, "AssignedTo")%></td>
        <td nowrap><%=lang(cnnDB, "DateSubmitted")%></td>
        <td nowrap><%=lang(cnnDB, "Status")%></td>
      </tr>
      <%
        Do While Not (rstProbList.EOF) AND (Counter <= (numToDisplay + start - 1))
          If Counter >= start Then
        %>
          <tr class="Body1" align="center">
            <td nowrap><% = rstProbList("id") %></td>
            <td><A HREF="details.asp?id=<% = rstProbList("id") %>"><% = rstProbList("title") %></A></td>
            <td nowrap><A HREF="mailto:<% = rstProbList("uemail") %>?Subject=HELPDESK: Problem <% = rstProbList("id") %>"><% = rstProbList("uid") %></A></td>
            <td nowrap><% = rstProbList("ruid") %></td>
            <td nowrap><% = DisplayDate(rstProbList("start_date"), lhdDateOnly) %></td>
            <td nowrap><% = rstProbList("sname") %></td>
          </tr>
        <%
          End If
          Counter = Counter + 1
          rstProbList.MoveNext
        Loop
        Response.Write("</table></center>")

        ' Calculate prev/next page links
        Dim startP, StartN
        startP = start - numToDisplay
        If startP < 1 Then
          startP = 1
        End if
        startN = start + numToDisplay
      %>
        <div align="center">
        <% If start > 1 Then %>
          <form method="POST" action="results.asp">
          <input type="hidden" name="start" value="<% = startP %>">
          <input type="hidden" name="num" value="<% = numToDisplay %>">
          <input type="hidden" name="uid" value="<% = uid %>">
          <input type="hidden" name="id" value="<% = id %>">
          <input type="hidden" name="rep" value="<% = rep %>">
          <input type="hidden" name="category" value="<% = category %>">
          <input type="hidden" name="department" value="<% = department %>">
          <input type="hidden" name="start_date" value="<% = start_date %>">
          <input type="hidden" name="end_date" value="<% = end_date %>">
          <input type="hidden" name="keywords" value="<% = Request("keywords") %>">
          <input type="hidden" name="title" value="<% = Request("title") %>">
          <input type="hidden" name="description" value="<% = Request("description") %>">
          <input type="hidden" name="solution" value="<% = Request("solution") %>">
          <input type="hidden" name="status" value="<% = status %>">
          <input type="hidden" name="priority" value="<% = priority %>">
          <input type="hidden" name="order" value="<% = order %>">
          <input type="submit" value="<%=lang(cnnDB, "Previous")%>">
          </form>
        <% End If
          If Not (rstProbList.EOF) Then
        %>
          <form method="POST" action="results.asp">
          <input type="hidden" name="start" value="<% = startN %>">
          <input type="hidden" name="num" value="<% = numToDisplay %>">
          <input type="hidden" name="uid" value="<% = uid %>">
          <input type="hidden" name="id" value="<% = id %>">
          <input type="hidden" name="rep" value="<% = rep %>">
          <input type="hidden" name="category" value="<% = category %>">
          <input type="hidden" name="department" value="<% = department %>">
          <input type="hidden" name="start_date" value="<% = start_date %>">
          <input type="hidden" name="end_date" value="<% = end_date %>">
          <input type="hidden" name="keywords" value="<% = Request("keywords") %>">
          <input type="hidden" name="title" value="<% = Request("title") %>">
          <input type="hidden" name="description" value="<% = Request("description") %>">
          <input type="hidden" name="solution" value="<% = Request("solution") %>">
          <input type="hidden" name="status" value="<% = status %>">
          <input type="hidden" name="priority" value="<% = priority %>">
          <input type="hidden" name="order" value="<% = order %>">
          <input type="submit" value="<%=lang(cnnDB, "Next")%>">
          </form>
        <%	End If %>
        <p>
        <a href="search.asp"><%=lang(cnnDB, "SearchAgain")%></a>
        </div>
      <%

      ' If no results returned:
      Else
      %>
        <div align="center">
          <table class="Wide">
            <tr class="Head1">
              <td colspan="6">
                <%=lang(cnnDB, "SearchResults")%>
              </td>
            </tr>
            <tr class="Head2" align="center">
              <td nowrap><%=lang(cnnDB, "ID")%></td>
              <td><%=lang(cnnDB, "Title")%></td>
              <td nowrap><%=lang(cnnDB, "UserName")%></td>
              <td nowrap><%=lang(cnnDB, "AssignedTo")%></td>
              <td nowrap><%=lang(cnnDB, "DateSubmitted")%></td>
              <td nowrap><%=lang(cnnDB, "Status")%></td>
            </tr>
            <tr class="Body1">
              <td colspan="6">
                <div align="center">
                  <%=lang(cnnDB, "Noresultsfound")%>.
                </div>
              </td>
            </tr>
          </table>
          <p>
          <a href="search.asp"><%=lang(cnnDB, "SearchAgain")%></a>
        </div>
  <%	End If

    ' Close results
    rstProbList.Close

    Call DisplayFooter(cnnDB, sid)
    cnnDB.Close
  %>
</body>
</html>


