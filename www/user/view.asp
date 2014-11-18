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

  Filename: view.asp
  Date:     $Date: 2002/06/15 23:49:20 $
  Version:  $Revision: 1.50.4.1 $
  Purpose:  Lists recent problems for the user, or a specific problem if
    a problem id is given.
  -->

  <!-- 	#include file = "../public.asp" -->
  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title><%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "ProblemList")%></title>
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>

    <%
      ' See if authenticated
      Call CheckUser(cnnDB, sid)

      ' Get the user's username
      Dim uid
      uid = Usr(cnnDB, sid, "uid")

      ' Get the problem ID from either a form entry or
      ' the URL entered.  If none exists then get the
      ' top fifty problem entered.
      Dim listStr, rstProbList
      listStr = "SELECT TOP 50 p.id, p.title, p.start_date, r.fname, r.email1 As remail, s.sname " & _
      "FROM (problems AS p " & _
      "INNER JOIN tblUsers AS r ON p.rep = r.sid) " & _
      "INNER JOIN status AS s ON p.status = s.status_id "

      Dim id
      If Len(Request.QueryString("id")) > 0 Then
        id = Request.QueryString("id")
      Else
        id = Request.Form("id")
      End If

      id = Trim(id)

      If IsNumeric(id) Then
        listStr = listStr & "WHERE p.uid='" & uid & "' AND p.id=" & Cint(id)
      Else
        listStr = listStr & "WHERE p.uid='" & uid & "'"
      End If

      ' Determine Sort Order
      Dim intSort, intOrder, intIDOrder, intTitleOrder, intRepOrder, intDateOrder, intStatusOrder
      intSort = Cint(Request.QueryString("sort"))
      If Len(Request.QueryString("order")) > 0 Then
        intOrder = Cint(Request.QueryString("order"))
      Else
        intOrder = 0
      End If
      Select Case intSort
        Case 1  ' id
          listStr = listStr & " ORDER BY p.id"
          If intOrder = 0 Then
            listStr = listStr & " DESC"
            intIDOrder = 1
          Else
            listStr = listStr & " ASC"
            intIDOrder = 0
          End If
        Case 2  ' title
          listStr = listStr & " ORDER BY p.title"
          If intOrder = 0 Then
            listStr = listStr & " ASC"
            intTitleOrder = 1
          Else
            listStr = listStr & " DESC"
            intTitleOrder = 0
          End If
        Case 3  ' uid
          listStr = listStr & " ORDER BY r.fname"
          If intOrder = 0 Then
            listStr = listStr & " ASC"
            intRepOrder = 1
          Else
            listStr = listStr & " DESC"
            intRepOrder = 0
          End If
        Case 4  ' start_date
          listStr = listStr & " ORDER BY p.start_date"
          If intOrder = 0 Then
            listStr = listStr & " DESC"
            intDateOrder = 1
          Else
            listStr = listStr & " ASC"
            intDateOrder = 0
          End If
        Case 5  ' status
          listStr = listStr & " ORDER BY p.status"
          If intOrder = 0 Then
            listStr = listStr & " DESC"
            intStatusOrder = 1
          Else
            listStr = listStr & " ASC"
            intStatusOrder = 0
          End If
        Case Else ' id again
          listStr = listStr & " ORDER BY p.id"
          If intOrder = 0 Then
            listStr = listStr & " DESC"
            intIDOrder = 1
          Else
            listStr = listStr & " ASC"
            intIDOrder = 0
          End If
      End Select

      Set rstProbList = SQLQuery(cnnDB, listStr)

      ' Set up the variables used to make the list run multiple pages.
      Dim Counter, numToDisplay, startNum, start
      Counter = 1
      If Len(Request.QueryString("num")) > 0 Then
        numToDisplay = CInt(Request.QueryString("num"))
      Else
        numToDisplay = 10
      End if
      If Len(Request.QueryString("start")) > 0 Then
        start = CInt(Request.QueryString("start"))
      Else
        start = 1
      End if
    %>

    <div align="center">
      <table class="Wide">
        <tr class="Head1">
          <td colspan="5">
            <%=lang(cnnDB, "ProblemListingfor")%>&nbsp;<% = uid %>
          </td>
        </tr>
        <tr class="Head2" align="center">
          <td nowrap><a href="view.asp?start=<% = start %>&num=<% = numToDisplay %>&sort=1&order=<% = intIDOrder %>" class="HeadLink"><%=lang(cnnDB, "ID")%></a></td>
          <td><a href="view.asp?start=<% = start %>&num=<% = numToDisplay %>&sort=2&order=<% = intTitleOrder %>" class="HeadLink"><%=lang(cnnDB, "Title")%></a></td>
          <td nowrap><a href="view.asp?start=<% = start %>&num=<% = numToDisplay %>&sort=3&order=<% = intRepOrder %>" class="HeadLink"><%=lang(cnnDB, "AssignedTo")%></a></td>
          <td nowrap><a href="view.asp?start=<% = start %>&num=<% = numToDisplay %>&sort=4&order=<% = intDateOrder %>" class="HeadLink"><%=lang(cnnDB, "DateSubmitted")%></a></td>
          <td nowrap><a href="view.asp?start=<% = start %>&num=<% = numToDisplay %>&sort=5&order=<% = intStatusOrder %>" class="HeadLink"><%=lang(cnnDB, "Status")%></a></td>
        </tr>
        <%
        ' Look for no results
        If Not rstProbList.EOF Then
          Do While Not (rstProbList.EOF) AND (Counter <= (numToDisplay + start - 1))
            If Counter >= start Then
        %>
              <tr class="Body1">
                <td nowrap><div align="center"><% = rstProbList("id") %></div></td>
                <td><div align="center"><A HREF="details.asp?id=<% = rstProbList("id") %>"><% = rstProbList("title") %></A></div></td>
                <td nowrap><div align="center"><A HREF="mailto:<% = rstProbList("remail") %>?Subject=HELPDESK: Problem <% = rstProbList("id") %>"><% = rstProbList("fname") %></A></div></td>
                <td nowrap><div align="center"><% = DisplayDate(rstProbList("start_date"), lhdDateOnly) %></div></td>
                <td nowrap><div align="center"><% = rstProbList("sname") %></div></td>
              </tr>
          <%
            End If
            Counter = Counter + 1
            rstProbList.MoveNext
          Loop
          
          If start > 1 or Not rstProbList.EOF Then
          %>
            <tr class="Head2">
              <td colspan="5">
                <%
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
                    <A HREF="view.asp?start=<% = startP %>&num=<% = numToDisplay %>&sort=<% = intSort %>&order=<% = intOrder %>"><%=lang(cnnDB, "Previous")%></A>&nbsp;
                  <% End If
                    If Not (rstProbList.EOF) Then
                  %>
                    <A HREF="view.asp?start=<% = startN %>&num=<% = numToDisplay %>&sort=<% = intSort %>&order=<% = intOrder %>"><%=lang(cnnDB, "Next")%></A>
                  <% End If %>
                </div>
              </td>
            </tr>
        <% End If %>

        <%
        ' If no results returned:
        Else
        %>

          <tr class="Body1">
            <td colspan="5">
              <div align="center">
                <b><%=lang(cnnDB, "Noresultsfound")%>.</b>
              </div>
            </td>
          </tr>
        <% End If %>
      </table>
    </div>
    <%
      ' Close results
      rstProbList.Close

      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>