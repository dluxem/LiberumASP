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
  Version:  $Revision: 1.50.2.2.2.1 $
  Purpose:  This page lists the open problem for the rep or another selected rep.
  -->

  <!-- 	#include file = "../public.asp" -->
  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

<head>
  <title><%=lang(cnnDB, "HelpDesk")%> - <%=lang(cnnDB, "OpenProblems")%></title>
  <link rel="stylesheet" type="text/css" href="../default.css">
</head>
<body>

<%
	' Check if user has permissions for this page
	Call CheckRep(cnnDB, sid)

	' Determine if a rep_id has been entered by form or in the
	' URL.  If not, just query on the rep's own id.
	Dim rep_id
	if Len(Request.QueryString("rep_id")) > 0 Then
		rep_id = Cint(Request.QueryString("rep_id"))
	Elseif Len(Request.Form("rep_id")) > 0 Then
		rep_id = Cint(Request.Form("rep_id"))
	Else
		rep_id = sid
	End If

	' Get the rep's name that is being queryed on
	Dim ruid, repRes
  Set repRes = SQLQuery(cnnDB, "SELECT uid FROM tblUsers WHERE sid=" & rep_id )
	ruid = repRes("uid")

	' Query the database for the problems.
	Dim listStr, cntStr, listRes
	listStr = "SELECT TOP 100 p.id, p.title, p.start_date, p.uid, p.uemail, r.uid AS ruid, pri.pname, s.sname " & _
	"FROM ((problems AS p " & _
	"INNER JOIN tblUsers AS r ON p.rep = r.sid) " & _
	"INNER JOIN priority AS pri ON p.priority = pri.priority_id) " & _
	"INNER JOIN status AS s ON p.status = s.status_id " & _
	"WHERE"

	cntStr = "SELECT count(*) AS total FROM problems WHERE"
	Dim disp_total

	' If a problem ID is entered, search only for that
	If Len(Request.QueryString("id"))>0 Then
		listStr = listStr & " p.id=" & Request.QueryString("id")
		disp_total = FALSE
	Else
		listStr = listStr & " p.status<>" & Cfg(cnnDB, "CloseStatus") & " AND p.rep=" & rep_id
		cntStr = cntStr & " status<>" & Cfg(cnnDB, "CloseStatus") & " AND rep=" & rep_id
		disp_total = TRUE
	End If

  ' Determine Sort Order
  Dim intSort, intOrder, intIDOrder, intTitleOrder, intUIDOrder, intDateOrder, intPriOrder, intStatusOrder
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
      listStr = listStr & " ORDER BY p.uid"
      If intOrder = 0 Then
        listStr = listStr & " ASC"
        intUIDOrder = 1
      Else
        listStr = listStr & " DESC"
        intUIDOrder = 0
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
    Case 5  ' priority
      listStr = listStr & " ORDER BY p.priority"
      If intOrder = 0 Then
        listStr = listStr & " DESC"
        intPriOrder = 1
      Else
        listStr = listStr & " ASC"
        intPriOrder = 0
      End If
    Case 6  ' status
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

  Set listRes = SQLQuery(cnnDB, listStr)

	' Get a total number of problems returned
  Dim cntRes, start
	If disp_total Then
		Set cntRes = SQLQuery(cnnDB, cntStr)
	End If

	' If not empty results, set up the page.  Only display
	' 10 results per page.
	If Not listRes.EOF Then
	Dim Counter, numToDisplay, startNum
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

  Dim strColumns, intUseInoutBoard
  intUseInoutBoard = Cfg(cnnDB, "UseInoutBoard")
  If intUseInoutBoard = 1 Then
    strColumns = 7
  Else
    strColumns = 6
  End If
  
%>
<div align="center">
  <table class="Wide">
  <tr class="Head1">
    <td colspan="<%=strColumns%>">
      <%=lang(cnnDB, "Problemsfor")%>&nbsp;<% = ruid %>
      <%
        If disp_total Then
          Response.Write("&nbsp;(" & lang(cnnDB, "Total") & ":" & cntRes("total") & ")")
          cntRes.Close
        End If
      %>
    </td>
  </tr>
  <tr align="center" Class="Head2">
    <td nowrap><a href="view.asp?rep_id=<% = rep_id %>&start=<% = start %>&num=<% = numToDisplay %>&sort=1&order=<% = intIDOrder %>" class="HeadLink"><%=lang(cnnDB, "ID")%></a></td>
    <td><a href="view.asp?rep_id=<% = rep_id %>&start=<% = start %>&num=<% = numToDisplay %>&sort=2&order=<% = intTitleOrder %>" class="HeadLink"><%=lang(cnnDB, "Title")%></a></td>
    <td nowrap><a href="view.asp?rep_id=<% = rep_id %>&start=<% = start %>&num=<% = numToDisplay %>&sort=3&order=<% = intUIDOrder %>" class="HeadLink"><%=lang(cnnDB, "UserName")%></a></td>
 <% If intUseInoutBoard = 1 Then %>
     <td><div class="info">i</div></td>
 <% End If %>
    <td nowrap><a href="view.asp?rep_id=<% = rep_id %>&start=<% = start %>&num=<% = numToDisplay %>&sort=4&order=<% = intDateOrder %>" class="HeadLink"><%=lang(cnnDB, "DateSubmitted")%></a></td>
    <td><a href="view.asp?rep_id=<% = rep_id %>&start=<% = start %>&num=<% = numToDisplay %>&sort=5&order=<% = intPriOrder %>" class="HeadLink"><%=lang(cnnDB, "Priority")%></a></td>
    <td><a href="view.asp?rep_id=<% = rep_id %>&start=<% = start %>&num=<% = numToDisplay %>&sort=6&order=<% = intStatusOrder %>" class="HeadLink"><%=lang(cnnDB, "Status")%></a></td>
  </tr>
  <%
    Do While Not (listRes.EOF) AND (Counter <= (numToDisplay + start - 1))
    If Counter >= start Then
  %>
    <tr align="center" valign="center" class="Body1">
      <td nowrap><% = listRes("id") %></td>
      <td><A HREF="details.asp?id=<% = listRes("id") %>"><% = listRes("title") %></A></td>
      <td nowrap><A HREF="mailto:<% = listRes("uemail") %>?Subject=HELPDESK: Problem <% = listRes("id") %>"><% = listRes("uid") %></A></td>
  <%
    If intUseInoutBoard = 1 Then 
      dim rstUser
      set rstUser = sqlQuery(cnnDB, "SELECT sid FROM tblUsers WHERE uid='" & listRes("uid") & "'")
      If not rstUser.EOF then
        Response.Write "<td><a href=""../inout/details.asp?id=" & rstUser("sid") & """><div class=""info"">i</div></a></td>"
      Else
        Response.Write "<td>&nbsp;</td>"
      End If
      rstUser.Close
    End If
  %>
      <td nowrap><% = DisplayDate(listRes("start_date"), lhdDateOnly) %></td>
      <td nowrap><% = listRes("pname") %></td>
      <td nowrap><% = listRes("sname") %></td>
    </tr>
  <%
    End If
    Counter = Counter + 1
    listRes.MoveNext
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
      <A HREF="view.asp?rep_id=<% = rep_id %>&start=<% = startP %>&num=<% = numToDisplay %>&sort=<% = intSort %>&order=<% = intOrder %>">Previous</A>&nbsp;
    <% End If
      If Not (listRes.EOF) Then
    %>
      <A HREF="view.asp?rep_id=<% = rep_id %>&start=<% = startN %>&num=<% = numToDisplay %>&sort=<% = intSort %>&order=<% = intOrder %>">Next</A>
    <% End If %>
    </div>
  <%

    ' If no results returned:
    Else
  %>
  <div align="center">
  <table border="0" cellspacing="3" cellpadding="5" width="600">
  <tr class="Head1">
    <td colspan="6" valign="center">
      <font size="+2"><b><%=lang(cnnDB, "OpenProblemsfor")%>&nbsp;<% = ruid %></b></font>
    </td>
  </tr>
  <tr align="center" Class="Head2">
    <td nowrap><%=lang(cnnDB, "ID")%></td>
    <td><%=lang(cnnDB, "Title")%></td>
    <td nowrap><%=lang(cnnDB, "UserName")%></td>
 <% If intUseInoutBoard = 1 Then %>
     <td><div class="info">i</div></td>
 <% End If%>
    <td nowrap><%=lang(cnnDB, "DateSubmitted")%></td>
    <td><%=lang(cnnDB, "Priority")%></td>
    <td><%=lang(cnnDB, "Status")%></td>
  </tr>
  <tr align="center" class="Body1">
    <td colspan="6">
      <%=lang(cnnDB, "Noresultsfound")%>.
    </td>
  </tr>
  </table>
  </div>
<%	End If

	' Close results
	listRes.Close
	repRes.Close

	Call DisplayFooter(cnnDB, sid)
	cnnDB.Close
%>

</body>
</html>


