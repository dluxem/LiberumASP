<%@ LANGUAGE="VBScript" %>
<% 
  Option Explicit
  Response.Buffer = TRUE
%>

<?xml version="1.0"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

  <!--
  Liberum Help Desk, Copyright (C) 2000-2001 Doug Luxem
  Liberum Help Desk comes with ABSOLUTELY NO WARRANTY
  Please view the license.html file for the full GNU General Public License.

  Filename: default.asp
  Date:     $Date: 2002/06/15 23:49:20 $
  Version:  $Revision: 1.50.4.1 $
  Purpose:  Form for submitting keywords to search knowledge base.
  -->

  <!-- 	#include file = "../public.asp" -->
  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title><%=lang(cnnDB, "HelpDesk")%> - <%=lang(cnnDB, "KnowledgeBase")%></title>
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>

    <%
      ' See if authenticated
      Call CheckKB(cnnDB, sid)
      Dim blnSearch
      blnSearch = False
      
      If (Request.Form("search") = "1") And Len(Request.Form("keywords")) > 0 Then
        blnSearch = True
      End If

      If blnSearch Then
        Dim queryStr, strWhere1, strWhere2, keywords, srchRes, varWordList, blAllOff
        keywords = Trim(Request.Form("keywords"))
        varWordList = Split(keywords, " ")

        queryStr = "SELECT id, title, start_date, close_date FROM problems"
        strWhere1 = "WHERE (kb=1 AND status=" & Cfg(cnnDB, "CloseStatus") & ")"

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

        queryStr = queryStr & " " & strWhere1 & strWhere2 & " ORDER BY start_date ASC"
        Set srchRes = SQLQuery(cnnDB, queryStr)

      End If


    If Not blnSearch Then
    %>
    <div align="center">
    <table class="Normal">
      <tr class="Head1">
        <td>
          <%=lang(cnnDB, "KnowledgeBase")%>
        </td>
      </tr>
      <tr class="Body1">
        <td>
          <form method="post" action="default.asp">
            <input type="hidden" name="search" value="1">
            <div align="center">
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
              <input type="submit" value="<%=lang(cnnDB, "Search")%>">
            </div>
          </form>
        </td>
      </tr>
    </table>
    </div>
    <p>
    <% Else

      If srchREs.EOF Then
    %>
      <div align="center">
        <table class="Wide">
          <tr class="Head1">
            <td>
              <%=lang(cnnDB, "KnowledgeBase")%>
            </td>
          </tr>
          <tr class="Body1">
            <td>
              <div align="center">
                <b><%=lang(cnnDB, "Noresultsfound")%>.</b>
              </div>
            </td>
          </tr>
        </table>
      </div>
      <p>
    <%	Else %>
      <div align="center">
      <table class="Wide">
      <tr class="Head1">
        <td colspan="3">
          <%=lang(cnnDB, "KnowledgeBase")%>
        </td>
      </tr>
      <tr class="Head2">
        <td width="250">
          <div align="center">
            <%=lang(cnnDB, "Title")%>
          </div>
        </td>
        <td>
          <div align="center">
            <%=lang(cnnDB, "StartDate")%>
          </div>
        </td>
        <td>
          <div align="center">
            <%=lang(cnnDB, "CloseDate")%>
          </div>
        </td>
      </tr>
      <%	Do While Not srchRes.EOF %>
        <tr class="Body1">
          <td>
            <div align="center">
              <a href="details.asp?id=<% = srchRes("id") %>"><% = srchRes("title") %></a>
            </div>
          </td>
          <td>
            <div align="center">
              <% = DisplayDate(srchRes("start_date"), lhdDateOnly) %>
            </div>
          </td>
          <td>
            <div align="center">
              <% = DisplayDate(srchRes("close_date"), lhdDateOnly) %>
            </div>
          </td>
        </tr>
      <%
          srchRes.MoveNext
        Loop
       %>
      </table>
      </div>
      <p>
    <%
      End If	' EOF
      Response.Write("<div align=""center""><a href=""default.asp"">" & lang(cnnDB, "SearchAgain") & "</a></div>")
      End If 	' search=1

    Call DisplayFooter(cnnDB, sid)
    cnnDB.Close

    %>
  </body>
</html>
