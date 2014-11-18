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
    <title><%=lang(cnnDB, "HelpDesk")%> - <%=lang(cnnDB, "ProblemDetails")%></title>
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>
    <%
      ' Check permissions
      Call CheckKB(cnnDB, sid)

      ' Get the problem ID
      Dim id
      id = Cint(Request.QueryString("id"))

      ' Generate a query, making sure to search for id and uid so a user
      ' cannot view someone else's problem.
      Dim queryStr, rstProbList, rstNotes

      queryStr = _
      "SELECT p.id, p.uid, p.uemail, p.uphone, p.ulocation, d.dname, p.start_date, p.status, s.sname, " & _
      "p.close_date, c.cname, r.uid As ruid, r.email1 As remail, r.fname, p.title, p.solution, p.description " & _
      "FROM (((problems AS p " & _
      "INNER JOIN departments AS d ON p.department = d.department_id) " & _
      "INNER JOIN status AS s ON p.status = s.status_id) " & _
      "INNER JOIN tblUsers AS r ON p.rep = r.sid) " & _
      "INNER JOIN categories AS c ON p.category = c.category_id " & _
      "WHERE p.id=" & id & " AND p.kb=1 AND p.status=" & Cfg(cnnDB, "CloseStatus")
      Set rstProbList = SQLQuery(cnnDB, queryStr)
      Set rstNotes = SQLQuery(cnnDB, "SELECT * FROM tblNotes WHERE id=" & id & " AND private=0 ORDER BY addDate ASC")

      ' If no results are returned, display an error
      If rstProbList.EOF Then
        cnnDB.Close
        Call DisplayError(3, "Problem ID " & id & " was not found in the database.")
      Else

      Dim description, solution
      description = Replace(rstProbList("description"), vbNewLine, "<br>")
      description = Replace(description, "[", "<b>[")
      description = Replace(description, "]", "]</b>")

      ' If it is a closed problem, get the solution
      If rstProbList("status") = Cfg(cnnDB, "CloseStatus") Then
        Dim solRes
        Set solRes = SQLQuery(cnnDB, "SELECT solution FROM problems WHERE id=" & id)
        solution = Replace(solRes("solution"), vbNewLine, "<br>")
        solution = Replace(solution, "[", "<b>[")
        solution = Replace(solution, "]", "]</b>")
      End If



      ' Display The problem info, and if OPEN allow some updates
    %>

    <div align="center">
      <table Class="Wide">
        <tr class="Head1">
          <td colspan="2">
            <b><%=lang(cnnDB, "DetailsforProblem")%>&nbsp;<% = id %></b>
          </td>
        </tr>
        <tr class="Body1" >
          <td>
            <table border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="45%">
                  <b><%=lang(cnnDB, "ProblemID")%>:</b>
                </td>
                <td>
                  <% = id %>
                </td>
              </tr>
              <tr>
                <td>
                  <b><%=lang(cnnDB, "Title")%>:</b>
                </td>
                <td>
                  <% = rstProbList("title") %>
                </td>
              </tr>
              <tr>
                <td>
                  <b><%=lang(cnnDB, "StartDate")%>:</b>
                </td>
                <td>
                  <% = DisplayDate(rstProbList("start_date"), lhdDateTime) %>
                </td>
              </tr>
              <tr>
                <td>
                  <b><%=lang(cnnDB, "CloseDate")%>:</b>
                </td>
                <td>
                  <% = DisplayDate(rstProbList("close_date"), lhdDateTime) %>
                </td>
              </tr>
              <tr>
                <td>
                  <b><%=lang(cnnDB, "Category")%>:</b>
                </td>
                <td>
                  <% = rstProbList("cname") %>
                </td>
              </tr>
            </table>
          </td>
        <tr class="Head2">
          <td colspan="2">
            <%=lang(cnnDB, "Description")%>:
          </td>
        </tr>
        <tr class="Body1">
          <td colspan="2">
            <% = description %>
          </td>
        </tr>
        <tr class="Head2">
          <td colspan="2">
            <%=lang(cnnDB, "Notes")%>:
          </td>
        </tr>          
        <tr class="Body1">
          <td colspan="2">
            <% If rstNotes.EOF Then %>
               <%=lang(cnnDB, "NoAvailableNotes")%>
             <% 
              Else
              Do While Not rstNotes.EOF
              
                Response.Write("<b>[")
                Response.Write(rstNotes("addDate") & " - " & rstNotes("uid") & "]")
                Response.Write("</b><br />" & vbNewLine)
                Response.Write(Replace(rstNotes("note"), vbNewLine, "<br />"))
                Response.Write("<p>" & vbNewLine)

                rstNotes.MoveNext
              Loop
              End If %>
          </td>
        </tr>
        <% If rstProbList("status") = Cfg(cnnDB, "CloseStatus") Then %>
          <tr class="Head2">
            <td colspan="2">
              <%=lang(cnnDB, "Solution")%>:
            </td>
          </tr>
          <tr class="Body1">
            <td colspan="2">
              <% = solution %>
            </td>
          </tr>
        <% SolRes.Close 
           End If %>
      </table>      <form>
      <input type=button value="<%=lang(cnnDB, "CloseThisWindow")%>" onClick="javascript:window.close();">
      </form>
    </div>

    <%
      End If

      ' Close Results
      rstProbList.Close

      cnnDB.Close

    %>
  </body>
</html>
