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

  Filename: details.asp
  Date:     $Date: 2002/06/15 23:49:20 $
  Version:  $Revision: 1.50.4.1 $
  Purpose:  This page displays the details of a problem.  The problem id is taken
  is an input from the URL. (details.asp?id=xx)  If the problem is not closed, the
  update form posts to update.asp.
  -->

  <!-- 	#include file = "../public.asp" -->
  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
    Session.Timeout = 40
  %>

  <head>
    <title><%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "ProblemDetails")%></title>
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>

    <%
      ' See if authenticated
      Call CheckUser(cnnDB, sid)
      ' Get the username

      dim uid
      uid = Usr(cnnDB, sid, "uid")

      ' Get the problem ID
      Dim id
      id = Cint(Request.QueryString("id"))
      If id = 0 Then
        If Len(Request.Form("id")) = 0 Then
          Call DisplayError(3, lang(cnnDB, "AproblemIDnumberisrequired"))
        End If
        id = Cint(Request.Form("id"))
      End If

      ' Generate a query, making sure to search for id and uid so a user
      ' cannot view someone else's problem.
      Dim queryStr, rstProblem, rstSolution

      queryStr = _
      "SELECT p.id, p.uid, p.uemail, p.uphone, p.ulocation, d.dname, p.start_date, p.status, s.sname, " & _
      "p.close_date, c.cname, r.uid As ruid, r.email1 As remail, r.fname, p.title, p.solution, p.description " & _
      "FROM (((problems AS p " & _
      "INNER JOIN departments AS d ON p.department = d.department_id) " & _
      "INNER JOIN status AS s ON p.status = s.status_id) " & _
      "INNER JOIN tblUsers AS r ON p.rep = r.sid) " & _
      "INNER JOIN categories AS c ON p.category = c.category_id " & _
      "WHERE p.id=" & id & " AND p.uid='" & uid & "'"

      Set rstProblem = SQLQuery(cnnDB, queryStr)

      ' If no results are returned, display an error
      If rstProblem.EOF Then
        cnnDB.Close
        Call DisplayError(3, lang(cnnDB, "ProblemID") & "&nbsp;" & id & "&nbsp;" & lang(cnnDB, "wasnotfoundinthedatabase") & ".")
      End If

      ' If it is a closed problem, get the solution
      Set rstSolution = SQLQuery(cnnDB, "SELECT solution FROM problems WHERE id=" & id & " AND uid='" & uid & "'")

      ' Get the Notes for this problem
      Dim rstNotes
      Set rstNotes = SQLQuery(cnnDB, "SELECT * FROM tblNotes WHERE id=" & id & " AND private=0 ORDER BY addDate ASC")

      ' Display The problem info, and if OPEN allow some updates
    %>

    <div align="center">
      <table Class="Normal">
        <tr>
          <td colspan="2" align="right">
            <em>*</em> = <%=lang(cnnDB, "Required")%> |
            <a href="print.asp?id=<% = id %>" target="printwindow"><%=lang(cnnDB, "PrinterFriendly")%></a>
          </td>
        </tr>
        <tr class="Head1">
          <td colspan="2">
            <%=lang(cnnDB, "DetailsforProblem")%>&nbsp;<% = id %>
          </td>
        </tr>
        <tr class="Body1">
          <td colspan="2">
            <table class="Normal" border="0" cellspacing="0">
              <tr>
                <td width="125" valign="top">
                  <b><%=lang(cnnDB, "ProblemID")%>:</b>
                </td>
                <td>
                  <% = id %>
                </td>
              </tr>
              <tr>
                <td>
                  <b><%=lang(cnnDB, "UserName")%>:</b>
                </td>
                <td>
                  <% = uid %>
                </td>
              </tr>
              <tr>
                <td>
                  <b><%=lang(cnnDB, "EMail")%>:</b>
                </td>
                <td>
                  <% = rstProblem("uemail") %>
                </td>
              </tr>
              <tr>
                <td>
                  <b><%=lang(cnnDB, "Phone")%>:</b>
                </td>
                <td>
                  <% = rstProblem("uphone") %>
                </td>
              </tr>
              <tr>
                <td>
                  <b><%=lang(cnnDB, "Location")%>:</b>
                </td>
                <td>
                  <% = rstProblem("ulocation") %>
                </td>
              </tr>
              <tr>
                <td>
                  <b><%=lang(cnnDB, "StartDate")%>:</b>
                </td>
                <td>
                  <% = DisplayDate(rstProblem("start_date"), lhdDateTime) %>
                </td>
              </tr>
              <% If rstProblem("status") = Cfg(cnnDB, "CloseStatus") Then %>
                <tr>
                  <td>
                    <b><%=lang(cnnDB, "CloseDate")%>:</b>
                  </td>
                  <td>
                    <% = DisplayDate(rstProblem("close_date"), lhdDateTime) %>
                  </td>
                </tr>
              <% End If %>
              <tr>
                <td>
                  <b><%=lang(cnnDB, "Department")%>:</b>
                </td>
                <td>
                  <% = rstProblem("dname") %>
                </td>
              </tr>
              <tr>
                <td>
                  <b><%=lang(cnnDB, "Category")%>:</b>
                </td>
                <td>
                  <% = rstProblem("cname") %>
                </td>
              </tr>
              <tr>
                <td>
                  <b><%=lang(cnnDB, "AssignedTo")%>:</b>
                </td>
                <td>
                  <a href="mailto:<% = rstProblem("remail") %>?Subject=<%=lang(cnnDB, "HelpDesk")%>:&nbsp;<%=lang(cnnDB, "Problem")%>&nbsp;<% = id %>"><% = rstProblem("fname") %></a>
                </td>
              </tr>
              <tr>
                <td>
                  <b><%=lang(cnnDB, "Status")%>:</b>
                </td>
                <td>
                  <% = rstProblem("sname") %>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr class="Head2">
          <td>
            <%=lang(cnnDB, "ProblemInformation")%>:
          </td>
        </tr>
        <tr class="Body1">
          <td>
            <form>
              <b><%=lang(cnnDB, "Title")%>:</b><br>
              <input type="text" size="50" readonly name="display_title" value="<% = rstProblem("title") %>">
              <p>
              <b><%=lang(cnnDB, "Description")%>:</b><br>
              <textarea readonly name="display_desc" rows="8" cols="80"><% = rstProblem("description") %></textarea>
            </form>
          </td>
        </tr>
        <tr class="Head2">
          <td>
            <%=lang(cnnDB, "Notes")%>:
          </td>
        </tr>
        <tr class="Body1">
          <td colspan="2" >
           <% If rstNotes.EOF Then 
               Response.Write lang(cnnDB, "NoAvailableNotes")
              Else
              Do While Not rstNotes.EOF
              
                Response.Write ("<b>[") & _
                  (DisplayDate(rstNotes("addDate"), lhdDateTime)) & _
                  (" - " & rstNotes("uid") & "]") & _
                  ("</b><br />" & vbNewLine) & _
                  (Replace(rstNotes("note"), vbNewLine, "<br />")) & _
                  ("<p>&nbsp;</p>" & vbNewLine)

                rstNotes.MoveNext
              Loop
              End If %>
          </td>
        </tr>
        <% If rstProblem("status") <> Cfg(cnnDB, "CloseStatus") Then %>
          <form method="post" action="update.asp">
            <input type="hidden" name="id" value="<% = id %>">
            <tr class="Head2">
              <td>
                <%=lang(cnnDB, "EnterAdditionalNotes")%>:
              </td>
            </tr>
            <tr class="Body1">
              <td>
                <div align="center">
                  <textarea name="notes" rows="8" cols="80" wrap="on"></textarea>
                </div>
              </td>
            </tr>
            <tr class="Head2">
              <td>
                <div align="center">
                  <input type="submit" value="<%=lang(cnnDB, "UpdateProblem")%>" name="B1">&nbsp;<input type="reset" value="<%=lang(cnnDB, "ClearNotes")%>" name="B2">
                </div>
              </td>
            </tr>
          </form>
        <% Else %>
          <tr class="Head2">
            <td>
              <%=lang(cnnDB, "Solution")%>:
            </td>
          </tr>
          <tr class="Body1">
            <td>
              <div align="center">
                <form><textarea name="display_solution" rows="8" cols="80" wrap="on"><% = rstSolution("solution") %></textarea></form>
              </div>
            </td>
          </tr>
        <% End If %>
      </table>
      <p><a href="view.asp"><%=lang(cnnDB, "ProblemListing")%></a></p>
   </div>

    <%
      ' Close Results
      rstProblem.Close
      rstNotes.Close
      rstSolution.Close

      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
