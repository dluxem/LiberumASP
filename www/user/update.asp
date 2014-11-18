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

  Filename: update.asp
  Date:     $Date: 2002/06/15 23:49:20 $
  Version:  $Revision: 1.50.4.1 $
  Purpose:  Updates the notes in the database.
  -->

  <!-- 	#include file = "../public.asp" -->
  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title><%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "ProblemUpdated")%></title>
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>
    <%
      ' See if authenticated
      Call CheckUser(cnnDB, sid)

      ' Get the username, problem id and the additional
      ' notes that will be entered.
      Dim uid, id, notes

      uid = 	Usr(cnnDB, sid, "uid")
      id = Request.Form("id")
      notes = Request.Form("notes")

    ' Check for required fields (uemail, category, department, title, description)

      if Len(notes)=0 Then
        cnnDB.Close
        Call DisplayError(1, lang(cnnDB, "AdditionalNotes"))
      End if


    ' Get old description
      Dim description, rep, rstOldRep, rstUpdateNotes, dtNoteDate
      Set rstOldRep = SQLquery(cnnDB, "SELECT rep FROM problems WHERE id=" & id)
      rep = rstOldRep("rep")

      dtNoteDate = SQLDate(Now, lhdAddSQLDelim)

    ' Update description, making sure only valid characters are in the description
      notes = Replace(notes, "'", "''")

      Set rstUpdateNotes = SQLQuery(cnnDB, "INSERT INTO tblNotes (id, [note], addDate, uid, private) " & _
        "VALUES (" & id & ", '" &  notes & "', " & dtNoteDate & ", '" & _
        Usr(cnnDB, sid, "uid") & "', 0)")

    Call eMessage(cnnDB, "repupdate", id, Usr(cnnDB, rep, "email1"))

    %>
    <div align="center">
      <table class="Normal">
      <tr class="Head1">
        <td colspan="2">
          <%=lang(cnnDB, "Problem")%>&nbsp;<% = id %>&nbsp;<%=lang(cnnDB, "isUpdated")%>
        </td>
      </tr>
      <tr class="Body1">
        <td>
          <div align="center">
            <%=lang(cnnDB, "Viewthedetailsof")%> <a href="details.asp?id=<% = id %>"><%=lang(cnnDB, "Problem")%>&nbsp;<% = id %></a>.
          </div>
        </td>
      </tr>
      </table>
    </div>

    <%
      ' Close records
      rstOldRep.Close

      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>

  </body>
</html>
