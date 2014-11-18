<%@ LANGUAGE="VBScript" %>
<% 
  Option Explicit
  'Buffer the response, so Response.Expires can be used
  Response.Buffer = TRUE
  Response.Expires = -1
%>


<?xml version="1.0"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

  <!--
  Liberum Help Desk, Copyright (C) 2000-2001 Doug Luxem
  Liberum Help Desk comes with ABSOLUTELY NO WARRANTY
  Please view the license.html file for the full GNU General Public License.

  Filename: details.asp
  Date:     $Date: 2002/08/28 15:31:15 $
  Version:  $Revision: 1.50.4.2 $
  Purpose:  This page displays the problem details and allows reps to
            modify them.
  -->

  <!-- 	#include file = "../public.asp" -->
  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

<head>
  <title><%=lang(cnnDB, "HelpDesk")%> - <%=lang(cnnDB, "EditProblem")%></title>
  <link rel="stylesheet" type="text/css" href="../default.css">
</head>
<body>

<%
	' Check if user has permissions for this page
	Call CheckRep(cnnDB, sid)

	' Get the problem ID
	Dim id, blnUpdate, strUpdateMessage
	id = Cint(Request.QueryString("id"))
  If Cint(Request.Form("update")) = 1 Then
    blnUpdate = True
  Else 
    blnUpdate = False
  End If

  Dim uid, uemail, uphone, ulocation, category, department, title, description, kb
  Dim priority, status, rep, time_spent, solution, notes, clost_date, start_date, close_date

  ' ===============================
  ' Update the problem
  If blnUpdate Then
    Dim oldrep

    ' Get the problem data from the form fields.
    id = Request.Form("id")
    uid = Request.Form("uid")
    uemail = Request.Form("uemail")
    uphone = Request.Form("uphone")
    ulocation = Request.Form("ulocation")
    category = Cint(Request.Form("category"))
    department = Cint(Request.Form("department"))
    title = Request.Form("title")
    priority = Cint(Request.Form("priority"))
    status = Cint(Request.Form("status"))
    rep = Cint(Request.Form("rep"))
    oldrep = Cint(Request.Form("oldrep"))
    time_spent = Request.Form("time_spent")
    solution = Request.Form("solution")
    notes = Request.Form("notes")

    If Request.Form("kb") = "on" Then
      kb = 1
    Else
      kb = 0
    End If

  ' Check for required fields (uemail, category, department, title, description)

    if Len(uid)=0 Then
      Call DisplayError(1, Lang(cnnDB, "UserName"))
    End if

    if (uemail = Cfg(cnnDB, "BaseEmail")) OR (InStr(uemail, "@")=0) Then
      Call DisplayError(1, Lang(cnnDB, "EMailAddress"))
    End if

    if Len(title)=0 Then
      Call DisplayError(1, Lang(cnnDB, "Title"))
    End if

    if (status=Cfg(cnnDB, "CloseStatus")) and (Len(solution)=0) Then
      Call DisplayError(1, Lang(cnnDB, "Solution"))
    End if
  ' Clean up fields
    uemail = Left(Trim(uemail), 50)
    uemail = Replace(uemail, "'", "''")
    uphone = Left(Trim(uphone), 50)
    uphone = Replace(uphone, "'", "''")
    ulocation = Left(Trim(ulocation), 50)
    ulocation = Replace(ulocation, "'", "''")
    title = Left(Trim(title), 50)
    title = Replace(title, "'", "''")

    time_spent = Trim(time_spent)
    If IsNumeric(time_spent) Then
      time_spent = Cint(time_spent)
    Else
      time_spent = 0
    End If


  ' Grab original description
    Dim rstDesc, strChangeNotes
    Set rstDesc = SQLQuery(cnnDB, "SELECT category, department, rep, status, priority FROM problems WHERE id=" & id)

  ' Insert actions into strChangeNotes
    If (category <> rstDesc("category")) OR (department <> rstDesc("department")) OR _
      (rep <> rstDesc("rep")) OR (status <> rstDesc("status")) OR _
      (priority <> rstDesc("priority")) Then

        If priority <> rstDesc("priority") Then
          Dim newPri, oldPri
          Set newPri = SQLQuery(cnnDB, "SELECT pname FROM priority WHERE priority_id=" & priority)
          Set oldPri = SQLQuery(cnnDB, "SELECT pname FROM priority WHERE priority_id=" & rstDesc("priority"))
          strChangeNotes = strChangeNotes & Lang(cnnDB, "PRIORITY_2") & ": " & oldPri("pname") & " => " & newPri("pname") & vbNewLine
          newPri.Close
          oldPri.Close
        End If

        If rep <> rstDesc("rep") Then
          Dim newRep, oldRep2
          Set newRep = SQLQuery(cnnDB, "SELECT uid FROM tblUsers WHERE sid=" & rep)
          Set oldRep2 = SQLQuery(cnnDB, "SELECT uid FROM tblUsers WHERE sid=" & rstDesc("rep"))
          strChangeNotes = strChangeNotes & Lang(cnnDB, "TRANSFERREPS") & ": " & oldRep2("uid") & " => " & newRep("uid") & vbNewLine
          newRep.Close
          oldRep2.Close
        End If

        If category <> rstDesc("category") Then
          Dim newCat, oldCat
          Set newCat = SQLQuery(cnnDB, "SELECT cname FROM categories WHERE category_id=" & category)
          Set oldCat = SQLQuery(cnnDB, "SELECT cname FROM categories WHERE category_id=" & rstDesc("category"))
          strChangeNotes = strChangeNotes & Lang(cnnDB, "CATEGORY_2") & ": " & oldCat("cname") & " => " & newCat("cname") & vbNewLine
          newCat.Close
          oldCat.Close
        End If

        if department <> rstDesc("department") Then
          Dim newDep, oldDep
          Set newDep = SQLQuery(cnnDB, "SELECT dname FROM departments WHERE department_id=" & department)
          Set oldDep = SQLQuery(cnnDB, "SELECT dname FROM departments WHERE department_id=" & rstDesc("department"))
          strChangeNotes = strChangeNotes & Lang(cnnDB, "DEPARTMENT_2") & ": " & oldDep("dname") & " => " & newDep("dname") & vbNewLine
          newDep.Close
          oldDep.Close
        End If

        If status <> rstDesc("status") Then
          Dim newStat, oldStat
          Set newStat = SQLQuery(cnnDB, "SELECT sname FROM status WHERE status_id=" & status)
          Set oldStat = SQLQuery(cnnDB, "SELECT sname FROM status WHERE status_id=" & rstDesc("status"))
          strChangeNotes = strChangeNotes & Lang(cnnDB, "STATUS_2") & ": " & oldStat("sname") & " => " & newStat("sname") & vbNewLine
          newStat.Close
          oldStat.Close
        End If
    End If


    ' Update the Notes
    Dim rstUpdateNotes, intPrivate, dtNoteDate, blnSendUpdateMsg
    blnSendUpdateMsg = False

    dtNoteDate = SQLDate(Now, lhdAddSQLDelim)

    If Len(notes)>0 Then
      notes = Replace(notes, "'", "''")
      If Request.Form("hidenotes") = "on" Then
        intPrivate = 1
      Else
        intPrivate = 0
        blnSendUpdateMsg = True
      End If    
      Set rstUpdateNotes = SQLQuery(cnnDB, "INSERT INTO tblNotes (id, [note], addDate, uid, private) " & _
        "VALUES (" & id & ", '" &  notes & "', " & dtNoteDate & ", '" & _
        Usr(cnnDB, sid, "uid") & "', " & intPrivate & ")")
    End If

    If Len(strChangeNotes)> 0 Then
      Set rstUpdateNotes = SQLQuery(cnnDB, "INSERT INTO tblNotes (id, [note], addDate, uid, private) " & _
        "VALUES (" & id & ", '" &  strChangeNotes & "', " & dtNoteDate & ", '" & _
        Usr(cnnDB, sid, "uid") & "', 1)")
    End If

  ' Get missing variables to enter problem

    Dim cname, catRes
    Set catRes = SQLQuery(cnnDB, "SELECT cname, rep_id FROM categories WHERE category_id=" & Request.Form("category"))
    cname = catRes("cname")

    Dim probRes
    Set probRes = SQLQuery(cnnDB, "SELECT start_date FROM problems WHERE id=" & id)
    start_date = probRes("start_date")

  ' Get the old priority
    Dim old_priority, oldPriRes
    Set oldPriRes = SQLQuery(cnnDB, "SELECT priority FROM problems WHERE id=" & id)
    old_priority = oldPriRes("priority")
    oldPriRes.Close

  ' Remove apostrophes
    description = Replace(description, "'", "''")
    solution = Replace(solution, "'", "''")

  ' All data is present
  ' Write problem into database

    Dim probStr
    probStr = "UPDATE problems SET " & _
      "uid='" & uid & "', " & _
      "uemail='" & uemail & "', " & _
      "uphone='" & uphone & "', " & _
      "ulocation='" & ulocation & "', " & _
      "category=" & category & ", " & _
      "department=" & department & ", " & _
      "title='" & title & "', " & _
      "priority=" & priority & ", " & _
      "status=" & status & ", " & _
      "rep=" & rep & ", " & _
      "kb=" & kb & ", " & _
      "time_spent=" & time_spent & ", " & _
      "solution='" & solution & "'"

    ' Add the closed date/time if the problem is closed
    If status = Cfg(cnnDB, "CloseStatus") Then
      probStr = probStr & ", close_date=" & SQLDate(Now, lhdAddSQLDelim)
    End If
    strUpdateMessage = Lang(cnnDB, "Theproblemhasbeensaved") & "."

    probStr = probStr & " WHERE id=" & id

    Set probRes = SQLQuery(cnnDB, probStr)

    If status = Cfg(cnnDB, "CloseStatus") Then
      If Not (Request.Form("noemail")="on") Then
      ' Send mail to the user'
        Call eMessage(cnnDB, "userclose", id, uemail)
      End If
    Else

      ' Notify user of update
      If (Cfg(cnnDB, "Notifyuser") = 1) And Not (Request.Form("noemail")="on") And blnSendUpdateMsg Then
        Call eMessage(cnnDB, "userupdate", id, uemail)
      End If

      'Send mail to the appropriate rep for transfered problems
      If (rep <> oldrep) Then
        Call eMessage(cnnDB, "repnew", id, Usr(cnnDB, rep, "email1"))

        'Page Rep if enabled
        If (priority >= Cfg(cnnDB, "EnablePager")) And (Len(Usr(cnnDB, rep, "email2")) > 0) Then
          Call eMessage(cnnDB, "reppager", id, Usr(cnnDB, rep, "email2"))
        End If

      ElseIf priority <> old_priority Then
        'Page Rep if enabled
        If (priority >= Cfg(cnnDB, "EnablePager")) And (Len(Usr(cnnDB, rep, "email2")) > 0) Then
          Call eMessage(cnnDB, "reppager", id, Usr(cnnDB, rep, "email2"))
        End If
      End If
    End If
  End If

  ' ===============================
  
  If Cint(Request.QueryString("reopen")) = 1 Then
    Dim strSQLOpen, rstOpenProbUpd, rstOpenNotes, dtOpenNoteDate, strOpenNote
    Dim rstOpenOldStat, rstOpenNewStat, rstOpenUserEmail

    strSQLOpen = "UPDATE problems SET " & _
      "rep = " & sid & ", " & _
      "status = " & Cfg(cnnDB, "DefaultStatus") & ", " & _
      "close_date = NULL"
    strSQLOpen = strSQLOpen & " WHERE id = " & id
    Set rstOpenProbUpd = SQLQuery(cnnDB, strSQLOpen)
    Set rstOpenNewStat = SQLQuery(cnnDB, "SELECT sname FROM status WHERE status_id=" & Cfg(cnnDB, "DefaultStatus"))
    Set rstOpenOldStat = SQLQuery(cnnDB, "SELECT sname FROM status WHERE status_id=" & Cfg(cnnDB, "CloseStatus"))
    strOpenNote = Lang(cnnDB, "STATUS_2") & ": " & rstOpenOldStat("sname") & " => " & rstOpenNewStat("sname") & vbNewLine
    dtOpenNoteDate = SQLDate(Now, lhdAddSQLDelim)
    Set rstOpenNotes = SQLQuery(cnnDB, "INSERT INTO tblNotes (id, [note], addDate, uid, private) " & _
        "VALUES (" & id & ", '" &  strOpenNote & "', " & dtOpenNoteDate & ", '" & _
        Usr(cnnDB, sid, "uid") & "', 1)")
    Set rstOpenUserEmail = SQLQuery(cnnDB, "SELECT uemail FROM problems WHERE id = " & id)
    Call eMessage(cnnDB, "usernew", id, rstOpenUserEmail("uemail"))
    rstOpenUserEmail.Close
    rstOpenNewStat.Close
    rstOpenOldStat.Close

  End If

  ' ===============================

	' Query the database for the problem info
  Dim rstProb, rstSol, rstNotes, entered_by, strProbQuery
  strProbQuery = "SELECT uid, uemail, uphone, ulocation, time_spent, department, " & _
		"category, status, priority, entered_by, rep, kb, start_date, close_date, title, description " & _
		"FROM problems WHERE id=" & id

  ' Make sure rep is viewing own problems.
  If Usr(cnnDB, sid, "RepAccess")=1 Then
    strProbQuery = strProbQuery & " AND rep=" & sid
  End If
  
	Set rstProb = SQLQuery(cnnDB, strProbQuery)
  
  If rstProb.EOF Then
    Call DisplayError(3, "Problem " & id & " could not be found in the database.")
  End If

	' Query for the solution seperately becuase SQL only
	' supports 1 blob per query
	Set rstSol = SQLQuery(cnnDB, "SELECT solution FROM problems WHERE id=" & id)

	uid = rstProb("uid")
	uemail = rstProb("uemail")
	uphone = rstProb("uphone")
	ulocation = rstProb("ulocation")
	time_spent = Cint(rstProb("time_spent"))
	department = Cint(rstProb("department"))
	category = Cint(rstProb("category"))
	status = Cint(rstProb("status"))
	priority = Cint(rstProb("priority"))
	rep = Cint(rstProb("rep"))
  kb = Cint(rstProb("kb"))
  entered_by = Cint(rstProb("entered_by"))
	start_date = rstProb("start_date")
	close_date = rstProb("close_date")
	title = rstProb("title")
	description = rstProb("description")

  ' Get the Notes for this problem
  Set rstNotes = SQLQuery(cnnDB, "SELECT * FROM tblNotes WHERE id=" & id & " ORDER BY addDate ASC")

  ' Get the solution and replace characters to make
	' it more readable.
	solution = rstSol("solution")

  ' If the rep has readonly access, or the problem is closed
  ' disable the fields
  Dim strTextDisable, strListDisable
  If Usr(cnnDB, sid, "RepAccess")=2 Or status = Cfg(cnnDB, "CloseStatus") Then
    strTextDisable = "readonly"
    strListDisable = "disabled"
  Else
    strTextDisable = ""
    strListDisable = ""
  End If
%>


<form action="details.asp" method="POST">
<input type="hidden" name="uid" value="<% = uid %>">
<input type="hidden" name="id" value="<% = id %>">
<input type="hidden" name="oldrep" value="<% = rep %>">
<input type="hidden" name="update" value="1">

<div align="center">
<table class="Normal">
<tr>
	<td colspan="2" align="right">
    <em>*</em> = <%=lang(cnnDB, "Required")%> |
	  <a href="print.asp?id=<% = id %>" target="printwindow"><%=lang(cnnDB, "PrinterFriendly")%></a>
	</td>
</tr>
<tr class="Head1">
	<td colspan="2">
		<%=lang(cnnDB, "EditProblem")%>&nbsp;<% = id %>
	</td>
</tr>

<% If blnUpdate Then %>
    <tr class="Head2">
      <td colspan="2">
        <div align="center">
          <% = strUpdateMessage %>
        </div>
      </td>
    </tr>
<% End If %>

<tr class="Body1">
	<td valign="top" width="50%" >
    <div align="center">
      <table class="Narrow">
      <tr>
        <td align="center" colspan="2">
          <b><%=lang(cnnDB, "ContactInformation")%></b>
        </td>
      </tr>
      <tr>
        <td>
          <b><%=lang(cnnDB, "UserName")%>:</b>
        </td>
        <td>
          <a href="mailto:<% = uemail %>?Subject=HELPDESK: Problem <% = id %>"><% = uid %></a>
        </td>
      </tr>
      <tr>
        <td>
          <b><%=lang(cnnDB, "EMail")%>:</b>
        </td>
        <td>
          <input type="text" name="uemail" size="20" value="<% = uemail %>" <% = strTextDisable %>><em>*</em>
        </td>
      </tr>
      <tr>
        <td>
          <b><%=lang(cnnDB, "Department")%>:</b>
        </td>
        <td>
          <SELECT NAME="department" <% = strListDisable %>>
          <%
            Dim rstDep
            Set rstDep = SQLQuery(cnnDB, "SELECT * From departments WHERE department_id > 0 ORDER BY dname ASC")
            If Not rstDep.EOF Then
            Do While Not rstDep.EOF
            If rstDep("department_id") = department Then
            %>
            <OPTION VALUE="<% = rstDep("department_id")%>" SELECTED>
            <% = rstDep("dname") %></OPTION>
            <% Else %>
            <OPTION VALUE="<% = rstDep("department_id")%>">
            <% = rstDep("dname") %></OPTION>

          <% 	End If
            rstDep.MoveNext
            Loop
            End If
          %>
          </SELECT><em>*</em>
        </td>
      </tr>
      <tr>
        <td>
          <b><%=lang(cnnDB, "Location")%>:</b>
        </td>
        <td>
          <input type="text" name="ulocation" size="20" value="<% = ulocation %>" <% = strTextDisable %>>
        </td>
      </tr>
      <tr>
        <td>
          <b><%=lang(cnnDB, "Phone")%>:</b>
        </td>
        <td>
          <input type="text" name="uphone" size="20"value="<% = uphone %>"  <% = strTextDisable %>>
        </td>
      </tr>
      <tr>
        <td>
          <b><%=lang(cnnDB, "EnteredBy")%>:</b>
        </td>
        <td>
          <% = Usr(cnnDb, entered_by, "uid") %>
        </td>
      </tr>
    </table>
	</td>
	<td valign="top" >
    <table class="Narrow">
      <tr>
        <td align="center" colspan="2">
          <b><%=lang(cnnDB, "Classification")%></b>
        </td>
      </tr>
      <tr>
        <td>
          <b><%=lang(cnnDB, "Category")%>:</b>
        </td>
        <td>
          <SELECT NAME="category" <% = strListDisable %>>
          <%
            Dim rstCat
            Set rstCat = SQLQuery(cnnDB, "SELECT * From categories WHERE category_id > 0 ORDER BY category_id ASC")
            If Not rstCat.EOF Then
            Do While Not rstCat.EOF
            If rstCat("category_id") = category Then
            %>
            <OPTION VALUE="<% = rstCat("category_id")%>" SELECTED>
            <% = rstCat("cname") %></OPTION>
            <% Else %>
            <OPTION VALUE="<% = rstCat("category_id")%>">
            <% = rstCat("cname") %></OPTION>

          <% 	End If
            rstCat.MoveNext
            Loop
            End If
          %>
          </SELECT><em>*</em>
        </td>
      </tr>
      <tr>
        <td>
          <b><%=lang(cnnDB, "Status")%>:</b>
        </td>
        <td>
          <SELECT NAME="status" <% = strListDisable %>>
          <%
            Dim rstStat
            Set rstStat = SQLQuery(cnnDB, "SELECT * From status WHERE status_id > 0 ORDER BY status_id ASC")
            If Not rstStat.EOF Then
            Do While Not rstStat.EOF
            If rstStat("status_id") = status Then
            %>
            <OPTION VALUE="<% = rstStat("status_id")%>" SELECTED>
            <% = rstStat("sname") %></OPTION>
            <% Else %>
            <OPTION VALUE="<% = rstStat("status_id")%>">
            <% = rstStat("sname") %></OPTION>

          <% 	End If
            rstStat.MoveNext
            Loop
            End If
          %>
          </SELECT><em>*</em>
        </td>
      </tr>
      <tr>
        <td>
          <b><%=lang(cnnDB, "Priority")%>:</b>
        </td>
        <td>
          <SELECT NAME="priority" <% = strListDisable %>>
          <%
            Dim rstPri
            Set rstPri = SQLQuery(cnnDB, "SELECT * From priority WHERE priority_id > 0 ORDER BY priority_id ASC")
            If Not rstPri.EOF Then
            Do While Not rstPri.EOF
            If rstPri("priority_id") = priority Then
            %>
            <OPTION VALUE="<% = rstPri("priority_id")%>" SELECTED>
            <% = rstPri("pname") %></OPTION>
            <% Else %>
            <OPTION VALUE="<% = rstPri("priority_id")%>">
            <% = rstPri("pname") %></OPTION>

          <% 	End If
            rstPri.MoveNext
            Loop
            End If
          %>
          </SELECT><em>*</em>
        </td>
      </tr>
      <tr>
        <td>
          <b><%=lang(cnnDB, "AssignTo")%>:</b>
        </td>
        <td>
          <SELECT NAME="rep" <% = strListDisable %>>
          <%
            Dim rstRes
            Set rstRes = SQLQuery(cnnDB, "SELECT * From tblUsers WHERE IsRep=1 AND RepAccess<>2 ORDER BY uid ASC")
            If Not rstRes.EOF Then
            Do While Not rstRes.EOF
            If rstRes("sid") = rep Then
            %>
            <OPTION VALUE="<% = rstRes("sid")%>" SELECTED>
            <% = rstRes("uid") %></OPTION>
            <% Else %>
            <OPTION VALUE="<% = rstRes("sid")%>">
            <% = rstRes("uid") %></OPTION>

          <% 	End If
            rstRes.MoveNext
            Loop
            End If
          %>
          </SELECT><em>*</em>
        </td>
      </tr>
      <tr>
        <td>
          <b><%=lang(cnnDB, "TimeSpent")%>:</b>
        </td>
        <td>
          <input type="text" size="4" name="time_spent" value="<% = time_spent %>"  <% = strTextDisable %>>(<%=lang(cnnDB, "minutes")%>)
        </td>
      </tr>
      <tr>
        <td>
          <b><%=lang(cnnDB, "StartDate")%>:</b>
        </td>
        <td>
          <% = DisplayDate(start_date, lhdDateTime) %>
        </td>
      </tr>
      <tr>
        <td>
          <b><%=lang(cnnDB, "CloseDate")%>:</b>
        </td>
        <td>
          <% = DisplayDate(close_date, lhdDateTime) %>
        </td>
      </tr>
    </table>
    <% If status = Cfg(cnnDB, "CloseStatus") Then %>
      <div align="center">
        <b><a href="details.asp?id=<% = id %>&reopen=1"><%=lang(cnnDB, "ReopenProblem")%></a></b>
      </div>
    <% End If %>
	</td>
</tr>
<tr class="Head2">
  <td colspan="2" >
    <%=lang(cnnDB, "ProblemInformation")%>:
  </td>
</tr>
<tr class="Body1">
	<td colspan="2" >
		<b><%=lang(cnnDB, "Title")%>:</b><em>*</em><br />
    <input type="text" name="title" size="50" value="<% = title %>"  <% = strTextDisable %>>
		
    <p>
		<b><%=lang(cnnDB, "Description")%>:</b><br />
		<textarea readonly rows="8" cols="80" name="disp_description"><% = description %></textarea>
	</td>
</tr>
<tr class="Head2">
  <td colspan="2" >
    <%=lang(cnnDB, "Notes")%>:
  </td>
</tr>
<tr class="Body1">
  <td colspan="2" >
    <% If rstNotes.EOF Then %>
       <%=lang(cnnDB, "NoAvailableNotes")%>
     <% 
      Else
      Do While Not rstNotes.EOF
      
        Response.Write("<b>[")
        Response.Write(DisplayDate(rstNotes("addDate"), lhdDateTime) & " - " & rstNotes("uid") & "]")
        If rstNotes("private") = 1 Then
          Response.Write (" - PRIVATE")
        End If
        Response.Write("</b><br />" & vbNewLine)
        Response.Write(Replace(rstNotes("note"), vbNewLine, "<br />"))
        Response.Write("<p>" & vbNewLine)

        rstNotes.MoveNext
      Loop
      End If %>
  </td>
</tr>
<% If status <> Cfg(cnnDB, "CloseStatus") Then %>
<tr class="Head2">
  <td colspan="2" >
    <%=lang(cnnDB, "EnterAdditionalNotes")%>:
  </td>
</tr>
<tr class="Body1">
	<td colspan="2" >
		<textarea rows="8" cols="80" name="notes"  <% = strTextDisable %>></textarea><br />
    <input type="checkbox" name="hidenotes" <% = strListDisable %>>&nbsp;<%=lang(cnnDB, "HideFromEndUser")%>
	</td>
</tr>
<% End If %>
<tr class="Head2">
	<td colspan="2" >
		<%=lang(cnnDB, "Solution")%>:
  </td>
</tr>
<tr class="Body1">
  <td colspan="2">
    <textarea rows="8" cols="80" name="solution"  <% = strTextDisable %>><% = solution %></textarea>
    <% If Cfg(cnnDB, "EnableKB") <> 0 Then
         If kb = 0 Then %>
          <input type="checkbox" name="kb" <% = strListDisable %>>&nbsp;<%=lang(cnnDB, "EnterinKnowledgeBase")%>
        <% Else %>
          <input type="checkbox" name="kb" checked <% = strListDisable %>>&nbsp;<%=lang(cnnDB, "EnterinKnowledgeBase")%>
        <% End If %>
    <% End If %>
  </td>
</tr>
</table>
<% If status <> Cfg(cnnDB, "CloseStatus") And Usr(cnnDB, sid, "RepAccess") <> 2 Then
     Response.Write("<tr class=""Head2"" align=""center""><td colspan=""2"">")
     If Cfg(cnnDB, "EmailType") <> 0 Then
%>
      <input type="checkbox" name="noemail" <% = strListDisable %>>&nbsp;<%=lang(cnnDB, "Dontsendemailtouser")%>
      <p>
    <% End If %>
      <input type="submit" value="<%=lang(cnnDB, "SaveProblem")%>" name="B1">
      </td></tr>
    <% End If %>
</div>
</form>
<%
	' close record sets
	rstCat.Close
	rstDep.Close
	rstStat.Close
	rstPri.Close
	rstRes.Close

  rstProb.Close
  rstSol.Close
  rstNotes.Close

	Call DisplayFooter(cnnDB, sid)
	cnnDB.Close
%>

</body>

</html>
