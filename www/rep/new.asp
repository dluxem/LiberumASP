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

  Filename: new.asp
  Date:     $Date: 2002/08/28 15:30:54 $
  Version:  $Revision: 1.51.2.1.2.2 $
  Purpose:  This page displays the form used for entering new problems.
  -->

  <!-- 	#include file = "../public.asp" -->
  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

<head>
<title><%=lang(cnnDB, "HelpDesk")%> - <%=lang(cnnDB, "NewProblem")%></title>
<link rel="stylesheet" type="text/css" href="../default.css">
</head>
<body>

<%
	' Check if user has permissions for this page
	Call CheckRep(cnnDB, sid)

  Dim blnSubmitNew, strSubmitResults
  If Cint(Request.Form("save")) = 1 Then
    blnSubmitNew = True
  Else
    blnSubmitNew = False
  End If

  ' ==============================================
  ' Save problem
  If blnSubmitNew Then
    ' Get the information from the form fields
    Dim uid, uemail, uphone, ulocation, category, department, title, description
    Dim priority, status, rep, time_spent, solution, entered_by, uselectid, kb

    uselectid = Request.Form("uselectid")
    uid = Request.Form("uid")
    uemail = Request.Form("uemail")
    uphone = Request.Form("uphone")
    ulocation = Request.Form("ulocation")
    category = Cint(Request.Form("category"))
    department = Request.Form("department")
    title = Request.Form("title")
    description = Request.Form("description")
    priority = Cint(Request.Form("priority"))
    status = Cint(Request.Form("status"))
    rep = Cint(Request.Form("rep"))
    time_spent = Cint(Request.Form("time_spent"))
    solution = Request.Form("solution")

    If Request.Form("kb") = "on" Then
      kb = 1
    Else
      kb = 0
    End If

    ' Check for required fields (uemail, category, department, title, description)
    If uselectid <> 0 then 
      uid = usr(cnnDB, uselectid, "uid") 
      uemail = usr(cnnDB, uselectid, "email1") 
      uphone = usr(cnnDB, uselectid, "phone") 
      ulocation = usr(cnnDB, uselectid, "location1") 
      department = usr(cnnDB, uselectid, "department") 
    Else 
      If Len(uid)=0 Then
        Call DisplayError(1, lang(cnnDB, "UserName"))
      End if

      if uemail = Cfg(cnnDB, "BaseEmail") Then
        Call DisplayError(1, lang(cnnDB, "EMailAddress"))
      End if
    End If

    if category = 0 Then
      Call DisplayError(1, lang(cnnDB, "Category"))
    End if

    if (department = 0) And (uselectid = 0) Then
      Call DisplayError(1, lang(cnnDB, "Department"))
    End if

    if Len(title)=0 Then
      Call DisplayError(1, lang(cnnDB, "Title"))
    Elseif Len(title) > 50 Then
      title = Trim(title)
      title = Left(title, 50)
    End if

    if Len(description)=0 Then
      Call DisplayError(1, lang(cnnDB, "Description"))
    End if

    if (status=Cfg(cnnDB, "CloseStatus")) and (Len(solution)=0) Then
      Call DisplayError(1, lang(cnnDB, "Solution"))
    End if

    ' Get missing variables to enter problem
    Dim id

    Dim dname, depRes
    Set depRes = SQLQuery(cnnDB, "SELECT dname FROM departments WHERE department_id=" & Request.Form("department"))
    dname = depRes("dname")

    Dim cname, catRes
    Set catRes = SQLQuery(cnnDB, "SELECT cname, rep_id FROM categories WHERE category_id=" & Request.Form("category"))
    cname = catRes("cname")

    entered_by = sid

    ' Get the problem ID number then immediately update it
    id = GetUnique(cnnDB, "problems")

    ' Convert strings to valid SQL strings
    On Error Resume Next
    uphone = Replace(uphone,"'","''")
    ulocation = Replace(ulocation,"'","''")
    title = Replace(title,"'","''")
    description = Replace(description,"'","''")
    solution = Replace(solution,"'","''")
    On Error Goto 0

    ' All data is present
    ' Write problem into database
    Dim probStr

    ' If status is closed, then include the closed date/time
      If status = Cfg(cnnDB, "CloseStatus") Then
        probStr = "INSERT INTO problems (id, uid, uemail, uphone, ulocation, " & _
        "category, department, title, description, priority, status, start_date, rep, time_spent, " & _
        "close_date, entered_by, solution, kb) " & _
        "VALUES (" & id & ",'" & uid & "','" & uemail & "','" & uphone & "','" & _
        ulocation & "'," & category & "," & department & ",'" & title & "','" & _
        description & "'," & priority & "," & status & "," & SQLDate(Now, lhdAddSQLDelim) & "," & rep & "," & time_spent & _
        "," & SQLDate(Now, lhdAddSQLDelim) & "," & entered_by & ",'" & solution & "', " & kb & ")"
      Else
        probStr = "INSERT INTO problems (id, uid, uemail, uphone, ulocation, " & _
        "category, department, title, description, priority, status, start_date, rep, time_spent, " & _
        "entered_by, solution, kb) " & _
        "VALUES (" & id & ",'" & uid & "','" & uemail & "','" & uphone & "','" & _
        ulocation & "'," & category & "," & department & ",'" & title & "','" & _
        description & "'," & priority & "," & status & "," & SQLDate(Now, lhdAddSQLDelim) & "," & rep & "," & time_spent & _
        "," & entered_by & ",'" & solution & "'," & kb & ")"
      End If

      Dim problemRes
      Set problemRes = SQLQuery(cnnDB, probStr)

    ' Get support rep information for later
      Dim remail, repRes
      Set repRes = SQLQuery(cnnDB, "SELECT * FROM tblUsers WHERE sid=" & rep)
      remail = repRes("email1")

    ' Send mail to the user and support rep if the problem
    ' was not closed at the time of being entered.
    Dim strSubject, strBody

    If status<>Cfg(cnnDB, "CloseStatus") Then
      ' Send mail to the user
      If Not (Request.Form("noemail")="on") Then
        Call eMessage(cnnDB, "usernew", id, uemail)
      End If

      ' Send mail to the Rep
      Call eMessage(cnnDB, "repnew", id, remail)

      'Page Rep if enabled
      If (priority >= Cfg(cnnDB, "EnablePager")) And (Len(Usr(cnnDB, rep, "email2")) > 0) Then
        Call eMessage(cnnDB, "reppager", id, Usr(cnnDB, rep, "email2"))
      End If

    End If

    ' If the problem is closed when being entered, send
    ' a different email to the user.
    If status=Cfg(cnnDB, "CloseStatus") Then
      If Not (Request.Form("noemail")="on") Then
        ' Send mail to the user
        Call eMessage(cnnDB, "userclose", id, uemail)
      End If
    End If
    strSubmitResults = Lang(cnnDB, "Problem") & " " & id & " " & Lang(cnnDB, "hasbeenentered") & "."
  End If
  ' ==============================================

  Dim rstDepList, rstCatList, rstStatList, rstPriList, rstRepList
%>
<form action="new.asp" method="POST">
<input type="hidden" name="save" value="1">

<div align="center">
  <table class="Normal">
    <tr>
      <td colspan="2" align="right">
        <em>*</em> = <%=lang(cnnDB, "Required")%>
      </td>
    </tr>
    <tr class="Head1">
      <td colspan="2">
        <%=lang(cnnDB, "SubmitANewProblem")%>
      </td>
    </tr>
    <% If blnSubmitNew Then %>
      <tr class="Head2">
        <td colspan="2">
          <div align="center">
            <% = strSubmitResults %>
          </div>
        </td>
      </tr>
    <% End If %>
    <tr class="Body1">
      <td valign="top" align="center" width="50%">
        <table class="narrow" border="0">
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
              <input type="text" name="uid" size="20"><em>*</em>
            </td>
          </tr>
          <tr>
            <td>
              <b><%=lang(cnnDB, "EMail")%>:</b>
            </td>
            <td>
              <input type="text" name="uemail" size="20" value="<% = Cfg(cnnDB, "BaseEmail") %>"><em>*</em>
            </td>
          </tr>
          <tr>
            <td>
              <b><%=lang(cnnDB, "Department")%>:</b>
            </td>
            <td>
              <SELECT NAME="department">
              <OPTION VALUE="0" SELECTED><%=lang(cnnDB, "SelectDepartment")%></OPTION>
              <%
                Set rstDepList = SQLQuery(cnnDB, "SELECT * From departments WHERE department_id > 0 ORDER BY dname ASC")
                If not rstDepList.EOF Then
                Do While Not rstDepList.EOF
              %>
              <OPTION VALUE="<% = rstDepList("department_id")%>">
              <% = rstDepList("dname") %></OPTION>

              <% 		rstDepList.MoveNext
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
              <input type="text" name="ulocation" size="20">
            </td>
          </tr>
          <tr>
            <td>
              <b><%=lang(cnnDB, "Phone")%>:</b>
            </td>
            <td>
               <input type="text" name="uphone" size="20">
          </tr>
          <% 
          If cfg(cnnDB, "useSelectUser") = 1 Then 
          %>
            <tr>
              <td colspan="2">
                <div align="center">
                  <b>--- <%=lang(cnnDB, "Or")%> ---</b>
                </div>
              </td>
            </tr>
            <tr>
              <td>
                <b><%=lang(cnnDB, "SelectUser")%>:</b>
              </td>
              <td>
                <select name="uselectid">
                  <option value="0" selected><%=lang(cnnDB, "SelectUser")%></option>
                  <%
                    ' Get list of users to diplay
                    Dim rstUser
                    Set rstUser = SQLQuery(cnnDB, "SELECT * FROM tblUsers WHERE sid > 0 ORDER BY uid ASC")
                    If not rstUser.EOF Then
                      Do While Not rstUser.EOF 
                      %>
                        <option value="<% = rstUser("sid")%>">
                        <% = rstUser("uid") %>&nbsp;(<% = rstUser("fname") %>)
                        </option>
                      <% 
                      rstUser.MoveNext
                      Loop
                    End If
                  %>
                </select>
              </td>
            </tr>
          <% 
          end if
          %>
        </table>
      </td>
      <td valign="top">
        <table class="narrow" border="0">
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
              <SELECT NAME="category">
              <OPTION VALUE="0" SELECTED><%=lang(cnnDB, "SelectCategory")%></OPTION>
              <%
                Set rstCatList = SQLQuery(cnnDB, "SELECT * From categories WHERE category_id > 0 ORDER BY cname ASC")
                If Not rstCatList.EOF Then
                Do While Not rstCatList.EOF
              %>
              <OPTION VALUE="<% = rstCatList("category_id")%>">
              <% = rstCatList("cname") %></OPTION>

              <% 		rstCatList.MoveNext
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
              <SELECT NAME="status">
                <%
                  Set rstStatList = SQLQuery(cnnDB, "SELECT * From status WHERE status_id > 0 ORDER BY status_id ASC")
                  If Not rstStatList.EOF Then
                  Do While Not rstStatList.EOF
                  If rstStatList("status_id") = Cfg(cnnDB, "DefaultStatus") Then
                  %>
                  <OPTION VALUE="<% = rstStatList("status_id")%>" SELECTED>
                  <% = rstStatList("sname") %></OPTION>
                  <% Else %>
                  <OPTION VALUE="<% = rstStatList("status_id")%>">
                  <% = rstStatList("sname") %></OPTION>

                <% 	End If
                  rstStatList.MoveNext
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
              <SELECT NAME="priority">
              <%
                Set rstPriList = SQLQuery(cnnDB, "SELECT * From priority WHERE priority_id > 0 ORDER BY priority_id ASC")
                If Not rstPriList.EOF Then
                Do While Not rstPriList.EOF
                If rstPriList("priority_id") = Cfg(cnnDB, "DefaultPriority") Then
                %>
                <OPTION VALUE="<% = rstPriList("priority_id")%>" SELECTED>
                <% = rstPriList("pname") %></OPTION>
                <% Else %>
                <OPTION VALUE="<% = rstPriList("priority_id")%>">
                <% = rstPriList("pname") %></OPTION>

              <% 	End If
                rstPriList.MoveNext
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
              <SELECT NAME="rep">
              <%
                Set rstRepList = SQLQuery(cnnDB, "SELECT * From tblUsers WHERE IsRep = 1 AND RepAccess <> 2 AND sid > 0 ORDER BY uid ASC")
                If Not rstRepList.EOF Then
                Do While Not rstRepList.EOF
                If rstRepList("sid") = sid Then
                %>
                <OPTION VALUE="<% = rstRepList("sid")%>" SELECTED>
                <% = rstRepList("uid") %></OPTION>
                <% Else %>
                <OPTION VALUE="<% = rstRepList("sid")%>">
                <% = rstRepList("uid") %></OPTION>

              <% 	End If
                rstRepList.MoveNext
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
              <input type="text" size="4" name="time_spent" value="0">(<%=lang(cnnDB, "minutes")%>)
            </td>
          </tr>
        </table>
      </td>
    <tr class="Head2">
      <td colspan="2">
        <%=lang(cnnDB, "ProblemInformation")%>:
      </td>
    </tr>
    <tr class="Body1">
      <td colspan="2">
        <b><%=lang(cnnDB, "Title")%>:</b><em>*</em><br />
        <input type="text" name="title" size="50">
        <p>
        <b><%=lang(cnnDB, "Description")%>:</b><em>*</em><br />
        <textarea rows="8" cols="80" name="description"></textarea>
      </td>
    </tr>
    <tr class="Head2">
      <td colspan="2">
        <%=lang(cnnDB, "Solution")%>:
      </td>
    </tr>
    <tr class="Body1">
      <td colspan="2">
        <textarea rows="8" cols="80" name="solution"></textarea>
        <% If Cfg(cnnDB, "EnableKB") <> 0 Then %>
          <input type="checkbox" name="kb">&nbsp;<%=lang(cnnDB, "EnterinKnowledgeBase")%>
        <% End If %>
      </td>
    </tr>
    <tr class="Head2">
      <td colspan="2" align="center">
        <% If Cfg(cnnDB, "EmailType") <> 0 Then %>
          <input type="checkbox" name="noemail">&nbsp;<%=lang(cnnDB, "Dontsendemailtouser")%>
        <% End If %>
        <p>
        <input type="submit" value="<%=lang(cnnDB, "SubmitProblem")%>" name="B1">&nbsp;<input type="reset" value="<%=lang(cnnDB, "ClearForm")%>" name="B2">
      </td>
    </tr>
  </table>
</div>
</form>

<%
	' close record sets
	rstCatList.Close
	rstDepList.Close
	rstStatList.Close
	rstPriList.Close
	rstRepList.Close

	Call DisplayFooter(cnnDB, sid)
	cnnDB.Close

%>

</body>

</html>
