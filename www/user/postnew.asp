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

  Filename: postnew.asp
  Date:     $Date: 2002/06/15 23:49:20 $
  Version:  $Revision: 1.51.4.1 $
  Purpose:  Takes the input from new.asp, checks for errors and enters the
  problem into the database.
  -->

  <!-- 	#include file = "../public.asp" -->
  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title><%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "ProblemSubmitted")%></title>
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>

    <%
      ' See if authenticated
      Call CheckUser(cnnDB, sid)

      ' Get the information from the form fields
      Dim uid, uemail, uphone, ulocation, category, department, title, description, entered_by
      Dim kb
      
      uid = Request.Form("uid")
      uemail = Request.Form("uemail")
      uphone = Request.Form("uphone")
      ulocation = Request.Form("ulocation")
      category = Cint(Request.Form("category"))
      department = Request.Form("department")
      title = Request.Form("title")
      description = Request.Form("description")
      entered_by = sid
      kb = 0

    ' Check for required fields (uemail, category, department, title, description)

      if InStr(uemail, "@")=0 Then
        cnnDB.Close
        Call DisplayError(1, lang(cnnDB, "Emailaddress"))
      End if

      if category = 0 Then
        cnnDB.Close
        Call DisplayError(1, lang(cnnDB, "Category"))
      End if

      if department = 0 Then
        cnnDB.Close
        Call DisplayError(1, lang(cnnDB, "Department"))
      End if

      if Len(title)=0 Then
        cnnDB.Close
        Call DisplayError(1, lang(cnnDB, "Title"))
      Elseif Len(title) > 50 Then
        title = Trim(title)
        title = Left(title, 50)
      End if

      if Len(description)=0 Then
        cnnDB.Close
        Call DisplayError(1, lang(cnnDB, "Description"))
      End if

    ' Get missing variables to enter problem
      Dim id, priority, status, start_date, rep, time_spent
      priority = Cfg(cnnDB, "DefaultPriority")
      status = Cfg(cnnDB, "DefaultStatus")
      time_spent = 0
      start_date = SQLDate(Now, lhdAddSQLDelim)

      ' Get the department name by querying on department_id
      Dim dname, rstDept
      Set rstDept = SQLQuery(cnnDB, "SELECT dname FROM departments WHERE department_id=" & Request.Form("department"))
      dname = rstDept("dname")

      ' Get the category name by querying on category_id
      Dim cname, catRes
      Set catRes = SQLQuery(cnnDB, "SELECT cname, rep_id FROM categories WHERE category_id=" & Request.Form("category"))
      rep = catRes("rep_id")
      cname = catRes("cname")

    ' Get the problem ID number then immediately update it
      id = GetUnique(cnnDB, "problems")

    ' Clean up variables
      uemail = Left(Trim(uemail),50)
      uemail = Replace(uemail,"'","''")
      uphone = Left(Trim(uphone),50)
      uphone = Replace(uphone,"'","''")
      ulocation = Left(Trim(ulocation),50)
      ulocation = Replace(ulocation,"'","''")
      title = Replace(title,"'","''")
      description = Replace(description,"'","''")

    ' All data is present
    ' Write problem into database
      Dim strProblemQry, rstProbInsert
      strProblemQry = "INSERT INTO problems (id, uid, uemail, uphone, ulocation, entered_by, " & _
        "category, department, title, description, priority, status, start_date, rep, time_spent, kb) " & _
        "VALUES (" & id & ",'" & uid & "','" & uemail & "','" & uphone & "','" & _
        ulocation & "'," & entered_by & "," & category & "," & department & ",'" & title & "','" & _
        description & "'," & priority & "," & status & "," & start_date & "," & rep & "," & time_spent & _
        "," & kb & ")"

      Set rstProbInsert = SQLQuery(cnnDB, strProblemQry)

    ' Send mail to the user and support rep if email is enabled.
    Call eMessage(cnnDB, "usernew", id, uemail)

    Call eMessage(cnnDB, "repnew", id, Usr(cnnDB, rep, "email1"))

    'Page Rep if enabled
    If (priority >= Cfg(cnnDB, "EnablePager")) And (Len(Usr(cnnDB, rep, "email2")) > 0) Then
      Call eMessage(cnnDB, "reppager", id, Usr(cnnDB, rep, "email2"))
    End If

    ' Convert the strings back to display them
      title = Replace(title,"''","'")
      description = Replace(description,"''","'")
    %>

    <div align="center">
      <table class="Wide">
        <tr class="Head1">
          <td>
            <%=lang(cnnDB, "Problem")%>&nbsp;<% = id %>&nbsp;<%=lang(cnnDB, "Submitted")%>
          </td>
        </tr>
        <tr class="Body1">
          <td>
            <table class="Wide">
              <tr>
                <td width="125">
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
                  <% = uemail %>
                </td>
              </tr>
              <tr>
                <td>
                  <b><%=lang(cnnDB, "Phone")%>:</b>
                </td>
                <td>
                  <% = uphone %>
                </td>
              </tr>
              <tr>
                <td>
                  <b><%=lang(cnnDB, "Location")%>:</b>
                </td>
                <td>
                  <% = ulocation %>
                </td>
              </tr>
              <tr>
                <td>
                  <b><%=lang(cnnDB, "StartDate")%>:</b>
                </td>
                <td>
                  <% = start_date %>
                </td>
              </tr>
              <tr>
                <td>
                  <b><%=lang(cnnDB, "Department")%>:</b>
                </td>
                <td>
                  <% = rstDept("dname") %>
                </td>
              </tr>
              <tr>
                <td>
                  <b><%=lang(cnnDB, "Category")%>:</b>
                </td>
                <td>
                  <% = catRes("cname") %>
                </td>
              </tr>
              <tr>
                <td>
                  <b><%=lang(cnnDB, "AssignedTo")%>:</b>
                </td>
                <td>
                  <a href="mailto:<% = Usr(cnnDB, rep, "email1") %>?Subject=<%=lang(cnnDB, "HELPDESK")%>: <%=lang(cnnDB, "Problem")%> <% = id %>"><% = Usr(cnnDB, rep, "fname") %></a>
                </td>
              </tr>
              <tr>
                <td>
                  <b><%=lang(cnnDB, "Title")%>:</b>
                </td>
                <td>
                  <% = title %>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr class="Head2">
          <td>
            <%=lang(cnnDB, "Description")%>:
          </td>
        </tr>
        <tr class="Body1">
          <td>
            <center><form><textarea name="display_desc" rows="10" cols="80"><% = description %></textarea></form></center>
          </td>
        </tr>
      </table>
    </div>

    <%
      ' Close records
      rstDept.Close
      catRes.Close

      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>

  </body>
</html>
