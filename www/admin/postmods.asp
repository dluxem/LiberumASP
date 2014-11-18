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

  Filename: postmods.asp
  Date:     $Date: 2001/12/09 02:01:24 $
  Version:  $Revision: 1.50 $
  Purpose:  Adds or modifies an item.  Takes data from the form
  on MODIFY.ASP, checks to make sure valid before
  entering in the database.
  -->
  
  <!-- 	#include file = "../public.asp" -->

  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title>
      <%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "ModificationDone")%>
    </title>
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>

    <%
      ' Check for perms to view this page
      Call CheckAdmin

      ' Get control values
      Dim mType, strLanguageName, numDataFields, data_id, intLangID
      numDataFields = Cint(Request.Form("numdatafields"))
      intLangID = Request.Form("mLangID")
      strLanguageName = Request.Form("mLanguage")
      mType = Cint(Request.Form("mtype"))
      data_id = Cint(Request.Form("data_id"))


      Dim data1, data2, data3, data4
      data1 = Request.Form("data1")
      data2 = Request.Form("data2")
      data3 = Request.Form("data3")
      data4 = Request.Form("data4")

      'Check required fields
      If (numDataFields >= 1) AND (Len(data1) = 0) Then
        Call DisplayError(1, lang(cnnDB, "Field") & " 1")
      End If
      If (numDataFields >= 2) AND (Len(data2) = 0) Then
        Call DisplayError(1, lang(cnnDB, "Field") & " 2")
      End If
      If (numDataFields >= 3) AND (Len(data3) = 0) Then
        Call DisplayError(1, lang(cnnDB, "Field") & " 3")
      End If
      If (numDataFields >= 4) AND (Len(data4) = 0) Then
        Call DisplayError(1, lang(cnnDB, "Field") & " 4")
      End If

      'Convert data fields to integers where needed
      Select Case mType
        Case 1	' rep
            data1 = Lcase(data1)
            data4 = Replace(data4, "'", "`")
        Case 2	' category
            data2 = Cint(data2)
        Case 3	' department
        Case 4	' priority
            data1 = Cint(data1)
        Case 5	' status
            data1 = Cint(data1)
        Case 6  ' Language
      End Select

      Dim sqlString, rstSQL, sqlString2
      ' If data_id = 0 then creating a new entry in the
      ' database, otherwise we are modifying an old one

      if data_id = 0 Then
        Select Case mType
          Case 1	' rep
            Dim repRes
            Set repRes = SQLQuery(cnnDB, "SELECT rep_id FROM reps WHERE ruid='" & data1 & "'")
            If (Not repRes.EOF) Then
              repRes.Close
              Call DisplayError(3, lang(cnnDB, "This username already exists")& ".")
            End If
            repRes.Close
            data_id = GetUnique(cnnDB, "reps")
            sqlString = "INSERT INTO reps (rep_id, ruid, rname, remail, rdescription) " & _
              "VALUES (" & data_id & ",'" & data1 & "','" & data2 & "','" & data3 & "','" & data4 & "')"

          Case 2	' category
            data_id = GetUnique(cnnDB, "categories")
            sqlString = "INSERT INTO categories (category_id, rep_id, cname) " & _
              "VALUES (" & data_id & "," & data2 & ",'" & data1 & "')"

          Case 3	' department
            data_id = GetUnique(cnnDB, "departments")
            sqlString = "INSERT INTO departments (department_id, dname) " & _
              "VALUES (" & data_id & ",'" & data1 & "')"

          Case 4	'  priorities
            If data1 < 0 Then
              Call DisplayError(3, "Enter a positive priority number.")
            End If
            Dim priRes
            Set priRes = SQLQuery(cnnDB, "SELECT priority_id FROM priority WHERE priority_id=" & data1)
            If Not priRes.EOF Then
              priRes.Close
              Call DisplayError(3, "Enter a unique priority number.")
            End If
            priRes.Close
            sqlString = "INSERT INTO priority (priority_id, pname) " & _
              "VALUES (" & data1 & ",'" & data2 & "')"

          Case 5	' status
            If data1 < 0 Then
              Call DisplayError(3, lang(cnnDB, "Enterapositivestatusnumber") & ".")
            End If
            Dim statusRes
            Set statusRes = SQLQuery(cnnDB, "SELECT status_id FROM status WHERE status_id=" & data1)
            If Not statusRes.EOF Then
              statusRes.Close
              Call DisplayError(3, lang(cnnDB, "Enterauniquestatusnumber") & ".")
            End If
            statusRes.Close
            sqlString = "INSERT INTO status (status_id, sname) " & _
              "VALUES (" & data1 & ",'" & data2 & "')"
              
          Case 6  ' Language
            data_id = GetUnique(cnnDB, "lang")
            Dim rstLanguage
            Set rstLanguage = SQLQuery(cnnDB, "SELECT id FROM tblLanguage WHERE langname='" & data1 & "' AND localized='" & data2 & "'")
            If Not rstLanguage.EOF Then
              rstLanguage.Close
              Call DisplayError(3, lang(cnnDB, "Enterauniquelanguageandlocalizedname") & ".")
            End If
            rstLanguage.Close
            sqlString = "INSERT INTO tblLanguage (id, langname, localized) " & _
              "VALUES (" & data_id & ",'" & data1 & "','" & data2 & "')"
              
        End Select
      Else
          Select Case mType
            Case 1	' rep
              Set repRes = SQLQuery(cnnDB, "SELECT rep_id FROM reps WHERE ruid='" & data1 & "'")
              If (Not repRes.EOF) Then
                If Cint(repRes("rep_id")) <> data_id Then
                  repRes.Close
                  Call DisplayError(3,lang(cnnDB, "Thisusernamealreadyexists") & ".")
                End If
              End If
              repRes.Close
              sqlString = "UPDATE reps SET " & _
                "ruid='" & data1 & "', " & _
                "rname='" & data2 & "', " & _
                "remail='" & data3 & "', " & _
                "rdescription='" & data4 & "' " & _
                "WHERE rep_id=" & data_id
  
            Case 2	' category
              sqlString = "UPDATE categories SET " & _
                "cname='" & data1 & "', " & _
                "rep_id=" & data2 & " " & _
                "WHERE category_id=" & data_id
  
            Case 3	' department
              sqlString = "UPDATE departments SET dname='" & data1 & "' " & _
                "WHERE department_id=" & data_id
  
            Case 4	' priority
              sqlString = "UPDATE priority SET pname='" & data2 & "' " & _
                "WHERE priority_id=" & data_id
  
            Case 5	' status
              sqlString = "UPDATE status SET sname='" & data2 & "' " & _
                "WHERE status_id=" & data_id
               
            Case 6	' Language
              sqlString = "UPDATE tblLanguage SET " & _
                "langname='" & data1 & "', " & _
                "localized='" & data2 & "' " & _
                "WHERE id=" & data_id
          End Select
        End If
      

      ' All data is present
      ' Write into database
      Set rstSQL = SQLQuery(cnnDB, sqlString)

      If mType = 6 and len(sqlString2) > 0 then
        Set rstSQL = SQLQuery(cnnDB, sqlString2)
      End If

    %>

    <div align="center">
      <table class="Normal">
        <tr class="Head1">
          <td>
            <%=lang(cnnDB, "OperationComplete")%>
          </td>
        </tr>
        <tr class="Body1">
          <td>
            <div align="center">
            <%=lang(cnnDB, "Thedatabasehasbeenupdated")%>.
            </div>
          </td>
        </tr>
      </table>
      <p><a href="default.asp">Administrative Menu</a></p>
      <%
        ' Get the correct URL to return to
        Select Case mType
          Case 1 %>
            <a href="viewrep.asp"><%=lang(cnnDB, "Manage")%>&nbsp;<%=lang(cnnDB, "SupportReps")%></a>
      <%		Case 2 %>
            <a href="viewcat.asp"><%=lang(cnnDB, "Manage")%>&nbsp;<%=lang(cnnDB, "Categories")%></a>
      <%		Case 3 %>
            <a href="viewdep.asp"><%=lang(cnnDB, "Manage")%>&nbsp;<%=lang(cnnDB, "Departments")%></a>
      <%		Case 4 %>
            <a href="viewpri.asp"><%=lang(cnnDB, "Manage")%>&nbsp;<%=lang(cnnDB, "Priorities")%></a>
      <%		Case 5 %>
            <a href="viewstatus.asp"><%=lang(cnnDB, "Manage")%>&nbsp;<%=lang(cnnDB, "Statuses")%></a>
      <%		Case 6 %>
            <a href="viewlang.asp"><%=lang(cnnDB, "Manage")%>&nbsp;<%=lang(cnnDB, "Languages")%></a>
      <%	End Select %>
    </div>
    
    <%
      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
