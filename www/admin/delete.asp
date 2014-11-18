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

  Filename: delete.asp.asp
  Date:     $Date: 2001/12/09 02:01:24 $
  Version:  $Revision: 1.50 $
  Purpose:  Removes the selected item from the database.
  -->
  
  <!-- 	#include file = "../public.asp" -->

  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title>
      <%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "ItemDeleted")%>
    </title>
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>

    <%
      ' Check for perms to view this page
      Call CheckAdmin

      ' get field/type and id number
      Dim mType, id, delRes, probRes, strLanguageName, intLangID
      
      intLangID = Request.QueryString("mLangID")
      strLanguageName = Request.QueryString("mLanguage")
      mType = Cint(Request.QueryString("mtype"))
      id = Cint(Request.QueryString("id"))


      ' look for missing data
      If (mType < 1) OR (mType > 6) Then
        Call DisplayError(3, lang(cnnDB, "InvalidIDordatatype"))
      End If

      ' Generate SQL strings to delete the item and update
      ' the problems that reference it.  Problems will use
      ' a 0 where the item is referenced.  The database
      ' should have a dummy entry with id of 0.
      Dim delStr, probStr, delStr2
      Select Case mType
        Case 1	' Rep
          delStr = "UPDATE tblUsers SET IsRep = 0 WHERE sid =" & id
          probStr = "UPDATE problems SET rep=0 WHERE rep=" & id
        Case 2	' category
          delStr = "DELETE FROM categories WHERE category_id=" & id
          probStr = "UPDATE problems SET category=0 WHERE category=" & id
        Case 3	' department
          delStr = "DELETE FROM departments WHERE department_id=" & id
          probStr = "UPDATE problems SET department=0 WHERE department=" & id
        Case 4	' priority
          delStr = "DELETE FROM priority WHERE priority_id=" & id
          probStr = "UPDATE problems SET priority=0 WHERE priority=" & id
        Case 5	' status
          delStr = "DELETE FROM status WHERE status_id=" & id
          probStr = "UPDATE problems SET status=0 WHERE status=" & id
        Case 6  ' Language
          delStr = "DELETE FROM tblLanguage WHERE id=" & id
          delStr2 = "DELETE FROM tblLangStrings WHERE id=" & id
      End Select

      ' Delete the item
      Set delRes = SQLQuery(cnnDB, delStr)
      If mType = 6 Then
        Set delRes = SQLQuery(cnnDB, delStr2)
      End If

      ' Update closed problems with Unknown (0) value
      Select Case mType
        case 1
          Set probRes = SQLQuery(cnnDB, probStr)
        case 2
          Set probRes = SQLQuery(cnnDB, probStr)
        case 3
          Set probRes = SQLQuery(cnnDB, probStr)
        case 4
          Set probRes = SQLQuery(cnnDB, probStr)
        case 5
          Set probRes = SQLQuery(cnnDB, probStr)
        case 6
        ' Does not affect problems
      End Select

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
              <%=lang(cnnDB, "Theitemhasbeendeletedfromthedatabase")%>
            </div>
          </td>
        </tr>
      </table>
      <p><a href="default.asp"><%=lang(cnnDB, "AdministrativeMenu")%></a></p>
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
