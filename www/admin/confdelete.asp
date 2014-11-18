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

  Filename: confdelete.asp
  Date:     $Date: 2001/12/09 02:01:24 $
  Version:  $Revision: 1.50 $
  Purpose:  Confirms the deletetion of a rep, category ,etc.  It looks for
  any open problem that use the item and gives a warning if
  any exist. If not, then displays a yes/no question to
  confirm the deletion.
  -->
  
  <!-- 	#include file = "../public.asp" -->

  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title>
      <%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "ConfirmDeletion")%>
    </title>
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>

    <%
      ' Check for perms to view this page
      Call CheckAdmin

      ' get field/type and id number
      Dim mType, id, strLanguageName, intLangID
      
      intLangID = Request.QueryString("mLangID")
      strLanguageName = Request.QueryString("mLanguage")
      mType = Cint(Request.QueryString("mtype"))
      if mType = 7 then
        id = Request.QueryString("id")
      else
        id = Cint(Request.QueryString("id"))
      end if

      'Do not allow admins to delete english language
      If mType = 6 and id = 1 Then
        Call DisplayError(3, lang(cnnDB, "Sorryyoucannotdeletethislanguage"))
      End If
      
      'Only allow admins to delete language variables from default language
      If mType = 7 and id > 0 Then
        Call DisplayError(3, lang(cnnDB, "Youmustdeletevariablesfromenglishlanguage"))
      End If
      
      ' look for missing data
      If (mType < 1) OR (mType > 7)Then
        Call DisplayError(3, lang(cnnDB, "InvalidIDordatatype"))
      End If
      If mType < 7 Then
        If id = 0 Then
        Call DisplayError(3, lang(cnnDB, "InvalidIDordatatype"))
        End If
      End If

      Dim okToDelete, catExists, catRes, probRes
      okToDelete = TRUE
      catExists = FALSE

      ' look for OPEN problems using this
      ' data.  You cannot delete the rep, department, etc
      ' if an open problem is using it.
      Select Case mType
        Case 1	' rep - look for problems and categories
          Set probRes = SQLQuery(cnnDB, "SELECT id FROM problems WHERE rep=" & id & " AND status<>" & Cfg(cnnDB, "CloseStatus"))
          Set catRes = SQLQuery(cnnDB, "SELECT category_id, cname FROM categories WHERE rep_id=" & id)
          If Not catRes.EOF Then
            catExists = TRUE
            okToDelete = FALSE
          End If

        Case 2	' category
          Set probRes = SQLQuery(cnnDB, "SELECT id FROM problems WHERE category=" & id & " AND status<>" & Cfg(cnnDB, "CloseStatus"))

        Case 3	' department
          Set probRes = SQLQuery(cnnDB, "SELECT id FROM problems WHERE department=" & id & " AND status<>" & Cfg(cnnDB, "CloseStatus"))

        Case 4	' priority
          Set probRes = SQLQuery(cnnDB, "SELECT id FROM problems WHERE priority=" & id & " AND status<>" & Cfg(cnnDB, "CloseStatus"))

        Case 5	' status
          Set probRes = SQLQuery(cnnDB, "SELECT id FROM problems WHERE status=" & id & " AND status<>" & Cfg(cnnDB, "CloseStatus"))
          If id = Cfg(cnnDB, "CloseStatus") Then
            probRes.Close
            Call DisplayError(3, lang(cnnDB, "TheCLOSEDstatuscannotbedeleted"))
          End If
          
        Case 6  ' Language
        ' Does not effect problems
        
        Case 7  ' Language strings
        ' Does not effect problems

      End Select

      ' Problems exists, so can't delete.
      if mType < 6 then
        If Not probRes.EOF Then
          okToDelete = FALSE
        End If
      End If

    %>
    <div align="center">
      <table class="Normal">
        <% If okToDelete Then %>
          <tr class="Head1">
            <td>
              <%=lang(cnnDB, "Areyousure")%>
            </td>
          </tr>
          <tr class="Body1">
            <td>
              <div align="center">
                <a href="<% = Request.ServerVariables("HTTP_REFERER") %>"><%=lang(cnnDB, "NO")%></a>
                <p><a href="delete.asp?mtype=<% = mtype %>&id=<% = id %>&mLangID=<% = intLangID %>&mLanguage=<% = strLanguageName %>"><%=lang(cnnDB, "YES")%></a></p>
              </div>
            </td>
          </tr>
        <% Else  %>
          <tr class="Head1">
            <td>
              <%=lang(cnnDB, "UnableToDelete")%>
            </td>
          </tr>
          <%
          If catExists Then
          %>
            <tr class="Body1">
              <td>
                <%=lang(cnnDB, "catExistsText")%>
                <p>
                <div align="center">
                  <% Do While Not catRes.EOF
                      Response.Write("<a href=""modify.asp?mtype=2&id=" & catRes("category_id") & """>")
                      Response.Write(catRes("cname") & "</a><BR>")
                    catRes.MoveNext
                    Loop
                  %>
                </div>
              </td>
            </tr>
          <%
          End If
          If NOT probRes.EOF Then
          %>
            <tr class="Body1">
              <td>
                <%=lang(cnnDB, "probResText")%>
                <p>
                <div align="center">
                  <% Do While Not probRes.EOF
                      Response.Write("<a href=""../rep/view.asp?id=" & probRes("id") & """>")
                      Response.Write(probRes("id") & "</a><BR>")
                    probRes.MoveNext
                    Loop
                  %>
                </div>
              </td>
            </tr>
        <% 	
          End If
        End If
        %>
      </table>
      <p><a href="default.asp"><%=lang(cnnDB, "AdministrativeMenu")%></a></p>
      <%
        ' Get the correct URL to return to
        Select Case mType
          Case 1 %>
            <a href="viewrep.asp"><%=lang(cnnDB, "Manage")%> <%=lang(cnnDB, "Support Reps")%></a>
      <%		Case 2 %>
            <a href="viewcat.asp"><%=lang(cnnDB, "Manage")%> <%=lang(cnnDB, "Categories")%></a>
      <%		Case 3 %>
            <a href="viewdep.asp"><%=lang(cnnDB, "Manage")%> <%=lang(cnnDB, "Departments")%></a>
      <%		Case 4 %>
            <a href="viewpri.asp"><%=lang(cnnDB, "Manage")%> <%=lang(cnnDB, "Priorities")%></a>
      <%		Case 5 %>
            <a href="viewstatus.asp"><%=lang(cnnDB, "Manage")%> <%=lang(cnnDB, "Statuses")%></a>
      <%		Case 6 %>
            <a href="viewlang.asp"><%=lang(cnnDB, "Manage")%> <%=lang(cnnDB, "Languages")%></a>
      <%		Case 7 %>
              <a href="viewlangstring.asp?id=<% = intLangID %>"><%=lang(cnnDB, "Manage")%> <%=lang(cnnDB, "LanguageStrings")%></a>
      <%	End Select %>
    </div>

    <%
      If mType = 1 Then
        catRes.Close
      End If
      If mType < 6 then
        probRes.Close
      End If

      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
