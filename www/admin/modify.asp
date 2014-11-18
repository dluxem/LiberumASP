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

  Filename: modify.asp
  Date:     $Date: 2001/12/09 02:01:24 $
  Version:  $Revision: 1.50 $
  Purpose:  This page is used to add/modify depts, categories, etc.
  -->
  
  <!-- 	#include file = "../public.asp" -->

  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title><%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "CreateModify")%></title>
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>

    <%
      ' Check for perms to view this page
      Call CheckAdmin

      ' Get the type of add/modify
      ' 1 - Rep
      ' 2 - Category
      ' 3 - Department
      ' 4 - Priority
      ' 5 - Status
      ' 6 - Language
      Dim mType, strLanguageName, intLangID
      intLangID = Request.QueryString("mLangID")
      strLanguageName = Request.QueryString("mLanguage")
      mType = Cint(Request.QueryString("Mtype"))
      If (mType < 1) or (mType > 6) Then
        Call DisplayError(3, "Invalid type to create/modify")
      End If

      Dim data_id, data1, data2, data3, data4, numDataFields
      Dim data1Name, data2Name, data3Name, data4Name, title

      If Not Len(Request.QueryString("id")) = 0 Then
        ' An ID is sent, so this is a modify of that ID
        data_id = Cint(Request.QueryString("id"))
        Select Case mType
          Case 1	' Modify Rep
            Dim repRes
            Set repRes = SQLQuery(cnnDB, "SELECT ruid, rname, remail, rdescription FROM reps WHERE rep_id=" & data_id)
            data1 = repRes("ruid")
            data2 = repRes("rname")
            data3 = repRes("remail")
            data4 = repRes("rdescription")
            repRes.Close
          Case 2	' Modify Category
            Dim catRes
            Set catRes = SQLQuery(cnnDB, "SELECT cname, rep_id FROM categories WHERE category_id=" & data_id)
            data1 = catRes("cname")
            data2 = catRes("rep_id")
            catRes.Close
          Case 3	' Modify Department
            Dim depRes
            Set depRes = SQLQuery(cnnDB, "SELECT dname FROM departments WHERE department_id=" & data_id)
            data1 = depRes("dname")
            depRes.Close
          Case 4	' Modify Priority
            Dim priRes
            Set priRes = SQLQuery(cnnDB, "SELECT pname FROM priority WHERE priority_id=" & data_id)
            data1 = data_id
            data2 = priRes("pname")
            priRes.Close
          Case 5	' Modify Status
            Dim statRes
            Set statRes = SQLQuery(cnnDB, "SELECT sname FROM status WHERE status_id=" & data_id)
            data1 = data_id
            data2 = statRes("sname")
            statRes.Close
          Case 6  ' Modify Language
            Dim rstLanguage
            set rstLanguage = SQLQuery(cnnDB, "SELECT * FROM tblLanguage WHERE id=" & data_id)
            data1 = rstLanguage("LangName")
            data2 = rstLanguage("Localized")
            rstLanguage.Close
        End Select
      Else

        ' No ID is sent, so this is an add
        data_id = 0
        
        Select Case mType
          Case 1	' Modify Rep
            data3Name = Cfg(cnnDB, "BaseEmail")
        End Select
      End If

      'Fill in the names of the data fields
      Select Case mType
        Case 1	' Modify Rep
          title = lang(cnnDB, "SupportRepresentatives")
          data1Name = lang(cnnDB, "Username")
          data2Name = lang(cnnDB, "FullName")
          data3Name = lang(cnnDB, "Email")
          data4Name = lang(cnnDB, "Description")
          numDataFields = 4
        Case 2	' Modify Category
          title = lang(cnnDB, "Category")
          data1Name = lang(cnnDB, "CategoryName")
          data2Name = lang(cnnDB, "PrimaryRep")
          numDataFields = 2
        Case 3  ' Modify Departments
          title = lang(cnnDB, "Department")
          data1Name = lang(cnnDB, "Department")
          numDataFields = 1
        Case 4	' Modify Priority
          title = lang(cnnDB, "Priority")
          data1Name = lang(cnnDB, "PriorityNumber")
          data2Name = lang(cnnDB, "PriorityName")
          numDataFields = 2
        Case 5	' Modify Status
          title = lang(cnnDB, "Status")
          data1Name = lang(cnnDB, "StatusNumber")
          data2Name = lang(cnnDB, "StatusName")
          numDataFields = 2
        Case 6  ' Modify language
          title = lang(cnnDB, "Language")
          data1Name = lang(cnnDB, "LanguageName")
          data2Name = lang(cnnDB, "LocalizedName")
          numDataFields = 2
      End Select


    %>
    <form method="post" action="postmods.asp" id=form1 name=form1>
      <div align="center">
        <table class="Normal">
          <tr class="Head1">
            <td>
              <%=lang(cnnDB, "CreateModify")%>&nbsp;<% = title %>
            </td>
          </tr>
          <tr class="Body1">
            <td>
              <input type="hidden" name="data_id" value="<% = data_id %>">
              <input type="hidden" name="numdatafields" value="<% = numDataFields %>">
              <input type="hidden" name="mtype" value="<% = mType %>">
              <input type="hidden" name="mLanguage" value="<% = strLanguageName %>">
              <input type="hidden" name="mLangID" value="<% = intLangID %>">

              <% If numDataFields >= 1 Then 
                Response.Write data1Name & ": " & _
                  "<input type=""text"" size=""25"" name=""data1"" value=""" & data1 & """>" & _
                  "<p>"
                End If
                If numDataFields >= 2 Then
                  Response.Write data2Name & ": "
                  ' This is a special case for categories, which require a pulldown menu
                  If mType <> 2 Then 
                    Response.Write "<input type=""text"" size=""25"" name=""data2"" value=""" & data2 & """>"
                  Else
                    Response.Write "<SELECT NAME=""data2"">"
                    Dim replRes
                    Set replRes = SQLQuery(cnnDB, "SELECT * FROM tblUsers WHERE IsRep=1 AND RepAccess <> 2 ORDER BY uid ASC")
                    If Not replRes.EOF Then
                      Do While Not replRes.EOF
                        If Cint(replRes("sid")) = data2 Then
                          Response.Write "<option value=""" & replRes("sid") & """ SELECTED>" & replRes("uid") & "</OPTION>"
                        Else
                          Response.Write "<OPTION VALUE=""" & replRes("sid") & """>" & replRes("uid") & "</OPTION>"
                        End If
                        replRes.MoveNext
                      Loop
                    End If
                    replRes.Close
                    Response.Write "</SELECT>"
                  End If
                  Response.Write "<p>"
                End If
                If numDataFields >= 3 Then
                  Response.Write data3Name & ": " & _
                    "<input type=""text"" size=""25"" name=""data3"" value=""" & data3 & """><p>"
                End If
                If numDataFields >= 4 Then
                  Response.Write data4Name & ": " & _
                    "<input type=""text"" size=""25"" name=""data4"" value=""" & data4 & """><p>"
                End If
                Response.Write "<div align=""center"">" & _
                  "<input type=""submit"" value=""" & lang(cnnDB, "CreateModify") & """><br />"
                If mType < 6 Then
                  If data_id > 0 Then 
                    Response.Write "<font size=""-1""><i>" & _
                      "(" & lang(cnnDB, "StatusandPrioritynumberscannotbechanged") & ")</i></font>"
                  End If
                End If %>
              </div>
            </td>
          </tr>
        </table>
        <p>
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
        <br />
        <a href="default.asp"><%=lang(cnnDB, "AdministrativeMenu")%></a>
      </div>
    </form>
    <%
      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
