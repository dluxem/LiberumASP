<%@ LANGUAGE="VBScript" %>
<% 
  Option Explicit
  'Buffer the response, so Response.Expires can be used
  Response.Buffer = TRUE
  Server.ScriptTimeOut = 600  ' Wait 10 minute to time out script
%>

<?xml version="1.0"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

  <!--
  Liberum Help Desk, Copyright (C) 2000-2001 Doug Luxem
  Liberum Help Desk comes with ABSOLUTELY NO WARRANTY
  Please view the license.html file for the full GNU General Public License.

  Filename: viewlangstring.asp
  Date:     $Date: 2002/01/23 02:10:00 $
  Version:  $Revision: 1.51.2.1 $
  Purpose:  Manage the text strings for a selected language
  -->

  <!--  #include file = "../settings.asp" -->
  <!-- 	#include file = "../public.asp" -->

  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title>
      <%=lang(cnnDB, "HelpDesk")%>&nbsp;<%=lang(cnnDB, "Manage")%>&nbsp;<%=lang(cnnDB, "LanguageStrings")%>
    </title>
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>

    <%
      ' Check for perms to view this page
      Call CheckAdmin
      
      Dim intLangID, intCompLangID
      intLangID = request.querystring("lang_id")
      if len(intLangID) = 0 then
        Call DisplayError(3, lang(cnnDB, "NolanguageIDgiven"))
      end if
      intCompLangID = 1 ' English

      Dim blnSaveSuccess, blnAddSuccess
      blnSaveSuccess = False
      blnAddSuccess = False

      If Request.Form("frm_save") = "1" Then
        Dim intCounter, rstLangFields, rstVarIsNull, rstUpdateLang, strUpdLangName, strVarName
        Set rstLangFields = SQLQuery(cnnDB, "SELECT variable FROM tblLangStrings WHERE id=" & intCompLangID & " ORDER BY variable ASC")
        For intCounter = 2  To Request.Form.Count
          strUpdLangName = Request.Form(intCounter)
          strUpdLangName = Replace(strUpdLangName, "'", "''")
          strVarName = rstLangFields("variable")
          Set rstVarIsNull = SQLQuery(cnnDB, "SELECT id FROM tblLangStrings WHERE id=" & intLangID & " AND variable='" & strVarName & "'")
          If rstVarISNull.EOF Then
            Set rstUpdateLang = SQLQuery(cnnDB, "INSERT INTO tblLangStrings (id, variable, LangText) VALUES " & _
              "(" & intLangID & ", '" & strVarName & "', '" & strUpdLangName & "')")
          Else
            Set rstUpdateLang = SQLQuery(cnnDB, "UPDATE tblLangStrings SET " & _
              "id = " & intLangID & ", " & _
              "variable = '" & strVarName & "', " & _
              "LangText = '" & strUpdLangName & "' " & _
              "WHERE id=" & intLangID & " AND variable='" & strVarName & "'")
          End If
          rstLangFields.MoveNext
        Next
        rstLangFields.Close
        blnSaveSuccess = True

        ' Remove cached language strings
        Call ClearLangCache(cnnDB)
      End If

      If Request.Form("frm_add") = "1" Then
        Dim intLang1ID, intLang2ID, strNew1, strNew2, strNewVarName
        Dim rstCheckVarName, rstInsert1, rstInsert2
        strNewVarName = Request.Form("varname")
        Set rstCheckVarName = SQLQuery(cnnDB, "SELECT id FROM tblLangStrings WHERE variable='" & strNewVarName & "'")
        If Not rstCheckVarName.EOF Then
          rstCheckVarName.Close
          Call DisplayError(3, lang(cnnDB, "Variablenameisalreadyinuse"))
        End If
        intLang1ID = Cint(Request.Form("string1_id"))
        intLang2ID = Cint(Request.Form("string2_id"))
        strNew1 = Request.Form("string1_value")
        strNew2 = Request.Form("string2_value")

        strNewVarName = replace(strNewVarName, "'", "''")
        strNew1 = replace(strNew1, "'", "''")
        strNew2 = replace(strNew2, "'", "''")

        Set rstInsert1 = SQLQuery(cnnDB, "INSERT INTO tblLangStrings (id, variable, LangText) VALUES " & _
          "(" & intLang1ID & ", '" & strNewVarName & "', '" & strNew1 & "')")
        If intLang1ID <> intLang2ID Then
          Set rstInsert1 = SQLQuery(cnnDB, "INSERT INTO tblLangStrings (id, variable, LangText) VALUES " & _
            "(" & intLang2ID & ", '" & strNewVarName & "', '" & strNew2 & "')")
        End If

        blnAddSuccess = True

        ' Remove cached language strings
        Call ClearLangCache(cnnDB)
      End If
      
      Dim rstLangString, rstLangNameDefault, rstLangNameCurrent
      Set rstLangString = SQLQuery(cnnDB, "SELECT * FROM tblLangStrings WHERE id=" & intCompLangID & " ORDER BY variable ASC")
      Set rstLangNameDefault = SQLQuery(cnnDB, "SELECT * FROM tblLanguage WHERE id=" & intCompLangID)
      Set rstLangNameCurrent = SQLQuery(cnnDB, "SELECT * FROM tblLanguage WHERE id=" & intLangID)
      
      Function GetEditLangString(strEditVar)
        Dim rstEditString
        Set rstEditString = SQLQuery(cnnDB, "SELECT LangText FROM tblLangStrings WHERE id=" & intLangID & " AND variable='" & strEditVar & "'")
        If rstEditString.EOF Then
          GetEditLangString=""
        Else
          Dim strEditLang
          strEditLang = rstEditString("LangText")
          strEditLang = Replace(strEditLang, "<", "&lt;")
          strEditLang = Replace(strEditLang, ">", "&gt;")
          strEditLang = Replace(strEditLang, """", "&quot;")
          GetEditLangString = strEditLang
        End If
        rstEditString.Close
      End Function

    %>
    <div align="center">
      <table class="Wide">
        <tr class="Head1">
          <td colspan="3">
            <%=lang(cnnDB, "LanguageStrings")%>
          </td>
        </tr>
        <% If blnSaveSuccess Then %>
          <tr class="Head2">
            <td colspan="3">
              <div align="center">
                <%=lang(cnnDB, "ChangesSaved")%>
              </div>
            </td>
          </tr>
        <% End If %>
        <% If blnAddSuccess Then %>
          <tr class="Head2">
            <td colspan="3">
              <div align="center">
                <%=lang(cnnDB, "StringAdded")%>
              </div>
            </td>
          </tr>
        <% End If %>
        <tr class="Head2">
          <td>
            <div align="center">
              <%=lang(cnnDB, "Variable")%>
            </div>
          </td>
          <td>
            <div align="center">
              <% = rstLangNameDefault("LangName") %>(<% = rstLangNameDefault("Localized") %>)
            </div
          </td>
          <td>
            <div align="center">
              <% = rstLangNameCurrent("LangName") %>(<% = rstLangNameCurrent("Localized") %>)
            </div>
          </td>
        </tr>
        <form action="viewlangstring.asp?lang_id=<% = intLangID %>" method="POST" name="change">
          <input type="hidden" name="frm_save" value="1">
          <%
            Do While Not rstLangString.EOF
          %>
            <tr class="Body1">
              <td align="center">
                <% = rstLangString("variable") %>
              </td>
              <td align="center">
                 <% = rstLangString("LangText") %>
              </td>
              <td align="center">
                <input type="text" size="30" name="<% = rstLangString("variable") %>" value="<% = GetEditLangString(rstLangString("variable")) %>">
              </td>
            </tr>
          <%
            rstLangString.MoveNext
            Loop
          %>
          <tr class="Head2">
            <td colspan="3">
              <div align="center">
                <input type="submit" value="<%=lang(cnnDB, "Save")%>">
              </div>
            </td>
          </tr>
        </form>
      </table>
      <p>
      <table class="Wide">
        <tr class="Head1">
          <td colspan="3">
            <%=lang(cnnDB, "AddLanguageString")%>
          </td>
        </tr>
        <tr class="Head2">
          <td>
            <div align="center">
              <%=lang(cnnDB, "Variable")%>
            </div>
          </td>
          <td>
            <div align="center">
              <% = rstLangNameDefault("LangName") %>(<% = rstLangNameDefault("Localized") %>)
            </div
          </td>
          <td>
            <div align="center">
              <% = rstLangNameCurrent("LangName") %>(<% = rstLangNameCurrent("Localized") %>)
            </div>
          </td>
        </tr>
        <form action="viewlangstring.asp?lang_id=<% = intLangID %>" method="POST" name="add">
          <input type="hidden" name="frm_add" value="1">
          <tr class="Body1">
            <td align="center">
              <input type="text" size="20" name="varname">
            </td>
            <td align="center">
               <input type="hidden" name="string1_id" value="<% = rstLangNameDefault("id") %>">
               <input type="text" size="30" name="string1_value">
            </td>
            <td align="center">
                               <input type="hidden" name="string2_id" value="<% = rstLangNameCurrent("id") %>">
               <input type="text" size="30" name="string2_value">
            </td>
          </tr>
          <tr class="Head2">
            <td colspan="3">
              <div align="center">
                <input type="submit" value="<%=lang(cnnDB, "AddNew")%>">
              </div>
            </td>
          </tr>
        </form>
      </table>
      <p>
      <a href="viewlang.asp"><%=lang(cnnDB, "Manage")%>&nbsp;<%=lang(cnnDB, "Languages")%></a><br />
      <a href="default.asp"><%=lang(cnnDB, "AdministrativeMenu")%></a>
      </p>
    </div>

    <%
      rstLangString.Close
      rstLangNameCurrent.Close
      rstLangNameDefault.Close
      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
