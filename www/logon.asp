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

  Filename: logon.asp
  Date:     $Date: 2002/06/15 23:49:20 $
  Version:  $Revision: 1.52.4.1 $
  Purpose:  Get username and password from user.  Will redirect them
  back to where they entered.
  -->
  <!--	#include file = "settings.asp" -->
  <!-- 	#include file = "public.asp" -->

  <%
  Call SetAppVariables

  Dim cnnDB, sid
  Set cnnDB = CreateCon
  sid = GetSid

  %>

  <head>
    <title>
      <% = Cfg(cnnDB, "SiteName") %>&nbsp;<%=lang(cnnDB, "HelpDesk")%>
    </title>
    <link rel="stylesheet" type="text/css" href="default.css">
  </head>
  <body>

    <%
      ' **************************************
      ' Get logon info.
      ' **************************************

      ' By default, the user will not have admin access.
      ' User must enter password on admin page.
      Dim username, url, invalid

      url = Trim(Request.QueryString("URL"))
      If Len(url) = 0 Then
        url = "default.asp"
      End If

      Select Case Cfg(cnnDB, "AuthType")

        Case 1	' NT Authentication
          Dim domainLen
          username = Request.ServerVariables("AUTH_USER")

          If Len(username) = 0 Then
            Call DisplayError(3, lang(cnnDB, "UnabletoobtainusernamewithNTauthentication"))
          End If

          domainLen = InStr(username, "\")
          username = Mid(username, (domainLen+1), (Len(username)-domainLen))
          username = Lcase(username)

          If CBool(InStr(username, "'")) Then
            Call DisplayError (3, lang(cnnDB, "Username") & "&nbsp;" & Lang(cnnDB, "containsinvalidcharacters") & ".")
          End If

          Dim ntUserRes
          Set ntUserRes = SQLQuery(cnnDB, "SELECT sid FROM tblUsers WHERE uid='" & username & "'")
          If ntUserRes.EOF Then
            Dim ntUpdRes, ntSidRes
            ntUpdRes = SQLQuery(cnnDB, "INSERT INTO tblUsers (sid, uid) VALUES (" & GetUnique(cnnDB, "users") & ", '" & username & "')")
            ntSidRes = SQLQuery(cnnDB, "SELECT sid FROM tblUsers WHERE uid='" & username & "'")
            Session("lhd_sid") = ntSidRes("sid")
            url = "register.asp?edit=1&new=1"
          Else
            Session("lhd_sid") = ntUserRes("sid")
          End If
          ntUserRes.Close

        Case 2	' DB Authentication
          If Request.Form("logon") = 1 Then
            Dim password
            username = Left(Lcase(Trim(Request.Form("uid"))), 50)
            username = Replace(username, "'", "''")
            password = Left(Trim(Request.Form("password")), 50)
            
            Dim userRes
            Set userRes = SQLQuery(cnnDB, "SELECT sid, password FROM tblUsers WHERE uid='" & username & "'")
            If userRes.EOF Then
              invalid = TRUE
              url=""
            ElseIf userRes("password") <> password Then
              invalid = TRUE
              frm_url = url
              url=""
            Else
              Session("lhd_sid") = userRes("sid")
            End If
            userRes.Close
          Else
            Dim frm_url
            frm_url = url
            url = ""
          End If

        Case 3	' External Authentication
          If Len(Session("lhd_ext_uid")) > 0 Then
            username = Lcase(Trim(Session("lhd_ext_uid")))
          ElseIf Len(Request.Form("lhd_ext_uid")) > 0 Then
            username = Lcase(Trim(Request.Form("lhd_ext_uid")))
          ElseIf Len(Request.QueryString("lhd_ext_uid")) > 0 Then
            username = Lcase(Trim(Request.QueryString("lhd_ext_uid")))
          Else
            Call DisplayError (3, lang(cnnDB, "Nousernamewasspecifiedbytheexternalauthenication") & ".")
          End If

          If CBool(InStr(username, "'")) Then
            Call DisplayError (3, lang(cnnDB, "Username") & "&nbsp;" & Lang(cnnDB, "containsinvalidcharacters") & ".")
          End If

          Dim extUserRes
          Set extUserRes = SQLQuery(cnnDB, "SELECT sid FROM tblUsers WHERE uid='" & username & "'")
          If extUserRes.EOF Then
            Dim extUpdRes, extSidRes
            extUpdRes = SQLQuery(cnnDB, "INSERT INTO tblUsers (sid, uid) VALUES (" & GetUnique(cnnDB, "users") & ", '" & username & "')")
            extSidRes = SQLQuery(cnnDB, "SELECT sid FROM tblUsers WHERE uid='" & username & "'")
            Session("lhd_sid") = extSidRes("sid")
            url = "register.asp?edit=1&new=1"
          Else
            Session("lhd_sid") = extUserRes("sid")
          End If
          extUserRes.Close

      End Select

      Session("lhd_IsAdmin") = FALSE

      If Len(url) > 0 Then
        ' Update logon time
        ' Update logon time
        Dim updTimeRes
        Set updTimeRes = SQLQuery(cnnDB, "UPDATE tblUsers SET dtLastAccess=" & SQLDate(Now, lhdAddSQLDelim) & " WHERE sid=" & Session("lhd_sid"))
        cnnDB.Close
        Response.Redirect url
      End If
    %>
    <div align="center">
      <table Class="Narrow">
        <tr class="Head1">
          <td>
            <% = Cfg(cnnDB, "SiteName") %>
            <br />
            <%=lang(cnnDB, "HelpDesk")%>
          </td>
        </tr>
        <tr class="Body1">
          <td>
            <form action="logon.asp?URL=<% = frm_url %>" method="POST">
              <input type="hidden" name="logon" value="1">
              <div align="center">
                <u><h3><%=lang(cnnDB, "Logon")%></h3></u>
              </div>
           <% If invalid Then 
                Response.Write "<div align=""center"">" & _
                  "<i>" & lang(cnnDB, "Invalidusernameorpassword") & ".</i></div>"
              End If %>
              <table class="Narrow" border="0">
                <tr>
                  <td>
                    <b><%=lang(cnnDB, "UserName")%>:</b>
                  </td>
                  <td>
                    <input type="text" name="uid" size="20">
                  </td>
                </tr>
                <tr>
                  <td>
                    <b><%=lang(cnnDB, "Password")%>:</b>
                  </td>
                  <td>
                    <input type="password" name="password" size="20">
                  </td>
                </tr>
                <tr><td>&nbsp;</td></tr>
              </table>
              <div align="center">
                <input type="submit" value="<%=lang(cnnDB, "Logon")%>">
              </div>
            </form>
            <p>
            <div align="center">
              <% If Cint(Cfg(cnnDB, "EnableKB")) = 3 Then %>
                <a href="kb/"><% = lang(cnnDB, "SearchtheKnowledgeBase") %></a>
                <p>
              <% End If %>
              <a href="register.asp"><%=lang(cnnDB, "NewUser")%></a>
              <% 
              If Cfg(cnnDB, "EmailType") <> 0 Then 
                Response.Write "| <a href=""forgotpass.asp"">" & lang(cnnDB, "EmailMyPassword") & "</a>"
              End If 
              %>
            </div>
            </p>
          </td>
        </tr>
      </table>
    </div>
    <%
      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
