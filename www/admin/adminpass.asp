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

  Filename: adminpass.asp
  Date:     $Date: 2001/12/09 02:01:24 $
  Version:  $Revision: 1.50 $
  Purpose:  Form to set the admin password.
  -->
  
  <!-- 	#include file = "../public.asp" -->

  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title>
      <%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "ChangeAdminPassword")%>
    </title>
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>

    <%
      ' Check for perms to view this page
      Call CheckAdmin

      ' Save Results
      If Request.Form("save") = "1" Then
        Dim strSQL, AdminPass1, AdminPass2, CurrPass, OldPass, strMessage, updRes
        AdminPass1 = Left(Trim(Request.Form("AdminPass1")), 50)
        AdminPass2 = Left(Trim(Request.Form("AdminPass2")), 50)
        CurrPass = Trim(Request.Form("CurrPass"))
        OldPass = Cfg(cnnDB, "AdminPass")

        strSQL = "UPDATE tblConfig SET " & _
          "AdminPass = '" & AdminPass1 & "'"

        If (AdminPass1 = AdminPass2) And (CurrPass = OldPass) Then
          Set updRes = SQLQuery(cnnDB, strSQL)
          strMessage = lang(cnnDB, "PasswordChanged")
        Else
          strMessage = lang(cnnDB, "PasswordChangeFailed")
        End If

      End If

    %>

    <div align="center">
      <form method="post" action="adminpass.asp">
        <input type="hidden" name="save" value="1">
        <table class="Normal">
          <tr class="Head1">
            <td>
              <%=lang(cnnDB, "ChangeAdminPassword")%>
            </td>
          </tr>
          <% If Request.Form("save") = "1" Then %>
            <tr class="Head2">
              <td>
                <div align="center">
                  <% = strMessage %>
                </div>
              </td>
            </tr>
          <% End If %>
          <tr class="Body1">
            <td>
              <table class="Normal">
                <tr>
                  <td width="120">
                    
                    <b><%=lang(cnnDB, "CurrentPassword")%>:</b>
                  </td>
                  <td>
                    <input type="password" size="30" name="CurrPass">
                  </td>
                </tr>
                <tr>
                  <td width="120">
                    <b><%=lang(cnnDB, "NewPassword")%>:</b>
                  </td>
                  <td>
                    <input type="password" size="30" name="AdminPass1">
                  </td>
                </tr>
                <tr>
                  <td width="120">
                    <b><%=lang(cnnDB, "ConfirmPassword")%>:</b>
                  </td>
                  <td>
                    <input type="password" size="30" name="AdminPass2">
                  </td>
                </tr>
              </table>
              <p>
              <div align="center">
                <input type="submit" value="<%=lang(cnnDB, "Save")%>">
              </div>
            </td>
          </tr>
        </table>
      </form>
      <p>
      <a href="default.asp"><%=lang(cnnDB, "AdministrativeMenu")%></a>
    </div>

    <%

      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
