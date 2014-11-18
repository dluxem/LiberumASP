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

  Filename: forgotpass.asp
  Date:     $Date: 2002/01/03 15:40:49 $
  Version:  $Revision: 1.51 $
  Purpose:  Emails the user their password.

  -->
  <!-- 	#include file = "public.asp" -->
  <%

    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid

    Response.Write "<head><title>" & lang(cnnDB, "HelpDesk") & " - " & lang(cnnDB, "EmailPassword") & _
      "</title>" & vbnewline & "<link rel=""stylesheet"" type=""text/css"" href=""default.css"">" & _
      "</head><body>"

    Dim success, invalidUID
    success = FALSE
    invalidUID = FALSE

    If Request.Form("email") = 1 Then
      Dim uid, userRes
      uid = Lcase(Trim(Request.Form("uid")))

      Set userRes = SQLQuery(cnnDB, "SELECT email1, password FROM tblUsers WHERE uid='" & uid & "'")

      If userRes.EOF Then
        invalidUID = TRUE
      Else
        Dim strSubject, strBody
        ' Send password to the user
        strSubject = lang(cnnDB, "HELPDESK") & " : " & lang(cnnDB, "password")
        strBody = _
        "Username: " & uid & vbNewLine & _
        "Password: " & userRes("password") & vbNewLine & _
        vbNewLine & _
        "Log in to the help desk @ " & Cfg(cnnDB, "BaseURL")

        Call SendMail(userRes("email1"), strSubject, strBody, cnnDB)

        success = TRUE
      End IF
    End If
  %>

  <div align="center">
    <table class="Narrow">
      <tr class="Head1">
        <td>
          <%=lang(cnnDB, "EmailPassword")%>
        </td>
      </tr>
      <tr class="Body1">
        <td>
       <% If success Then 
            response.write "<div align=""center"">" & vbnewline & _
              "<u><h3>" & lang(cnnDB, "PasswordSent") & "</h3></u>" & _
              "<p>" & lang(cnnDB, "PasswordSentText") & ".</p>" & _
              "<p><b><a href=""logon.asp"">" & lang(cnnDB, "HelpDesk") & _
              "&nbsp;" & lang(cnnDB, "Logon") & "</a></b></p>" & vbnewline & _
              "</div>"
           Else %>
            <form action="forgotpass.asp" method="POST">
              <input type="hidden" name="email" value="1">
              <div align="center">
                <u><h3><%=lang(cnnDB, "EnterUsername")%></h3></u>
                <% If invalidUID Then %>
                  <i><%=lang(cnnDB, "Invalidusername")%>.</i>
                  <br /><br />
                <% End If %>
                <input type="text" name="uid" size="15">
                <br /><br />
                <input type="submit" value="<%=lang(cnnDB, "EmailPassword")%>">
              </div>
            </form>
          <% End If %>
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
