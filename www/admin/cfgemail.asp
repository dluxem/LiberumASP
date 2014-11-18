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

  Filename: cfgemail.asp
  Date:     $Date: 2001/12/09 02:01:24 $
  Version:  $Revision: 1.50 $
  Purpose:  Form to configure the email messages sent to users and reps.
  -->
  
  <!-- 	#include file = "../public.asp" -->

  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title>
      <%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "MessageConfiguration")%>
    </title>
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>

    <%
      ' Check for perms to view this page
      Call CheckAdmin

      Dim eType, displayMenu

      eType = Left(Trim(Request.Form("type")), 50)
      If Len(eType) > 0 Then
        displayMenu = False
      Else
        displayMenu = True
      End If

      ' Save Results
      If Request.Form("save") = "1" Then
        Dim strSQL, updRes
        Dim Subject, Body

        Subject = Left(Trim(Request.Form("subject")), 50)
        Body = Trim(Request.Form("body"))


        strSQL = "UPDATE tblEmailMsg SET " & _
          "subject = '" & Subject & "', " & _
          "body = '" & Body & "' " & _
          "WHERE type='" & eType & "'"

        Set updRes = SQLQuery(cnnDB, strSQL)

      End If

      If Not displayMenu Then
        Dim cfgRes
        ' Get current configuration
        Set cfgRes = SQLQuery(cnnDB, "Select Subject, Body From tblEmailMsg WHERE type='" & eType & "'")

        If cfgRes.EOF Then
          Call DisplayError(3, lang(cnnDB, "Unable to read message from the database") & ".")
        End If
      End If
    %>

    <div align="center">
      <table class="Normal">
        <tr class="Head1">
          <td>
            <%=lang(cnnDB, "EmailMessages")%>
          </td>
        </tr>
        <% If displayMenu Then %>
          <tr class="Body1">
            <td>
              <form method="post" action="cfgemail.asp">
                <table class="Normal">
                  <tr>
                    <td width="120">
                      <b><%=lang(cnnDB, "Message")%>:</b>
                    </td>
                    <td>
                      <select name="type">
                        <option value="usernew"><%=lang(cnnDB, "User")%> - <%=lang(cnnDB, "New")%></option>
                        <option value="userupdate"><%=lang(cnnDB, "User")%> - <%=lang(cnnDB, "Update")%></option>
                        <option value="userclose"><%=lang(cnnDB, "User")%> - <%=lang(cnnDB, "Close")%></option>
                        <option value="repnew"><%=lang(cnnDB, "Rep")%> - <%=lang(cnnDB, "New")%></option>
                        <option value="repupdate"><%=lang(cnnDB, "Rep")%> - <%=lang(cnnDB, "Update")%></option>
                        <option value="repclose"><%=lang(cnnDB, "Rep")%> - <%=lang(cnnDB, "Close")%></option>
                        <option value="reppager"><%=lang(cnnDB, "Rep")%> - <%=lang(cnnDB, "Pager")%></option>
                      </select>
                    </td>
                  </tr>
                </table>
                <p>
                <div align="center">
                  <input type="submit" value="<%=lang(cnnDB, "EditMessage")%>">
                </div>
              </form>
            </td>
          </tr>
        <% Else %>
          <% 	If Request.Form("save") = "1" Then %>
            <tr class="Head2">
              <td>
                <div align="center">
                  <i><%=lang(cnnDB, "MessageSaved")%>.</i>
                </div>
              </td>
            </tr>
         <% End If %>
          <tr class="Body1">
            <td>
              <form method="post" action="cfgemail.asp">
                <input type="hidden" name="save" value="1">
                <input type="hidden" name="type" value="<% = eType %>">
                <table class="Normal">
                  <tr>
                    <td width="120">
                      <b><%=lang(cnnDB, "Subject")%>:</b>
                    </td>
                    <td>
                      <input type="text" size="50" name="subject" value="<% = cfgRes("Subject") %>">
                    </td>
                  </tr>
                </table>
                <b><%=lang(cnnDB, "Body")%>:</b><br>
                <div align="center">
                  <textarea rows="8" cols="80" name="body"><% = cfgRes("body") %></textarea>
                  <p>
                  <a href="cfgemail_help.asp" target="#"><%=lang(cnnDB, "SyntaxHelp")%></a>
                  <p>
                  <input type="submit" value="<%=lang(cnnDB, "Save")%>">
                </div>
              </form>
            </td>
          </tr>
        <% End If %>
      </table>
      <p>
      <% If Not displayMenu Then %>
        <a href="cfgemail.asp"><%=lang(cnnDB, "Chooseanothermessage")%></a><br>
      <% End If %>
      <a href="default.asp"><%=lang(cnnDB, "AdministrativeMenu")%></a>
    </div>
    <%
    '	cfgRes.Close

      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
