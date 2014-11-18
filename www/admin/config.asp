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

  Filename: config.asp
  Date:     $Date: 2001/12/09 02:01:24 $
  Version:  $Revision: 1.50 $
  Purpose:  Allows modifications to the site's settings.
  -->
  
  <!-- 	#include file = "../public.asp" -->

  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title>
      <%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "Configuration")%>
    </title>
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>

    <%
      ' Check for perms to view this page
      Call CheckAdmin

      ' Save Results
      If Request.Form("save") = "1" Then
        Dim strSQL, updRes
        Dim SiteName, BaseURL, HDName, HDReply, BaseEmail, EmailType, intNotifyUser
        Dim DefaultPriority, DefaultStatus, KBFreeText
        Dim CloseStatus, AuthType, DefaultLanguage
        Dim EnableKB, SMTPServer, EnablePager, UseSelectUser, UseInoutBoard
        Dim AllowImageUpload, MaxImageSize

        SiteName = Left(Trim(Request.Form("sitename")), 50)
        BaseURL = Left(Trim(Request.Form("baseurl")), 50)
        HDName = Left(Trim(Request.Form("hdname")), 50)
        HDReply = Left(Trim(Request.Form("hdreply")), 50)
        BaseEmail = Left(Trim(Request.Form("baseemail")), 50)
        EmailType = Cint(Request.Form("emailtype"))
        intNotifyUser = Cint(Request.Form("notifyuser"))
        DefaultPriority = Cint(Request.Form("defaultpriority"))
        DefaultStatus = Cint(Request.Form("defaultstatus"))
        CloseStatus = Cint(Request.Form("closestatus"))
        AuthType = Cint(Request.Form("authtype"))
        EnableKB = Cint(Request.Form("enablekb"))
        SMTPServer = Left(Trim(Request.Form("smtpserver")), 50)
        EnablePager = Cint(Request.Form("enablepager"))
        UseSelectUser = Cint(Request.Form("UseSelectUser"))
        UseInoutBoard = Cint(Request.Form("UseInoutBoard"))
        KBFreeText = Cint(Request.Form("KBFreeText"))
        DefaultLanguage = Cint(Request.Form("DefaultLanguage"))
        AllowImageUpload = Cint(Request.Form("AllowImageUpload"))
        MaxImageSize = Left(Trim(Request.Form("MaxImageSize")), 20)

        strSQL = "UPDATE tblConfig SET " & _
          "SiteName = '" & SiteName & "', " & _
          "BaseURL = '" & BaseURL & "', " & _
          "HDName = '" & HDName & "', " & _
          "HDReply = '" & HDReply & "', " & _
          "BaseEmail = '" & BaseEmail & "', " & _
          "NotifyUser = " & intNotifyUser & ", " & _
          "EmailType = " & EmailType & ", " & _
          "EnableKB = " & EnableKB & ", " & _
          "KBFreeText = " & KBFreeText & ", " & _
          "DefaultPriority = " & DefaultPriority & ", " & _
          "DefaultStatus = " & DefaultStatus & ", " & _
          "CloseStatus = " & CloseStatus & ", " & _
          "AuthType = " & AuthType & ", " & _
          "SMTPServer = '" & SMTPServer & "', " & _
          "UseSelectUser = " & UseSelectUser & ", " & _
          "UseInoutBoard = " & UseInoutBoard & ", " & _
          "DefaultLanguage = " & DefaultLanguage & ", " & _
          "AllowImageUpload = '" & AllowImageUpload & "', " & _
          "MaxImageSize = " & MaxImageSize & ", " & _
          "EnablePager = " & EnablePager

        Set updRes = SQLQuery(cnnDB, strSQL)

      End If

      Dim rstConfig
      ' Get current configuration
      Set rstConfig = SQLQuery(cnnDB, "Select * From tblConfig")

      If rstConfig.EOF Then
        Call DisplayError(3, lang(cnnDB, "Unabletoreadconfigurationfromthedatabase"))
      End If

    %>

    <div align="center">
      <form method="post" action="config.asp">
        <input type="hidden" name="save" value="1">
        <table class="Normal">
          <tr>
            <td>
              <div align="right">
                <a href="default.asp"><%=lang(cnnDB, "AdministrativeMenu")%></a> | 
                <b><a href="config_help.asp" target="#"><%=lang(cnnDB, "Help")%></a></b>
              </div>
            </td>
          </tr>
          <tr class="Head1">
            <td>
              <%=lang(cnnDB, "Configuration")%>
            </td>
          </tr>
          <% 	If Request.Form("save") = "1" Then %>
            <tr class="Head2">
              <td>
                <div align="center">
                  <%=lang(cnnDB, "ConfigurationSaved")%>
                </div>
              </td>
            </tr>
          <% End If %>
          <tr>
            <td>
              <table class="Normal">
                <tr class="body1">
                  <td width="120">
                    <b><%=lang(cnnDB, "SiteName")%>:</b>
                  </td>
                  <td>
                    <input type="text" size="30" name="sitename" value="<% = rstConfig("SiteName") %>">
                  </td>
                </tr>
                <tr class="body2">
                  <td>
                    <b><%=lang(cnnDB, "BaseURL")%>:</b>
                  </td>
                  <td>
                    <input type="text" size="30" name="baseurl" value="<% = rstConfig("BaseURL") %>">
                    <font size="-2"><i>(http://intranet/helpdesk)</i></font>
                  </td>
                </tr>
                <tr class="body1">
                  <td>
                    <b><%=lang(cnnDB, "AdministratorsName")%>:</b>
                  </td>
                  <td>
                    <input type="text" size="30" name="hdname" value="<% = rstConfig("HDName") %>">
                  </td>
                </tr>
                <tr class="body2">
                  <td>
                    <b><%=lang(cnnDB, "AdministratorsEmail")%>:</b>
                  </td>
                  <td>
                    <input type="text" size="30" name="hdreply" value="<% = rstConfig("HDReply") %>">
                  </td>
                </tr>
                <tr class="body1">
                  <td>
                    <b><%=lang(cnnDB, "BaseEmail")%>:</b>
                  </td>
                  <td>
                    <input type="text" size="30" name="baseemail" value="<% = rstConfig("BaseEmail") %>">
                    <font size="-2"><i>(@company.com)</i></font>
                  </td>
                </tr>
                <tr class="body2">
                  <td>
                    <b><%=lang(cnnDB, "EmailType")%>:</b>
                  </td>
                  <td>
                    <select name="emailtype">
                    <%
                      Dim optRes
                      Set optRes = SQLQuery(cnnDB, "SELECT * From tblConfig_Email")
                      If Not optRes.EOF Then
                        Do While Not optRes.EOF
                          If optRes("id") = rstConfig("EmailType") Then
                    %>
                            <option value="<% = optRes("id")%>" selected>
                            <% = optRes("type") %></OPTION>
                    <% 			Else %>
                            <option value="<% = optRes("id")%>">
                            <% = optRes("type") %></OPTION>

                    <% 			End If

                        optRes.MoveNext
                        Loop
                      End If
                      optRes.Close
                    %>
                    </select>
                  </td>
                </tr>
                <tr class="body1">
                  <td>
                    <b><%=lang(cnnDB, "SMTPServer")%>:</b>
                  </td>
                  <td>
                    <input type="text" size="30" name="smtpserver" value="<% = rstConfig("SMTPServer") %>">
                    <font size="-2"><i>(<%=lang(cnnDB, "JMailorASPEmail")%>)</i></font>
                  </td>
                </tr>
                <tr class="body2">
                  <td>
                    <b><%=lang(cnnDB, "PagerPriorityLevel")%>:</b>
                  </td>
                  <td>
                    <select name="enablepager">
                    <% If rstConfig("EnablePager") = 0 Then %>
                      <option value="0" selected><%=lang(cnnDB, "Disabled")%></option>
                    <% Else %>
                      <option value="0"><%=lang(cnnDB, "Disabled")%></option>
                    <% End If %>
                    <%
                      Set optRes = SQLQuery(cnnDB, "SELECT * FROM priority WHERE priority_id > 0")
                      If Not optRes.EOF Then
                        Do While Not optRes.EOF
                          If optRes("priority_id") = rstConfig("EnablePager") Then
                    %>
                            <option value="<% = optRes("priority_id")%>" selected>
                            <% = optRes("pname") %></OPTION>
                    <% 			Else %>
                            <option value="<% = optRes("priority_id")%>">
                            <% = optRes("pname") %></OPTION>

                    <% 			End If

                        optRes.MoveNext
                        Loop
                      End If
                      optRes.Close
                    %>
                    </select>
                  </td>
                </tr>
                <tr class="body1">
                  <td>
                    <b><%=lang(cnnDB, "EmailUseronUpdate")%>:</b>
                  </td>
                  <td>
                    <select name="notifyuser">
                    <% If rstConfig("NotifyUser") = "0" Then %>
                      <option value="0" selected><%=lang(cnnDB, "NO")%></option>
                      <option value="1"><%=lang(cnnDB, "YES")%></option>
                    <% Else %>
                      <option value="0"><%=lang(cnnDB, "NO")%></option>
                      <option value="1" selected><%=lang(cnnDB, "YES")%></option>
                    <% End If %>
                    </select>
                  </td>
                </tr>
                <tr class="body2">
                  <td>
                    <b><%=lang(cnnDB, "EnableUserKB")%>:</b>
                  </td>
                  <td>
                    <select name="enablekb">
                    <% Select Case rstConfig("EnableKB")
                        Case 0 %>
                          <option value="0" selected><%=lang(cnnDB, "Disable")%></option>
                          <option value="1"><%=lang(cnnDB, "RepsOnly")%></option>
                          <option value="2"><%=lang(cnnDB, "UsersReps")%></option>
                          <option value="3"><%=lang(cnnDB, "Anyone")%></option>
                     <% Case 1 %>
                          <option value="0"><%=lang(cnnDB, "Disable")%></option>
                          <option value="1" selected><%=lang(cnnDB, "RepsOnly")%></option>
                          <option value="2"><%=lang(cnnDB, "UsersReps")%></option>
                          <option value="3"><%=lang(cnnDB, "Anyone")%></option>
                     <% Case 2 %>
                          <option value="0"><%=lang(cnnDB, "Disable")%></option>
                          <option value="1"><%=lang(cnnDB, "RepsOnly")%></option>
                          <option value="2" selected><%=lang(cnnDB, "UsersReps")%></option>
                          <option value="3"><%=lang(cnnDB, "Anyone")%></option>
                     <% Case 3 %>
                          <option value="0"><%=lang(cnnDB, "Disable")%></option>
                          <option value="1"><%=lang(cnnDB, "RepsOnly")%></option>
                          <option value="2"><%=lang(cnnDB, "UsersReps")%></option>
                          <option value="3" selected><%=lang(cnnDB, "Anyone")%></option>
                    <% End Select%>
                    </select>
                  </td>
                </tr>
                <tr class="body1">
                  <td>
                    <b><%=lang(cnnDB, "KBSQLFreeTextSearches")%>:</b>
                  </td>
                  <td>
                    <select name="kbfreetext">
                    <% If rstConfig("KBFreeText") = "0" Then %>
                      <option value="0" selected><%=lang(cnnDB, "Disable")%></option>
                      <option value="1"><%=lang(cnnDB, "Enable")%></option>
                    <% Else %>
                      <option value="0"><%=lang(cnnDB, "Disable")%></option>
                      <option value="1" selected><%=lang(cnnDB, "Enable")%></option>
                    <% End If %>
                    </select>
                  </td>
                </tr>
                <tr class="body2">
                  <td>
                    <b><%=lang(cnnDB, "DefaultPriority")%>:</b>
                  </td>
                  <td>
                    <select name="defaultpriority">
                    <%
                      Set optRes = SQLQuery(cnnDB, "SELECT * FROM priority WHERE priority_id > 0")
                      If Not optRes.EOF Then
                        Do While Not optRes.EOF
                          If optRes("priority_id") = rstConfig("DefaultPriority") Then
                    %>
                            <option value="<% = optRes("priority_id")%>" selected>
                            <% = optRes("pname") %></OPTION>
                    <% 			Else %>
                            <option value="<% = optRes("priority_id")%>">
                            <% = optRes("pname") %></OPTION>

                    <% 			End If

                        optRes.MoveNext
                        Loop
                      End If
                      optRes.Close
                    %>
                    </select>
                  </td>
                </tr>
                <tr class="body1">
                  <td>
                    <b><%=lang(cnnDB, "DefaultStatus")%>:</b>
                  </td>
                  <td>
                    <select name="defaultstatus">
                    <%
                      Set optRes = SQLQuery(cnnDB, "SELECT * FROM status WHERE status_id > 0 AND status_id <>" & rstConfig("CloseStatus") & " ORDER BY status_id ASC")
                      If Not optRes.EOF Then
                        Do While Not optRes.EOF
                          If optRes("status_id") = rstConfig("DefaultStatus") Then
                    %>
                            <option value="<% = optRes("status_id")%>" selected>
                            <% = optRes("sname") %></OPTION>
                    <% 			Else %>
                            <option value="<% = optRes("status_id")%>">
                            <% = optRes("sname") %></OPTION>

                    <% 			End If

                        optRes.MoveNext
                        Loop
                      End If
                      optRes.Close
                    %>
                    </select>
                  </td>
                </tr>
                <tr class="body2">
                  <td>
                    <b><%=lang(cnnDB, "CloseStatus")%>:</b>
                  </td>
                  <td>
                    <select name="closestatus">
                    <%
                      Set optRes = SQLQuery(cnnDB, "SELECT * FROM status WHERE status_id > 0 ORDER BY status_id ASC")
                      If Not optRes.EOF Then
                        Do While Not optRes.EOF
                          If optRes("status_id") = rstConfig("CloseStatus") Then
                    %>
                            <option value="<% = optRes("status_id")%>" selected>
                            <% = optRes("sname") %></OPTION>
                    <% 			Else %>
                            <option value="<% = optRes("status_id")%>">
                            <% = optRes("sname") %></OPTION>

                    <% 			End If

                        optRes.MoveNext
                        Loop
                      End If
                      optRes.Close
                    %>
                    </select>
                  </td>
                </tr>
                <tr class="body1">
                  <td>
                    <b><%=lang(cnnDB, "AuthenticationType")%>:</b>
                  </td>
                  <td>
                    <select name="authtype">
                    <%
                      Set optRes = SQLQuery(cnnDB, "SELECT * From tblConfig_Auth")
                      If Not optRes.EOF Then
                        Do While Not optRes.EOF
                          If optRes("id") = rstConfig("AuthType") Then
                    %>
                            <option value="<% = optRes("id")%>" selected>
                            <% = optRes("type") %></OPTION>
                    <% 			Else %>
                            <option value="<% = optRes("id")%>">
                            <% = optRes("type") %></OPTION>

                    <% 			End If

                        optRes.MoveNext
                        Loop
                      End If
                      optRes.Close
                    %>
                    </select>
                  </td>
                </tr>
                <tr class="body2">
                  <td>
                    <b><%=lang(cnnDB, "UseSelectUser")%>:</b>
                  </td>
                  <td>
                    <select name="useSelectUser">
                    <% If rstConfig("useSelectUser") = "0" Then %>
                      <option value="0" selected><%=lang(cnnDB, "NO")%></option>
                      <option value="1"><%=lang(cnnDB, "YES")%></option>
                    <% Else %>
                      <option value="0"><%=lang(cnnDB, "NO")%></option>
                      <option value="1" selected><%=lang(cnnDB, "YES")%></option>
                    <% End If %>
                    </select>
                  </td>
                </tr>
                <tr class="body1">
                  <td>
                    <b><%=lang(cnnDB, "UseInOutBoard")%>:</b>
                  </td>
                  <td>
                    <select name="useInoutBoard">
                    <% If rstConfig("useInoutBoard") = "0" Then %>
                      <option value="0" selected><%=lang(cnnDB, "NO")%></option>
                      <option value="1"><%=lang(cnnDB, "YES")%></option>
                    <% Else %>
                      <option value="0"><%=lang(cnnDB, "NO")%></option>
                      <option value="1" selected><%=lang(cnnDB, "YES")%></option>
                    <% End If %>
                    </select>
                  </td>
                </tr>
                <tr class="body2">
                  <td>
                    <b><%=lang(cnnDB, "AllowImageUpload")%>:</b>
                  </td>
                  <td>
                    <select name="AllowImageUpload">
                    <% If rstConfig("AllowImageUpload") = "0" Then %>
                      <option value="0" selected><%=lang(cnnDB, "NO")%></option>
                      <option value="1"><%=lang(cnnDB, "YES")%></option>
                    <% Else %>
                      <option value="0"><%=lang(cnnDB, "NO")%></option>
                      <option value="1" selected><%=lang(cnnDB, "YES")%></option>
                    <% End If %>
                    </select>
                  </td>
                </tr>
                <tr class="body1">
                  <td>
                    <b><%=lang(cnnDB, "MaxImageSize")%>:</b>
                  </td>
                  <td>
                    <input type="text" size="30" name="MaxImageSize" value="<% = rstConfig("MaxImageSize") %>">
                  </td>
                </tr>
                <tr class="body2">
                  <td>
                    <b><%=lang(cnnDB, "Defaultlanguage")%>:</b>
                  </td>
                  <td>
                    <select name="DefaultLanguage">
                    <%
                      Set optRes = SQLQuery(cnnDB, "SELECT * From tblLanguage")
                      If Not optRes.EOF Then
                        Do While Not optRes.EOF
                          If optRes("id") = rstConfig("DefaultLanguage") Then
                    %>
                            <option value="<% = optRes("id")%>" selected>
                            <% = optRes("LangName") %> (<% = optRes("Localized") %>)</OPTION>
                    <% 			Else %>
                            <option value="<% = optRes("id")%>">
                            <% = optRes("LangName") %> (<% = optRes("Localized") %>)</OPTION>

                    <% 			End If

                        optRes.MoveNext
                        Loop
                      End If
                      optRes.Close
                    %>
                    </select>
                  </td>
                </tr>
              </table>
              <div align="right">
                <a href="default.asp"><%=lang(cnnDB, "AdministrativeMenu")%></a> |
                <b><a href="config_help.asp" target="#"><%=lang(cnnDB, "Help")%></a></b>
              </div>
            </td>
          </tr>
        </table>
        <input type="submit" value="<%=lang(cnnDB, "Save")%>">
      </form>
    </div>
    <%
    	rstConfig.Close

      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
