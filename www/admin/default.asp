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

  Filename: default.asp
  Date:     $Date: 2002/01/24 14:57:50 $
  Version:  $Revision: 1.52.2.1 $
  Purpose:  This is the main administrative menu.  A password prompt will be
  displayed before the menu is shown.
  -->
  
  <!--  #include file = "../settings.asp" -->
  <!-- 	#include file = "../public.asp" -->

  <% 
    Call SetAppVariables

    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title>
      <%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "AdministrativeMenu")%>
    </title>
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>

    <%
      ' Ask for password if the user has not already entered it.
      If Not Session("lhd_IsAdmin") Then

        ' If the password was submitted to the page via a form, check
        ' it and allow access, setting IsAdmin to TRUE
        If Trim(Request.Form("password")) = Cfg(cnnDB, "AdminPass") Then
          Session("lhd_IsAdmin") = TRUE
        Else	' Else display form
        %>
        <div align="center">
          <table class="Normal">
            <tr class="Head1">
              <td>
                <%=lang(cnnDB, "AdministrativeLogon")%>
              </td>
            </tr>
            <tr class="Body1">
              <td>
                <p>&nbsp;
                <% 'Look for a wrong password and display an error
                   If Len(Request.Form("password")) > 0 Then %>
                     <div align="center">
                      <%=lang(cnnDB, "Passwordisincorrect")%>
                     </div>
                <% End If %>
                <form method="post" action="default.asp">
                  <p>
                  <div align="center">
                    <b><%=lang(cnnDB, "Password")%>:</b> <input type="password" name="password" size="20">
                    <input type="submit" Value="<%=lang(cnnDB, "Logon")%>">
                  </div>
                </form>
              </td>
            </tr>
          </table>
        </div>

        <%
        Call DisplayFooter(cnnDB, sid)

        ' Don't allow the browser to cache the logon page
        Response.Expires = -1
        Response.End

        End If
      End If

    ' User is logged in with admin privs, now
    ' display the normal menu

    %>
    <div align="center">
      <table class="Normal">
        <tr class="Head1">
          <td>
            <%=lang(cnnDB, "AdministrativeMenu")%>
          </td>
        </tr>
        <tr class="Body1" align="center">
          <td>
            <a href="config.asp"><%=lang(cnnDB, "Configure")%>&nbsp;<%=lang(cnnDB, "Site")%></a>
          </td>
        </tr>
        <tr class="Body1" align="center">
          <td>
            <a href="test.asp"><%=lang(cnnDB, "TestConfiguration")%></a>
          </td>
        </tr>
        <tr class="Body1" align="center">
          <td>
            <a href="cfgemail.asp"><%=lang(cnnDB, "Configure")%>&nbsp;<%=lang(cnnDB, "EmailMessages")%></a>
          </td>
        </tr>
        <tr class="Body1" align="center">
          <td>
            <a href="adminpass.asp"><%=lang(cnnDB, "ChangeAdminPassword")%></a>
          </td>
        </tr>
        <tr class="Body1" align="center">
          <td>
            <a href="viewusers.asp"><%=lang(cnnDB, "Manage")%>&nbsp;<%=lang(cnnDB, "Users")%></a>
          </td>
        </tr>
        <tr class="Body1" align="center">
          <td>
            <a href="viewcat.asp"><%=lang(cnnDB, "Manage")%>&nbsp;<%=lang(cnnDB, "Categories")%></a>
          </td>
        </tr>
        <tr class="Body1" align="center">
          <td>
            <a href="viewdep.asp"><%=lang(cnnDB, "Manage")%>&nbsp;<%=lang(cnnDB, "Departments")%></a>
          </td>
        </tr>
        <tr class="Body1" align="center">
          <td>
            <a href="viewpri.asp"><%=lang(cnnDB, "Manage")%>&nbsp;<%=lang(cnnDB, "Priorities")%></a>
          </td>
        </tr>
        <tr class="Body1" align="center">
          <td>
            <a href="viewstatus.asp"><%=lang(cnnDB, "Manage")%>&nbsp;<%=lang(cnnDB, "Statuses")%></a>
          </td>
        </tr>
        <tr class="Body1" align="center">
          <td>
            <a href="viewlang.asp"><%=lang(cnnDB, "Manage")%>&nbsp;<%=lang(cnnDB, "Languages")%></a>
          </td>
        </tr>
        <tr class="Body1" align="center">
          <td>
            <a href="reports.asp"><%=lang(cnnDB, "Reports")%></a>
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
