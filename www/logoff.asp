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

  Filename: logoff.asp
  Date:     $Date: 2002/01/24 15:10:05 $
  Version:  $Revision: 1.50.2.1 $
  Purpose:  Logs off the user by removing session variables.

  -->
  <!-- 	#include file = "public.asp" -->
  <%

    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title>
      <%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "LogOff")%>
    </title>
    <link rel="stylesheet" type="text/css" href="default.css"> 
  </head>
  <body>

    <div align="center">
      <table Class="Normal">
        <tr class="Head1">
          <td>
            <% = Cfg(cnnDB, "SiteName") %>
            <br />
            <%=lang(cnnDB, "HelpDesk")%>
          </td>
        </tr>
        <tr Class="Body1">
          <td align="center">
            <%=lang(cnnDB, "Yoursessionhasbeenloggedoff")%>.
            <p>
            <b><a href="<% = Cfg(cnnDB, "BaseURL") %>/logon.asp"><%=lang(cnnDB, "Clickheretologin")%>.</a></b></p>
          </td>
        </tr>
      </table>
    </div>

    <%
      Session("lhd_LanguageID") = Empty
      Session("lhd_IsAdmin") = False
      Session("lhd_sid") = 0
      sid = 0

      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
