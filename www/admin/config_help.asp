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

  Filename: config_help.asp
  Date:     $Date: 2001/12/09 02:01:24 $
  Version:  $Revision: 1.50 $
  Purpose:  Describes the various settings.
  -->
  
  <!-- 	#include file = "../public.asp" -->

  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title><%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "ConfigurationHelp")%></title>
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>

    <%
      ' Check for perms to view this page
      Call CheckAdmin

    %>

    <div align="center">
      <table class="Normal">
        <tr class="Head1">
          <td colspan="2" >
            <%=lang(cnnDB, "ConfigurationHelp")%>
          </td>
        </tr>
        <tr class="Body1" valign="top" valign="top">
          <td><b><%=lang(cnnDB, "SiteName")%></b></td>
          <td><%=lang(cnnDB, "SiteNameHelp")%></td>
        </tr>
        <tr class="Body2" valign="top">
          <td><b><%=lang(cnnDB, "BaseURL")%></b></td>
          <td><%=lang(cnnDB, "BaseURLHelp_2")%></td>
        </tr>
        <tr class="Body1" valign="top">
          <td><b><%=lang(cnnDB, "AdministratorsName")%></b></td>
          <td><%=lang(cnnDB, "AdministratorsNameHelp")%></td>
        </tr>
        <tr class="Body2" valign="top">
          <td><b><%=lang(cnnDB, "AdministratorsEmail")%></b></td>
          <td><%=lang(cnnDB, "AdministratorsEmailHelp")%></td>
        </tr>
        <tr class="Body1" valign="top">
          <td><b><%=lang(cnnDB, "BaseEmail")%></b></td>
          <td><%=lang(cnnDB, "BaseEmailHelp")%></td>
        </tr>
        <tr class="Body2" valign="top">
          <td><b><%=lang(cnnDB, "EmailType")%></b></td>
          <td><%=lang(cnnDB, "EmailTypeHelp")%></td>
        </tr>
        <tr class="Body1" valign="top">
          <td><b><%=lang(cnnDB, "SMTPServer")%></b></td>
          <td><%=lang(cnnDB, "SMTPServerHelp")%></td>
        </tr>
        <tr class="Body2" valign="top">
          <td><b><%=lang(cnnDB, "PagerPriorityLevel")%></b></td>
          <td><%=lang(cnnDB, "PagerPriorityLevelHelp")%></td>
        </tr>
        <tr class="Body1" valign="top">
          <td><b><%=lang(cnnDB, "EmailUseronUpdate")%></b></td>
          <td><%=lang(cnnDB, "EmailUseronUpdateHelp")%></td>
        </tr>
        <tr class="Body2" valign="top">
          <td><b><%=lang(cnnDB, "EnableUserKB")%></b></td>
          <td><%=lang(cnnDB, "EnableUserKBHelp")%></td>
        </tr>
        <tr class="Body1" valign="top">
          <td><b><%=lang(cnnDB, "KBSQLFreeTextSearches")%></b></td>
          <td><%=lang(cnnDB, "KBSQLFreeTextSearchesHelp")%></td>
        </tr>
        <tr class="Body2" valign="top">
          <td><b><%=lang(cnnDB, "DefaultPriority")%></b></td>
          <td><%=lang(cnnDB, "DefaultPriorityHelp")%></td>
        </tr>
        <tr class="Body1" valign="top">
          <td><b><%=lang(cnnDB, "DefaultStatus")%></b></td>
          <td><%=lang(cnnDB, "DefaultStatusHelp")%></td>
        </tr>
        <tr class="Body2" valign="top">
          <td><b><%=lang(cnnDB, "CloseStatus")%></b></td>
          <td><%=lang(cnnDB, "CloseStatusHelp")%></td>
        </tr>
        <tr class="Body1" valign="top">
          <td><b><%=lang(cnnDB, "AuthenticationType")%></b></td>
          <td><%=lang(cnnDB, "AuthenticationTypeHelp")%></td>
        </tr>
        <tr class="Body2" valign="top">
          <td><b><%=lang(cnnDB, "UseSelectUser")%></b></td>
          <td><%=lang(cnnDB, "UseSelectUserHelp")%></td>
        </tr>
        <tr class="Body1" valign="top">
          <td><b><%=lang(cnnDB, "UseInOutBoard")%></b></td>
          <td><%=lang(cnnDB, "UseInOutBoardHelp")%></td>
        </tr>
        <tr class="Body2" valign="top">
          <td><b><%=lang(cnnDB, "AllowImageUpload")%></b></td>
          <td><%=lang(cnnDB, "AllowImageUploadHelp")%></td>
        </tr>
        <tr class="Body1" valign="top">
          <td><b><%=lang(cnnDB, "MaxImageSize")%></b></td>
          <td><%=lang(cnnDB, "MaxImageSizeHelp")%></td>
        </tr>
        <tr class="Body2" valign="top">
          <td><b><%=lang(cnnDB, "Defaultlanguage")%></b></td>
          <td><%=lang(cnnDB, "DefaultlanguageHelp")%></td>
        </tr>
      </table>
      <form>
        <input type=button value="<%=lang(cnnDB, "CloseThisWindow")%>" onClick="javascript:window.close();">
      </form>
    </div>

    <%
    '	cfgRes.Close

      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
