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

  Filename: cfgemail_help.asp
  Date:     $Date: 2001/12/09 02:01:24 $
  Version:  $Revision: 1.50 $
  Purpose:  Help page for configuring email.
  -->
  
  <!-- 	#include file = "../public.asp" -->

  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title>
      <%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "MessageConfigurationHelp")%>
    </title>
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
        <td colspan="2">
          <%=lang(cnnDB, "MessageConfigurationHelp")%>
        </td>
      </tr>
      <tr class="Body1">
        <td colspan="2">
          <%=lang(cnnDB, "MessageConfigurationHelpText")%>
          <p>
            <%=lang(cnnDB, "ProblemID")%>: <b>[problemid]</b><br>
            <%=lang(cnnDB, "Title")%>: <b>[title]</b>
          </p><br />
        </td>
      </tr>
      <tr class="Body1">
        <td>
          <b><%=lang(cnnDB, "Variable")%></b>
        </td>
        <td>
          <b><%=lang(cnnDB, "Definition")%></b>
        </td>
      </tr>
      <tr class="Body1">
        <td>
          problemid
        </td>
        <td>
          <%=lang(cnnDB, "problemidHelp")%>
        </td>
      </tr>
      <tr class="Body1">
        <td>
          title
        </td>
        <td>
          <%=lang(cnnDB, "titleHelp")%>
        </td>
      </tr>
      <tr class="Body1">
        <td>
          category
        </td>
        <td>
          <%=lang(cnnDB, "categoryHelp")%>
        </td>
      </tr>
      <tr class="Body1">
        <td>
          status
        </td>
        <td>
          <%=lang(cnnDB, "statusHelp")%>
        </td>
      </tr>
      <tr class="Body1">
        <td>
          priority
        </td>
        <td>
          <%=lang(cnnDB, "priorityHelp")%>
        </td>
      </tr>
      <tr class="Body1">
        <td>
          startdate
        </td>
        <td>
          <%=lang(cnnDB, "startdateHelp")%>
        </td>
      </tr>
      <tr class="Body1">
        <td>
          closedate
        </td>
        <td>
          <%=lang(cnnDB, "closedateHelp")%>
        </td>
      </tr>
      <tr class="Body1">
        <td>
          department
        </td>
        <td>
          <%=lang(cnnDB, "departmentHelp")%>
        </td>
      </tr>
      <tr class="Body1">
        <td>
          description
        </td>
        <td>
          <%=lang(cnnDB, "descriptionHelp")%>
        </td>
      </tr>
      <tr class="Body1">
        <td>
          notes
        </td>
        <td>
          <%=lang(cnnDB, "notesHelp")%>
        </td>
      </tr>
      <tr class="Body1">
        <td>
          solution
        </td>
        <td>
          <%=lang(cnnDB, "solutionHelp")%>
        </td>
      </tr>
      <tr class="Body1">
        <td>
          uid
        </td>
        <td>
          <%=lang(cnnDB, "uidHelp")%>
        </td>
      </tr>
      <tr class="Body1">
        <td>
          ufname
        </td>
        <td>
          <%=lang(cnnDB, "ufnameHelp")%>
        </td>
      </tr>
      <tr class="Body1">
        <td>
          uemail
        </td>
        <td>
          <%=lang(cnnDB, "uemailHelp")%>
        </td>
      </tr>
      <tr class="Body1">
        <td>
          phone
        </td>
        <td>
          <%=lang(cnnDB, "phoneHelp")%>
        </td>
      </tr>
      <tr class="Body1">
        <td>
          location
        </td>
        <td>
          <%=lang(cnnDB, "locationHelp")%>
        </td>
      </tr>
      <tr class="Body1">
        <td>
          uurl
        </td>
        <td>
          <%=lang(cnnDB, "uurlHelp")%>
        </td>
      </tr>
      <tr class="Body1">
        <td>
          rid
        </td>
        <td>
          <%=lang(cnnDB, "ridHelp")%>
        </td>
      </tr>
      <tr class="Body1">
        <td>
          rfname
        </td>
        <td>
          <%=lang(cnnDB, "rfnameHelp")%>
        </td>
      </tr>
      <tr class="Body1">
        <td>
          remail
        </td>
        <td>
          <%=lang(cnnDB, "remailHelp")%>
        </td>
      </tr>
      <tr class="Body1">
        <td>
          rurl
        </td>
        <td>
          <%=lang(cnnDB, "rurlHelp")%>
        </td>
      </tr>
      <tr class="Body1">
        <td>
          baseurl
        </td>
        <td>
          <%=lang(cnnDB, "baseurlHelp")%>
        </td>
      </tr>
    </table>
    <p>
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
