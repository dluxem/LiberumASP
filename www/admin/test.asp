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

  Filename: test.asp
  Date:     $Date: 2002/01/24 14:57:50 $
  Version:  $Revision: 1.1.2.1 $
  Purpose:  Misc commands for testing purpose
  -->
  
  <!-- 	#include file = "../public.asp" -->

  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title>
      <%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=Lang(cnnDB, "TestConfiguration")%>
    </title>
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>
    <div align="center">
      <table class="Normal">
        <tr class="Head1">
          <td>
            <%=Lang(cnnDB, "TestConfiguration")%>
          </td>
        </tr>
        <%
          ' Check for perms to view this page
          Call CheckAdmin

          Dim doit
          doit = request.querystring("doit")

          If doit = 1 Then
            Dim rstConfig
            ' Get current configuration
            Set rstConfig = SQLQuery(cnnDB, "Select * From tblConfig")
      
            If rstConfig.EOF Then
              Call DisplayError(3, lang(cnnDB, "Unabletoreadconfigurationfromthedatabase"))
            End If

            Dim strTo, strSubject, StrBody            
            strTo = rstConfig("HDReply")
            strSubject = "Test Message"
            strBody = "This is a test message from Liberum Help Desk"
            
            Call SendMail (strTo, strSubject, strBody, cnnDB)
            Response.Write "<tr class=""Head2"" align=""center""><td>" & Lang(cnnDB,"Messagesentto") & " " & strTo & "</td></tr>"
           	rstConfig.Close
          End if
        %>
        <tr class="body1c">
          <td><a href="sysinfo.asp"><%=Lang(cnnDB, "ShowSystemInformation")%></a></td>
        </tr>

        <tr class="body1c">
          <td><a href="test.asp?doit=1"><%=Lang(cnnDB, "SendTestEmail")%></a></td>
        </tr>

        <tr class="body2">
          <td align="right"><a href="default.asp"><%=lang(cnnDB, "AdministrativeMenu")%></a></td>
        </tr>
      </table>
    </div>

    <%

      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
