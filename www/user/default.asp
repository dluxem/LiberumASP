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
  Date:     $Date: 2001/12/09 02:01:24 $
  Version:  $Revision: 1.50 $
  Purpose:  This is the main menu for users.
  -->

  <!-- 	#include file = "../public.asp" -->
  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title>
      <% = Cfg(cnnDB, "SiteName") %>&nbsp;<%=lang(cnnDB, "HelpDesk")%>
    </title>
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>
    <%
      ' See is user is validated
      Call CheckUser(cnnDB, sid)

      Call DisplayHeader(cnnDB, sid)
    %>

    <div align="center">
      <table class="Normal">
        <tr class="Head1">
          <td>
            <% = Cfg(cnnDB, "SiteName") %>
          </td>
        </tr>
        <tr class="Head2">
          <td>
            <div align="center">
              <%=lang(cnnDB, "HelpDesk")%>
            </div>
          </td>
        </tr>
        <tr class="Body1">
          <td>
            <div align="center">
              <b><a href="new.asp"><%=lang(cnnDB, "SubmitNewProblem")%></a></b>
            </div>
          </td>
        </tr>
        <tr class="Body1">
          <td>
            <div align="center">
              <a href="view.asp"><%=lang(cnnDB, "ViewProblemList")%></a>
            </div>
          </td>
        </tr>
        <tr class="Body1">
          <td valign="center">
            <div align="center">
              <br />
              <form method="POST" action="details.asp">
                <input type="text" size="6" name="id"> <input type="submit" value="<%=lang(cnnDB, "OpenspecificID")%>">
              </form>
            </div>
          </td>
        </tr>
        <% If Cfg(cnnDB, "EnableKB") >= 2 Then %>
          <tr class="Head2">
            <td>
              <div align="center">
                <%=lang(cnnDB, "KnowledgeBase")%>
              </div>
            </td>
          </tr>
          <tr class="Body1">
            <td>
              <div align="center">
                <a href="../kb/default.asp"><%=lang(cnnDB, "SearchtheKnowledgeBase")%></a>
              </div>
            </td>
          </tr>
          <tr class="Body1">
            <td valign="center">
              <div align="center">
                <br />
                <form method="POST" action="../kb/details.asp">
                  <input type="text" size="6" name="id"> <input type="submit" value="<%=lang(cnnDB, "LookupbyID")%>">
                </form>
              </div>
            </td>
          </tr>
        <% End If %>
        <tr class="Head2">
          <td>
            <div align="center">
              <%=lang(cnnDB, "Other")%>
            </div>
          </td>
        </tr>
        <% If Cfg(cnnDB, "UseInoutBoard") = 1 Then %>
          <tr class="Body1">
            <td align="center">
              <a href="../inout/default.asp"><%=lang(cnnDB, "InOutBoard")%></a>
            </td>
          </tr>
        <% End If %>
        <tr class="Body1">
          <td>
            <div align="center">
              <a href="../register.asp?edit=1"><%=lang(cnnDB, "EditInformation")%></a>
            </div>
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
