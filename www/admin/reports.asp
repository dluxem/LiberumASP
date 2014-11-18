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

  Filename: reports.asp
  Date:     $Date: 2001/12/09 02:01:24 $
  Version:  $Revision: 1.50 $
  Purpose:  Form for admins to submit report queries.
  -->
  
  <!-- 	#include file = "../public.asp" -->

  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title>
      <%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "Reports")%>
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
          <td>
            <%=lang(cnnDB, "Reports")%>
          </td>
        </tr>
        <tr class="Body1c">
          <td>
            <form method="post" action="viewreports.asp">
              <input type="radio" name="type" value="0" CHECKED> <%=lang(cnnDB, "Department")%>&nbsp;<%=lang(cnnDB, "Report")%><br />
              <input type="radio" name="type" value="1"> <%=lang(cnnDB, "Category")%>&nbsp;<%=lang(cnnDB, "Report")%><br />
              <input type="radio" name="type" value="2"> <%=lang(cnnDB, "SupportRep")%>&nbsp;<%=lang(cnnDB, "Report")%><br />
              <hr width="75%">
              <%=lang(cnnDB, "StartDate")%>:
              <select name="s_month" size="1">
              <%
                Dim temp_date, year_adj, count
                year_adj = 0
                For count = 1 to 12
                  temp_date = Month(now) - 1
                  If temp_date = 0 Then
                    temp_date = 12
                    year_adj = -1
                  End If
                  If count = temp_date Then
                    Response.Write("<option value=""" & count & """ selected>" & count & "</option>")
                  Else
                    Response.Write("<option value=""" & count & """>" & count & "</option>")
                  End If
                Next
                Response.Write ("</select> / <select name=""s_day"" size=""1"">")
                For count = 1 to 31
                  If count = Day(now) Then
                    Response.Write("<option value=""" & count & """ selected>" & count & "</option>")
                  Else
                    Response.Write("<option value=""" & count & """>" & count & "</option>")
                  End If
                Next
                Response.Write ("</select> / <select name=""s_year"" size=""1"">")
                year_adj = Year(now) + year_adj
                For count = 2000 to 2010
                  If count = year_adj Then
                    Response.Write("<option value=""" & count & """ selected>" & count & "</option>")
                  Else
                    Response.Write("<option value=""" & count & """>" & count & "</option>")
                  End If
                Next
                Response.Write ("</select><br />" & _
                 lang(cnnDB, "EndDate") & ":" & _
                 "<select name=""e_month"" size=""1"">")
                For count = 1 to 12
                  If count = Month(now) Then
                    Response.Write("<option value=""" & count & """ selected>" & count & "</option>")
                  Else
                    Response.Write("<option value=""" & count & """>" & count & "</option>")
                  End If
                Next
                Response.Write ("</select> / <select name=""e_day"" size=""1"">")
                For count = 1 to 31
                  If count = Day(now) Then
                    Response.Write("<option value=""" & count & """ selected>" & count & "</option>")
                  Else
                    Response.Write("<option value=""" & count & """>" & count & "</option>")
                  End If
                Next
                Response.Write ("</select> / <select name=""e_year"" size=""1"">")
                For count = 2000 to 2010
                  If count = Year(now) Then
                    Response.Write("<option value=""" & count & """ selected>" & count & "</option>")
                  Else
                    Response.Write("<option value=""" & count & """>" & count & "</option>")
                  End If
                Next
              %>
              </select>
              <br /><br />
              <center><input type="submit" value="<%=lang(cnnDB, "ViewReport")%>"></center>
            </form>
          </td>
        </tr>
      </table>
      <p><a href="default.asp"><%=lang(cnnDB, "AdministrativeMenu")%></a></p>
    </div>

    <%
      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close

    %>
  </body>
</html>
