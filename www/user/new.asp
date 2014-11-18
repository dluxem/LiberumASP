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

  Filename: new.asp
  Date:     $Date: 2001/12/09 02:01:24 $
  Version:  $Revision: 1.50 $
  Purpose:  A form for users to enter new problems.
  -->

  <!-- 	#include file = "../public.asp" -->
  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid

    Session.timeout = 40
  %>

  <head>
    <title><%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "NewProblem")%></title>
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>

    <%
      ' See if authenticated
      Call CheckUser(cnnDB, sid)

      ' Determine the username and email address for the
      ' user
      dim uid, uemail
      uid = Usr(cnnDB, sid, "uid")
      uemail = Usr(cnnDB, sid, "email1")
    %>

    <form action="postnew.asp" method="POST">
      <input type="hidden" name="uid" value="<% = uid %>">
      <div align="center">
        <table class="Wide">
          <tr>
            <td colspan="2">
              <div align="right">
                <em>*</em> - <%=lang(cnnDB, "Required")%>
              </div>
            </td>
          </tr>
          <tr class="Head1">
            <td colspan="2">
              <%=lang(cnnDB, "SubmitANewProblem")%>
            </td>
          </tr>
          <tr Class="Body1">
            <td valign="top">
              <div align="center">
                <table class="narrow" border="0">
                  <tr>
                    <td align="center" colspan="2">
                      <b><%=lang(cnnDB, "ContactInformation")%></b>
                    </td>
                  </tr>
                  <tr>
                    <td>
                      <b><%=lang(cnnDB, "UserName")%>:</b>
                    </td>
                    <td>
                      <% = uid %>
                    </td>
                  </tr>
                  <tr>
                    <td>
                      <b><%=lang(cnnDB, "EMail")%>:</b>
                    </td>
                    <td>
                      <input type="text" name="uemail" size="20" value="<% = uemail %>"><em>*</em>
                    </td>
                  </tr>
                  <tr>
                    <td>
                      <b><%=lang(cnnDB, "Location")%>:</b>
                    </td>
                    <td>
                      <input type="text" name="ulocation" size="20" value="<% = Usr(cnnDB, sid, "location1") %>">
                    </td>
                  </tr>
                  <tr>
                    <td>
                      <b><%=lang(cnnDB, "Phone")%>:</b>
                    </td>
                    <td>
                      <input type="text" name="uphone" size="20" value="<% = Usr(cnnDB, sid, "phone") %>">
                    </td>
                  </tr>
                </table>
              </div>
            </td>
            <td valign="top">
              <div align="center">
                <table class="narrow" border="0">
                  <tr>
                    <td colspan="2">
                      <div align="center">
                        <b><%=lang(cnnDB, "ProblemClassification")%></b>
                      </div>
                    </td>
                  </tr>
                  <tr>
                    <td>
                      <b><%=lang(cnnDB, "Department")%>:</b>
                    </td>
                    <td>
                      <SELECT NAME="department">
                      <OPTION VALUE="0" >Select Department</OPTION>
                      <%

                        ' Get list of departments to diplay
                        Dim rstDepList
                        Set rstDepList = SQLQuery(cnnDB, "SELECT * From departments WHERE department_id > 0 ORDER BY dname ASC")
                        If not rstDepList.EOF Then
                        Do While Not rstDepList.EOF

                        If rstDepList("department_id") = Usr(cnnDB, sid, "department") Then
                        %>
                          <option value="<% = rstDepList("department_id")%>" selected>
                          <% = rstDepList("dname") %></OPTION>
                        <%
                        Else
                        %>
                          <OPTION VALUE="<% = rstDepList("department_id")%>">
                          <% = rstDepList("dname") %></OPTION>

                      <%
                        End IF
                        rstDepList.MoveNext
                        Loop
                        End If
                      %>
                      </SELECT><em>*</em>
                    </td>
                  </tr>
                  <tr>
                    <td>
                      <b><%=lang(cnnDB, "Category")%>:</b>
                    </td>
                    <td>
                      <SELECT NAME="category">
                      <OPTION VALUE="0" SELECTED><%=lang(cnnDB, "SelectCategory")%></OPTION>
                      <%
                        ' Get list of categories to display
                        Dim rstCatList
                        Set rstCatList = SQLQuery(cnnDB, "SELECT * From categories WHERE category_id > 0 ORDER BY cname ASC")
                        If Not rstCatList.EOF Then
                        Do While Not rstCatList.EOF
                      %>
                      <OPTION VALUE="<% = rstCatList("category_id")%>">
                      <% = rstCatList("cname") %></OPTION>

                      <% 		rstCatList.MoveNext
                        Loop
                        End If
                      %>
                      </SELECT><em>*</em>
                    </td>
                  </tr>
                </table>
              </div>
            </td>
          </tr>
          <tr class="Head2">
            <td colspan="2">
              <%=lang(cnnDB, "ProblemInformation")%>:
            </td>
          </tr>
          <tr class="Body1">
            <td colspan="2">
              <b><%=lang(cnnDB, "Title")%>:</b> <em>*</em><br>
              <input type="text" name="title" size="50">
              <p>
              <b><%=lang(cnnDB, "Description")%>:</b> <em>*</em><br />
              <textarea rows="12" cols="80" name="description"></textarea>
            </td>
          </tr>
          <tr class="Head2">
            <td colspan="2">
              <div align="center">
                <input type="submit" value="<%=lang(cnnDB, "SubmitProblem")%>" name="B1">&nbsp;<input type="reset" value="<%=lang(cnnDB, "ClearForm")%>" name="B2">
              </div>
            </td>
          </tr>
        </table>
      </div>
    </form>

    <%
      ' close record sets
      rstCatList.Close
      rstDepList.Close

      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>