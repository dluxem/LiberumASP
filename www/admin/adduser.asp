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

  Filename: adduser.asp
  Date:     $Date: 2002/06/15 23:49:20 $
  Version:  $Revision: 1.50.4.1 $
  Purpose:  Form to add new users.
  -->
  
  <!-- 	#include file = "../public.asp" -->

  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title><%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "AddNewUsers")%></title>
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>

    <%
      Call CheckAdmin


      Dim success
      success = FALSE

      If Request.Form("save") = 1 Then
        Dim uid, email, fname, phone, location, department, pager
        Dim firstname, lastname, ListOnInoutboard, phone_home, phone_mobile
        Dim jobfunction, userresume, statuscode, statustext, statusdate
        Dim usrLanguage, IsRep, RepAccess, intNewSid, InoutAdmin

        intNewSid = GetUnique(cnnDB, "users")
        
        uid = Left(Lcase(Trim(Request.Form("uid"))), 50)

        email = Left(Lcase(Trim(Request.Form("email"))), 50)
        email = Replace(email, "'", "''")

        pager = Left(Lcase(Trim(Request.Form("pager"))), 50)
        pager = Replace(pager, "'", "''")

        phone = Left(Trim(Replace(Request.Form("phone"), "'", "''")), 50)
        location = Left(Trim(Replace(Request.Form("location"), "'", "''")), 50)
        department = Cint(Request.Form("department"))

        ListOnInoutBoard = Cint(Request.Form("ListOnInoutBoard"))
        firstname = Left(Trim(Replace(Request.Form("firstname"), "'", "''")), 25)
        lastname = Left(Trim(Replace(Request.Form("lastname"), "'", "''")), 24)
        fname = firstname & " " & lastname

        phone_home = Left(Trim(Replace(Request.Form("phone_home"), "'", "''")), 50)
        phone_mobile = Left(Trim(Replace(Request.Form("phone_mobile"), "'", "''")), 50)
        jobfunction = Request.form("jobfunction")
        jobfunction = Replace(jobfunction , "'", "''")
        userresume = Request.form("userresume")
        userresume = Replace(userresume , "'", "''")
        statuscode = Cint(Request.Form("statuscode"))
        statustext = Request.form("statustext")
        statustext = Replace(statustext , "'", "''")
    		statusdate = SQLDate(Now, lhdAddSQLDelim)
    		usrLanguage = Cint(Request.form("usrLanguage"))
        RepAccess = Cint(Request.Form("repaccess"))

        If CBool(InStr(uid, "'")) Then
          Call DisplayError (3, lang(cnnDB, "Username") & "&nbsp;" & Lang(cnnDB, "containsinvalidcharacters") & ".")
        End If
        If Len(email) = 0 Then
          Call DisplayError (3, lang(cnnDB, "Emailaddress") & " " & lang(cnnDB, "isarequiredfield") & ".")
        End If
        If Len(firstname) = 0 Then
          Call DisplayError (3, lang(cnnDB, "FirstName") & " " & lang(cnnDB, "isarequiredfield") & ".")
        End IF
        If Len(lastname) = 0 Then
          Call DisplayError (3, lang(cnnDB, "LastName") & " " & lang(cnnDB, "isarequiredfield") & ".")
        End IF

        If Request.Form("InoutAdmin") = "on" Then
          InoutAdmin = 1
        Else
          InoutAdmin = 0
        End If

        Dim blnRepProbs
        blnRepProbs = False
        If Request.Form("isrep") = "on" Then
          IsRep = 1
        Else
          IsRep = 0
        End If
        
        Dim sqlString, updRes

        Dim newpassword
        newpassword = Left(Trim(Request.Form("newpassword")), 50)
        newpassword = Replace(newpassword, "'", "''")
        sqlString = "INSERT INTO tblUsers (sid, uid, email1, email2, fname, firstname, lastname, phone, " & _
          "phone_home, phone_mobile, location1, department, InoutAdmin, IsRep, RepAccess, statuscode, " & _
          "statustext, statusdate, jobfunction, userresume, ListOnInoutboard, [language], [password])" & _
          " VALUES (" & intNewSid & ", '" & uid & "', '" & email & "', '" & pager & "', '" & fname & "', '" & _
          firstname & "', '" & lastname & "', '" & phone & "', '" & phone_home & "', '" & _
          phone_mobile & "', '" & location & "', " & department & ", " & InoutAdmin & ", " & IsRep & ", " & _
          RepAccess & ", " & statuscode & ", '" & statustext & "', " & statusdate & ", '" & _
          jobfunction & "', '" & userresume & "', " & ListOnInoutboard & ", " & usrLanguage & ", '" & _
          newpassword & "')"

          Set updRes = SQLQuery(cnnDB, sqlString)
          success = True
      End If
    %>

    <div align="center">
      <table class="Normal">
        <tr class="Head1">
          <td>
            <%=lang(cnnDB, "AddNewUsers")%>
          </td>
        </tr>
        <% If success Then %>
          <tr class="Head2">
            <td>
              <div align="center">
                <%=lang(cnnDB, "AccountCreated")%>: '<% = uid %>'
              </div>
            </td>
          </tr>
        <% End If %>
        <tr class="Body1">
          <td>
            <form name="upduser" action="adduser.asp" method="POST">
              <input type="hidden" name="save" value="1">
              <p>
              <table class="Normal">
                <tr>
                  <td colspan="2">
                    <div align="right">
                      <i><em>*</em> = <%=lang(cnnDB, "Required")%></i>
                    </div>
                  </td>
                </tr>
                <tr>
                  <td width="150">
                    <b><%=lang(cnnDB, "Username")%>: </b>
                  </td>
                  <td>
                    <input type="text" name="uid" size="30"><em>*</em>
                  </td>
                </tr>
                <tr>
                  <td width="150">
                    <b><%=lang(cnnDB, "FirstName")%>: </b>
                  </td>
                  <td>
                    <input type="text" name="firstname" size="30"><em>*</em>
                  </td>
                </tr>
                <tr>
                  <td width="150">
                    <b><%=lang(cnnDB, "LastName")%>: </b>
                  </td>
                  <td>
                    <input type="text" name="lastname" size="30"><em>*</em>
                  </td>
                </tr>
                <tr>
                  <td width="150">
                    <b><%=lang(cnnDB, "EmailAddress")%>: </b>
                  </td>
                  <td>
                    <input type="text" name="email" size="30"><em>*</em>
                  </td>
                </tr>
                <% If (Cfg(cnnDB, "EnablePager") > 0) Then %>
                <tr>
                  <td width="150">
                    <b><%=lang(cnnDB, "PagerAddress")%>: </b>
                  </td>
                  <td>
                    <input type="text" name="pager" size="30">
                  </td>
                </tr>
                <% Else %>
                  <input type="hidden" name="pager" value="">
                <% End If %>
                <tr>
                  <td width="150">
                    <b><%=lang(cnnDB, "PhoneNumber")%>: </b>
                  </td>
                  <td>
                    <input type="text" name="phone" size="30">
                  </td>
                </tr>
                <% if cfg(cnnDB,"useinoutboard") = 1 then %>
                  <tr>
                    <td width="150">
                      <b><%=lang(cnnDB, "HomePhone")%>: </b>
                    </td>
                    <td>
                      <input type="text" name="phone_home" size="30">
                    </td>
                  </tr>
                  <tr>
                    <td width="150">
                      <b><%=lang(cnnDB, "MobilePhone")%>: </b>
                    </td>
                    <td>
                      <input type="text" name="phone_mobile" size="30">
                    </td>
                  </tr>
                <% Else %>
                  <input type="hidden" name="phone_home" value="">
                  <input type="hidden" name="phone_mobile" value="">
                <% End If %>
                <tr>
                  <td width="150">
                    <b><%=lang(cnnDB, "Location")%>: </b>
                  </td>
                  <td>
                    <input type="text" name="location" size="30">
                  </td>
                </tr>
                <tr>
                  <td width="150">
                    <b><%=lang(cnnDB, "Department")%>: </b>
                  </td>
                  <td>
                    <select name="department">
                      <option value="0"><%=lang(cnnDB, "NotSpecified")%></option>
                  <%

                      ' Get list of departments to diplay
                      Dim depRes
                      Set depRes = SQLQuery(cnnDB, "SELECT * From departments WHERE department_id > 0 ORDER BY dname ASC")
                      If not depRes.EOF Then
                        Do While Not depRes.EOF
                   %>
                          <option value="<% = depRes("department_id")%>">
                          <% = depRes("dname") %>
                          </option>

                  <%	    depRes.MoveNext
                        Loop
                        depRes.Close
                      End If
                  %>
                    </select>
                  </td>
                </tr>
                <tr>
                  <td width="150">
                    <b><%=lang(cnnDB, "Language")%>:</b>
                  </td>
                  <td>
                    <select name="usrLanguage">
                  <%
                      ' Get list of languages to diplay
                      Dim rstLang
                      Set rstLang = SQLQuery(cnnDB, "SELECT * From tblLanguage")
                      If not rstLang.EOF Then
                        Do While Not rstLang.EOF

                        If rstLang("id") = Cfg(cnnDB, "DefaultLanguage") Then
                  %>
                          <option value="<% = rstLang("id")%>" selected>
                          <% = rstLang("LangName") %> (<% = rstLang("Localized") %>)
                          </option>
                  <%			Else %>
                          <option value="<% = rstLang("id")%>">
                          <% = rstLang("LangName") %> (<% = rstLang("Localized") %>)
                          </option>

                  <%			End If

                        rstLang.MoveNext
                        Loop
                        rstLang.Close
                      End If
                  %>
                    </select>
                  </td>
                </tr>

                <tr class="Head2">
                  <td colspan="2">
                    <%=lang(cnnDB, "SupportRep")%>:
                  </td>
                </tr>
                <tr>
                  <td>
                    <b><%=lang(cnnDB, "Enable")%>:</b>
                  </td>
                  <td>
                    <input type="checkbox" name="isrep">
                  </td>
                </tr>
                <tr>
                  <td>
                    <b><%=lang(cnnDB, "AccessLevel")%>:</b>
                  </td>
                  <td>
                    <select name="repaccess">
                      <option value="0" selected><%=lang(cnnDB, "Normal")%></option>
                      <option value="1"><%=lang(cnnDB, "Restricted")%></option>
                      <option value="2"><%=lang(cnnDB, "ReadOnly")%></option>
                    </select>
                  </td>
                </tr>
                <% if cfg(cnnDB,"useinoutboard") = 1 then %>
                  <tr class="Head2">
                    <td colspan="2">
                      <%=lang(cnnDB, "InOutBoard")%>:
                    </td>
                  </tr>
                  <tr>
                    <td width="150">
                      <b><%=lang(cnnDB, "Status")%>: </b>
                    </td>
                    <td>
                      <select name="statuscode">
                        <option value="0" selected><%=lang(cnnDB, "In")%></option>
                        <option value="1"><%=lang(cnnDB, "Out")%></option>
                        <option value="2"><%=lang(cnnDB, "Leave")%></option>
                      </select>
                    </td>
                  </tr>
                  <tr>
                    <td width="150">
                      <b><%=lang(cnnDB, "Status")%>&nbsp;<%=lang(cnnDB, "text")%>: </b>
                    </td>
                    <td>
                      <input type="text" name="statustext" size="30">
                    </td>
                  </tr>
                  <tr>
                    <td width="150">
                      <b><%=lang(cnnDB, "JobFunction")%>: </b>
                    </td>
                    <td>
                      <textarea name="jobfunction" rows="4" cols="30"></textarea>
                    </td>
                  </tr>
                  <tr>
                    <td width="150">
                      <b><%=lang(cnnDB, "Resume")%>: </b>
                    </td>
                    <td>
                      <textarea name="userresume" rows="4" cols="30"></textarea>
                    </td>
                  </tr>
                  <tr>
                    <td width="150">
                      <b><%=lang(cnnDB, "ListonBoard")%>: </b>
                    </td>
                    <td>
                      <select name="listoninoutboard">
                        <option value="0"><%=lang(cnnDB, "NO")%></option>
                        <option value="1" selected><%=lang(cnnDB, "YES")%></option>
                      </select>
                    </td>
                  </tr>
                  <tr>
                    <td>
                      <b><%=lang(cnnDB, "Administrator")%>:</b>
                    </td>
                    <td>
                        <input type="checkbox" name="inoutadmin">
                    </td>
                  </tr>
                <% Else %>
                  <input type="hidden" name="statuscode" value="0">
                  <input type="hidden" name="listoninoutboard" value="0">
                  <input type="hidden" name="statustext" value="">
                  <input type="hidden" name="jobfunction" value="">
                  <input type="hidden" name="userresume" value="">
                  <input type="hidden" name="inoutadmin" value="0">
                <% End If %>
                <tr class="Head2">
                  <td colspan="2">
                    <%=lang(cnnDB, "Password")%>:
                  </td>
                </tr>
                <tr>
                  <td width="150">
                    <b><%=lang(cnnDB, "NewPassword")%>: </b>
                  </td>
                  <td>
                    <input type="password" name="newpassword" size="30"><em>*</em>
                  </td>
                </tr>
              </table>
              <p>
              <div align="center">
                <p>
                <input type="submit" value="<%=lang(cnnDB, "Save")%>">
              </div>
            </form>
          </td>
        </tr>
      </table>
      <p>
      <a href="viewusers.asp"><%=lang(cnnDB, "ManageUsers")%></a><br>
      <a href="default.asp"><%=lang(cnnDB, "AdministrativeMenu")%></a>
    </div>
    <%
      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
