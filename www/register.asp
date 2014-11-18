<%@ LANGUAGE="VBScript" %>
<% 
  Option Explicit
  'Buffer the response, so Response.Expires can be used
  Response.Buffer = TRUE
  Response.Expires = -1
%>

<?xml version="1.0"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

  <!--
  Liberum Help Desk, Copyright (C) 2000-2001 Doug Luxem
  Liberum Help Desk comes with ABSOLUTELY NO WARRANTY
  Please view the license.html file for the full GNU General Public License.

  Filename: register.asp
  Date:     $Date: 2002/01/21 20:35:22 $
  Version:  $Revision: 1.50.2.1 $
  Purpose:  Page for registering, editing user information and changing passwords.

  -->
  <!-- 	#include file = "public.asp" -->
  <%

    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>
  <head>
    <title>
      <%=lang(cnnDB, "HelpDesk")%> &nbsp;-&nbsp; <%=lang(cnnDB, "Register")%>
    </title>
    <link rel="stylesheet" type="text/css" href="default.css">  
  </head>
  <body>

  <%
    Dim success, edit
    success = FALSE
    If (Request.QueryString("edit") = 1) OR (Request.Form("edit") = 1) Then
      edit = True
    Else
      edit = False
    End If

    If Request.Form("create") = 1 Then
      Dim uid, email, password1, password2, fname, phone, location, department, pager
      Dim firstname, lastname, phone_home, phone_mobile, usrLanguage
      uid = Left(Lcase(Trim(Request.Form("uid"))), 50)

      email = Left(Lcase(Trim(Request.Form("email"))), 50)
      email = Replace(email, "'", "''")

      pager = Left(Lcase(Trim(Request.Form("pager"))), 50)
      pager = Replace(pager, "'", "''")

      password1 = Left(Trim(Replace(Request.Form("password1"), "'", "''")), 50)
      password2 = Left(Trim(Replace(Request.Form("password2"), "'", "''")), 50)

      phone = Left(Trim(Replace(Request.Form("phone"), "'", "''")), 50)
      
      location = Left(Trim(Replace(Request.Form("location"), "'", "''")), 50)
      department = Cint(Request.Form("department"))
      usrLanguage = Cint(Request.Form("usrLanguage"))

      firstname = Left(Trim(Replace(Request.Form("firstname"), "'", "''")), 25)
      lastname = Left(Trim(Replace(Request.Form("lastname"), "'", "''")), 24)
      fname = firstname & " " & lastname
      if cfg(cnnDB,"useinoutboard") = 1 then
        phone_home = Left(Trim(Replace(Request.Form("phone_home"), "'", "''")), 50)
        phone_mobile = Left(Trim(Replace(Request.Form("phone_mobile"), "'", "''")), 50)
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

      Dim sqlString

      If Edit Then
        Dim oldpassword
        oldpassword = Left(Trim(Request.Form("oldpassword")), 50)
        If (Len(oldpassword) > 0) or (Len(password1) > 0) Then
          If password1 <> password2 Then
            Call DisplayError (3, lang(cnnDB, "Passwordsdonotmatch") & ".")
          End If
          If oldpassword <> Usr(cnnDB, sid, "password") Then
            Call DisplayError (3, lang(cnnDB, "Passwordisincorrect") & ".")
          End If
          sqlString = "UPDATE tblUsers SET " & _
            "email1 = '" & email & "', " & _
            "email2 = '" & pager & "', " & _
            "fname = '" & fname & "', " & _
            "firstname = '" & firstname & "', " & _
            "lastname = '" & lastname & "', " & _
            "phone = '" & phone & "', " & _
            "phone_home = '" & phone_home & "', " & _
            "phone_mobile = '" & phone_mobile & "', " & _
            "location1 = '" & location & "', " & _
            "department = " & department & ", " & _
            "[language] = " & usrLanguage & ", " & _
            "[password] = '" & password1 & "'"
        Else
          sqlString = "UPDATE tblUsers SET " & _
            "email1 = '" & email & "', " & _
            "email2 = '" & pager & "', " & _
            "fname = '" & fname & "', " & _
            "firstname = '" & firstname & "', " & _
            "lastname = '" & lastname & "', " & _
            "phone = '" & phone & "', " & _
            "phone_home = '" & phone_home & "', " & _
            "phone_mobile = '" & phone_mobile & "', " & _
            "location1 = '" & location & "', " & _
            "[language] = " & usrLanguage & ", " & _
            "department = " & department
        End If

        sqlString = sqlString & " WHERE sid=" & sid

        Dim updRes, checkNameRes
        Set updRes = SQLQuery(cnnDB, sqlString)
        success = True

        ' Remove Language ID
        Session("lhd_LanguageID") = Empty
      Else
        If Len(uid) = 0 Then
          Call DisplayError (3, lang(cnnDB, "Username") & "&nbsp;" & lang(cnnDB, "isarequiredfield") & ".")
        End If
        If CBool(InStr(uid, "'")) Then
          Call DisplayError (3, lang(cnnDB, "Username") & "&nbsp;" & Lang(cnnDB, "containsinvalidcharacters") & ".")
        End If
        If password1 <> password2 Then
          Call DisplayError (3, lang(cnnDB, "Passwordsdonotmatch") & ".")
        End If

        If Len(password1) = 0 Then
          Call DisplayError (3, lang(cnnDB, "Password") &"&nbsp;" & lang(cnnDB, "isarequiredfield") & ".")
        End IF

        Set checkNameRes = SQLQuery(cnnDB, "SELECT uid FROM tblUsers WHERE uid='" & uid & "'")
        If Not checkNameRes.EOF Then
          checkNameRes.Close
          Call DisplayError (3, lang(cnnDB, "Username") & "&nbsp;" & lang(cnnDB, "alreadyinuse") & ".")
        Else
          checkNameRes.Close
        End If

        Dim newSid
        newSid = GetUnique(cnnDB, "users")

        sqlString = "INSERT INTO tblUsers " & _
          "(sid, uid, [password], email1, email2, fname, firstname, lastname, phone, phone_home, phone_mobile, location1, [language], department) VALUES (" & _
          newSid & ", " & _
          "'" & uid & "', " & _
          "'" & password1 & "', " & _
          "'" & email & "', " & _
          "'" & pager & "', " & _
          "'" & fname & "', " & _
          "'" & firstname & "', " & _
          "'" & lastname & "', " & _
          "'" & phone & "', " & _
          "'" & phone_home & "', " & _
          "'" & phone_mobile & "', " & _
          "'" & location & "', " & _
          usrLanguage & ", " & _
          department & _
          ")"

        Dim insertRes
        Set insertRes = SQLQuery(cnnDB, sqlString)
        success = TRUE

      End If
    End If

    Dim frm_email, frm_fname, frm_phone, frm_location, frm_department, frm_pager
    Dim frm_firstname, frm_lastname, frm_phone_home, frm_phone_mobile, frm_usrLanguage
    If Edit Then
      frm_email = Usr(cnnDB, sid, "email1")
      frm_phone = Usr(cnnDB, sid, "phone")
      frm_location = Usr(cnnDB, sid, "location1")
      frm_department = Usr(cnnDB, sid, "department")
      frm_pager = Usr(cnnDB, sid, "email2")
      frm_firstname = Usr(cnnDB, sid, "firstname")
      frm_lastname = Usr(cnnDB, sid, "lastname")
      frm_usrLanguage = Usr(cnnDB, sid, "[language]")
      if cfg(cnnDB,"useinoutboard") = 1 then
        frm_phone_home = Usr(cnnDB, sid, "phone_home")
        frm_phone_mobile = Usr(cnnDB, sid, "phone_mobile")
      End If
    End If
  %>

  <div align="center">
    <table Class="Normal">
      <tr Class="Head1">
        <td>
       <% If Edit Then 
            Response.Write lang(cnnDB, "UpdateInformation")
           Else 
            Response.Write lang(cnnDB, "Registration") & "<br>" & _
            lang(cnnDB, "fornewusers")
           End If %>
        </td>
      </tr>
      <tr class="body1">
        <td>
          <% If success Then
              If Edit Then
          %>

              <div align="center">
                <u><h3><%=lang(cnnDB, "AccountUpdated")%></h3></u>
                <p><%=lang(cnnDB, "accountupdatedtext")%>.</p>
                <p><b><a href="default.asp"><%=lang(cnnDB, "MainMenu")%></a></b></p>
              </div>
          <%
              Else
          %>
              <div align="center">
                <u><h3><%=lang(cnnDB, "AccountCreated")%></h3></u>
                <p><%=lang(cnnDB, "AccountCreatedText")%></p>
                <p><b><a href="logon.asp"><%=lang(cnnDB, "HelpDesk")%>&nbsp;<%=lang(cnnDB, "Logon")%></a></b></p>
              </div>
          <%
              End If
          Else		' New user form%>

             <form action="register.asp" method="POST">
              <%	If Edit Then %>
                <input type="hidden" name="edit" value="1">
                <input type="hidden" name="create" value="1">
              <% 	Else %>
                <input type="hidden" name="create" value="1">
                <div align="center">
                  <u><h3><%=lang(cnnDB, "Register")%></h3></u>
                </div>
              <% End If %>

              <p>&nbsp;</p>
              <table class="normal">
              <tr>
                <td width="150">
                  <b><%=lang(cnnDB, "Username")%>: </b>
                </td>
                <td>
                  <% If Edit Then %>
                    <b><% = Usr(cnnDB, sid, "uid") %></b>
                  <% Else %>
                    <input type="text" name="uid" size="30"><em>*</em>
                  <% End If %>
                </td>
              </tr>
              <tr>
                <td width="150">
                  <b><%=lang(cnnDB, "FirstName")%>: </b>
                </td>
                <td>
                  <input type="text" name="firstname" size="30" value="<% = frm_firstname %>"><em>*</em>
                </td>
              </tr>
              <tr>
                <td width="150">
                  <b><%=lang(cnnDB, "LastName")%>: </b>
                </td>
                <td>
                  <input type="text" name="lastname" size="30" value="<% = frm_lastname %>"><em>*</em>
                </td>
              </tr>
              <tr>
                <td width="150">
                  <b><%=lang(cnnDB, "Emailaddress")%>: </b>
                </td>
                <td>
                  <input type="text" name="email" size="30" value="<% = frm_email %>"><em>*</em>
                </td>
              </tr>
              <% If (Cfg(cnnDB, "EnablePager") > 0) And (Usr(cnnDB, sid, "IsRep") = 1) Then %>
                <tr>
                  <td width="150">
                    <b><%=lang(cnnDB, "PagerAddress")%>: </b>
                  </td>
                  <td>
                    <input type="text" name="pager" size="30" value="<% = frm_pager %>">
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
                  <input type="text" name="phone" size="30" value="<% = frm_phone %>">
                </td>
              </tr>
              <% if cfg(cnnDB, "useinoutboard") = 1 then %>
                <tr>
                  <td width="150">
                    <b><%=lang(cnnDB, "HomePhone")%>: </b>
                  </td>
                  <td>
                    <input type="text" name="phone_home" size="30" value="<% = frm_phone_home %>">
                  </td>
                </tr>
                <tr>
                  <td width="150">
                    <b><%=lang(cnnDB, "MobilePhone")%>: </b>
                  </td>
                  <td>
                    <input type="text" name="phone_mobile" size="30" value="<% = frm_phone_mobile %>">
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
                  <input type="text" name="location" size="30" value="<% = frm_location %>">
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

                      If depRes("department_id") = frm_department Then
                %>
                        <option value="<% = depRes("department_id")%>" selected>
                        <% = depRes("dname") %>
                        </option>
                <%			Else %>
                        <option value="<% = depRes("department_id")%>">
                        <% = depRes("dname") %>
                        </option>

                <%			End If

                      depRes.MoveNext
                      Loop
                      depRes.Close
                    End If
                %>
                  </select>
                </td>
                </tr>
                    <tr>
                      <td width="150">
                        <b><%=lang(cnnDB, "Language")%>: </b>
                      </td>
                      <td>
                        <select name="usrLanguage">
                      <%
                          ' Get list of languages to diplay
                          Dim rstLang
                          Set rstLang = SQLQuery(cnnDB, "SELECT * From tblLanguage")
                          If not rstLang.EOF Then
                            Do While Not rstLang.EOF
  
                            If rstLang("id") = frm_usrLanguage Then
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
              </table>
            <% If Cfg(cnnDB, "AuthType") = 2 Then %>
              <hr width="80%">
              <table class="normal">
              <% If Edit Then %>
              <tr>
                <td width="150">
                  <b><%=lang(cnnDB, "OldPassword")%>: </b>
                </td>
                <td>
                  <input type="password" name="oldpassword" size="30"><em>*</em>
                </td>
              </tr>
              <% End If %>
              <tr>
                <td width="150">
                  <b><%=lang(cnnDB, "Password")%>: </b>
                </td>
                <td>
                  <input type="password" name="password1" size="30"><em>*</em>
                </td>
              </tr>
              <tr>
                <td width="150">
                  <b><%=lang(cnnDB, "ConfirmPassword")%>: </b>
                </td>
                <td>
                  <input type="password" name="password2" size="30"><em>*</em>
                </td>
              </tr>
              </table>
            <% Else %>
              <input type="hidden" name="oldpassword" value="">
              <input type="hidden" name="password1" value="">
              <input type="hidden" name="password2" value="">
            <% End If %>
              <p><div align="center">
                <i><em>*</em> = <%=lang(cnnDB, "Required")%></i>
                <p><input type="submit" value="<%=lang(cnnDB, "Submit")%>"></p>
              </div></p>
              </form>
          <% End If %>
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
