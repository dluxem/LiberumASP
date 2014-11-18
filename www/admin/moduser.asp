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

  Filename: moduser.asp
  Date:     $Date: 2002/08/28 15:30:07 $
  Version:  $Revision: 1.52.4.2 $
  Purpose:  Form to modify user account info.
  -->
  
  <!-- 	#include file = "../public.asp" -->

  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title><%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "ModifyUser")%></title>
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>

    <%
      Call CheckAdmin

      Dim modSid
      modSid = Cint(Request.Form("usersid"))

      Dim success
      success = FALSE

      If Request.Form("save") = 1 Then
        Dim uid, email, fname, phone, location, department, pager
        Dim firstname, lastname, ListOnInoutboard, phone_home, phone_mobile
        Dim jobfunction, userresume, statuscode, statustext, statusdate
        Dim usrLanguage, IsRep, RepAccess, InoutAdmin
        
        uid = Left(Lcase(Trim(Request.Form("uid"))), 50)
        uid = Replace(uid, "'", "''")

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

        If Request.Form("InoutAdmin") = "on" Then
          InoutAdmin = 1
        Else
          InoutAdmin = 0
        End If
        
        If Len(email) = 0 Then
          Call DisplayError (3, lang(cnnDB, "Emailaddress") & lang(cnnDB, "is a required field") & ".")
        End If
        'Changed into computed field
        'If Len(fname) = 0 Then
        '  Call DisplayError (3, "Full Name is a required field.")
        'End IF
        If Len(firstname) = 0 Then
          Call DisplayError (3, lang(cnnDB, "FirstName") & lang(cnnDB, "is a required field") & ".")
        End IF
        If Len(lastname) = 0 Then
          Call DisplayError (3, lang(cnnDB, "LastName") & lang(cnnDB, "is a required field") & ".")
        End IF

        Dim blnRepProbs
        blnRepProbs = False
        If Request.Form("isrep") = "on" Then
          IsRep = 1
        Else
          If Usr(cnnDB, modSid, "IsRep") = 1 Then
            Dim rstRepProbs
            Set rstRepProbs = SQLQuery(cnnDB, "SELECT id FROM problems WHERE rep=" & modSid & " AND status<>" & Cfg(cnnDB, "CloseStatus"))
            If Not rstRepProbs.EOF Then
              blnRepProbs = True
              IsRep = 1
            Else
              IsRep = 0
            End If
            rstRepProbs.Close
          Else
            IsRep = 0
          End If
        End If
        
        Dim sqlString, updRes

        Dim newpassword
        newpassword = Left(Trim(Request.Form("newpassword")), 50)
        newpassword = Replace(newpassword, "'", "''")
        If Len(newpassword) > 0 Then
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
            "IsRep = " & IsRep & ", " & _
            "RepAccess = " & RepAccess & ", " & _
            "InoutAdmin = " & InoutAdmin & ", " & _
            "statuscode = " & statuscode & ", " & _
            "statustext = '" & statustext & "', " & _
            "statusdate = " & statusdate & ", " & _
            "jobfunction = '" & jobfunction & "', " & _
            "userresume = '" & userresume & "', " & _
            "ListOnInoutboard = " & ListOnInoutboard & ", " & _ 
            "[Language] = " & usrLanguage & ", " & _ 
            "[password] = '" & newpassword & "'"
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
            "IsRep = " & IsRep & ", " & _
            "RepAccess = " & RepAccess & ", " & _
            "InoutAdmin = " & InoutAdmin & ", " & _
            "statuscode = " & statuscode & ", " & _
            "statustext = '" & statustext & "', " & _
            "statusdate = " & statusdate & ", " & _
            "jobfunction = '" & jobfunction & "', " & _
            "userresume = '" & userresume & "', " & _
            "ListOnInoutboard = " & ListOnInoutboard & ", " & _ 
            "[Language] = " & usrLanguage & ", " & _ 
            "department = " & department
        End If

        sqlString = sqlString & " WHERE sid=" & modSid
        Set updRes = SQLQuery(cnnDB, sqlString)
        success = True


      End If

      If Request.Form("delete") = 1 Then
        Dim delRes, rstProblemUpd1, rstProblemUpd2, strUserId
        strUserId = usr(cnnDB, modSid, "uid")
        If Usr(cnnDB, modSid, "IsRep") = 0 Then
          Set delRes = SQLQuery(cnnDB, "DELETE FROM tblUsers WHERE sid = " & modSid)
          Set rstProblemUpd1 = SQLQuery(cnnDB, "UPDATE problems SET entered_by=0 WHERE entered_by = " & modSid)
          Set rstProblemUpd2 = SQLQuery(cnnDB, "UPDATE problems SET rep=0 WHERE rep = " & modSid)
          success = True
        Else
          Dim rstRepCats
          Set rstRepCats = SQLQuery(cnnDB, "SELECT category_id FROM categories WHERE rep_id=" & modSid)
          If Not rstRepCats.EOF Then
            Call DisplayError(3, "Please reassign categories to a different support rep.")
          End If
          rstRepCats.Close
          Set rstRepProbs = SQLQuery(cnnDB, "SELECT id FROM problems WHERE rep=" & modSid & " AND status<>" & Cfg(cnnDB, "CloseStatus"))
          If Not rstRepProbs.EOF Then
            blnRepProbs = True
          Else
            Set delRes = SQLQuery(cnnDB, "DELETE FROM tblUsers WHERE sid = " & modSid)
            Set rstProblemUpd1 = SQLQuery(cnnDB, "UPDATE problems SET entered_by=0 WHERE entered_by = " & modSid)
            success = True
          End If
          rstRepProbs.Close
        End If
        'Code to delete user image file if In/Out Board is activated
        If cfg(cnnDB, "UseInoutBoard") = 1 Then
          If success = True Then
            Dim objFSO, strUserImage
            strUserImage = Server.MapPath("..\image\" & strUserId & ".jpg")
            Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
            If objFSO.FileExists(strUserImage) Then
              objFSO.DeleteFile strUserImage, False
            End If
            Set objFSO = Nothing
          End If
        End If
      Else
        Dim frm_email, frm_fname, frm_phone, frm_location, frm_department, frm_pager
        Dim frm_firstname, frm_lastname, frm_phone_home, frm_phone_mobile
        Dim frm_jobfunction, frm_userresume, frm_listoninoutboard, frm_usrLanguage
        Dim frm_statuscode, frm_statustext, frm_IsRep, frm_RepAccess, frm_InoutAdmin
        frm_email = Usr(cnnDB, modSid, "email1")
        frm_firstname = Usr(cnnDB, modSid, "firstname")
        frm_lastname = Usr(cnnDB, modSid, "lastname")
        frm_phone = Usr(cnnDB, modSid, "phone")
        frm_phone_home = Usr(cnnDB, modSid, "phone_home")
        frm_phone_mobile = Usr(cnnDB, modSid, "phone_mobile")
        frm_location = Usr(cnnDB, modSid, "location1")
        frm_department = Usr(cnnDB, modSid, "department")
        frm_pager = Usr(cnnDB, modSid, "email2")
        frm_statuscode = Usr(cnnDB, modSid, "statuscode")
        frm_statustext = Usr(cnnDB, modSid, "statustext")
        frm_jobfunction = Usr(cnnDB, modSid, "jobfunction")
        frm_userresume = Usr(cnnDB, modSid, "userresume")
        frm_listoninoutboard = Usr(cnnDB, modSid, "listoninoutboard")
        frm_usrLanguage = Usr(cnnDB, modSid, "[language]")
        frm_IsRep = Usr(cnnDB, modSid, "IsRep")
        frm_RepAccess = Usr(cnnDB, modSid, "RepAccess")
        frm_InoutAdmin = Usr(cnnDB, modSid, "InoutAdmin")
      End If

    %>

    <div align="center">
      <table class="Normal">
        <tr class="Head1">
          <td>
            <%=lang(cnnDB, "UpdateInformation")%>
          </td>
        </tr>
        <% If blnRepProbs Then %>
          <tr class="Head2">
            <td>
              <div align="center">
                <%=lang(cnnDB, "ErrorUpdatingAccount")%>
              </div>
            </td>
          </tr>
          <tr class="Body1">
            <td>
              <%=lang(cnnDB, "ErrorUpdatingAccountText")%>.
            </td>
          </tr>
        <% ElseIf Request.Form("delete") = 1 Then %>
          <tr class="Head2">
            <td>
              <div align="center">
                <%=lang(cnnDB, "AccountDeleted")%>
              </div>
            </td>
          </tr>
          <tr class="Body1">
            <td>
              <div align="center">
                <%=lang(cnnDB, "Theaccounthasbeenremoved")%>. 
              </div>
            </td>
          </tr>
        <% Else ' user form %>
          <% If success Then %>
            <tr class="Head2">
              <td>
                <div align="center">
                  <%=lang(cnnDB, "AccountUpdated")%>
                </div>
              </td>
            </tr>
          <% End If %>
          <tr class="Body1">
            <td>
              <form name="upduser" action="moduser.asp" method="POST">
                <input type="hidden" name="usersid" value="<% = modSid %>">
                <input type="hidden" name="save" value="1">
                <p>
                <table class="Normal">
                  <tr>
                    <td colspan="2">
                      <div align="right"><i><em>*</em> = <%=lang(cnnDB, "Required")%></i></div>
                    </td>
                  </tr>
                  <tr>
                    <td width="150">
                      <b><%=lang(cnnDB, "Username")%>: </b>
                    </td>
                    <td>
                      <b><% = Usr(cnnDB, modSid, "uid") %></b>
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
                      <b><%=lang(cnnDB, "EmailAddress")%>: </b>
                    </td>
                    <td>
                      <input type="text" name="email" size="30" value="<% = frm_email %>"><em>*</em>
                    </td>
                  </tr>
                  <% If (Cfg(cnnDB, "EnablePager") > 0) And (Usr(cnnDB, modSid, "IsRep") = 1) Then %>
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
                  <% if cfg(cnnDB,"useinoutboard") = 1 then %>
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
                      <% If frm_IsRep = 0 Then %>
                        <input type="checkbox" name="isrep">
                      <% Else %>
                        <input type="checkbox" name="isrep" checked>
                      <% End If %>
                    </td>
                  </tr>
                  <tr>
                    <td>
                      <b><%=lang(cnnDB, "AccessLevel")%>:</b>
                    </td>
                    <td>
                      <select name="repaccess">
                        <% Select Case frm_RepAccess
                            Case 0
                              Response.Write("<option value=""0"" selected>" & lang(cnnDB, "Normal") & "</option>")
                              Response.Write("<option value=""1"">" & lang(cnnDB, "Restricted") & "</option>")
                              Response.Write("<option value=""2"">" & lang(cnnDB, "ReadOnly") & "</option>")
                            Case 1
                              Response.Write("<option value=""0"">" & lang(cnnDB, "Normal") & "</option>")
                              Response.Write("<option value=""1"" selected>" & lang(cnnDB, "Restricted") & "</option>")
                              Response.Write("<option value=""2"">" & lang(cnnDB, "ReadOnly") & "</option>")
                            Case 2
                              Response.Write("<option value=""0"">" & lang(cnnDB, "Normal") & "</option>")
                              Response.Write("<option value=""1"">" & lang(cnnDB, "Restricted") & "</option>")
                              Response.Write("<option value=""2"" selected>" & lang(cnnDB, "ReadOnly") & "</option>")
                            Case Else
                              Response.Write("<option value=""0"">" & lang(cnnDB, "Normal") & "</option>")
                              Response.Write("<option value=""1"">" & lang(cnnDB, "Restricted") & "</option>")
                              Response.Write("<option value=""2"">" & lang(cnnDB, "ReadOnly") & "</option>")
                           End Select
                         %>
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
                        <% If frm_statuscode = "2" Then %>
                          <option value="0"><%=lang(cnnDB, "In")%></option>
                          <option value="1"><%=lang(cnnDB, "Out")%></option>
                          <option value="2" selected><%=lang(cnnDB, "Leave")%></option>
                        <% Else if frm_statuscode = "1" Then%>
                          <option value="0"><%=lang(cnnDB, "In")%></option>
                          <option value="1" selected><%=lang(cnnDB, "Out")%></option>
                          <option value="2"><%=lang(cnnDB, "Leave")%></option>
                        <% Else %>
                          <option value="0" selected><%=lang(cnnDB, "In")%></option>
                          <option value="1"><%=lang(cnnDB, "Out")%></option>
                          <option value="2"><%=lang(cnnDB, "Leave")%></option>
                        <% End If 
                          End If
                          %>
                        </select>
                      </td>
                    </tr>
                    <tr>
                      <td width="150">
                        <b><%=lang(cnnDB, "Status")%>&nbsp;<%=lang(cnnDB, "text")%>: </b>
                      </td>
                      <td>
                        <input type="text" name="statustext" size="30" value="<% = frm_statustext %>">
                      </td>
                    </tr>
                    <tr valign="top">
                      <td width="150">
                        <b><%=lang(cnnDB, "JobFunction")%>: </b>
                      </td>
                      <td>
                        <textarea name="jobfunction" rows="4" cols="30"><% = frm_jobfunction %></textarea>
                      </td>
                    </tr>
                    <tr valign="top">
                      <td width="150">
                        <b><%=lang(cnnDB, "Resume")%>: </b>
                      </td>
                      <td>
                        <textarea name="userresume" rows="4" cols="30"><% = frm_userresume %></textarea>
                      </td>
                    </tr>
                    <tr>
                      <td width="150">
                        <b><%=lang(cnnDB, "ListonBoard")%>: </b>
                      </td>
                      <td>
                        <select name="listoninoutboard">
                        <% If frm_listoninoutboard = "0" Then %>
                          <option value="0" selected><%=lang(cnnDB, "NO")%></option>
                          <option value="1"><%=lang(cnnDB, "YES")%></option>
                        <% Else %>
                          <option value="0"><%=lang(cnnDB, "NO")%></option>
                          <option value="1" selected><%=lang(cnnDB, "YES")%></option>
                        <% End If %>
                        </select>
                      </td>
                    </tr>
                    <tr>
                      <td>
                        <b><%=lang(cnnDB, "Administrator")%>:</b>
                      </td>
                      <td>
                        <% If frm_InoutAdmin = 0 Then %>
                          <input type="checkbox" name="inoutadmin">
                        <% Else %>
                          <input type="checkbox" name="inoutadmin" checked>
                        <% End If %>
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
                      <input type="password" name="newpassword" size="30">
                    </td>
                  </tr>
                </table>
                <p>
                <div align="center">
                  <p>
                  <input type="submit" value="<%=lang(cnnDB, "Save")%>">
                </div>
              </form>
              <hr width="80%">
              <p>
              <div align="center">
                <form name="deluser" action="moduser.asp" method="POST">
                  <input type="hidden" name="usersid" value="<% = modSid %>">
                  <input type="hidden" name="delete" value="1">
                  <input type="submit" value="<%=lang(cnnDB, "DeleteAccount")%>">
                </form>
              </div>
            </td>
          </tr>
        <% End If %>
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
