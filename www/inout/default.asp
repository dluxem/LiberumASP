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
  Purpose:  This is the main menu for the In/Out Board.
  -->

  <!-- 	#include file = "../public.asp" -->
  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title>
      <% = Cfg(cnnDB, "SiteName") %>&nbsp;<%=lang(cnnDB, "InOutBoard")%>
    </title>
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>
  <%
    ' See is user is validated
    Call CheckUser(cnnDB, sid)

  	'Check if user is In/Out Board admin
    dim inoutadmin
    inoutadmin = Usr(cnnDB, sid, "inoutadmin")

    'Call DisplayHeader(cnnDB, sid)
  %>

  <div align="center">
  <table class="wide">
      <tr class="Head1">
    <td>
	    <% = Cfg(cnnDB, "SiteName") %>
      <br />
      <%=lang(cnnDB, "InOutBoard")%>
    </td>
  </tr>
  </table>

  <%
    Dim lhd_row_back, lhd_row_back_0, lhd_row_back_1, counter, lhd_row
    Dim orderBy, strSql, sqlWhere, srtOrder
    Dim mFirstname, mLastname, mUid, mDept, mInoutStatus, mPhone, check
    
    lhd_row_back_0 = "Body1"
    lhd_row_back_1 = "Body2"
    lhd_row_back = lhd_row_back_0
    lhd_row = 0
    counter = 0
    check = ""
    
    sqlWhere = "(tblUsers.department = departments.department_id)"
    sqlWhere = sqlWhere & " AND (ListOnInoutBoard = 1)"
    sqlWhere = sqlWhere & " AND (sid > 0)"
    
    if request.form("button") = lang(cnnDB, "Search") then
    	mFirstname = request.form("mFirstname")
    	mLastname= request.form("mLastname")
    	mUid = request.form("mUid")
    	mDept = request.form("mDept")
    	mPhone = request.form("mPhone")
    	mInoutStatus = request.form("mInoutStatus")
    else
    	mFirstname = ""
    	mLastname = ""
    	mUid = ""
    	mDept = ""
    	mPhone = ""
    	mInoutStatus = 0
    end if
  
    Dim intSort, intOrder, intFirstNameOrder, intLastNameOrder
    Dim intStatusOrder, intPhoneOrder, intUserIDOrder, intDeptOrder
    intSort = Cint(Request.QueryString("sort"))
    If Len(Request.QueryString("order")) > 0 Then
      intOrder = Cint(Request.QueryString("order"))
    Else
      intOrder = 0
    End If
    select case intSort
    	case 1 'Firstname
        If intOrder = 0 Then
      	  orderby = "firstname ASC, lastname ASC"
          intFirstNameOrder = 1
        Else
      	  orderby = "firstname DESC, lastname DESC"
          intFirstNameOrder = 0
        End If
    	case 2 'Lastname
        If intOrder = 0 Then
      		orderby = "lastname ASC, firstname ASC"
          intLastNameOrder = 1
        Else
      		orderby = "lastname DESC, firstname DESC"
          intLastNameOrder = 0
        End If
    	case 3 'Status Code
        If intOrder = 0 Then
      	  orderby = "statuscode ASC, firstname ASC, lastname ASC"
          intStatusOrder = 1
        Else
      	  orderby = "statuscode DESC, firstname DESC, lastname DESC"
          intStatusOrder = 0
        End If
    	case 4 'Phone
    		orderby = "phone"
        If intOrder = 0 Then
          orderby = orderby & " ASC"
          intPhoneOrder = 1
        Else
          orderby = orderby & " DESC"
          intPhoneOrder = 0
        End If
    	case 5 'User ID
    		orderby = "uid"
        If intOrder = 0 Then
          orderby = orderby & " ASC"
          intUserIDOrder = 1
        Else
          orderby = orderby & " DESC"
          intUserIDOrder = 0
        End If
    	case 6 'Department
        If intOrder = 0 Then
    		  orderby = "dname ASC, firstname ASC, lastname ASC"
          intDeptOrder = 1
        Else
    		  orderby = "dname DESC, firstname DESC, lastname DESC"
          intDeptOrder = 0
        End If
      Case Else 'firstname again
        If intOrder = 0 Then
      	  orderby = "firstname ASC, lastname ASC"
          intFirstNameOrder = 1
        Else
      	  orderby = "firstname DESC, lastname DESC"
          intFirstNameOrder = 0
        End If
    end select
    
    if len(mFirstname) > 0 then
    	sqlwhere = sqlwhere & " AND (Firstname like '" & mFirstname & "%')"
    end if
    if len(mLastname) > 0 then
    	sqlwhere = sqlwhere & " AND (Lastname like '" & mLastname & "%')"
    end if
    if len(mUid) > 0 then
    	sqlwhere = sqlwhere & " AND (uid like '" & mUid & "%')"
    end if
    if len(mPhone) > 0 then
    	sqlwhere = sqlwhere & " AND (phone like '" & mPhone & "%')"
    end if
    if mInoutStatus = 1 then
    	check = "checked"
    	sqlwhere = sqlwhere & " AND (statuscode >= 1 " & ")"
    end if
    if len(mDept) > 0 then
    	sqlwhere = sqlwhere & " AND (departments.dname like '" & mDept & "%')"
    end if
    strSql = "SELECT tblUsers.*, departments.dname" &_
    	" FROM tblUsers, departments " &_
    	" WHERE " & sqlWhere &_
    	" ORDER BY " & orderby
  %>

  <br />
  <table class="inoutwide">
  <form method="post" action="default.asp?sort=<% = intSort %>&order=<% =intOrder %>">
  <tr class="normal"><td colspan=7 align="right">
  	<input type="submit" name="button" value="<%=lang(cnnDB, "Search")%>">
  	<input type="submit" name="button" value="<%=lang(cnnDB, "ClearForm")%>">
  	<input type="submit" name="button" value="<%=lang(cnnDB, "ShowAll")%>"></td>
  </tr>

  <tr class="head2" align="center">
    <td><a href="default.asp?sort=1&order=<% = intFirstNameOrder %>" class="HeadLink"><%=lang(cnnDB, "FirstName")%></td>
    <td><a href="default.asp?sort=2&order=<% = intLastNameOrder %>" class="HeadLink"><%=lang(cnnDB, "LastName")%></td>
    <td><a href="default.asp?sort=3&order=<% = intStatusOrder %>" class="HeadLink">X</td>
    <td><a href="default.asp?sort=4&order=<% = intPhoneOrder %>" class="HeadLink"><%=lang(cnnDB, "Phone")%></td>
    <td><a href="default.asp?sort=5&order=<% = intUserIDOrder %>" class="HeadLink"><%=lang(cnnDB, "UserName")%></td>
    <td><a href="default.asp?sort=6&order=<% = intDeptOrder %>" class="HeadLink"><%=lang(cnnDB, "Department")%></td>
    <td>&nbsp;</td>
  </tr>
  
  <tr class="head2" align="center">
    <td><input type="text" name="mFirstname" size=16 value="<% = mFirstname %>"></td>
    <td><input type="text" name="mLastname" size=16 value="<% = mLastname %>"></td>
    <td><input type="checkbox" name="mInoutStatus" value=1 <% = check %>></td>
    <td><input type="text" name="mPhone" size=9 value="<% = mPhone %>"></td>
    <td><input type="text" name="mUid" size=8 value="<% = mUid %>"></td>
    <td><input type="text" name="mDept" size=16 value="<% = mDept %>"></td>
    <td>&nbsp;</td>
  </tr>
  </form>
  
  <%
  
  ' Display a list of all users
  dim rstUser
  Set rstUser = SQLQuery(cnnDB, strSql)
  If Not rstUser.EOF and request.form("button") <> lang(cnnDB, "ClearForm") Then
  	Do While Not rstUser.EOF
  		response.write "<tr class=""" & lhd_row_back & """>" & _
  		  "<td>" & rstUser("firstname") & "</td>" & _
  		  "<td>" & rstUser("lastname") & "</td>"
		

  		select case rstUser("statuscode")
        case 2 
          response.write "<td><a href=""details.asp?id=" & rstUser("sid") & """>" &_
          "<img src=""../image/yellow_pin.gif"" border=0 alt=""(" &rstUser("statusdate") & ") " & rstUser("statustext") & " - " & lang(cnnDB, "clickfordetails") & """></a></td>" 
        case 1
          response.write "<td><a href=""details.asp?id=" & rstUser("sid") & """>" &_
          "<img src=""../image/red_pin.gif"" border=0 alt=""(" &rstUser("statusdate") & ") "  & rstUser("statustext") & " - " & lang(cnnDB, "clickfordetails") & """></a></td>" 
        case 0
          response.write "<td><a href=""details.asp?id=" & rstUser("sid") & """>" &_
          "<img src=""../image/green_pin.gif"" border=0 alt=""" & lang(cnnDB, "Clickfordetails") & """></a></td>" 
      end select
          
  		response.write "<td>" & rstUser("phone") & "</td>" & _
  		  "<td>" & rstUser("uid") & "</td>" & _
 			  "<td>" & rstUser("dname") & "</td>"
  		if inoutadmin = 1 or (sid = rstUser("sid")) then
  			response.write "<td><a href=""update.asp?id=" & rstUser("sid") & """>" & lang(cnnDB, "Edit") & "</a></td>"
  		else
  			response.write "<td>&nbsp;</td>"
  		end if
  		response.write "</tr>" & vbNewLine
  
  		Counter = Counter + 1
  		If lhd_row = 0 then
  			lhd_row = 1
  			lhd_row_back = lhd_row_back_1
  		else
  			lhd_row = 0
  			lhd_row_back = lhd_row_back_0
  		end if
  		rstUser.MoveNext
  	Loop
  End If
  rstUser.close
  
  if counter > 0 then 
  	response.write "<tr class=""body3""><td><font size=-2>" & counter & "&nbsp;" & lang(cnnDB, "recordsfound") & ".</font></td></tr>" & _
      "<tr class=""body3c""><td colspan=""7""><br />" &_
  	  "<img src=""../image/green_pin.gif"" border=""0"">&nbsp;=&nbsp;" & lang(cnnDB, "In") & "&nbsp;&nbsp;&nbsp;" &_
  	  "<img src=""../image/red_pin.gif"" border=""0"">&nbsp;=&nbsp;" & lang(cnnDB, "Out") & "&nbsp;&nbsp;&nbsp;" &_
  	  "<img src=""../image/yellow_pin.gif"" border=0>&nbsp;=&nbsp;" & lang(cnnDB, "Leave") &_
  	  "<br /><br />" & lang(cnnDB, "inoutstatustext") & "</td></tr>"
  else
    if request.form("button") <> lang(cnnDB, "ClearForm") then
      response.write "<tr class=""body3""><td><font size=-2>" & lang(cnnDB, "Noresultsfound") & ".</font></td></tr>"
    end if
  end if
  %>
  </table>
  <%
    Call DisplayFooter(cnnDB, sid)
    cnnDB.Close
  %>
</div>
</body>
</html>
