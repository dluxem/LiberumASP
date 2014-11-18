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

  Filename: delrep.asp
  Date:     $Date: 2001/12/09 02:01:24 $
  Version:  $Revision: 1.50 $
  Purpose:  Removed the selected support rep from viewrep.asp
  -->
  
  <!-- 	#include file = "../public.asp" -->

  <% 
    Dim cnnDB, sid
    Set cnnDB = CreateCon
    sid = GetSid
  %>

  <head>
    <title><%=lang(cnnDB, "HelpDesk")%>&nbsp;-&nbsp;<%=lang(cnnDB, "RemoveSupportReps")%></title>
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>

    <%
      ' Check for perms to view this page
      Call CheckAdmin

      Dim sidList, blnActiveProblem, strActiveRepName, blnActiveCategory
      sidList = Split(Request.Form("rsid"), ",")
      blnActiveProblem = False
      blnActiveCategory = False

      Dim rep, probRes
      For Each Rep in sidList
        Set probRes = SQLQuery(cnnDB, "SELECT id FROM problems WHERE rep=" & Rep & " AND status<>" & Cfg(cnnDB, "CloseStatus"))
        If Not probRes.EOF Then
          blnActiveProblem = True
          strActiveRepName = Usr(cnnDB, Rep, "uid")
        End If
        probRes.Close
      Next

      Dim rstProb
      If Not blnActiveProblem Then
        For Each Rep in sidList
          Set rstProb = SQLQuery(cnnDB, "SELECT category_id FROM categories WHERE rep_id = " & Rep)
          If Not rstProb.EOF Then
            blnActiveCategory = True
            strActiveRepName = Usr(cnnDB, Rep, "uid")
          End If
        Next
        rstProb.Close
      End If

      Dim oldRep, OldProbRes, repRes, rstEnteredBy
      If Not (blnActiveProblem Or blnActiveCategory) Then
        For Each oldRep In sidList
          Set OldProbRes = SQLQuery(cnnDB, "UPDATE problems SET rep=0 WHERE rep=" & oldRep)
          Set repRes = SQLQuery(cnnDB, "UPDATE tblUsers SET IsRep=0 WHERE sid=" & oldRep)
        Next
      End If
    %>
    <div align="center">
      <table class="Normal">
        <tr class="Head1">
          <td>
            <font size="+2"><b><%=lang(cnnDB, "SupportRepresentatives")%></b></font>
          </td>
        </tr>
        <tr class="Body1">
          <td>
            <div align="center">
              <% If blnActiveProblem Then %>
                <%=lang(cnnDB, "Theuser")%>&nbsp;'<% = strActiveRepName %>'&nbsp;
                <%=lang(cnnDB, "DelrepOpenProblemsText")%>
              <% ElseIf blnActiveCategory Then %>
                <%=lang(cnnDB, "Theuser")%>&nbsp;'<% = strActiveRepName %>'&nbsp;
                <%=lang(cnnDB, "DelrepCategoryAssignedText")%>
              <% Else %>
                <%=lang(cnnDB, "Thesupportrephavebeenremoved")%>
              <% End If %>
            </div>
          </td>
        </tr>
      </table>
      <p>
      <a href="default.asp"><%=lang(cnnDB, "AdministrativeMenu")%></a><br>
      <a href="viewrep.asp"><%=lang(cnnDB, "Manage")%>&nbsp;<%=lang(cnnDB, "SupportReps")%></a>
    </div>

    <%

      Call DisplayFooter(cnnDB, sid)
      cnnDB.Close
    %>
  </body>
</html>
