<%@ LANGUAGE=VBScript LCID=2057%>
<%
  Option Explicit
  'Buffer the response, so Response.Expires can be used
  Response.Buffer = TRUE
%>
<?xml version="1.0"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"
"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

  <!--
  Liberum Help Desk, Copyright (C) 2000-2001 Doug Luxem
  Liberum Help Desk comes with ABSOLUTELY NO WARRANTY
  Please view the license.html file for the full GNU General Public License.

  Filename: sysinfo.asp
  Date:     $Date: 2002/01/24 14:57:50 $
  Version:  $Revision: 1.1.2.1 $
  Purpose:  Extract info from the server for debug purpose
  -->

  <!-- #include file = "../settings.asp" -->
  <%
    On Error Resume Next
    Call SetAppVariables
  %>

  <head>
    <title>
      System Information
    </title>
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>
  <%
  	Dim strConn, cnnDB
    Select Case Application("DBType")
      Case 1	' SQL Sec
        strConn = "Provider=SQLOLEDB.1;Data Source=" & Application("SQLServer") & _
          ";Initial Catalog=" & Application("SQLDBase") & _
          ";uid=" & Application("SQLUser") & ";pwd=" & Application("SQLPass")

      Case 2	' SQL Integrated Sec
        strConn = "Provider=SQLOLEDB.1;Data Source=" & Application("SQLServer") & _
          ";Initial Catalog=" & Application("SQLDBase") & ";Integrated Security=SSPI"

      Case 3	' Access
        strConn = "Provider=Microsoft.Jet.OLEDB.4.0" & _
          ";Data Source=" & Application("AccessPath")
      Case 4	' DSN
        strConn = "DSN=" & Application("DSN_Name")

    End Select

    ' Create and open the database connection and save it as a
    ' session variable
    Set cnnDB = Server.CreateObject("ADODB.Connection")
    cnnDB.Open(strConn)

    Dim blnConnError
    blnConnError = False
    If Err.Number <> 0 Then
      Err.Clear
      blnConnError = True
    End If

    Dim fsFileSys, fsFile, blnDBExists
    blnDBExists = True
    Set fsFileSys = Server.CreateObject("Scripting.FileSystemObject")
    If Not fsFileSys.FileExists(Application("AccessPath")) Then
      blnDBExists = False
    End If

    Dim blnUpdateError
    If blnConnError Then
      blnUpdateError = True
    Else
      Err.Clear
      Dim adoRec
      Set adoRec = Server.CreateObject("ADODB.Recordset")
      adoRec.ActiveConnection = cnnDB
      adoRec.Open("INSERT INTO tblConfig (version) VALUES ('_test')")
      adoRec.ActiveConnection = cnnDB
      adoRec.Open("DELETE FROM tblConfig WHERE version='_test'")
      If Err.Number <> 0 Then
        Err.Clear
        blnUpdateError = True
      Else
        blnUpdateError = False
      End If
    End If

    Function CfgLookup(strSetting)
      On Error Resume Next
      Err.Clear
      Dim confRes
      Set confRes = Server.CreateObject("ADODB.Recordset")
      confRes.ActiveConnection = cnnDB
      confRes.Open("SELECT " & strSetting & " From tblConfig")
      If Err.Number <> 0 Then
        Err.Clear
        CfgLookup = "<em>Error</em>"
      Else
        If confRes.EOF Then
          CfgLookup = "<em>Missing</em>"
        Else
          CfgLookup = confRes(strSetting)
        End If
      End If
      confRes.Close
    End Function



    Dim strKey, intN
    intN = 1
  %>

  <div align="center">
    <table class="normal">
      <tr class="Head1"><td colspan="2"><div align="center">System Information</div></td></tr>
      <tr class="head2">
        <td>Application Variable</td>
        <td>Value</td>
      </tr>
      <tr class="body1">
        <td>DBType</td>
        <td><% = Application("DBType") %></td>
      </tr>
      <tr class="body2">
        <td>SQLServer</td>
        <td><% = Application("SQLServer") %></td>
      </tr>
      <tr class="body1">
        <td>SQLDBase</td>
        <td><% = Application("SQLDBase") %></td>
      </tr>
      <tr class="body2">
        <td>AccessPath</td>
        <td><% = Application("AccessPath") %></td>
      </tr>
      <tr class="body1">
        <td>DNS_Name</td>
        <td><% = Application("DSN_Name") %></td>
      </tr>

      <tr class="head2">
        <td>Verify DB</td>
        <td>Value</td>
      </tr>
      <tr class="body1">
        <td>DB Connection</td>
        <% If blnConnError Then %>
          <td><em>Error Connecting</em></td>
        <% Else %>
          <td>Good</td>
        <% End If %>
      </tr>
      <tr class="body2">
        <td>Access .MDB Exists</td>
        <% If Not blnDBExists Then %>
          <td><em>Database does not exist.</em></td>
        <% Else %>
          <td>Good</td>
        <% End If %>
      </tr>
      <tr class="body1">
        <td>Update DB</td>
        <% If blnUpdateError Then %>
          <td><em>Error Updating Database</em></td>
        <% Else %>
          <td>Good</td>
        <% End If %>
      </tr>

      <tr class="head2">
        <td>Config Settings</td>
        <td>Value</td>
      </tr>
      <tr class="body1">
        <td>BaseURL</td>
        <td><% = CfgLookup("BaseURL") %></td>
      </tr>
      <tr class="body2">
        <td>EmailType</td>
        <td><% = CfgLookup("EmailType") %></td>
      </tr>
      <tr class="body1">
        <td>SMTPServer</td>
        <td><% = CfgLookup("SMTPServer") %></td>
      </tr>
      <tr class="body2">
        <td>EnablePager</td>
        <td><% = CfgLookup("EnablePager") %></td>
      </tr>
      <tr class="body1">
        <td>AuthType</td>
        <td><% = CfgLookup("AuthType") %></td>
      </tr>
      <tr class="body2">
        <td>Version</td>
        <td><% = CfgLookup("Version") %></td>
      </tr>
      <tr class="body1">
        <td>UseInOutBoard</td>
        <td><% = CfgLookup("UseInOutBoard") %></td>
      </tr>
      <tr class="body2">
        <td>KBFreeText</td>
        <td><% = CfgLookup("KBFreeText") %></td>
      </tr>

      <tr class="head2">
        <td>Server Variable</td>
        <td>Value</td>
      </tr>
      <%
        For Each strKey In Request.ServerVariables
          Response.write "<tr class=""body" & intN & """>" & _
            "<td>" & strKey & "</td>" & _
            "<td>" & Request.ServerVariables(strKey) & "</td>" & _
            "</tr>" &vbNewLine
          if intN = 1 then intN = 2 else intN = 1 end if
        Next
        Response.Write "</table>"
      %>
  </div>
  </body>

</html>
