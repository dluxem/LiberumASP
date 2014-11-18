<%@ LANGUAGE="VBScript" %>
<% 
  Option Explicit
  'Buffer the response, so Response.Expires can be used
  Response.Buffer = True
  Response.Expires = -1
  Server.ScriptTimeOut = 600  ' Wait 10 minutes to time out the script
%>


<?xml version="1.0"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

  <!--
  Liberum Help Desk, Copyright (C) 2000-2001 Doug Luxem
  Liberum Help Desk comes with ABSOLUTELY NO WARRANTY
  Please view the license.html file for the full GNU General Public License.

  Filename: setup.asp
  Date:     $Date: 2002/08/28 15:31:43 $
  Version:  $Revision: 1.57.2.1.2.2 $
  Purpose:  This page is used for upgrades and will update the database
            to the current version.
  -->

  <!--	#include file = "settings.asp" -->
  <!-- 	#include file = "public.asp" -->
  <% 
    Call SetAppVariables
    Dim cnnDB
    Set cnnDB = CreateCon
  %>

  <head>
    <title>Help Desk - Update Database</title>
    <link rel="stylesheet" type="text/css" href="default.css">
  </head>
  <body>

    <%
      Dim strCurVersion, strNewVersion, blnUpdate, blnIsCurrent, blnError, blnUpdateLang, blnOverWrite

      ' -------------------------------------
      ' Enter the latest version number below
      ' -------------------------------------
      strNewVersion = "0.97"
      
      Dim strLangDir
      strLangDir = "lang\"

      blnUpdate = False
      blnUpdateLang = False
      blnError = False
      blnOverWrite = False

      If Cint(Request.Form("update")) = 1 Then
        blnUpdate = True
      End If

      If Cint(Request.Form("updatelang")) = 1 Then
        blnUpdateLang = True
      End If

      If Request.Form("overwrite") = "on" Then
        blnOverWrite = True
      End If

      Dim rstVersion

      Set rstVersion = Server.CreateObject("ADODB.Recordset")
      rstVersion.ActiveConnection = cnnDB

      On Error Resume Next

      ' Retrieve the current version of the database
      ' If a version number does not exists, assume it is 0.95
      rstVersion.Open("SELECT Version From tblConfig")

      Select Case Err.Number
        Case 0      ' Successful
          strCurVersion = rstVersion("Version")
        Case Else  ' Column Doesn't Exist
          strCurVersion = "0.95"
          Err.Clear
        End Select
      rstVersion.Close
      If IsNull(strCurVersion) Then
        strCurVersion = "0.96"
      End If
            
      If Not Application("Debug") Then
        On Error Resume Next
      Else
        On Error Goto 0
      End If

      If strCurVersion = strNewVersion Then
        blnIsCurrent = True
      Else
        blnIsCurrent = False
      End If

      ' Do the updates
      If blnUpdate Then
        Dim strErrorCmd, strErrorDesc, strErrorSrc, intErrorNum, strConn
        cnnDB.Close
        If Application("DBType") = 1 Then
          strConn = "Provider=SQLOLEDB.1;Data Source=" & Application("SQLServer") & _
	  		    ";Initial Catalog=" & Application("SQLDBase") & _
  	  		  ";uid=" & Request.Form("sqluser") & ";pwd=" & Request.Form("sqlpass")
          cnnDB.Open(strConn)
          If Err.Number <> 0 Then
            blnError = True
            intErrorNum = Err.Number
            strErrorDesc = Err.Description
            strErrorSrc = Err.Source
            strErrorCmd = strConn
            Err.Clear
            Set cnnDB = CreateCon
          End If
        Else
          Set cnnDB = CreateCon
        End If
   
        ' Start a ADO transaction
        cnnDB.BeginTrans

        ' -------------------------------
        ' Update from 0.95 to 0.96
        ' -------------------------------
        If strCurVersion = "0.95" Then
          ' Add version field and value
          UpdateDB("ALTER TABLE tblConfig ADD Version varchar(6) NULL")
          UpdateDB("UPDATE tblConfig SET Version = '0.96'")
          
          ' Add field for enabling update emails to users
          UpdateDB("ALTER TABLE tblConfig ADD NotifyUser int")
          UpdateDB("UPDATE tblConfig SET NotifyUser=0")

          strCurVersion = "0.96"
        End If

        ' ------------------------------
        ' Update from 0.96 to 0.97
        ' ------------------------------
        If strCurVersion = "0.96" Then
          ' Add entered_by field in problems
          UpdateDB("ALTER TABLE problems ADD entered_by int NULL")
          UpdateDB("UPDATE problems SET entered_by=0")

          ' Create tblNotes table
          UpdateDB("CREATE TABLE tblNotes (" & _
            "id int NOT NULL, " & _
            "[note] text NULL, " & _
            "addDate datetime NULL, " & _
            "uid varchar(50) NULL, " & _
            "private int NULL)")

          ' Add field for enabling select user on rep/new form
          UpdateDB("ALTER TABLE tblConfig ADD UseSelectUser int NULL")
          UpdateDB("UPDATE tblConfig SET UseSelectUser=1")

          ' Create In/Out board fields in tblUsers
          UpdateDB("ALTER TABLE tblUsers " & _
            "ADD ListOnInoutBoard int NOT NULL DEFAULT 1, " & _
            "[firstname] varchar (50) NULL, " & _
            "[lastname] varchar (50) NULL, " & _
            "inoutadmin int NOT NULL DEFAULT 0, " & _
            "[phone_home] varchar (50) NULL, " & _
            "[phone_mobile] varchar (50) NULL, " & _
            "[jobfunction] text NULL, " & _
            "[userresume] text NULL, " & _
            "[statustext] varchar (255) NULL, " & _
            "statuscode int NOT NULL DEFAULT 0, " & _
            "statusdate datetime NULL")
            
          ' Set initial values for tblInout
          UpdateDB("UPDATE tblUsers SET InoutAdmin=0")
          UpdateDB("UPDATE tblUsers SET statuscode=0")
          UpdateDB("UPDATE tblUsers SET ListOnInoutBoard=1")
          UpdateDB("ALTER TABLE tblConfig ADD UseInoutBoard int NULL")
          UpdateDB("UPDATE tblConfig SET UseInoutBoard=0")

          ' Remove color fields in tblConfig
          UpdateDB("ALTER TABLE tblConfig DROP " & _
            "COLUMN Color1, Color2, BGColor, TextColor, LinkColor, VLinkColor, ALinkColor")

          ' Add kb field to problems table
          UpdateDB("ALTER TABLE problems ADD kb int NULL")
          UpdateDB("UPDATE problems SET kb=0")
          UpdateDB("UPDATE problems SET kb=1 WHERE status=" & Cfg(cnnDB, "CloseStatus"))

          ' SQL free text searching in config
          UpdateDB("ALTER TABLE tblConfig ADD KBFreeText int NULL")
          UpdateDB("UPDATE tblConfig SET KBFreeText=0")

          ' Change EnableKB value to match new values
          UpdateDB("UPDATE tblConfig SET EnableKB=2 WHERE EnableKB=1")

          'Add default language field to config table
          UpdateDB("ALTER TABLE tblConfig ADD DefaultLanguage int NULL")
          UpdateDB("UPDATE tblConfig SET DefaultLanguage=1")
          
          'Add default language field to user table
          UpdateDB("ALTER TABLE tblUsers ADD [Language] int NULL")
          UpdateDB("UPDATE tblUsers SET [Language]=1")

          'Add language table key
          UpdateDB("ALTER TABLE db_keys ADD Lang int NULL")
          UpdateDB("UPDATE db_keys SET Lang=2")

          'Add table for available languages
          UpdateDB("CREATE TABLE tblLanguage (" & _
            "id int NOT NULL, " & _
            "LangName varchar (50) NULL, " & _
            "Localized varchar (50) NULL)")
          UpdateDB("INSERT INTO tblLanguage (id, LangName, Localized) VALUES (1, 'English', 'English')")

          'Add table for language strings
          UpdateDB("CREATE TABLE tblLangStrings (" & _
            "id int NOT NULL, " & _
            "variable varchar (50) NOT NULL, " & _
            "LangText text NOT NULL)")
          
          ' Add Restricted/ReadOnly fields
          UpdateDB("ALTER TABLE tblUsers ADD " & _
            "RepAccess int NOT NULL DEFAULT 0")

          UpdateDB("UPDATE tblUsers SET RepAccess=0")

          ' Update userupate message
          Dim strUserUpdate, strRepUpdate
          strUserUpdate = "Your help desk problem has been updated.  You can view the problem at: [uurl]" & vbNewLine & vbNewLine & _
            "PROBLEM DETAILS" & vbNewLine & _
            "---------------" & vbNewLine & _
            "ID: [problemid]" & vbNewLine & _
            "User: [uid]" & vbNewLine & _
            "Date: [startdate]" & vbNewLine & _
            "Title: [title]" & vbNewLine & vbNewLine & _
            "DESCRIPTION" & vbNewLine & _
            "-----------" & vbNewLine & _
            "[description]" & vbNewLine & vbNewLine & _
            "NOTES" & vbNewLine & _
            "-----------" & vbNewLine & _
            "[notes]"
          UpdateDB("UPDATE tblEmailMsg SET body = '" & strUserUpdate & "' WHERE type='userupdate'")
          strRepUpdate = "The following problem has been updated.  You can view the problem at [rurl]" & vbNewLine & vbNewLine & _
            "PROBLEM DETAILS" & vbNewLine & _
            "---------------" & vbNewLine & _
            "ID: [problemid]" & vbNewLine & _
            "User: [uid]" & vbNewLine & _
            "Date: [startdate]" & vbNewLine & _
            "Title: [title]" & vbNewLine & vbNewLine & _
            "DESCRIPTION" & vbNewLine & _
            "-----------" & vbNewLine & _
            "[description]" & vbNewLine & vbNewLine & _
            "NOTES" & vbNewLine & _
            "-----------" & vbNewLine & _
            "[notes]"
          UpdateDB("UPDATE tblEmailMsg SET body = '" & strRepUpdate & "' WHERE type='repupdate'")

          ' On/Off config switch to allow upload of user images or not
          UpdateDB("ALTER TABLE tblConfig ADD AllowImageUpload int NULL")
          UpdateDB("UPDATE tblConfig SET AllowImageUpload=0")
          
          ' Set maximum filesize for uploaded user images
          UpdateDB("ALTER TABLE tblConfig ADD [MaxImageSize] varchar (20) NULL")
          UpdateDB("UPDATE tblConfig SET MaxImageSize='100000'")

          ' Add ASPEmail selection
          UpdateDB("INSERT INTO tblConfig_Email (id, type) VALUES (4, 'ASPMail')")

          ' Update version field
          UpdateDB("UPDATE tblConfig SET Version = '0.97'")           
          strCurVersion = "0.97"
        End If

        ' Language Updates (keep last)
        Call UpdateAllLanguages

        ' Check for errors and either roll back the transaction or commit it
        If blnError Then
          cnnDB.RollbackTrans
        Else
          cnnDB.CommitTrans
          blnIsCurrent = True
        End If
      End If

      ' Update just the language strings
      If blnUpdateLang Then

        Call UpdateAllLanguages

        ' Check for errors and either roll back the transaction or commit it
        If blnError Then
          cnnDB.RollbackTrans
        Else
          cnnDB.CommitTrans
        End If
      End If

      ' ---------------------------------------------------
      ' Subroutine that calls the updates for each language
      ' ---------------------------------------------------
      Sub UpdateAllLanguages
        'English
        If Request.Form("english") <> "" Then
          Call UpdateLang("English", "English", "English_English.txt")
        End If

        'Norwegian (Norsk)
        If Request.Form("norwegian") <> "" Then
          Call UpdateLang("Norwegian", "Norsk", "Norwegian_Norsk.txt")
        End If

        'Danish (Dansk)
        If Request.Form("danish") <> "" Then
          Call UpdateLang("Danish", "Dansk", "Danish_Dansk.txt")
        End If

        'Dutch (Nederlands)
        If Request.Form("dutch") <> "" Then
          Call UpdateLang("Dutch", "Nederlands", "Dutch_Nederlands.txt")
        End If

        'German (Deutsch)
        If Request.Form("german") <> "" Then
          Call UpdateLang("German", "Deutsch", "German_Deutsch.txt")
        End If

        'French
        If Request.Form("french") <> "" Then
          Call UpdateLang("French", "Français", "French_Français.txt")
        End If

        'Spanish
        If Request.Form("spanish") <> "" Then
          Call UpdateLang("Spanish", "Español", "Spanish_Español.txt")
        End If

        ' Remove cached language strings
        Call ClearLangCache(cnnDB)
      
      End Sub

      ' Subroutine to update languages in the database
      Sub UpdateLang(strLangName, strLocalized, strFileName)
        Dim rstGetLangID, intLangID
        Set rstGetLangID = SQLQuery(cnnDB, "SELECT id FROM tblLanguage WHERE LangName = '" & strLangName & "' AND localized = '" & strLocalized & "'")
        If rstGetLangID.EOF Then
          intLangID = GetUnique(cnnDB, "lang")
          UpdateDB("INSERT INTO tblLanguage (id, LangName, localized) VALUES " & _
            "(" & intLangID & ", '" & strLangName & "', '" & strLocalized & "')")
        Else 
          intLangID = rstGetLangID("id")
        End If
        
        Dim fsFileSys, fsFile, fsLine, strVarName, strLangText, rstCheckVar, strFullFileName, tsLangFile
        Set fsFileSys = Server.CreateObject("Scripting.FileSystemObject")
        strFullFileName = Server.MapPath(strLangDir & strFileName)
        If fsFileSys.FileExists(strFullFileName) Then
          Set fsFile = fsFileSys.GetFile(strFullFileName)
          Set tsLangFile = fsFileSys.OpenTextFile(strFullFileName, 1, 0) ' ForReading, ASCII format
          Do While Not tsLangFile.AtEndOfStream 
            fsLine = tsLangFile.ReadLine()
            fsLine = Trim(fsLine)
            If Not (InStr(fsLine, ";") = 1) and Not (Left(fsLine, 1) = "[")  and (len(fsline)>0) Then
              fsLine = Split(fsLine, "=", 2)
              strVarName = Trim(fsLine(0))
              strLangText = Trim(fsLine(1))
              If Len(strVarName) > 0 And Len(strLangText) > 0 Then
                strLangText = Replace(strLangText, "'", "''") 
                strVarName = Replace(strVarName, "'", "''") 
                Set rstCheckVar = SQLQuery(cnnDB, "SELECT * FROM tblLangStrings WHERE id=" & intLangID & " AND variable='" & strVarName & "'")
                If rstCheckVar.EOF Then
                  UpdateDB("INSERT INTO tblLangStrings (id, variable, LangText) VALUES " & _
                    "(" & intLangID & ", '" & strVarName & "', '" & strLangText & "')")
                Else
                  If (StrComp(rstCheckVar("variable"), strVarName, 0) = 0) And blnOverWrite Then  'Doing a binary compare
                    UpdateDB("UPDATE tblLangStrings SET LangText='" & strLangText & "' WHERE " & _
                      "id = " & intLangID & " AND variable='" & strVarName & "'")
                   Else
                    UpdateDB("INSERT INTO tblLangStrings (id, variable, LangText) VALUES " & _
                      "(" & intLangID & ", '" & strVarName & "', '" & strLangText & "')")
                   End If
                End If
                rstCheckVar.Close
                Set rstCheckVar = Nothing
              End If
            End If
          Loop
          tsLangFile.Close
          Set tsLangFile = Nothing
          Set fsFile = Nothing
        End If
        rstGetLangID.Close
        Set rstGetLangID = Nothing
      End Sub

      ' Subroutine to update the database and check for any errors that occured
      Sub UpdateDB(strSQLCommand)
        If Not Application("Debug") Then
          On Error Resume Next
        End If
        If Not blnError Then
          cnnDB.Execute (strSQLCommand)
          If Err.Number <> 0 Then
            blnError = True
            intErrorNum = Err.Number
            strErrorDesc = Err.Description
            strErrorSrc = Err.Source
            strErrorCmd = strSQLCommand
            Err.Clear
          End If
        End If
      End Sub

      ' Subroutine to print out the language form
      Sub PrintLangForm
        Response.Write("<b>Select languages to install:</b><br>")
        ' Four languages per line
        Response.Write("English:<input type=""checkbox"" name=""english"" checked> | ")
        Response.Write("Danish:<input type=""checkbox"" name=""danish""> | ")
        Response.Write("Dutch:<input type=""checkbox"" name=""dutch"">")
        Response.Write("<br />")
        Response.Write("French:<input type=""checkbox"" name=""french""> | ")
        Response.Write("German:<input type=""checkbox"" name=""german""> | ")
        Response.Write("Norwegian:<input type=""checkbox"" name=""norwegian"">")
        Response.Write("Spanish:<input type=""checkbox"" name=""spanish"">")
        Response.Write("<br />")
        Response.Write("Overwrite any existing language strings:<input type=""checkbox"" name=""overwrite"" checked>")
        Response.Write("<br />")
        Response.Write("<i>Processing may take several minutes.</i>")

      End Sub
    %>

    <div align="center">
      <table Class="Normal">
        <tr Class="Head1">
          <td>
            Update Database
          </td>
        </tr>
        <tr Class="Body1">
          <td>
            <% If blnError Then %>
                <b>Error!</b><p>
                An error has occured during the update process.
                <p>
                <b>Command:</b> <% = strErrorCmd%><br>
                <b>Number:</b> <% = intErrorNum %><br>
                <b>Source:</b> <% = strErrorSrc %><br>
                <b>Description:</b> <% = strErrorDesc %>
            <% Elseif blnUpdate or blnUpdateLang Then %>
                <b>Successfully Updated!</b><p>
                Your database has been successfully update to version <% = strNewVersion %>.  You should
                now remove setup.asp from your web server to prevent unauthorized attempts to manipulate
                your database.
                <p>
                <div align="center">
                  <a href="admin/">Configure your help desk.</a>
                </div>
            <% Elseif blnIsCurrent Then%>
                <b>Database Is Current.</b><br>
                <b>Version: <% = strCurVersion %></b>
                <p>
                Your database configuration is current; however, if you are doing a new installation of
                Liberum Help Desk then you will need to install the language strings using the button below.
                <p>
                <div align="center">
                  <form method="post" action="setup.asp">
                    <input type="hidden" name="updatelang" value="1">
                    <input type="submit" value="Install/Upgrade Language Strings"><br>
                    <% Call PrintLangForm %>
                  </form>
                </div>
                <p>
                <div align="center">
                  <a href="default.asp">Logon to the help desk.</a>
                </div>
              <% Else ' database need updating%>
                <form method="post" action="setup.asp">
                  <input type="hidden" name="update" value="1">
                  <% If strCurVersion = "0.95" Then %>
                    <b>Warning:</b> Setup was unable to detect which version of the database you are running.
                    This is normal if you are running version 0.95 and you may continue with the update;
                    however, if you are running a version previous to 0.95, then you must manually update the
                    database to 0.95 or higher.
                    <p>
                    <b>New Version: <% = strNewVersion %></b>
                  <% Else %>
                    <b>Version Detected: <% = strCurVersion %></b><br>
                    <b>New Version: <% = strNewVersion %></b><p>
                    You may update your current version of the database to the new one.
                  <% End If %>
                  <% If Application("DBType") = 1 Then %>
                    <p>
                    <b>Enter an account with sysadmin or dbowner roles:</b><br>
                    User: <input type="text" name="sqluser" size="20" value="sa"><br>
                    Password: <input type="password" name="sqlpass" size="20">
                  <% ElseIf Application("DBType") = 2 Then %>
                    <p>
                    <b>Please make sure that your account has SQL sysadmin or dbowner roles before
                    continuing.</b>
                  <% End If %>
                  <p>
                  <div align="center">
                    <input type="Submit" value="Update to v<% = strNewVersion %>"><br>
                    <% Call PrintLangForm %>
                  </div>
                </form>
            <% End If %>
          </td>
        </tr>
      </table>
    </div>
    <%
      cnnDB.Close
    %>
  </body>
</html>
