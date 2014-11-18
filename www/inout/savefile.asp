<%@ LANGUAGE="VBScript" %>
<% 
  Option Explicit
  With Response
    .Buffer = True
    .Expires = 0
    .Clear
  End With
%>


<?xml version="1.0"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

  <!--
  Liberum Help Desk, Copyright (C) 2000-2001 Doug Luxem
  Liberum Help Desk comes with ABSOLUTELY NO WARRANTY
  Please view the license.html file for the full GNU General Public License.

  Filename: savefile.asp
  Date:     $Date: 2002/01/24 14:19:18 $
  Version:  $Revision: 1.51.2.1 $
  Purpose:  Save uploaded user image to disk
  -->

  <!--  #include file = "../public.asp" -->
  <!--  #include file = "../uploadClass.asp" -->

  <% 
    Dim cnnDB, sid, usid, uid
    Set cnnDB = CreateCon
    sid = GetSid
   	usid = request.QueryString("uid")
   	uid = usr(cnnDB, usid, "uid")
  %>

  <head>
    <title>
      <% = Cfg(cnnDB, "SiteName") %>&nbsp;<%=lang(cnnDB, "Uploadimage")%>
    </title>
    <link rel="stylesheet" type="text/css" href="../default.css">
  </head>
  <body>

  <%
    ' See is user is validated
    Call CheckUser(cnnDB, sid)
  %>

  <div align="center">
    <table class="normal">
      <tr class="Head1"><td><% = Cfg(cnnDB, "SiteName") %></td></tr>
      <tr class="Head2" align="center"><td><%=lang(cnnDB, "Uploadimage")%></td></tr>
    </table>
    <table class="normal">
      <%
        Dim objFO, objProps, objFile
        Set objFO = New FileUpload
        Set objProps = objFO.GetUploadSettings
        with objProps
          .Extensions = Array("jpg", "jpeg")
          .UploadDirectory = Server.Mappath("../image/")
          .AllowOverWrite = true
          .MaximumFileSize = cfg(cnnDB, "MaxImageSize")
          .MininumFileSize = 1000
          .UploadDisabled = false
        End with
        set objProps = nothing
        objFO.ProcessUpload
        set objFile = objFO.File(1)
        if objFile.ErrorMessage <> "" then
          response.write "<tr class=""body1c""><td>" & lang(cnnDB, "Anerroroccurreduploadingafile") & ": " & _
            objFile.ErrorMessage & "</td></tr>"
        else
          objFile.FileName = uid & ".jpg"
          objFile.SaveAsFile
          if objFile.UploadSuccessful then
            response.write "<tr class=""body1c""><td>" & lang(cnnDB, "fileuploadedsuccessfully") & "</td></tr>"
          else
            response.write "<tr class=""body1c""><td>" & lang(cnnDB, "Anerroroccurredsavingfiletodisk") & ": " & _
              objFile.ErrorMessage & "</td></tr>"
          end if
        end if
        
        Response.Write"</table><br />" & _
          "<a href=""details.asp?id=" & usid & """>" & lang(cnnDB, "Details") & _
          "</div>"

        Call DisplayFooter(cnnDB, sid)
        'clean up
        set objFile = Nothing
        set objFO = Nothing
        cnnDB.Close
      %>
  </body>
</html>
