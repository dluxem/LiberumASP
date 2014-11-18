<%
'***********************************************************************************************
' FileUpload Object v2.5
'   support@aspemporium.com
'
'This multi-object system has complete documentation which comes in the download package as
'a compiled HTML Help reference (*.chm) that will run on any Windows machine. It is fully indexed, 
'searchable and you can bookmark pages. Every single public property/method and class is documented. 
'If you're ever scratching your head wondering what one of these functions returns... you should just 
'look it up in the reference. There are 6 objects containing a total of 60 properties and methods for 
'version 2.5. Once again, the reference documents everything that's new, everything that's changed 
'and everything else, just for fun. There's also sections on installation, usage, common problems,
'hardware/software requirements, etc. There's also code all over the place in there showing how to
'use the objects and code tricks for using the objects more efficiently as well. 
'
'***********************************************************************************************
'object summary:
'	FileUpload Class
'	FO_Processor Class
'	FO_File Class
'	FO_Properties Class
'	FO_FileChecker Class
'	Base64Encoder Class
'***********************************************************************************************






Class FileUpload
	Private UploadRequest, oProps, iFrmCt
	Private iKnownFileCount, iKnownFormCount	
	Private oOutFiles

	Private Sub Class_Initialize
		iFrmCt = 0
		Set oProps = New FO_Properties
		Set UploadRequest = Server.CreateObject("Scripting.Dictionary")
		iKnownFileCount = 0
		iKnownFormCount = 0
		set oOutFiles = Server.CreateObject("Scripting.Dictionary")
	End Sub

	Private Sub Class_Terminate
		set oOutFiles = Nothing
		Set UploadRequest = Nothing
		Set oProps = Nothing
	End Sub

	Public Property Get Version()
		Version = "2.5"
	End Property

	Public Function GetUploadSettings()
		Set GetUploadSettings = oProps
	End Function

	Public Property Get FormCount
		FormCount = iKnownFormCount
	End Property

	Public Property Get FileCount
		FileCount = iKnownFileCount
	End Property

	Public Property Get TotalFormCount
		TotalFormCount = iFrmCt
	End Property

	Private Function GetFormEncType()
		Dim sContType, hCutOff

		sContType = request.servervariables("CONTENT_TYPE")
		hCutOff = instr(sContType, ";")
		if hCutOff > 0 then
			sContType = UCase(Trim(Left(sContType, hCutOff - 1)))
		else
			sContType = UCase(Trim(sContType))
		end if
		GetFormEncType = sContType
	End Function

	Public Default Sub ProcessUpload
	'after processupload is called, totalformcount property, formcount and 
	'filecount properties are filled, form method returns entered data
		Dim RequestBin, oProcess, iTotBytes, key, arr, iKnownProps, oFile
		Dim fofilecheck, sEncType, sReqMeth

		iTotBytes = Request.TotalBytes
		if iTotBytes = 0 then
			iFrmCt = 0
			exit sub
		end if

		 ' read posted content(s)
		RequestBin = Request.BinaryRead(iTotBytes)





		'11/14/2001 - test request method and encoding
		'*********************************************************************
		'- You can add your own parsers here by following the same format below.
		'  if the input is a POST, you can add parsing methods to use
		'  by entering a new enctype in the inner select case statement below.
		'
		'  If the input is a GET, you can also add a parser for that condition or
		'  any other request method below by expanding the outer select case statement.
		'
		'- see appendix 1 in the docs for step by step instructions for adding
		'  your own input parsers
		'
		'*********************************************************************

		''''''''''''''''''''''''''''''''''''''''''''''''''
		'1.) request method check
		''''''''''''''''''''''''''''''''''''''''''''''''''
		'test request method
		sReqMeth = request.servervariables("REQUEST_METHOD")
		select case UCase(sReqMeth)
			case "POST"
				'determine enctype of form
				''''''''''''''''''''''''''''''''''''''''''''''''''
				'2.) form encoding method check
				''''''''''''''''''''''''''''''''''''''''''''''''''
				'test form encoding type
				sEncType = GetFormEncType
				select case sEncType
					case "MULTIPART/FORM-DATA"

						 ' call BuildUploadRequest to parse binary info
						Set oProcess = New FO_Processor
						oProcess.BuildUploadRequest  RequestBin, UploadRequest
						Set oProcess = Nothing

					case "APPLICATION/X-WWW-FORM-URLENCODED"

						 ' call ascii form processor
						Set oProcess = New FO_Processor
						oProcess.BuildUploadRequest_ASCII oProcess.getString(RequestBin), UploadRequest
						Set oProcess = Nothing

					case else

						'do nothing with unknown enc types
				end select

			case "GET"
				'do nothing with querystring inputs...

				'To create your own GET parser, let IIS do the hard work for you
				'and just retrieve the QUERY_STRING environment variable
				'and then pass it to a new method in the FO_Processor object
				'that will process it...
				'
				'    inputs_to_parse = Request.ServerVariables("QUERY_STRING")
				'     ' call my query string processor
				'    Set oProcess = New FO_Processor
				'    oProcess.MyQueryStringProcessor inputs_to_parse, UploadRequest
				'    Set oProcess = Nothing
				'

			case else
				'do nothing with other request methods
		end select











		arr = uploadrequest.keys

		if not isarray(arr) then
			iFrmCt = 0
			exit sub
		end if

		iFrmCt = ubound(arr)
		for each key in arr
			if isobject(uploadrequest.item(key)) then
				iKnownProps = ubound(uploadrequest.item(key).keys) + 1
				if iKnownProps = 4 then
					'it's a file
					iKnownFileCount = iKnownFileCount + 1

					set fofilecheck = new FO_FileChecker
					fofilecheck.SetCurrentProperties oProps
					fofilecheck.FileInput_NamePath = uploadrequest.item(key).item("FileName")
					fofilecheck.FileInput_ContentType = uploadrequest.item(key).item("ContentType")
					fofilecheck.FileInput_BinaryText = uploadrequest.item(key).item("Value")
					fofilecheck.FileInput_FormInputName = uploadrequest.item(key).item("InputName")
					set oFile = fofilecheck.ValidateVerifyReturnFile()
					set fofilecheck = nothing

					oOutFiles.add iKnownFileCount, oFile
					set oFile = nothing
					uploadrequest.remove key
				elseif iKnownProps = 1 then
					'it's a form input
					iKnownFormCount = iKnownFormCount + 1
				else
					'i have no idea what it is
				end if
			end if
		next
	End Sub

	Public Function File(ByVal blobName)
		'version 2.5 allows an input name as well as an integer between
		'1 and FileCount.

		Dim blobs, blob, subdict, tmpName

		'new addition for 2.5 adds inputname to internal blob number
		'processing step which searches all keys for the entered name
		'first. if found, substitutes the number of the blobname entered
		'for the ordinal internal blob number. If not found, processing
		'continues as usual.
		blobs = oOutFiles.Keys
		For Each blob In blobs
			'this is a FO_File object
			Set subdict = oOutFiles.Item(blob)
			tmpName = subdict.frmInputName
			If UCase(Trim(tmpName)) = UCase(Trim(blobName)) Then
				blobName = blob
				Exit For
			End If
		Next

		'old version 2.0 way
		if isobject(oOutFiles.Item(blobName)) then
			Set File = oOutFiles.Item(blobName)
		else
			Set File = Nothing
		end if
	End Function

	Public Function Form(ByVal inputName)
		if isobject(UploadRequest.Item(inputName)) then
			Form = UploadRequest.Item(inputName).Item("Value")
		else
			Form = ""
		end if
	End Function

	Public Function FormLen(ByVal inputName)
		if isobject(UploadRequest.Item(inputName)) then
			FormLen = Len(UploadRequest.Item(inputName).Item("Value"))
		else
			FormLen = 0
		end if
	End Function

	Public Function FormEx(ByVal inputName, ByVal vDefaultValue)
		dim vTmp

		if isobject(UploadRequest.Item(inputName)) then
			vTmp = UploadRequest.Item(inputName).Item("Value")
			if len(trim(CStr(vTmp))) = 0 then
				FormEx = vDefaultValue
				Exit Function
			end if

			FormEx = vTmp
			Exit Function
		end if

		FormEx = vDefaultValue
	End Function

	Public Function Inputs()
		if isobject(UploadRequest) then
			Inputs = UploadRequest.keys
		else
			Inputs = ""
		end if
	End Function

	Public Sub ShowUploadForm(ByVal sSubmitPage)
		 ' display the upload form and let the 
		 ' user know what they can and cannot upload
		Dim tmp, item

		With Response
			.Write("<P>You can currently add any file of type: ")
			tmp = ""
			If IsArray(oProps.Extensions) Then
				For Each Item In oProps.Extensions
					tmp = tmp & "<CODE>*." & Item & "</CODE>, "
				Next
				tmp = left( tmp, Len(tmp) - 2 )
			End If
			.Write(tmp & "<BR>")
			.Write("Each file must have a maximum size of: <CODE>~ ")
			.Write(Round( oProps.MaximumFileSize / 1024, 1 ) & " k</CODE> ")
			.Write("and a minimum size of: <CODE>~ ")
			.Write(FormatNumber(Round( oProps.MininumFileSize _
				/ 1024, 1 ), 1) & " k.</CODE></P>")
			.Write("</P>")

			.Write("<FORM ENCTYPE=""multipart/form-data"" ACTION=""")
			.Write(sSubmitPage & """ METHOD=""POST"">" & vbCrLf)

			.Write("Please select a file to upload ")
			if oProps.UploadDisabled Then
				.Write("from your computer [upload is disabled]:<BR>" & vbCrLf)
				.Write("<INPUT TYPE=FILE NAME=""blob"" DISABLED><BR><BR>" & vbCrLf)
			Else
				.Write("from your computer:")
				.Write(" [Upload is optional]")

				.Write("<BR>" & vbCrLf)
				.Write("<INPUT TYPE=FILE NAME=""blob""><BR><BR>" & vbCrLf)
			End If

			.Write("Please enter your full name:<BR>" & vbCrLf)
			.Write("<INPUT TYPE=TEXT NAME=""myName"" SIZE=35><BR><BR>" & vbCrLf)
			.Write("<INPUT TYPE=SUBMIT VALUE=""Upload File"">" & vbCrLf)
			.Write("</FORM>" & vbCrLf)
		End With
	End Sub
End Class



Class FO_FileChecker
	Private oProps, sFileName, hFileBinLen, sFileBin, sFileContentType, sFileFormInputName

	Private Sub Class_Initialize()
		'initialize everything to the "bad" settings
		sFileName = ""
		hFileBinLen = 0
		sFileBin = ""
		sFileContentType = ""
	End Sub

	Public Sub SetCurrentProperties(byref oPropertybag)
		Set oProps = oPropertybag
	End Sub

	Public Property Let FileInput_FormInputName(ByVal fname)
		sFileFormInputName = fname
	End Property

	Public Property Let FileInput_NamePath(ByVal fname)
		Dim realfilename

		'** parse the file name minus any directory path from the input path
		realfilename = Right(fname, Len(fname) - InstrRev(fname,"\"))

		sFileName = trim(realfilename)
	End Property

	Public Property Let FileInput_ContentType(ByVal conttype)
		sFileContentType = conttype
	End Property

	Public Property Let FileInput_BinaryText(ByVal binstring)
		Dim  binlen

		binlen = lenb(binstring)
		hFileBinLen = binlen
		sFileBin = binstring
	End Property

	Public Function ValidateVerifyReturnFile()	'As FO_File
		'call all the validation methods.
		'if any fail, fill the FO_File object
		'accordingly and stop processing

		if IllegalCharsFound then
			Set ValidateVerifyReturnFile = FillFOFileObj(false, "", "", "bad character in file name", "", "", "", sFileFormInputName)
			Exit Function
		end if

		if FileNameBadOrExists then
			Set ValidateVerifyReturnFile = FillFOFileObj(false, "", "", "file name bad or non-existent or file with same name already exists and overwrite disabled", "", "", "", sFileFormInputName)
			Exit Function
		end if

		If FileExtensionIsBad then
			Set ValidateVerifyReturnFile = FillFOFileObj(false, "", "", "file extension is not allowed or doesn't exist", "", "", "", sFileFormInputName)
			Exit Function
		End If

		If FileSizeIsBad then
			Set ValidateVerifyReturnFile = FillFOFileObj(false, "", "", "file size is either too large or too small", "", "", "", sFileFormInputName)
			Exit Function
		end if

		Set ValidateVerifyReturnFile = FillFOFileObj(false, "", "", "", sFileContentType, sFileName, sFileBin, sFileFormInputName)
	End Function

	Private Function FillFOFileObj(byval success, byval abspath, byval virpath, byval stderr, byval contenttype, byval fname, byval binarytext, byval forminputname)
		'create FO_File object	
		Dim oFile

		set oFile = New FO_File
		oFile.SetCurrentProperties oProps
		oFile.bSuccess = success
		oFile.sAbsPath = abspath
		oFile.sVirPath = virpath
		oFile.sStdErr = stderr
		oFile.sCType = contenttype
		oFile.sFileName = fname
		oFile.binValue = binarytext
		oFile.frmInputName = forminputname
		set FillFOFileObj = oFile
	End Function	

	'added illegal character check...
	Public Function IllegalCharsFound()
		'** test file name for illegal characters
		Dim re

		set re = new regexp
		re.pattern = "\\\/\:\*\?\""\<\>\|"
		re.global = true
		re.ignorecase = true
		if re.test(sFileName) then
			IllegalCharsFound = true
		else
			IllegalCharsFound = false
		end if
		set re = nothing
	End Function

	Public Function FileNameBadOrExists()
		Dim absuploaddirectory, oFSO

		'** test file name length
		if len(trim(sFileName)) = 0 then
			FileNameBadOrExists = true
			Exit Function
		end if
		
		'repaired this block to only get the file system involved if necessary.
		'if allowing overwrite, who cares. otherwise, see if file exists.
		'considered not valid if file exists
		if oProps.AllowOverWrite then
			FileNameBadOrExists = false
			Exit Function
		end if

		absuploaddirectory = oProps.uploaddirectory & "\" & trim(sFileName)

		'** test for file exists, if necessary
		set oFSO = server.createobject("Scripting.FileSystemObject")
		if oFSO.FileExists(absuploaddirectory) then
			FileNameBadOrExists = true
		else
			FileNameBadOrExists = false
		end if
		Set oFSO = Nothing
	End Function

	Public Function FileExtensionIsBad()
		Dim sFileExtension, bFileExtensionIsValid, sFileExt

		'** parse for file type extension
		if len(trim(sFileName)) = 0 then
			FileExtensionIsBad = true
			Exit Function
		end if

		sFileExtension = right(sFileName, len(sFileName) - instrrev(sFileName, "."))
		bFileExtensionIsValid = false	'assume extension is bad
		for each sFileExt in oProps.extensions
			if ucase(sFileExt) = ucase(sFileExtension) then
				'if the extensions match, it's good. stop checking
				bFileExtensionIsValid = True
				exit for
			end if
		next
		FileExtensionIsBad = not bFileExtensionIsValid
	End Function

	Public Function FileSizeIsBad()
		if hFileBinLen > oProps.MaximumFileSize then
			FileSizeIsBad = True
			Exit Function
		end if

		if hFileBinLen < oProps.MininumFileSize then
			FileSizeIsBad = True
			Exit Function
		end if

		FileSizeIsBad = False
	End Function
End Class



Class FO_Processor
	 ' #########################################################
	 ' # UPLOAD ROUTINES                                       #
	 ' # For detailed information about these routines, go to: #
	 ' # http://www.asptoday.com/articles/20000316.htm         #
	 ' #########################################################

	Private Function getByteString(byval StringStr)
		 ' For detailed information about this routine, go to:
		 ' http://www.asptoday.com/articles/20000316.htm
		dim char, i

		For i = 1 to Len(StringStr)
			char = Mid(StringStr, i, 1)
			getByteString = getByteString & chrB(AscB(char))
		Next
	End Function

	Public Function getString(byval StringBin)
		 ' For detailed information about this routine, go to:
		 ' http://www.asptoday.com/articles/20000316.htm
		dim intCount

		getString =""
		For intCount = 1 to LenB(StringBin)
			getString = getString & chr(AscB(MidB(StringBin, intCount, 1))) 
		Next
	End Function

	Public Sub BuildUploadRequest_ASCII(ByVal sPostStr, ByRef UploadRequest) 
		dim i, j, blast, sName, vValue

		blast = false
		i = -1
		do while i <> 0
			if i = -1 then
				i = 1
			else
				i = i + 1
			end if
			j = instr(i, sPostStr, "=") + 1
			sName = mid(sPostStr, i, j-i-1)
			i = instr(j, sPostStr, "&")
			if i = 0 then 
				vValue = mid(sPostStr, j)
			else
				vValue = mid(sPostStr, j, i - j)
			end if

			Dim uploadcontrol
			set uploadcontrol = createobject("Scripting.Dictionary")
			uploadcontrol.add "Value", vValue

			if not uploadrequest.exists(sName) then
				uploadrequest.add sName, uploadcontrol
			end if
		loop
	End Sub



	Public Sub BuildUploadRequest(byref RequestBin, byref UploadRequest)
		 ' For detailed information about this routine, go to:
		 ' http://www.asptoday.com/articles/20000316.htm
		dim PosBeg, PosEnd, boundary, boundaryPos, Pos, Name, PosFile
		dim PosBound, FileName, ContentType, Value, sEncType, sReqMeth

		'zero byte check
		if lenb(RequestBin) = 0 then 
			'7/23/01 - zero byte check
			'no form data posted
			exit sub
		end if

		PosBeg = 1
		PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))

		if posend = 0 then
			'7/23/01 - no binary input passed check
			'translate binary to ascii and transfer control
			'to the regular form parser.

			BuildUploadRequest_ASCII getString(requestbin), UploadRequest
			Exit Sub
		end if

		boundary = MidB(RequestBin,PosBeg,PosEnd-PosBeg)
		boundaryPos = InstrB(1,RequestBin,boundary)
		Do until (boundaryPos=InstrB(RequestBin,boundary & getByteString("--")))
			Dim UploadControl
			Set UploadControl = Server.CreateObject("Scripting.Dictionary")
			Pos = InstrB(BoundaryPos,RequestBin,getByteString("Content-Disposition"))
			Pos = InstrB(Pos,RequestBin,getByteString("name="))
			PosBeg = Pos+6
			PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(34)))
			Name = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
			PosFile = InstrB(BoundaryPos,RequestBin,getByteString("filename="))
			PosBound = InstrB(PosEnd,RequestBin,boundary)

			If  PosFile<>0 AND (PosFile<PosBound) Then
				PosBeg = PosFile + 10
				PosEnd =  InstrB(PosBeg,RequestBin,getByteString(chr(34)))
				FileName = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
				UploadControl.Add "FileName", FileName
				Pos = InstrB(PosEnd,RequestBin,getByteString("Content-Type:"))
				PosBeg = Pos+14
				PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))
				ContentType = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
				UploadControl.Add "ContentType",ContentType
				PosBeg = PosEnd+4
				PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
				Value = MidB(RequestBin,PosBeg,PosEnd-PosBeg)
			Else
				Pos = InstrB(Pos,RequestBin,getByteString(chr(13)))
				PosBeg = Pos+4
				PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
				Value = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
			End If
			UploadControl.Add "Value" , Value
			UploadControl.Add "InputName", Name
			if not uploadrequest.exists(name) then 
				'7/22/01 - added check to see if top level input name already
				'exists to prevent bombing if 2 inputs have the same name.
				'Now, if this situation occurs, the first input is always used
				'and any other inputs with the same name are discarded.
				UploadRequest.Add name, UploadControl	
			end if

			BoundaryPos=InstrB(BoundaryPos+LenB(boundary),RequestBin,boundary)
		Loop
	End Sub
End Class



Class FO_File
	Public bSuccess
	Public sAbsPath
	Public sVirPath
	Public sStdErr
	Public sCType
	Public frmInputName
	Public binValue
	Private hBtCt, sURiPath, sFiExt
	private sfinme

	Private oProps

	Public property let sFileName(byval filenameinput)
		'resolve extension
		sFiExt = right(filenameinput, len(filenameinput) - instrrev(filenameinput, "."))
		sfinme = filenameinput
	end property

	public property get sFileName()
		sFileName = sfinme
	end property

	Private Sub Class_Initialize()
		bSuccess = false
		sAbsPath = ""
		sVirPath = ""
		sStdErr = ""
		hBtCt = 0
		sCType = ""
		sFileName = ""
		binValue = ""
		sURiPath = ""
	End Sub

	Public Sub SetCurrentProperties(byref oPropertybag)
		Set oProps = oPropertybag
	End Sub

	Public Sub SaveAsRecord(byref oField)
		sAbsPath = ""
		sVirPath = ""
		sURiPath = ""
		bSuccess = false

		If LenB(binValue) = 0 Then 
			Exit Sub
		End If

		if oProps.UploadDisabled then
			sStdErr = "Uploading disabled by administrator"
			Exit Sub
		end if
		
		If IsObject(oField) Then
			'8/18/2001 - added some error handling to try to
			'catch errors when trying to add blobs to a
			'ms access 97 database (which doesn't support them)
			On Error Resume Next
			oField.AppendChunk binValue
			if Err Then
				sStdErr = Err.Description
				bBtCt = 0
				bSuccess = false
				Exit Sub
			end if
			On Error GoTo 0

			hBtCt = lenb(binValue)
			bSuccess = true
		End If
	End Sub

	Public Sub SaveAsFile()
		If sStdErr <> "" Then
			exit sub
		end if

		'upload file
		WriteUploadFile oProps.uploaddirectory & "\" & sFileName, binValue
	End Sub

	Public Function SaveAsBinaryString()
		If LenB(binValue) = 0 Then 
			bBtCt = 0
			bSuccess = false
			Exit Function
		End If

		if oProps.UploadDisabled then
			bBtCt = 0
			bSuccess = false
			sStdErr = "Uploading disabled by administrator"
			Exit Function
		end if

		SaveAsBinaryString = binValue
		hBtCt = lenb(binValue)
		bSuccess = true
	End Function

	Public Function SaveAsString()
		Dim outstr, i

		If LenB(binValue) = 0 Then 
			bBtCt = 0
			bSuccess = false
			Exit Function
		End If

		if oProps.UploadDisabled then
			bBtCt = 0
			bSuccess = false
			sStdErr = "Uploading disabled by administrator"
			Exit Function
		end if

		' translate binary data into ASCII 
		outstr = ""
		For i = 1 to LenB( binValue )
			outstr = outstr & chr( AscB( MidB( binValue, i, 1) ) )
		Next
		SaveAsString = outstr
		hBtCt = lenb(binValue)
		bSuccess = true
	End Function

	Public Function SaveAsBase64EncodedStr()
		Dim outstr, oEnc

		If LenB(binValue) = 0 Then 
			bBtCt = 0
			bSuccess = false
			Exit Function
		End If

		if oProps.UploadDisabled then
			bBtCt = 0
			bSuccess = false
			sStdErr = "Uploading disabled by administrator"
			Exit Function
		end if

		'base 64 encode ASCII
		Set oEnc = New Base64Encoder
		outstr = oEnc.EncodeStr(binValue)
		Set oEnc = Nothing
		SaveAsBase64EncodedStr = outstr
		hBtCt = lenb(binValue)
		bSuccess = true
	End Function

	Private Sub WriteUploadFile(byVal NAME, byVal CONTENTS)
		 ' create the file on the server
		dim ScriptObject, i, NewFile

		on error resume next

		if oProps.UploadDisabled then
			err.raise "31234", "FO Obj", "Uploading disabled by administrator"
		else
			Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
			Set NewFile = ScriptObject.CreateTextFile( NAME )
			For i = 1 to LenB( CONTENTS )
				 ' translate binary data into ASCII 
				 ' characters and write them into the file.
				NewFile.Write chr( AscB( MidB( CONTENTS, i, 1) ) )
			Next
			NewFile.Close
			Set NewFile = Nothing
			Set ScriptObject = Nothing
		end if
		if err.number <> 0 then
			sStdErr = Err.Description
			bSuccess = false
		else
			sAbsPath = NAME
			sVirPath = UnMappath(NAME)
			hBtCt = lenb(CONTENTS)
			sURiPath = "http://" & request.servervariables("HTTP_HOST") & sVirPath
			bSuccess = true
		end if
		on error goto 0
	End Sub

	Private Function UnMappath(byVal pathname)
		'http://aspemporium.com/aspEmporium/codelib/codelib.asp?pid=8&cid=8
		dim tmp, strRoot

		strRoot = Server.Mappath("/")
		tmp = replace( lcase( pathname ), lcase( strRoot ), "" )
		tmp = replace( tmp, "\", "/" )
		UnMappath = tmp
	End Function

	Public Property Get ContentType()
		ContentType = sCType
	End Property

	Public Property Let FileName(byval newfilename)
		'store in: sFileName
		'after validating

		'test new filename - on error, filename
		'remains what it was when entered if an
		'upload is attempted after an unsuccessful
		'rename.

		Dim oFileChk

		set oFileChk = New FO_FileChecker
		oFileChk.SetCurrentProperties oProps
		oFileChk.FileInput_NamePath = newfilename
		if oFileChk.IllegalCharsFound Then
			sStdErr = "illegal characters found in new file name"
			bSuccess = false
			set oFileChk = Nothing
			Exit Property
		end if
		if oFileChk.FileNameBadOrExists Then
			sStdErr = "file name is bad or file with same name already exists and overwrite disabled"
			bSuccess = false
			set oFileChk = Nothing
			Exit Property
		End If
		if oFileChk.FileExtensionIsBad Then
			sStdErr = "file extension is not allowed or doesn't exist"
			bSuccess = false
			set oFileChk = Nothing
			Exit Property
		End If
		Set oFileChk = Nothing

		'reset filename to new file name if passes all tests
		sStdErr = ""
		sFileName = newfilename
	End Property

	Public Property Get FileExtension()
		FileExtension = sFiExt
	End Property

	Public Property Get FileNameWithoutExtension()
		'chop any/all extensions from the filename and return just the file name without the extension

		FileNameWithoutExtension = StripFileExtensionFromFileName(sFileName)
	End Property

	Public Function StripFileExtensionFromFileName(ByVal filenametostrip)
		Dim hExtensionStart, tmpfilenametoalter

		tmpfilenametoalter = filenametostrip
		hExtensionStart = -1
		do while not hExtensionStart = 0
			hExtensionStart = instrrev(tmpfilenametoalter, ".")
			if hExtensionStart > 0 then
				tmpfilenametoalter = left(tmpfilenametoalter, hExtensionStart - 1)
			end if
		loop
		StripFileExtensionFromFileName = tmpfilenametoalter
	End Function

	Public Function JoinFileExtensionToFileName(ByVal filenametojoin, byval fileextensiontojoin)
		Dim strippedfilename

		strippedfilename = StripFileExtensionFromFileName(filenametojoin)
		JoinFileExtensionToFileName = strippedfilename & "." & fileextensiontojoin
	End Function

	Public Function GetFileNameFromFilePath(ByVal filewithpath)
		dim fileend

		fileend = instrrev(filewithpath, "\")
		GetFileNameFromFilePath = right(filewithpath, len(filewithpath) - fileend)
	End Function

	Public Property Get FileName()
		FileName = sFileName
	End Property

	Public Property Get UploadSuccessful()
		UploadSuccessful = bSuccess
	End Property

	Public Property Get AbsolutePath()
		AbsolutePath = sAbsPath
	End Property

	Public Property Get URLPath()
		URLPath = sURiPath
	End Property

	Public Property Get VirtualPath()
		VirtualPath = sVirPath
	End Property

	Public Property Get ErrorMessage()
		ErrorMessage = sStdErr
	End Property

	Public Property Get ByteCount()
		ByteCount = hBtCt
	End Property
End Class



Class FO_Properties
	Private sErrHead		'string
	Private sErrMsg			'string
	Private arrExt			'variant - array
	Private strUploadDir		'string
	Private boolAllowOverwrite	'boolean
	Private lngUploadSize		'long
	Private bMin			'long
	Private bByPass			'boolean

	Private Sub Class_Initialize()
		sErrHead = "FileUpload Object - Invalid Property Setting"
		sErrMsg = ""
		arrExt = Array("txt", "htm", "html", "zip", "inc")
		strUploadDir = Server.Mappath("/")
		boolAllowOverwrite = false
		lngUploadSize = 100000
		bMin = 1024
		bByPass = false
	End Sub

	Public Sub ResetAll()
		Class_Initialize
	End Sub

	Public Property LET Extensions(byVal arrayInput)
		dim item, bErr

		bErr = false
		if isarray(arrayInput) then
			'check array
			for each item in arrayInput
				if instr(item, ".") <> 0 then
					bErr = true
					exit for
				end if
			next
			if not bErr then
				arrExt = arrayInput
				Exit Property
			else
				arrayInput = ""
			end if
		end if

		sErrMsg = "Extensions property input must be an array of extensions without the dot(.)."
		if arrayInput = "*" then
			Err.Raise 21340, sErrHead, sErrMsg & _
				" The Wildcard is no longer supported as an option."
		else
			Err.Raise 21341, sErrHead, sErrMsg
		end if
	End Property

	Public Property LET UploadDirectory(byVal strInput)
		Dim oFSO, bDoesntExist

		bDoesntExist = false

		if instr(strInput, "/") <> 0 then
			strInput = ""
			Err.Raise 21342, sErrHead, _
				"UploadDirectory property - absolute path required for this property."
			exit property
		end if

		Set oFSO = CreateObject("Scripting.FileSystemObject")
		if not oFSO.FolderExists(strInput) then bDoesntExist = true
		set oFSO = Nothing
		if bDoesntExist then
			Err.Raise 21343, sErrHead, "UploadDirectory property - """ & _
				strInput & """ directory doesn't exist on the server."
			Exit Property
		end if

		strUploadDir = strInput
	End Property

	Public Property LET AllowOverWrite(byVal boolInput)
		on error resume next
		boolInput = cbool(boolInput)
		on error goto 0
		boolAllowOverwrite = boolInput
	End Property

	Public Property LET MaximumFileSize(byVal lngInput)
		if isnumeric(lngInput) then
			on error resume next
			lngInput = CLng( lngInput )
			on error goto 0

			lngUploadSize = lngInput
			exit property
		end if

		Err.Raise 21344, sErrHead, "MaximumFileSize Property must be a long integer."
	End Property

	Public Property LET MininumFileSize(byVal lngInput)
		if isnumeric(lngInput) then
			on error resume next
			lngInput = CLng( lngInput )
			on error goto 0

			bMin = lngInput
			exit property
		end if

		Err.Raise 21345, sErrHead, "MininumFileSize Property must be a long integer."
	End Property

	Public Property LET UploadDisabled(byval boolInput)
		on error resume next
		boolInput = cbool(boolInput)
		on error goto 0
		bByPass = boolInput
	End Property

	Public Property GET UploadDisabled()
		UploadDisabled = bByPass
	End Property

	Public Property GET MininumFileSize()
		MininumFileSize = bMin
	End Property

	Public Property GET Extensions()
		Extensions = arrExt
	End Property

	Public Property GET UploadDirectory()
		UploadDirectory = strUploadDir
	End Property

	Public Property GET AllowOverWrite()
		AllowOverWrite = boolAllowOverwrite
	End Property

	Public Property GET MaximumFileSize()
		MaximumFileSize = lngUploadSize
	End Property
End Class

Class Base64Encoder
	'written for vb by: webmaster@q-tec.org
	'and converted by bill <support@aspemporium.com> for
	'the CCVerification class and brought over to the
	'FileUpload class
	Private Base64Chars

	Private Sub Class_Initialize()
		Base64Chars =	"ABCDEFGHIJKLMNOPQRSTUVWXYZ" & _
				"abcdefghijklmnopqrstuvwxyz" & _
				"0123456789" & _
				"+/"
	End Sub

	Public Function EncodeStr(byVal strIn)
		Dim c1, c2, c3, w1, w2, w3, w4, n, strOut
		For n = 1 To Len(strIn) Step 3
			c1 = Asc(Mid(strIn, n, 1))
			c2 = Asc(Mid(strIn, n + 1, 1) + Chr(0))
			c3 = Asc(Mid(strIn, n + 2, 1) + Chr(0))
			w1 = Int(c1 / 4) : w2 = (c1 And 3) * 16 + Int(c2 / 16)
			If Len(strIn) >= n + 1 Then 
				w3 = (c2 And 15) * 4 + Int(c3 / 64) 
			Else 
				w3 = -1
			End If
			If Len(strIn) >= n + 2 Then 
				w4 = c3 And 63 
			Else 
				w4 = -1
			End If
			strOut = strOut + mimeencode(w1) + mimeencode(w2) + _
					  mimeencode(w3) + mimeencode(w4)
		Next
		EncodeStr = strOut
	End Function

	Private Function mimedecode(byVal strIn)
		If Len(strIn) = 0 Then 
			mimedecode = -1 : Exit Function
		Else
			mimedecode = InStr(Base64Chars, strIn) - 1
		End If
	End Function

	Public Function DecodeStr(byVal strIn)
		Dim w1, w2, w3, w4, n, strOut
		For n = 1 To Len(strIn) Step 4
			w1 = mimedecode(Mid(strIn, n, 1))
			w2 = mimedecode(Mid(strIn, n + 1, 1))
			w3 = mimedecode(Mid(strIn, n + 2, 1))
			w4 = mimedecode(Mid(strIn, n + 3, 1))
			If w2 >= 0 Then _
				strOut = strOut + _
					Chr(((w1 * 4 + Int(w2 / 16)) And 255))
			If w3 >= 0 Then _
				strOut = strOut + _
					Chr(((w2 * 16 + Int(w3 / 4)) And 255))
			If w4 >= 0 Then _
				strOut = strOut + _
					Chr(((w3 * 64 + w4) And 255))
		Next
		DecodeStr = strOut
	End Function


	Private Function mimeencode(byVal intIn)
		If intIn >= 0 Then 
			mimeencode = Mid(Base64Chars, intIn + 1, 1) 
		Else 
			mimeencode = ""
		End If
	End Function
End Class
%>