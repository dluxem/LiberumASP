<SCRIPT LANGUAGE = "VBScript" RUNAT="Server">

'	Liberum Help Desk, Copyright (C) 2000 Doug Luxem
'	Liberum Help Desk comes with ABSOLUTELY NO WARRANTY
'	Please veiw the license.html file for the full GNU
'	General Public License

' --------
' SETTINGS.ASP
'
' Loads the appplication variables.
' --------

' SetAppVariables:
' The procedure runs when the application is started or the file is changed
' Primary ojbectives are to set variables/constants used throughout the
' application.
Sub SetAppVariables


	'========================================
	' Database Information

		' Database Type
		' 1 - SQL Server with SQL security (set SQLUser/SQLPass)
		' 2 - SQL Server with integrated security
		' 3 - Access Database (set AccessPath)
		' 4 - DSN (An ODBC DataSource) (set DSN_Name)

	Application("DBType") = 3
	'========================================

	'============ SQL SETTINGS ==============
	Application("SQLServer") = "SQLSERVER"	' Server name (don't put the leading \\)
	Application("SQLDBase") = "HelpDesk"	' Database name
	Application("SQLUser") = "sa"			' Account to log into the SQL server with
	Application("SQLPass") = "sapass"		' Password for account
	' =======================================

	'=========== ACCESS SETTINGS ============
	'Physical path to database file
	Application("AccessPath") = "C:\Inetpub\Databases\helpdesk2000.mdb"
	'========================================

	'============= DSN SETTINGS =============
	Application("DSN_Name") = "HelpDeskDSN"
	'========================================

	' Enable Debugging:
	' Set to true to view full MS errors and other debug information
	' printed.  (This will disable most On Error Resume Next statements.)
	Application("Debug") = False


End Sub

</SCRIPT>