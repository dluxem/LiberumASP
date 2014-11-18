/* helpdesk.sql
**
** This is the SQL schema file for Liberum Helpdesk.
** It was created using SQL Server 2000 but should work fine under
** 6.5 or newer.
**
** Before using, change the FILENAME locations in the CREATE DATABASE
** statement to match your SQL server settings.
**
** Version: 0.97
**
** $Revision: 1.10 $
** $Date: 2002/01/02 21:18:56 $
*/

USE master
GO

CREATE DATABASE HelpDesk
  ON
  (
    NAME = HelpDesk,
    FILENAME = 'c:\mssql\data\helpdesk.mdf',
    SIZE = 5,
    FILEGROWTH = 5
  )
  LOG ON
  (
    NAME = 'HelpDesk_log',
    FILENAME = 'c:\mssql\data\helpdesk_log.ldf',
    SIZE = 5MB,
    FILEGROWTH = 5MB
  )
GO

USE HelpDesk
GO

CREATE TABLE [dbo].[categories] (
	[category_id] [int] NOT NULL ,
	[cname] [varchar] (50) NULL ,
	[rep_id] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[db_keys] (
	[problems] [int] NULL ,
	[departments] [int] NULL ,
	[categories] [int] NULL ,
	[users] [int] NULL ,
	[Lang] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[departments] (
	[department_id] [int] NOT NULL ,
	[dname] [varchar] (50) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[priority] (
	[priority_id] [int] NOT NULL ,
	[pname] [varchar] (50) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[problems] (
	[id] [int] NOT NULL ,
	[uid] [varchar] (50) NULL ,
	[uemail] [varchar] (50) NULL ,
	[ulocation] [varchar] (50) NULL ,
	[uphone] [varchar] (50) NULL ,
	[rep] [int] NULL ,
	[status] [int] NULL ,
	[time_spent] [int] NULL ,
	[category] [int] NULL ,
	[priority] [int] NULL ,
	[department] [int] NULL ,
	[title] [varchar] (50) NULL ,
	[description] [text] NULL ,
	[solution] [text] NULL ,
	[start_date] [datetime] NULL ,
	[close_date] [datetime] NULL ,
	[entered_by] [int] NULL ,
	[kb] [int] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[status] (
	[status_id] [int] NOT NULL ,
	[sname] [varchar] (50) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblConfig] (
	[SiteName] [varchar] (50) NULL ,
	[BaseURL] [varchar] (50) NULL ,
	[AdminPass] [varchar] (50) NULL ,
	[EmailType] [int] NULL ,
	[SMTPServer] [varchar] (50) NULL ,
	[HDName] [varchar] (50) NULL ,
	[HDReply] [varchar] (50) NULL ,
	[BaseEmail] [varchar] (50) NULL ,
	[EnablePager] [int] NULL ,
	[NotifyUser] [int] NULL ,
	[EnableKB] [int] NULL ,
	[DefaultPriority] [int] NULL ,
	[DefaultStatus] [int] NULL ,
	[CloseStatus] [int] NULL ,
	[AuthType] [int] NULL ,
	[Version] [varchar] (6) NULL ,
	[UseSelectUser] [int] NULL ,
	[UseInoutBoard] [int] NULL ,
	[KBFreeText] [int] NULL ,
	[DefaultLanguage] [int] NULL ,
	[AllowImageUpload] [int] NULL ,
	[MaxImageSize] [varchar] (20) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblConfig_Auth] (
	[ID] [int] NOT NULL ,
	[Type] [varchar] (50) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblConfig_Email] (
	[ID] [int] NOT NULL ,
	[Type] [varchar] (50) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblEmailMsg] (
	[type] [varchar] (50) NOT NULL ,
	[subject] [varchar] (50) NULL ,
	[body] [text] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblLangStrings] (
	[id] [int] NOT NULL ,
	[variable] [varchar] (50) NOT NULL ,
	[LangText] [text] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblLanguage] (
	[id] [int] NOT NULL ,
	[LangName] [varchar] (50) NULL ,
	[Localized] [varchar] (50) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblNotes] (
	[id] [int] NOT NULL ,
	[note] [text] NULL ,
	[addDate] [datetime] NULL ,
	[uid] [varchar] (50) NULL ,
	[private] [int] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblUsers] (
	[sid] [int] NOT NULL ,
	[uid] [varchar] (50) NULL ,
	[password] [varchar] (50) NULL ,
	[fname] [varchar] (50) NULL ,
	[email1] [varchar] (50) NULL ,
	[email2] [varchar] (50) NULL ,
	[phone] [varchar] (50) NULL ,
	[location1] [varchar] (50) NULL ,
	[location2] [varchar] (50) NULL ,
	[department] [int] NULL ,
	[IsRep] [int] NULL ,
	[dtCreated] [datetime] NULL ,
	[dtLastAccess] [datetime] NULL ,
	[ListOnInoutBoard] [int] NOT NULL ,
	[firstname] [varchar] (50) NULL ,
	[lastname] [varchar] (50) NULL ,
	[inoutadmin] [int] NOT NULL ,
	[phone_home] [varchar] (50) NULL ,
	[phone_mobile] [varchar] (50) NULL ,
	[jobfunction] [text] NULL ,
	[userresume] [text] NULL ,
	[statustext] [varchar] (255) NULL ,
	[statuscode] [int] NOT NULL ,
	[statusdate] [datetime] NULL ,
	[Language] [int] NULL ,
	[RepAccess] [int] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

ALTER TABLE [dbo].[categories] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[category_id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[departments] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[department_id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[priority] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[priority_id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[problems] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[status] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[status_id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblConfig_Auth] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblConfig_Email] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblEmailMsg] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[type]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblUsers] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[sid]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblUsers] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblUsers_department] DEFAULT (0) FOR [department],
	CONSTRAINT [DF_tblUsers_IsRep] DEFAULT (0) FOR [IsRep],
	CONSTRAINT [DF_tblUsers_dtCreated] DEFAULT (getdate()) FOR [dtCreated],
	CONSTRAINT [DF_tblUsers_dtLastAccess] DEFAULT (getdate()) FOR [dtLastAccess],
	CONSTRAINT [DF__tblUsers__ListOn__267ABA7A] DEFAULT (1) FOR [ListOnInoutBoard],
	CONSTRAINT [DF__tblUsers__inouta__276EDEB3] DEFAULT (0) FOR [inoutadmin],
	CONSTRAINT [DF__tblUsers__status__286302EC] DEFAULT (0) FOR [statuscode],
	CONSTRAINT [DF__tblUsers__RepAcc__2B3F6F97] DEFAULT (0) FOR [RepAccess]
GO

INSERT INTO tblconfig
  (
    SiteName, BaseURL, AdminPass, EmailType, SMTPServer, HDName, HDReply,
    BaseEmail, EnablePager, NotifyUser, EnableKB, DefaultPriority, DefaultStatus, CloseStatus, AuthType,
    Version, UseSelectUser, UseInoutBoard, KBFreeText, DefaultLanguage, AllowImageUpload, MaxImageSize
  )
  VALUES
  (
    'Company Name', 'http://www.company.com/helpdesk', 'admin', 1, 'smtp.company.com',
    'Consultant', 'helpdesk@company.com', '@company.com', 0, 0, 2, 1, 1, 100, 2,
    '0.97', 1, 0, 0, 1, 0, 100000
  )
GO

INSERT INTO tblConfig_Auth (ID, Type)
  VALUES (1, 'NT Authentication')
INSERT INTO tblConfig_Auth (ID, Type)
  VALUES (2, 'Database')
INSERT INTO tblConfig_Auth (ID, Type)
  VALUES (3, 'External Authentication')
GO

INSERT INTO tblConfig_Email (ID, Type)
  VALUES (0, 'Disabled')
INSERT INTO tblConfig_Email (ID, Type)
  VALUES (1, 'CDONTS')
INSERT INTO tblConfig_Email (ID, Type)
  VALUES (2, 'JMail')
INSERT INTO tblConfig_Email (ID, Type)
  VALUES (3, 'ASPEmail')
INSERT INTO tblConfig_Email (ID, Type)
  VALUES (4, 'ASPMail')
GO

INSERT INTO tblEmailMsg(type, subject, body)
  VALUES
  (
    'repclose', 'HELPDESK: Problem [problemid] Closed',
    'The following problem has been closed.  You can view the problem at [rurl]' + CHAR(13) + CHAR(13) +
    'PROBLEM DETAILS' + CHAR(13) +
    '---------------' + CHAR(13) +
    'ID: [problemid]' + CHAR(13) +
    'User: [uid]' + CHAR(13) +
    'Date: [startdate]' + CHAR(13) +
    'Title: [title]' + CHAR(13) +
    'Priority: [priority]' + CHAR(13) +
    'Category: [category]' + CHAR(13) + CHAR(13) +
    'SOLUTION' + CHAR(13) +
    '--------' + CHAR(13) +
    '[solution]'
  )
INSERT INTO tblEmailMsg(type, subject, body)
  VALUES
  (
    'repnew', 'HELPDESK: Problem [problemid] Assigned',
    'The following problem has been assigned to you.  You can update the problem at [rurl]' + CHAR(13) + CHAR(13) +
    'PROBLEM DETAILS' + CHAR(13) +
    '---------------' + CHAR(13) +
    'ID: [problemid]' + CHAR(13) +
    'Date: [startdate]' + CHAR(13) +
    'Title: [title]' + CHAR(13) +
    'Priority: [priority]' + CHAR(13) +
    'Category: [category]' + CHAR(13) + CHAR(13) +
    'USER INFORMATION' + CHAR(13) +
    '----------------' + CHAR(13) +
    'Username: [uid]' + CHAR(13) +
    'Email: [uemail]' + CHAR(13) +
    'Phone: [phone]' + CHAR(13) +
    'Location: [location]' + CHAR(13) +
    'Department: [department]' + CHAR(13) + CHAR(13) +
    'DESCRIPTION' + CHAR(13) +
    '-----------' + CHAR(13) +
    '[description]'
  )
INSERT INTO tblEmailMsg(type, subject, body)
  VALUES
  (
    'reppager', 'HELPDESK: Problem [problemid] Assigned', 
    'Title: [title]' + CHAR(13) +
    'Priority: [priority]' + CHAR(13) +
    'User: [uid]'
  )
INSERT INTO tblEmailMsg(type, subject, body)
  VALUES 
  (
    'repupdate', 'HELPDESK: Problem [problemid] Updated',
    'The following problem has been updated.  You can view the problem at [rurl]' + CHAR(13) + CHAR(13) +
    'PROBLEM DETAILS' + CHAR(13) +
    '---------------' + CHAR(13) +
    'ID: [problemid]' + CHAR(13) +
    'User: [uid]' + CHAR(13) +
    'Date: [startdate]' + CHAR(13) +
    'Title: [title]' + CHAR(13) +
    'Priority: [priority]' + CHAR(13) +
    'Category: [category]' + CHAR(13) + CHAR(13) +
    'DESCRIPTION' + CHAR(13) +
    '-----------' + CHAR(13) +
    '[description]' + CHAR(13) + CHAR(13) +
    'NOTES' + CHAR(13) +
    '-----------' + CHAR(13) +
    '[notes]'
  )
INSERT INTO tblEmailMsg(type, subject, body)
  VALUES 
  (
    'userclose', 'HELPDESK: Problem [problemid] Closed',
    'Your help desk problem has been closed.  You can view the solution below or at: [uurl]' + CHAR(13) + CHAR(13) +
    'PROBLEM DETAILS' + CHAR(13) +
    '---------------' + CHAR(13) +
    'ID: [problemid]' + CHAR(13) +
    'User: [uid]' + CHAR(13) +
    'Date: [startdate]' + CHAR(13) +
    'Title: [title]' + CHAR(13) + CHAR(13) +
    'SOLUTION' + CHAR(13) +
    '--------' + CHAR(13) +
    '[solution]'
  )
INSERT INTO tblEmailMsg(type, subject, body)
  VALUES ('usernew', 'HELPDESK: Problem [problemid] Created',
    'Thank you for submitting your problem to the help desk.  You can view or update the problem at: [uurl]' + CHAR(13) + CHAR(13) +
    'PROBLEM DETAILS' + CHAR(13) +
    '---------------' + CHAR(13) +
    'ID: [problemid]' + CHAR(13) +
    'User: [uid]' + CHAR(13) +
    'Date: [startdate]' + CHAR(13) +
    'Title: [title]' + CHAR(13) + CHAR(13) +
    'DESCRIPTION' + CHAR(13) +
    '-----------' + CHAR(13) +
    '[description]'
  )
INSERT INTO tblEmailMsg(type, subject, body)
  VALUES ('userupdate', 'HELPDESK: Problem [problemid] Updated',
    'Your help desk problem has been updated.  You can view the problem at: [uurl]' + CHAR(13) + CHAR(13) +
    'PROBLEM DETAILS' + CHAR(13) +
    '---------------' + CHAR(13) +
    'ID: [problemid]' + CHAR(13) +
    'User: [uid]' + CHAR(13) +
    'Date: [startdate]' + CHAR(13) +
    'Title: [title]' + CHAR(13) + CHAR(13) +
    'DESCRIPTION' + CHAR(13) +
    '-----------' + CHAR(13) +
    '[description]' + CHAR(13) + CHAR(13) +
    'NOTES' + CHAR(13) +
    '-----------' + CHAR(13) +
    '[notes]'
  )
GO

INSERT INTO tblUsers(sid, uid, fname, email1)
  VALUES (0, 'unknown', 'Unknown', 'none@localhost')
GO

INSERT INTO db_keys(problems, departments, categories, users, Lang)
  VALUES (1,2,1,1,2)  
GO

INSERT INTO status(status_id, sname)
  VALUES (0,'UNKNOWN')
INSERT INTO status(status_id, sname)
  VALUES (1,'OPEN')
INSERT INTO status(status_id, sname)
  VALUES (100,'CLOSED')
GO

INSERT INTO priority(priority_id, pname)
  VALUES (0,'UNKNOWN')
INSERT INTO priority(priority_id, pname)
  VALUES (1,'LOW')
INSERT INTO priority(priority_id, pname)
  VALUES (2,'HIGH')
GO

INSERT INTO departments(department_id, dname)
  VALUES (0,'UNKNOWN')
INSERT INTO departments(department_id, dname)
  VALUES (1,'Dept1')
GO

INSERT INTO tblLanguage (id, LangName, Localized)
  VALUES (1, 'English', 'English')
GO