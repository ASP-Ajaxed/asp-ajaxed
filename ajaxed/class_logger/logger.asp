<%
'**************************************************************************************************************

'' @CLASSTITLE:		Logger
'' @CREATOR:		Michal Gabrukiewicz
'' @CREATEDON:		28.04.2008
'' @CDESCRIPTION:	Provides the opportunity to log messages into text files. Logging can be done on different levels.
''					debug, warn, info or error. Which messages will actually be logged can be set with the logLevel property. 
''					- on dev env the logLevel is debug by default (so all levels are being logged)
''					- Logger supports ascii coloring => log files are easier to read. 
''					- Its recommended to download e.g. "cygwin" and hook up the log file with "tail -f file.log". This allows you to immediately follow the log and it supports coloring as well.
''					- its also possible to view the changes log files directly in the ajaxed console
''					- logfiles are named after the environment. dev.log and live.log
''					- within the ajaxed library lib.logger holds an instance of a ready-to-use logger
''					- some useful debug messages are already automatically logged within the ajaxed library. e.g. SQL queries, page requestes, ajaxed callbacks, emails, ...
''					- log files can also be found in the ajaxed console.
''					- it might be necessary that the logs path has write permission for the IIS user
'' @VERSION:		1.0

'**************************************************************************************************************
class Logger

	'private members
	private p_prefix, escapeChar, extension, p_path
	
	'public members
	public msgPrefix	''[string] the prefix for all log messages. by default its the users ip and time stamp followed by a tabulator
	public defaultStyle	''[string] the default appearance of log messages. check doc in log() method. leave empty if you dont want to use styles at all. This will keep your file clean if you prefer to read them with a common text editor
	public logLevel		''[int] indicates the logger level. debug = 1, info = 2, warn = 4, error = 8
						''- all messages with a higher level than the loglevel will be logged. lower number wont. so e.g. setting level to 2 (info) will log info-, warn-, and error messages
						''- on dev env its "debug" (1) by default.
						''- on live env its "error" (8) by default.
						''- set it to 0 to disable logging completely
	
	'**************************************************************************************************************
	'* constructor 
	'**************************************************************************************************************
	public sub class_initialize()
		msgPrefix = lib.init(AJAXED_LOGMSG_PREFIX, request.servervariables("remote_addr") & " " & now() & vbTab)
		defaultStyle = lib.init(AJAXED_LOGSTYLE, "0;37") 'the default style how text in the log appears.
		extension = ".log"
		p_path = lib.init(AJAXED_LOGSPATH, "/ajaxedLogs/")
		p_prefix = empty
		escapeChar = chr(27)
		logLevel = lib.init(AJAXED_LOGLEVEL, -1)
		'if not set by the config then set the defaults for each env
		if logLevel = -1 and lib.dev then logLevel = 1
		if logLevel = -1 and lib.live then logLevel = 8
	end sub
	
	public property get path ''[string] gets the virtual path where the logfiles are located. e.g. /ajaxedLogs/
		path = p_path
	end property
	
	public property let prefix(val) ''[string] sets the prefix for the log file. Useful if you have more applications and want to have the logs separated. e.g. your app is named "app1" then it will result in a logfile app1_dev.log on the DEV env. default = empty
		prefix = val & "_"
	end property
	
	public property get logfile ''[string] gets the full path and name of the logfile
		logfile = path & p_prefix & lCase(lib.env) & extension
	end property
	
	'**************************************************************************************************************
	'' @SDESCRIPTION: 	logs a message
	'' @PARAM:			level [int]: level of logging. possible: 0 (disabled), 1 (debug), 2 (info), 4 (warn), 8 (error)
	''					- its depending on the loggerLevel setting which level will be logged
	'' @PARAM:			msg [string], [array]: the message you want to log. if array then each field is treated as a line but only the first one has the message prefix
	'' @PARAM:			style [string]: ansi style and color code. check details at http://en.wikipedia.org/wiki/ANSI_escape_code
	''					some values: 31 red, 32 green, 33 yellow, 34 blue, 35 magenta, 36 cyan, 37 white
	''					41 red BG, 42, green BG, ....
	'**************************************************************************************************************
	public sub [log](level, byVal msg, byVal style)
		if not logsOnLevel(level) then exit sub
		if not lib.fso.folderExists(server.mapPath(path)) then lib.fso.createFolder(server.mapPath(path))
		set file = lib.fso.openTextFile(server.mapPath(logfile), 8, true)
		if isEmpty(style) then style = defaultStyle
		if not isArray(msg) then msg = array(msg)
		for i = 0 to uBound(msg)
			if i = 0 then file.write(msgPrefix & getStyle(style))
			file.write(msg(i))
			if i = uBound(msg) then file.write(getStyle(defaultStyle))
			file.write(vbNewLine)
		next
		file.close()
		set file = nothing
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	indicates if a given level would be logged by the logger or not
	'' @DESCRIPTION:	useful if your log message uses some resources which could be saved if the desired loglevel isnt even supported
	'' @PARAM:			level [int]: the level which should be checked.
	'' @RETURN:			[bool] true if it would be logged, otherwise false
	'******************************************************************************************************************
	public function logsOnLevel(level)
		logsOnLevel = level >= logLevel and logLevel > 0
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	helper to log a message as a debug message
	'' @PARAM:			msg [string], [array]: message(s) to debug. (if array then treated as lines)
	'******************************************************************************************************************
	public sub debug(msg)
		me.log 1, msg, "1;31"
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	helper to log a message as an error. 
	'' @PARAM:			msg [string], [array]: message(s) to log
	'******************************************************************************************************************
	public sub [error](msg)
		me.log 8, msg, "0;31"
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	helper to log a message as a warning
	'' @PARAM:			msg [string], [array]: message(s) to log
	'******************************************************************************************************************
	public sub warn(msg)
		me.log 4, msg, "1;33"
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION: 	helper to log a message as an information
	'' @PARAM:			msg [string], [array]: message(s) to log
	'******************************************************************************************************************
	public sub info(msg)
		me.log 2, msg, "1;37"
	end sub
	
	'******************************************************************************************************************
	'* getStyle 
	'******************************************************************************************************************
	private function getStyle(val)
		getStyle = ""
		if isEmpty(defaultStyle) then exit function
		getStyle = escapeChar & "[" & val & "m"
	end function
	
	'**************************************************************************************************************
	'' @SDESCRIPTION: 	Deletes all logfiles
	'**************************************************************************************************************
	public sub clearLogs()
		mPath = server.mappath(path)
		if not lib.fso.folderExists(mPath) then exit sub
		
		set folder = lib.fso.getFolder(mPath)
		for each f in folder.files
			if str.endsWith(f.name, extension) then f.delete()
		next
		set folder = nothing
	end sub

end class
%>