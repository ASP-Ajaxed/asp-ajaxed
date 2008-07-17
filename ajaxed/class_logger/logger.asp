<%
'**************************************************************************************************************

'' @CLASSTITLE:		Logger
'' @CREATOR:		Michal Gabrukiewicz
'' @CREATEDON:		28.04.2008
'' @CDESCRIPTION:	Provides the opportunity to log messages into text files. Logging can be done on different levels:
''					- <em>1</em> (debug): all kind of debugging messages
''					- <em>2</em> (info): messages which should inform the developer about something. e.g. Progress of a procedure
''					- <em>4</em> (warn): messages which warn the developer. e.g. a method is obsolete
''					- <em>8</em> (error): messages which contain an error
''					Which messages are actually logged can be set with the <em>logLevel</em> property directly or in the config with <em>AJAXED_LOGLEVEL</em>.
''					Simple logging within ajaxed:
''					<code>
''					<%
''					lib.logger.debug("some debug message")
''					lib.logger.warn("a warning like e.g. use method A instead of B")
''					lib.logger.info("user logged in")
''					lib.logger.error("some error happend")
''					% >
''					</code>
''					- Logging is disabled by default
''					- Logger supports ascii coloring => log files are easier to read. 
''					- Its recommended to download e.g. "cygwin" and hook up the log file with <em>"tail -f file.log"</em>. This allows you to immediately follow the log and it supports coloring as well.
''					- Its also possible to view the changes log files directly in the ajaxed console
''					- Logfiles are named after the environment. <em>dev.log</em> and <em>live.log</em>
''					- Within the ajaxed library <em>lib.logger</em> holds an instance of a ready-to-use logger
''					- Some useful debug messages are already automatically logged within the ajaxed library. e.g. SQL queries, page requestes, ajaxed callbacks, emails, ...
''					- Log files can also be found in the ajaxed console.
''					- It might be necessary that the logs path has write permission for the IIS user
''					- All non-ascii chars (<em>\u0100-\uFFFF</em>) will be converted to its hex notation (e.g. <em>\uF9F9</em>) as most shells cannot display unicode
'' @VERSION:		1.1

'**************************************************************************************************************
class Logger

	'private members
	private p_prefix, escapeChar, extension, p_path, regx
	
	'public members
	public msgPrefix	''[string] the prefix for all log messages. by default its the users ip and time stamp followed by a tabulator
	public defaultStyle	''[string] the default appearance of log messages. check doc in <em>log()</em> method. leave empty if you dont want to use styles at all. This will keep your file clean if you prefer to read them with a common text editor
	public logLevel		''[int] indicates the logger level. disabled = <em>0</em>, debug = <em>1</em>, info = <em>2</em>, warn = <em>4</em>, error = <em>8</em>
						''- all messages with a number than the <em>loglevel</em> will be logged as well. Lower numbers wont! So e.g. setting level to <em>2</em> (info) will log info-, warn-, and error messages
						''- by default it is set to <em>0</em> (disabled)
	
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
		'if not set by the config then deactivate the logger by default
		if logLevel = -1 then logLevel = 0
		set regx = new Regexp
		regx.global = true
		regx.ignoreCase = false
		'this pattern is being used to replace all non ascii chars.
		'NOTE: there are 4 chars which are UTF8 stooges (check http://www.xbeat.net/vbspeed/c_UCase.htm)
		'they behave differently to others when using ascw or chrw... This is considered in the pattern
		regx.pattern = "[\u0100-\uFFFF\u00" & hex(ascw(chrw(154))) & "\u00" & hex(ascw(chrw(156))) & "\u00" & hex(ascw(chrw(158))) & "\u00" & hex(ascw(chrw(159))) & "]"
	end sub
	
	public property get path ''[string] gets the virtual path where the logfiles are located. e.g. <em>/ajaxedLogs/</em>
		path = p_path
	end property
	
	public property let prefix(val) ''[string] sets the prefix for the log file. Useful if you have more applications and want to have the logs separated. e.g. your app is named <em>"app1"</em> then it will result in a logfile <em>app1_dev.log</em> on the <em>dev</em> env. default = EMPTY
		prefix = val & "_"
	end property
	
	public property get logfile ''[string] gets the full path and name of the logfile
		logfile = path & p_prefix & lCase(lib.env) & extension
	end property
	
	'**************************************************************************************************************
	'' @SDESCRIPTION: 	logs a message
	'' @PARAM:			level [int]: level of logging. possible: <em>0</em> (disabled), <em>1</em> (debug), <em>2</em> (info), <em>4</em> (warn), <em>8</em> (error)
	''					- its depending on the loggerLevel setting which level will be logged
	'' @PARAM:			msg [string], [array]: the message you want to log. if ARRAY then each field is treated as a line but only the first one has the message prefix
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
			file.write(encode(msg(i)))
			if i = uBound(msg) then file.write(getStyle(defaultStyle))
			file.write(vbNewLine)
		next
		file.close()
		set file = nothing
	end sub
	
	'******************************************************************************************************************
	'* encodes all non ascii chars within the logged string
	'******************************************************************************************************************
	private function encode(val)
		encode = val
		'replace all UTF-8 chars to its hex notation
		if regx.test(val) then
			'we store the chars in a dictionary to replace them all in one go later
			set chrs = lib.newDict(empty)
			for each m in regx.execute(val)
				'TODO: actually there is a bug with utf8 stooges (see constructor for more details)
				'there are 4 chars which are being displayed wrong
				if not chrs.exists(m.value) then chrs.add m.value, "\u" & str.padLeft(hex(ascw(m.value)), 4, 0)
			next
			for each k in chrs.keys
				encode = replace(encode, k, chrs(k))
			next
		end if
	end function
	
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
	'' @PARAM:			msg [string], [array]: message(s) to debug. (if ARRAY then treated as lines)
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