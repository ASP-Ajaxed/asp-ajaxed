<!--#include file="class_textTemplateBlock.asp"-->
<%
'**************************************************************************************************************

'' @CLASSTITLE:		TextTemplate
'' @CREATOR:		Michal Gabrukiewicz - gabru at grafix.at
'' @CREATEDON:		24.10.2003
'' @CDESCRIPTION:	Represents a textbased template which can be used as content for emails, etc.
''					It uses a file and replaces given placeholders with specific values. Placeholders can
''					be common name value pairs or even whole blocks which hold name value pairs and can be
''					duplicated several times. It's possible to create, modify and delete the templates.
''					Example for the usage as an email template (first line of the template is used as subject):
''					This is how a template would look like:
''					<code>
''					Welcome!
''					Dear <<< name >>>,
''					Thanks joining us!
''					We belive your age must be <<< age | not specified >>>.
''					</code>
''					The following code is required to us the template above:
''					<code>
''					<%
''					set t = new TextTemplate
''					t.filename = "/templateFileName.txt"
''					t.add "name", "John Doe"
''					email.subject = t.getFirstLine()
''					email.body = t.getAllButFirstLine()
''					% >
''					</code>
''					<strong>Note:</strong> The age placeholder has not been used in the code.
''					For that reason the <em>TextTemplate</em> will return its default value which can be
''					specified after the <em>|</em> (pipe) sign. It can be read as 'display age or "not specified" if no age available'
'' @REQUIRES:		-
'' @VERSION:		1.3

'**************************************************************************************************************

class TextTemplate

	private vars, blocks, p_content, regex, blockBegin, blockEnd, validVarnamesPattern
	private defaultValueSeparator
	
	public fileName				''[string] The virtual path including the filename of your template. e.g. <em>/userfiles/t.html</em>
	public placeHolderBegin		''[string] If you want to use your own placeholder characters. this is the beginning. e.g. <em>&lt;&lt;&lt;</em>
	public placeHolderEnd		''[string] If you want to use your own placeholder characters. this is the ending. e.g. <em>&gt;&gt;&gt;</em>
	public cleanParse			''[bool] should all unused blocks, vars, etc. been removed after parsing? default = TRUE
	public UTF8					''[bool] is the template saved as UTF8 and should it be stored as UTF8? default = TRUE
	
	public property get content	''[string] If no content is provided, we load the conents of the given file
		if p_content = "" then p_content = getContents()
		content = p_content
	end property
	
	public property let content(val) ''[string] Provide your own content.
		 p_content = val
	end property
	
	'******************************************************************************************************************
	'* constructor 
	'******************************************************************************************************************
	public sub class_Initialize()
		set vars = server.createObject("Scripting.Dictionary")
		set blocks = server.createObject("Scripting.Dictionary")
		fileName = empty
		p_content = ""
		placeHolderBegin = "<<< "
		placeHolderEnd = " >>>"
		defaultValueSeparator = " \| "
		blockBegin = "BLOCK"
		blockEnd = "ENDBLOCK"
		validVarnamesPattern = "[(a-zA-Z0-9)|_]+"
		set regex = new RegExp
		regex.ignoreCase = true
		regex.global = true
		cleanParse = true
		UTF8 = true
	end sub
	
	'******************************************************************************************************************
	'* destructor 
	'******************************************************************************************************************
	public sub class_terminate()
		set vars = nothing
		set blocks = nothing
		set regex = nothing
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	Adds a variable/block which should be replaced within the template
	'' @DESCRIPTION:	All placeholders in the template using this name (e.g. <em>&lt;&lt;&lt; VARNAME &gt;&gt;&gt;</em>) will be replaced by the
	''					value of the given variable. if the value was already added then it will be updated by the new value
	'' @PARAM:			varName [string]: The name of your variable
	''					when providing a Block the varname is the name of the block you want to add.
	'' @PARAM:			varValue [string], [TextTemplateBlock]: The value which will be used.
	''					if its a block then provide a <em>TextTemplateBlock</em> instance
	'******************************************************************************************************************
	public sub add(varName, varValue)
		'allow only letters, numbers and "_"
		regex.pattern = "^" & validVarnamesPattern & "$"
		if not regex.test(varName) then lib.error("varName (" & varname & ") can only contain letters.")
		
		'we store all vars in uppercase in our collections
		var = uCase(varName)
		
		'we check where we need to add the variable.
		'to the common ones or to the blocks
		if isObject(varValue) then
			set container = blocks
		else
			set container = vars
		end if
		
		'we need to update (remove and add again) the value if the name already exists
		if container.exists(var) then vars.remove(var)
		container.add var, varValue
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	Alias for <em>add()</em>.
	'******************************************************************************************************************
	public sub addVariable(varName, varValue)
		add varName, varValue
	end sub
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	Returns a string where the template and the placeholders are merged
	'******************************************************************************************************************
	public function returnString()
		'now replace the placeholders and return the whole String
		toReturn = parseTemplate(content)
		'we return what should be returned
		returnString = toReturn
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	Returns the first line of the template file
	'' @DESCRIPTION:	Returns the content of the first line of the template file. 
	'' 					The place holders will be parsed as well
	'' @RETURNS:		varValue [string]: The parsed first line of the file
	'******************************************************************************************************************
	public function getFirstLine()
		s = returnString()
		lines = split(s, chr(13))
		getFirstLine = lines(0)
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	Returns the parsed file without the first line
	'' @RETURNS:		varValue [string]: The parsed file without the first line
	'******************************************************************************************************************
	public function getAllButFirstLine()
		s = returnString()
		lines = split(s, chr(13))
		for i = 1 to UBound(lines)
  			getAllButFirstLine = getAllButFirstLine & lines(i)
		next
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	Saves the text template to the template directory
	'**********************************************************************************************************
	public function save()
		fileToOpen = server.MapPath(fileName)
		if UTF8 then
			with server.createobject("ADODB.Stream")
				.charset = "utf-8"
				.open()
				.writeText(p_content)
				.saveTofile fileToOpen, 2
				.close()
			end with
		else
			set myfile = lib.FSO.openTextfile(fileToOpen, 2, true)
			myFile.write(p_content)
			myFile.close()
			set myfile = nothing
		end if
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	Deletes the text template file
	'**********************************************************************************************************
	public function delete()
		lib.FSO.deleteFile server.MapPath(fileName), true
	end function
	
	'******************************************************************************************************************
	'* parses a string and replaces all placeholders, blocks, etc with the values of defined in the class.
	'******************************************************************************************************************
	private function parseTemplate(input)
		parseTemplate = input
		
		'first parse the blocks, that no vars will be mixed up with the common ones afterwards
		for each blockname in blocks.keys
			set block = blocks(blockname)
			regex.pattern = getBlockPattern(blockname)
			'extract the block from the string
			set matches = regex.execute(parseTemplate)
			'we walk through all matches and through all items of the block and concate the output to
			'one string so we can replace it then instead of the block
			for each m in matches
				parsedBlock = ""
				for each item in block.items.items
					parsedItem = m.value
					for i = 0 to uBound(item) step 2
						'concat up each item
						parsedItem = replacePlaceHolders(parsedItem, item(i), item(i + 1))
					next
					parsedBlock = parsedBlock & parsedItem
				next
				'remove the block definition (block and endblock)
				parsedBlock = replacePlaceHolders(parsedBlock, blockBegin & " " & blockname, "")
				parsedBlock = replacePlaceHolders(parsedBlock, blockEnd & " " & blockname, "")
				'write back ..
				parseTemplate = replace(parseTemplate, m.value, parsedBlock)
			next
		next
		
		'replace all variables not in any block
		for each varName in vars.keys
			parseTemplate = replacePlaceHolders(parseTemplate, varName, vars(varName) & "")
		next
		
		if cleanParse then
			'remove all unused blocks
			regex.pattern = getBlockPattern(validVarnamesPattern)
			parseTemplate = regex.replace(parseTemplate, "")
			'remove all unused vars
			parseTemplate = replacePlaceHolders(parseTemplate, validVarnamesPattern, "")
		end if
	end function
	
	'******************************************************************************************************************
	'* replaces a placeholder in a given string by a value. so from <<< NAME >>> will be made e.g. "Michal"
	'******************************************************************************************************************
	private function replacePlaceHolders(input, varName, varValue)
	    regex.pattern = placeHolderBegin & varName & "(" & defaultValueSeparator & "(.+))?" & placeHolderEnd
	    'replace the placeholder with the value (if available) otherwise replace it with the default value ($2)
	    replacePlaceHolders = regex.replace(input & "", str.parse(varValue, "$2") & "")
	end function
	
	'******************************************************************************************************************
	'* gets the pattern which matches block(s). blockName can be any specifc block or even a pattern..
	'* Example <<< BLOCK X >>> .. <<< ENDBLOCK X >>>
	'******************************************************************************************************************
	private function getBlockPattern(blockName)
		getBlockPattern = placeHolderBegin & blockBegin & _
							" (" & blockName & ")" & placeHolderEnd & "[\s\S]*?" & _
							placeHolderBegin & blockEnd & " (\1)" & placeHolderEnd
	end function
	
	'******************************************************************************************************************
	'* LoadContents 	Loads the contents of the given file
	'******************************************************************************************************************
	private function getContents()
		getContents = ""
		fileToOpen = server.MapPath(fileName)
		
		'we check if file exists
		if lib.FSO.FileExists(fileToOpen) then
			if UTF8 then
				with server.createObject("ADODB.Stream")
					.charset = "utf-8"
					.open()
					.loadFromFile(fileToOpen)
					getContents = .readText(-1)
					.close()
				end with
			else
				set myfile = lib.FSO.openTextfile(fileToOpen, 1, false)
				'now we read every line from the file
				getContents = myfile.readAll()
				myfile.close()
				set myfile = nothing
			end if
		end if
	end function

end class
%>