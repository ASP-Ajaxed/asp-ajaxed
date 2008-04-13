﻿<!--#include file="../../ajaxed.asp"-->
<%
'******************************************************************************************
'* Creator: 	David Rankin, adapted by Michal
'* Created on: 	2006-10-25 08:53
'* Description: ajaxed documentor. creates a documentation for a given folder on the server
'*				and saves it into that folder with the filename doc.html
'*				TODO: parse new lines and lists (so documentation is formatted nicer)
'* Input:		POST params: folder (folder to parse - if no given then "ajaxed" is used)
'******************************************************************************************

set	classNames = server.createObject("scripting.Dictionary")
set xml = server.createObject("MSXML2.DOMDocument.3.0")
set page = new AjaxedPage
with page
	.plain = true
	.onlyDev = true
	.draw()
end with
set page = nothing

'******************************************************************************************
'* main 
'******************************************************************************************
sub main()
	if not lib.fso.folderExists(getFolder(true)) then lib.error("Folder '" & getFolder(false) & "' could not be found on the server.")
	
	xml.async = false
	set pi = xml.createProcessingInstruction("xml", "version=""1.0""")
	xml.insertBefore pi, xml.childNodes.item(0)
	set pi = nothing
	set classesNodes = getNewNode("classes" , "")
	classesNodes.setAttribute "createdOn", now()
	xml.appendChild(classesNodes)
	
	readFolder(lib.fso.getFolder(getFolder(true)))
	
	findTypesIn("requires/types/type")
	findTypesIn("return/type")
	findTypesIn("parameter/types/type")
	findTypesIn("friendof/types/type")
	findTypesIn("property/types/type")
	
	set xsl = server.createObject("MSXML2.DOMDocument.3.0")
	xsl.async = false
	xsl.load(server.mapPath("style.xsl"))
	'if its the ajaxed folder, then we call the documentation index.html
	filename = lib.iif(getFolder(false) = "/ajaxed/", "index.html", "doc.html")
	set f = lib.fso.createTextFile(getFolder(true) & "\" & filename, true)
	f.write(xml.transformNode(xsl))
	f.close()
	set f = nothing
	set xsl = nothing
	set xml = nothing
	str.writeln("Documentation successfully created. It can be found <a href=""" & getFolder(false) & "" & filename & """>here</a>")
end sub

'******************************************************************************************
'* getFolder 
'******************************************************************************************
function getFolder(absolute)
	getFolder = page.RF("folder")
	if getFolder = "" then getFolder = "/ajaxed/"
	getFolder = str.ensureSlash(getFolder)
	if absolute then getFolder = server.mappath(getFolder)
end function

'******************************************************************************************
'* readFolder Reads all files in a folder, and its sub folders, subfolders, subfolders
'******************************************************************************************
function readFolder(folder)
	if not includeVersions() and str.startsWith(ucase(folder.name), "V")  then exit function
	for each file in folder.files
		parseFile file
 	next
	for each subfolder in folder.SubFolders
	 	readFolder(subfolder)
	next
end function

'******************************************************************************************
'* findTypesIn 
'******************************************************************************************
function findTypesIn(xPath)
	'read through xml doc, and find requires classes. 
	'when you find a class, if the name is in the dictionary, then add the id attribute
	set nList = xml.documentElement.getElementsByTagName(xPath)
	for i = 0 To (nList.length - 1)
		aName = ucase(Trim(nList.Item(i).text)) & ""
		if classNames.exists(aName) then
			nList.Item(i).setAttribute "id", classNames(aName)
		end if
  	next
end function

'******************************************************************************************
'* excludeVersions
'******************************************************************************************
function includeVersions()
	'this comes from the former gablibrary which had versions for components.
	'this is not yet supported by AJAXED
	includeVersions = false
end function

'******************************************************************************************
'* getNewNode	Returns a named node witht eh given text
'******************************************************************************************
function getNewNode(name, value)
	set node = xml.createNode(1, name, "")
	node.text = trim(value)
	set getNewNode = node
end function

'******************************************************************************************
'* parseFile	Parses the ASP class into an XML file
'******************************************************************************************
function parseFile(file)
	if not str.endsWith(file.name, ".asp") then exit function
	if file.size = 0 then exit function
	set reader = file.OpenAsTextStream(1)
	contents = reader.readAll()
	set reader = nothing
	if instr(contents, "'' @CLASSTITLE:") then
		enforceRulz = false
		contents = replace(contents, vbtab, "")
		lines = split(contents, vbcrlf)
		
		'Class details
		title = ""
		creator = ""
		createdon = ""
		description = ""
		staticname = ""
		postfix = ""
		version = ""
		compatibility = ""
		requires = ""
		obsolete = ""
		friendof = ""
		for i = 0 to ubound(lines)
			aLine = trim(lines(i)) & ""
			if str.startsWith(aLine, "'' @CLASSTITLE:") then
				title = replace(aLine, "'' @CLASSTITLE:", "")
				if instr(title, "<<< CLASSTITLE >>>") then exit function
				if not ClassNames.Exists(ucase(trim(title))) then 
					ClassNames.add ucase(trim(title)) & "", lib.getUniqueID()
				end if
			elseif str.startsWith(aLine, "'' @CREATOR:") then
				creator = trim(replace(aLine, "'' @CREATOR:", ""))
			elseif str.startsWith(aLine, "'' @CREATEDON:") then
				createdon = trim(replace(aLine, "'' @CREATEDON:", ""))
			elseif str.startsWith(aLine, "'' @CDESCRIPTION:") then
				' A long description so get comments in lines below
				description = trim(replace(aLine, "'' @CDESCRIPTION:", ""))
				for j = (i + 1) to ubound(lines)
					if str.startsWith(lines(j), "''") and not str.startsWith(lines(j), "'' @") then
						tmp = replace(lines(j), "''", "")
						description = description & vbcrlf & trim(tmp)
						else
						j = ubound(lines)
					end if
				next
			elseif str.startsWith(aLine, "'' @STATICNAME:") then
				staticname = trim(replace(aLine, "'' @STATICNAME:", ""))
			elseif str.startsWith(aLine, "'' @POSTFIX:") then
				postfix = trim(replace(aLine, "'' @POSTFIX:", ""))
			elseif str.startsWith(aLine, "'' @VERSION:") then
				version = trim(replace(aLine, "'' @VERSION:", ""))
			elseif str.startsWith(aLine, "'' @COMPATIBLE:") then
				compatibility = trim(replace(aLine, "'' @COMPATIBLE:", ""))
			elseif str.startsWith(aLine, "'' @REQUIRES:") then
				requires = trim(replace(aLine, "'' @REQUIRES:", ""))
			elseif str.startsWith(aLine, "'' @FRIENDOF:") then
				friendof = trim(replace(aLine, "'' @FRIENDOF:", ""))
			end if
		next
		
		if title = "" then exit function
			set classNode = getNewNode("class", "")
			with classNode
			.appendChild(getNewNode("name", title))
			.appendChild(getNewNode("author", creator))
			if instr(description,"OBSOLETE!") then classNode.setAttribute "obsolete", "1"
			if ClassNames.Exists(ucase(title)) then
				'Creates a unique ID for the class
				.setAttribute "id", ClassNames.item(uCase(title))
			end if
			if trim(description) = "" and EnforceRulz then 
				lib.error("Error in class:" & getVirtualPath(file.path) & "<br> - No Desciption for class <b>" & title & "</b>")
			end if
			.appendChild(getNewNode("description", description))
			.appendChild(getNewNode("staticname", staticname))
			.appendChild(getNewNode("postfix", postfix))
			.appendChild(getNewNode("version", version))
			if not trim(compatibility) = "" then
				set CompatibleNode = getNewNode("compatible","")
				browsers = split(compatibility,",")
				for a = 0 to ubound(browsers) 
					CompatibleNode.appendChild(getNewNode("browser", trim(browsers(a))))
				next
				.appendChild(CompatibleNode)
				set CompatibleNode = nothing
			end if
			requires = replace(requires, "-", "")
			if not trim(requires) = "" then
				set requiresNode = getNewNode("requires","")
				classes = split(requires,",")
				set rTypes = getNewNode("types","")
				for a = 0 to ubound(classes) 
					rTypes.appendChild(getNewNode("type", trim(classes(a))))
				next
				requiresNode.appendChild(rTypes)
				.appendChild(requiresNode)
				set rTypes = nothing
				set requiresNode = nothing
			end if
			Set fs = Server.CreateObject("Scripting.FileSystemObject")
			if fs.FolderExists(file.parentFolder & "\demo\") then
				.appendChild(getNewNode("demo", getVirtualPath(file.parentFolder & "\demo\")))
			end if
			.appendChild(getNewNode("path", getVirtualPath(file.path)))
			.appendChild(getNewNode("created", createdon))
			.appendChild(getNewNode("modified", file.dateLastModified))
			
			'Friend of type, only on type eyxpected ATM
			if not trim(friendof) = "" then
				set nodefriendof = getNewNode("friendof", "")
				set FTypes = getNewNode("types", "")
				set FType = getNewNode("type", trim(friendof))
				FTypes.appendChild(FType)
				nodefriendof.appendChild(FTypes)
				.appendChild(nodefriendof)
				set FTypes = nothing
				set FType = nothing
				set nodefriendof = nothing
			end if
			Set fs = nothing
		end with
		
		'read constructor
		set initsNode = getNewNode("initializations", "")
		for i = 0 to ubound(lines)
			aLine = trim(lines(i)) & "" 
			aLine = replace(aLine, vbtab, "")
			lineUpper = ucase(aLine)
			if (str.startsWith(lineUpper, "PUBLIC SUB CLASS_INITIALIZE")) then 
				'in constructor
				for j = i to ubound(lines)
					aLine = trim(lines(j)) & "" 
					aLine = replace(aLine, vbtab, "")
					lineUpper = ucase(aLine)
					if (InStr(lineUpper,"LIB.INIT")) then
						set iNode = getNewNode("init", "")
						'we have an init line
						aLine = replace(aLine, "=", "")
						lineElements = split(aLine, " ")
						'appends property, the first element in the contructor
						iNode.appendChild(getNewNode("property", lineElements(0)))
						Set regex = New RegExp
						regex.pattern = "\(.*\)"
						set matched = regex.execute(aLine)
						for each ma in matched
							inits = cstr(ma)
							inits = Mid(inits, 2, (len(inits)-2)) 'Remove fist and last bracket only ()
							inits = replace(inits, """", "")
							a = split(inits, ",")
							if ubound(a) > 0 then
								iNode.appendChild(getNewNode("constant", a(0)))
								iNode.appendChild(getNewNode("default", a(1)))
							end if
						next
						initsNode.appendChild(iNode)
					end if
					if (str.startsWith(lineUpper, "END SUB")) then 
						i = ubound(lines) 'exit contstructor loop
						exit for 'exit this loop
					end if
				next
			end if
		next
		classNode.appendChild(initsNode)
		
		'get the methods
		set methodsNode = getNewNode("methods", "")
		set classProperties = getNewNode("properties","")
		for i = 0 to ubound(lines)
			aLine = trim(lines(i)) & "" 
			aLine = replace(aLine, vbtab, "")
			if str.startsWith(aLine, "public") then
				aLine = replace(aLine ,"public ", "")
				if (str.startsWith(aLine, "function ")) or (str.startsWith(aLine, "sub ")) then
					ExpectReturn = false
					set initializeVars = server.createObject("Scripting.Dictionary")
					if str.startsWith(aLine, "function ") then ExpectReturn = true
					aLine = replace(aLine ,"function ", "")
					aLine = replace(aLine, "sub ", "")
					lineUpper = ucase(aLine)
					'if its the constructor or destructor then skip this line...
					if not (str.startsWith(lineUpper, "CLASS_TERMINATE") or str.startsWith(lineUpper, "CLASS_INITIALIZE")) then
						'replace some keywords from the params.
						Set regex = New RegExp
						regex.pattern = "\(.*\)"
						'get name
						methodName = regex.replace(aLine, "")
						'get params
						regex.pattern = "\(.*\)"
						set matched = regex.execute(aLine)
						 'Step through our matches 
						 set params = server.createObject("Scripting.Dictionary")
						 For Each item in matched 
						 	x = split(item.Value, ",")
							for a = 0 to ubound(x)
								x(a) = replace(x(a),"(","")
								x(a) = replace(x(a),")","")
								x(a) = trim(x(a))
								if not x(a) = "" then params.add trim(x(a)) & "", ""
							next
						 Next
						sdescription = ""
						parameter = ""
						ldescription = ""
						returns = ""
						'get the details about the method
						for j = (i - 1) to 0 Step -1
							aline = trim(lines(j))
							aLine = replace(aLine, vbtab, "")
							if str.startsWith(aLine, "'") then 
								if str.startsWith(aLine, "'' @PARAM:") then
									aLine = replace(aline,"'' @PARAM:", "")
									aLine = trim(replace(aLine, vbtab, ""))
									if str.startsWith(aLine, "-") then aLine = str.trimStart(aLine, 1)
									aLine = trim(aLine)
									'get all lines for the description, if the description is breaked up in new lines
									parameterFor = split(aLine," ")
									'remove the parameter from the start of the parameter description
									if ubound(parameterFor) >=0 then
										aLine = str.trimStart(aLine, len(trim(parameterFor(0))))
									end if
									parameter = trim(aLine)
									for k = (j + 1) to i
										addLine = trim(lines(k))
										if str.startsWith(addLine, "''") and not str.startsWith(addLine, "'' @") then
											lines(k) = replace(lines(k), vbtab, "")
											lines(k) = replace(lines(k),"''", "")
											parameter = parameter  & vbnewline & trim(lines(k))
										else
											exit for
										end if
									next
									'loop through parameters found in method, and add description
									for each key in params.keys
										akey = trim(replace(ucase(key),"BYREF",""))
										if ubound(parameterFor) >=0 then
											if akey = uCase(parameterFor(0)) then
												params(key) = parameter
											end if
										end if
									next
									
								end if
								if str.startsWith(aLine, "'' @RETURN:") then
									aLine = replace(aline,"'' @RETURN:", "")
									aLine = replace(aLine, vbtab, "")
									returns = trim(aLine)
									'get all lines for the description, if the description is breaked up in new lines
									for k = (j + 1) to i
										addLine = trim(lines(k))
										if str.startsWith(addLine, "''") and not str.startsWith(addLine, "'' @") then
											lines(k) = replace(lines(k), vbtab, "")
											lines(k) = replace(lines(k),"''", "")
											returns = returns  & vbnewline & trim(lines(k))
										else
											exit for
										end if
									next
								end if
								if str.startsWith(aLine, "'' @DESCRIPTION:") then
									aLine = replace(aline,"'' @DESCRIPTION:", "")
									
									aLine = replace(aLine, vbtab, "")
									ldescription = trim(aLine)
									'get all lines for the description, if the description is breaked up in new lines
									for k = (j + 1) to i
										addLine = trim(lines(k))
										addLine = replace(addLine, vbtab, "")
										if str.startsWith(addLine, "''") and not str.startsWith(addLine, "'' @") then
											lines(k) = replace(lines(k), vbtab, "")
											lines(k) = replace(lines(k),"''", "")
											ldescription = ldescription  & vbnewline & trim(lines(k))
										else
											exit for
										end if
									next
								end if
								if str.startsWith(aLine, "'' @SDESCRIPTION:") then
									aLine = replace(aline,"'' @SDESCRIPTION:", "")
									aLine = replace(aLine, vbtab, "")
									sdescription = trim(aLine)
									'get all lines for the description, if the description is breaked up in new lines
									for k = (j + 1) to i
										addLine = trim(lines(k))
										addLine = replace(addLine, vbtab, "")
										if str.startsWith(addLine, "''") and not str.startsWith(addLine, "'' @") then
											lines(k) = replace(lines(k), vbtab, "")
											lines(k) = replace(lines(k),"''", "")
											sdescription = sdescription  & vbnewline & trim(lines(k))
										else
											exit for
										end if
									next
								end if
							else
								exit for
							end if
						next
						set methodNode = getNewNode("method", "")
						with methodNode
							.appendChild(getNewNode("name", methodName))
							if instr(sdescription,"OBSOLETE!") then methodNode.setAttribute "obsolete", "1"
							if instr(sdescription,"[static]") or  instr(ldescription,"[static]") _
								or instr(sdescription,"STATIC!") or  instr(ldescription,"STATIC!") then methodNode.setAttribute "static", "1"
							.appendChild(getNewNode("shortdescription", sdescription))
							if instr(ldescription,"OBSOLETE!") then methodNode.setAttribute "obsolete", "1"
							.appendChild(getNewNode("longdescription", ldescription))
							if ExpectReturn and trim(returns) = "" and EnforceRulz then
								lib.error("Error in class:" & getVirtualPath(file.path) & "<br> - No return parameter provided for method<b> " & methodname & " </b>in class <b>" & title & "</b>")
							end if
							returnType = getVarType(returns)
							returns = removeSquareBrakets(returns)
						end with
						set Parameters = getNewNode("parameters", "")
						for each key in params.keys
							aType = ""
							
							set aParameter = getNewNode("parameter", "")
							Set aregex = New RegExp
							aregex.IgnoreCase = True
							aregex.pattern = "(^byRef )|(^byval )"
							akey = aregex.replace(key,"")
							set aName	= getNewNode("name", trim(akey))
							if  instr(ucase(key), "BYREF") > 0 then aName.setAttribute "passed", "byRef"
							if instr(ucase(key), "BYVAL") > 0 then aName.setAttribute "passed", "byVal"
							
							paramDesc = params(key)
							if trim(paramDesc) = "" and EnforceRulz then 
								errorMessage = "Error in class:" & getVirtualPath(file.path) & "<br> - No Desciption for parameter <b>" & akey &"</b> in class <b>" & title & "</b>"
								errorMessage = errorMessage & " method <b>" & methodName & "</b>"
								lib.error(errorMessage)
							end if
							set aDescription = getNewNode("description", removeSquareBrakets(paramDesc))
							aParameter.appendChild(aName)
							aParameter.appendChild(aDescription)
							'types
							set typesNode = getNewNode("types", "")
							types = split(getVarType(paramDesc),",")
							if ubound(types) >= 0 then
								
								for a = 0 to ubound(types)
									if not validateType(trim(types(a))) and EnforceRulz then
										errorMessage = "Error in class:" & getVirtualPath(file.path) & "<br>Invalid type name <b>" & types(a) &"</b> in class <b>" & title
										errorMessage = errorMessage & "</b> method <b>" & methodname & "</b>"
										lib.error(errorMessage)
									end if
									set typeNode = getNewNode("type", trim(types(a)))
									typesNode.appendChild(typeNode)
								next
							end if
							aParameter.appendChild(typesNode)
							Parameters.appendChild(aParameter)
							set aParameter = nothing
							set aName	= nothing
							set aDescription = nothing
						next
						methodNode.appendChild(parameters)
						set returnNode = getNewNode("return", "")
							
						if not validateType(returnType) and EnforceRulz and ExpectReturn then
							errorMessage = "Error in class:" & getVirtualPath(file.path) & "<br> - Invalid return type name<b>" & returnType & "</b> in class <b>" & title
							errorMessage = errorMessage & "</b> method <b>" & methodname & "</b>"
							lib.error(errorMessage)
						end if
						returnNode.appendChild(getNewNode("description", returns))
						returnNode.appendChild(getNewNode("type", returnType))
						methodNode.appendChild(returnNode)
						methodsNode.appendChild(methodNode)
					end if
					classNode.appendChild(methodsNode)
				else
					
					set classProperty = getNewNode("property","")
					'get the properties that start with public property
					if str.startsWith(aLine, "property get ") or str.startsWith(aLine, "property let ") or str.startsWith(aLine, "property set") or str.startsWith(aLine, "default property")then
						description = ""
						isLetProperty = str.startsWith(aLine, "property let ")
						isGetProperty = str.startsWith(aLine, "property get ")
						isSetProperty = str.startsWith(aLine, "property set ")
						aLine = replace(aLine, "property get ", "")
						aLine = replace(aLine, "property let ", "")
						aLine = replace(aLine, "property set ", "")
						propertyInfo = split(aLine, "''")
						propertyName = trim(propertyInfo(0))
						Set regex = New RegExp
						regex.pattern = "\(.*\)"
						'get name
						propertyName = regex.replace(propertyName, "")
						set matched = regex.execute(aLine)
						if uBound(propertyInfo) > 0 then
							if isLetProperty or isSetProperty then
								description = "SET: " + trim(propertyInfo(1))
							else
								description = "GET: " + trim(propertyInfo(1))
							end if
						end if
						'we check if a get or set property is already here with this name.
						'when yes then we update its description
						set pNode = classProperties.selectSingleNode("property[name='" & propertyName & "']")
						if not pNode is nothing then
							set anotherNode = pNode.getElementsByTagName("description")
							description = anotherNode(0).text & vbnewline &  description
							classProperties.removeChild(pnode)
						end if
					else
						'get the properties that are just "public"
						memberInfo = split(aLine,"''")
						propertyName = trim(memberInfo(0))
						propertyName = replace(propertyName, "property get", "")
						propertyName = replace(propertyName, "property let", "")
						propertyName = replace(propertyName, "property set", "")
						if ubound(memberInfo) > 0 then
							description = trim(memberInfo(1))
							'we check if there are any comments for the membervariable in the next lines
							'so the comment is breaked into new lines...
							for j = (i + 1) to  ubound(lines)
								aline = trim(lines(j))
								addLine = replace(addLine, vbtab, "")
								if str.startsWith(aline, "''") then
									description = description  & vbnewline & replace(aline, "''", "")
								else
									exit for
								end if
							next
						end if
					end if
					
					'Now set the details for the class property node
					with classProperty
					if str.startsWith(propertyName, "default ") then
						.setAttribute "defaultProperty", "1"
						propertyName = trim(str.trimStart(propertyName, 8))
					end if
					.appendChild(getNewNode("name", propertyName))
					if instr(description,"GET") then
						.setAttribute "readOnly", "1"
					elseif instr(description,"SET") then
						.setAttribute "writeOnly", "1"
					end if
					
					if instr(description,"OBSOLETE!") then classProperty.setAttribute "obsolete", "1"
					set typesNode = getNewNode("types", "")
					types = split(getVarType(description),",")
					for a=0 to ubound(types)
						if not trim(types(a)) = "" then
							if not validateType(types(a)) and enforceRulz then
								errorMessage = "Error in class:" & getVirtualPath(file.path) & "<br>Invalid type name <b>" & types(a) &"</b> in class <b>" & title
								errorMessage = errorMessage & "</b> method <b>" & methodname & "</b>"
								lib.error(errorMessage)
							end if
							set typeNode = getNewNode("type", trim(types(a)))
							typesNode.appendChild(typeNode)
							set typeNode = nothing
						end if
					next
					classProperty.appendChild(typesNode)
					set typesNode = nothing
					if trim(removeSquareBrakets(description)) = "" and enforceRules then 
						lib.error("Error in class:" & getVirtualPath(file.path) & "<br>No description for <b>" & propertyName & "</b> in class <b>" & title  & "</b>")
					end if
					.appendChild(getNewNode("description", removeSquareBrakets(description)))
					end with
					classProperties.appendChild(classProperty)
				end if
				classNode.appendChild(classProperties)
				
			end if
			xml.documentElement.appendChild(classNode)
		next
		set classProperties = nothing
		set classProperty = nothing
		set classNode = nothing
	end if
end function

'******************************************************************************************
'* removeSquareBrakets 	' removes all from and including the sqaure brakets
'******************************************************************************************
function removeSquareBrakets(value)
	removeSquareBrakets = ""
	if trim(value) = "" then exit function
	set regex = New RegExp
	regex.pattern = "\[.*\]"
	removeSquareBrakets = regex.replace(value,"")
	if str.startsWith(removeSquareBrakets, ":") then removeSquareBrakets = str.trimStart(removeSquareBrakets, 1)
	set matched = nothing
	set regex = nothing
end function

'******************************************************************************************
'* function 
'******************************************************************************************
function validateType(typeValue)
	validateType = false
	if trim(typeValue) = "" then exit function
	set regex = New RegExp
	regex.pattern = "\W+"
	validateType = not regex.test(trim(typeValue))
	set regex = nothing
end function

'******************************************************************************************
'* getVarType 	pass a string with the type in [] and get the type name back. 
'*				for example "- [int] num" returns "int"
'******************************************************************************************
function getVarType(value)
	getVarType = ""
	if trim(value) = "" then exit function
	set regex = New RegExp
	regex.pattern = "\[.*\]"
	set matched = regex.execute(value)
	For Each item in matched
		aType = item.Value
		aType = replace(aType,"[","")
		aType = replace(aType,"]","")
		aType = replace(aType,":","")
		getVarType = getVarType & aType & ", "
	Next
	if len(getVarType) > 2 then getVarType = str.trimEnd(getVarType, 2)
	set matched = nothing
	set regex = nothing
end function

'******************************************************************************************
'* getVirtualPath	Returns the virtual path, based on a given path
'******************************************************************************************
function getVirtualPath(path)
	getVirtualPath = ""
	if trim(path) = "" then exit function
	serverRoot = server.mapPath("/")
	physicalPath= replace(path,serverRoot,"")
	getVirtualPath = replace(physicalPath,"\","/")
	'if not str.endswith(getVirtualPath, "/") then getVirtualPath = getVirtualPath & "/"
end function
%>