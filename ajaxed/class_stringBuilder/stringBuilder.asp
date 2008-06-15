<%
'**************************************************************************************************************

'' @CLASSTITLE:		StringBuilder
'' @CREATOR:		m
'' @CREATEDON:		2008-05-07 15:14
'' @CDESCRIPTION:	Represents a string builder which handles string concatenation. If a supported stringbuilder
''					COM component can be found it is used and hence the concatenation is much faster than common string
''					concatening.
''					- check supported components with the supportedComponents property. 
''					- you should use it whereever there is a loads of output to be rendered on your page. Its faster than normal concatening.
''					- basically just intantiate it and use append() method for appending. In the end use toString() to output your string
''					- if there is no component found then it would be faster to directly write the output to the response. This can be achieved using the write() method
''					Best way to use the StringBuilder (always uses the fastest possible method):
''					<code>
''					<%
''					set output = new StringBuilder
''					output("some text")
''					output("some other text")
''					% >
''					<%= output.toString() % >
''					</code>
'' @REQUIRES:		-
'' @VERSION:		0.1

'**************************************************************************************************************
class StringBuilder

	'private members
	private p_component
	
	'public members
	public component	''[string] holds the component which should be used. empty = none (common concatination)
	
	public property get supportedComponents ''[array] gets the supported string builder COM components. the order represents the order which will be loaded first if available
		supportedComponents = array("system.io.stringwriter", "stringbuildervb.stringbuilder")
	end property
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	public sub class_initialize()
		component = lib.detectComponent(supportedComponents)
		if not isEmpty(component) then
			set p_component = server.createObject(component)
			if component = "stringbuildervb.stringbuilder" then p_component.init 40000, 7500
		end if
	end sub
	
	'**********************************************************************************************************
	'* desctructor 
	'**********************************************************************************************************
	public sub class_terminate()
		set p_component = nothing
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	writes a string with the string builder. the difference to append is that it will 
	''					output to the response if there is no component found
	'' @DESCRIPTION:	- its recommended to use this instead of append() when rendering html markup. it wil allways use this fastest method. if stringbuilder available then using stringbuilder otherwise direct output to response.
	''					- note: be sure that on the place you use write() it can output directly to the response if needed
	'' @PARAM:			val [string]: the string you want to write
	'**********************************************************************************************************
	public default sub write(val)
		if isEmpty(component) then
			str.write(val)
		else
			append(val)
		end if
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	appends a string to the builder
	'' @PARAM:			val [string]: the string you want to append
	'**********************************************************************************************************
	public sub append(val)
		if isEmpty(component) then
			p_component = p_component & val
		elseif component = "system.io.stringwriter" then
			p_component.write_12(val)
		elseif component = "stringbuildervb.stringbuilder" then
			p_component.append(val)
		else
			lib.throwError("unknown stringbuilder component")
		end if
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	returns the concatenated string
	'' @RETURN:			[string] concatenated string
	'**********************************************************************************************************
	public function toString()
		if isEmpty(component) then
			toString = p_component
		elseif component = "system.io.stringwriter" then
			toString = p_component.getStringBuilder().toString()
		elseif component = "stringbuildervb.stringbuilder" then
			toString = p_component.toString()
		else
			lib.throwError("unknown stringbuilder component")
		end if
	end function

end class
%>