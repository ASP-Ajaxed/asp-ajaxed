<%
'**************************************************************************************************************

'' @CLASSTITLE:		TextTemplateBlock
'' @CREATOR:		Michal Gabrukiewicz
'' @CREATEDON:		2006-10-28 14:36
'' @CDESCRIPTION:	Represents a block which is used within a <em>TextTemplate</em>.
''					Blocks are defined with <em>&lt;&lt;&lt; BLOCK NAME &gt;&gt;&gt;</em> ... <em>&lt;&lt;&lt; BLOCKEND NAME &gt;&gt;&gt;</em>. Placeholders may be defined between the
''					begining and the ending of the block. Example of a block
''					<code>
''					<<< BLOCK DETAILS >>>
''						Name: <<< NAME >>>
''					<<< BLOCKEND DETAILS >>>
''					</code>
'' @REQUIRES:		-
'' @FRIENDOF:		TextTemplate
'' @VERSION:		0.1

'**************************************************************************************************************
class TextTemplateBlock

	'public members
	public items		''[dictionary] Items of the block. <em>key</em> = autoID, <em>value</em> = ARRAY with vars.
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	public sub class_initialize()
		set items = ["D"](empty)
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	Adds an item to the block.
	'' @DESCRIPTION:	The number of items will result in the same number of copied blocks.
	'' @PARAM:			vars [array]: A paired ARRAY. so 1st value is the name of the 1st var and 2nd value is
	''					is the value of the 1st var, etc. therfore number of values must be even!
	''					Example: <code><% vars = array(var1, value1, var2, value2, ...) % ></code>
	'**********************************************************************************************************
	public sub addItem(vars)
		if not isArray(vars) then lib.throwError("TextTemplateBlock.addItem() requires an array.")
		if (uBound(vars) + 1) mod 2 <> 0 then
			lib.throwError("TextTemplateBlock.addItem() vars must be even when using addItem(). Example: (var1, value1, var2, value2, ...)")
		end if
		items.add lib.getUniqueID(), vars
	end sub
	
	'***********************************************************************************************************
	'' @SDESCRIPTION: 	Adds all rows of a given recordset to the block. 
	'' @DESCRIPTION:	- The recordset is traversed from its current position.
	'' @PARAM: 			dataRS [recordset]: The name of the placeHolders within the block match the recordsets field names.
	'***********************************************************************************************************
    public sub addRS(dataRS)
    	if dataRS.eof then exit sub
    	while not dataRS.eof
    		set dc = (new DataContainer)(array())
    		for each field in dataRS.fields
    			dc.add(field.name).add(dataRS.fields(field.name))
    		next
    		addItem(dc.data)
    		dataRS.movenext()
    	wend
    end sub

end class
%>