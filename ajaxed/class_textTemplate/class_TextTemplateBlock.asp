<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003 - This file is part of GAB_LIBRARY		
'* For license refer to the license.txt in the root    						
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		TextTemplateBlock
'' @CREATOR:		Michal Gabrukiewicz
'' @CREATEDON:		2006-10-28 14:36
'' @CDESCRIPTION:	represents a block which is used within the TextTemplate
''					Blocks are defined with <<< BLOCK NAME >>> ... <<< BLOCKEND NAME >>>. var can be defined between the
''					"begin" and the "end" of the block. Each block can hold items which will
''					be templated with the block
'' @REQUIRES:		-
'' @VERSION:		0.1

'**************************************************************************************************************
class TextTemplateBlock

	'public members
	public items		''[dictionary] items of the block. key = autoID, value = array with vars.
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	public sub class_initialize()
		set items = server.createObject("scripting.dictionary")
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	Adds an item to the block
	'' @DESCRIPTION:	the number of items will result in the same number of copied blocks.
	'' @PARAM:			vars [array]: a paired array. so 1st value is the name of the 1st var and 2nd value is
	''					is the value of the 1st var, etc. therfore number of values must be even!
	''					Example: (var1, value1, var2, value2, ...)
	'**********************************************************************************************************
	public sub addItem(vars)
		if (uBound(vars) + 1) mod 2 <> 0 then
			lib.error("Vars must be even when using addItem(). Example: (var1, value1, var2, value2, ...)")
		end if
		items.add lib.getUniqueID(), vars
	end sub

end class
%>