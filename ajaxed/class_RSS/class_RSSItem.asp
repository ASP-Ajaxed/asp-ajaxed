<%
'**************************************************************************************************************

'' @CLASSTITLE:		RSSItem
'' @CREATOR:		Michal Gabrukiewicz
'' @CREATEDON:		2006-11-08 18:15
'' @CDESCRIPTION:	used with the RSSReader
'' @REQUIRES:		-
'' @FRIENDOF:		RSS
'' @VERSION:		0.1

'**************************************************************************************************************
class RSSItem

	'public members
	public title			''[string] title of the item
	public description		''[string] content
	public category			''[string] name of the category
	public link				''[string] link to the item
	public comments			''[string] link to the comments of the item (if available)
	public GUID				''[string] an unique identifier for the item. usually the same as link
	public author			''[string] author of the item
	public publishedDate	''[date] date when the item has been published. your local timezone.
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	public sub class_initialize()
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	reflection of the properties and its values
	'' @RETURN:			[dictionary] <em>key</em> = property-name, <em>value</em> = property value
	'**********************************************************************************************************
	public function reflect()
		set reflect = server.createObject("scripting.dictionary")
		with reflect
			.add "title", title
			.add "description", description
			.add "category", category
			.add "link", link
			.add "comments", comments
			.add "GUID", guid
			.add "author", author
			.add "publishedDate", publishedDate
		end with
	end function

end class
%>