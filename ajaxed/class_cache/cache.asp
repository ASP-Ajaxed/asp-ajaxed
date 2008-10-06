<%
'**************************************************************************************************************
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		Cache
'' @CREATOR:		Michal Gabrukiewicz
'' @CREATEDON:		2006-11-10 15:46
'' @CDESCRIPTION:	Lets you cache values, objects, etc. It uses the application-variables and
''					therefore the cache will be shared with all other users. it can cash items which are
''					identified by an Identifier. The cache will be identified by its name
''					The cache is protected against memory problems. e.g. huge contents wont be cached,
''					cache has a limited size. so that the server is not to busy you can define how the caching
''					works for a specific thing. e.g. you want to cache RSS Feeds then you can setup that the 
''					cache for RSS will hold a maximum of 10 RSS feeds shared. The organisation of the cache is done
''					automatically if the maximum amount of items is reached
'' @REQUIRES:		-
'' @VERSION:		0.1

'**************************************************************************************************************
class Cache

	'private members
	private prefix
	
	private property get cachebase 'returns a dictionary which represents the cache. key = name of the cache. value = array(expires, HTML)
		set cachebase = nothing
		if not isArray(application(prefix & name)) then
			application.lock()
			application(prefix & name) = array(server.createObject("scripting.dictionary"))
			application.unlock()
		end if
		
		set cachebase = application(prefix & name)(0)
	end property
	
	'public members
	public name					''[string] name of the cache you want to create. normally a word which describes what it is caching
	public interval				''[string] what interval is used to define the expiration of the items. default = h (hour)
								''allowed values are all values which can be used with dateadd() function. e.g. m = month, etc.
	public intervalValue		''[string] which value of the interval. default = 1
	public maxSlots				''[string] the maximum amount of slots for caching items. default = 10
	public maxItemSize			''[string] the maximum size in bytes of an item which will be cached. all items with a size
								''bigger than this value wont be cached (saving memory on the server). default = 100000
								''works only if the items are no objects and no arrays
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	public sub class_initialize()
		name = ""
		interval = "h"
		intervalValue = 1
		maxSlots = 10
		maxItemSize = 100000
		prefix = "GL_Cache_"
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	stores an item into the cache. if it already exists then it gets overwritten
	'' @PARAM:			identifier [string]: the identifier for the item. with this you can get the item afterwards
	'' @PARAM:			item [variant]: something you want to store within the cache.
	''					Note: Objects does not work!!
	'**********************************************************************************************************
	public sub store(identifier, item)
		if name = "" then lib.error("ICE 4598: Name is needed for caching.")
		
		'check the size only for strings...
		if not isObject(item) and not isArray(item) then
			if len(item) > maxItemSize then exit sub
		end if
		
		set c = cachebase
		expires = dateadd(interval, intervalValue, now())
		packedItem = array(expires, item)
		'if it exists then remove and add to the end
		application.lock()
		if c.exists(identifier) then c.remove(identifier)
		c.add identifier, packedItem
		application.unlock()
		reOrganize()
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	clears the cache.
	'**********************************************************************************************************
	public sub clear()
		set c = cachebase
		application.lock()
		c.removeAll()
		application.unlock()
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION: 	gets an item from the cache. 
	'' @PARAM:			identifier [string]: identifier of the item within the cache
	'' @RETURN:			[variant] empty if no value found, otherwise the value.
	'**********************************************************************************************************
	public function getItem(identifier)
		set c = cachebase
		if (c.exists(identifier)) then
			cachedItem = c(identifier)
			if not itemExpired(cachedItem) then
				getItem = cachedItem(1)
			else
				'if expired then remove immediately
				application.lock()
				c.remove(identifier)
				application.unlock()
			end if
		end if
	end function
	
	'**********************************************************************************************************
	'* reOrganize 
	'**********************************************************************************************************
	private sub reOrganize()
		set c = cachebase
		'if the cache is full, then we remove all expired.
		'when no expired found we have to remove the first one
		if c.count > maxSlots then
			removedAtLeastOne = false
			for each identifier in c.keys
				rf = c(identifier)
				if itemExpired(rf) then
					application.lock()
					c.remove(identifier)
					application.unlock()
					removedAtLeastOne = true
				end if
			next
			'if no one could be removed because no one is expired yet,
			'then we have to remove the first one
			if not removedAtLeastOne then
				identifiers = c.keys
				application.lock()
				c.remove(identifiers(0))
				application.unlock()
			end if
		end if
	end sub
	
	'**********************************************************************************************************
	'* itemExpired 
	'**********************************************************************************************************
	private function itemExpired(cachedItemArray)
		itemExpired = (cachedItemArray(0) < now())
	end function

end class
%>