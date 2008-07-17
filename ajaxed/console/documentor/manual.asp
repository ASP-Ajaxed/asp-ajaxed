<!--#include file="../../ajaxed.asp"-->
<%
'******************************************************************************************
'* Creator: 	michal
'* Created on: 	2008-04-12 20:15
'* Description: help for the documentor
'* Input:		-
'******************************************************************************************

set page = new AjaxedPage
with page
	.title = "ASP Documentor Reference"
	.onlyDev = true
	.defaultStructure = true
	.draw()
end with
set page = nothing

'******************************************************************************************
'* header  
'******************************************************************************************
sub header()
	page.loadCSSFile "../std.css", empty
end sub

'******************************************************************************************
'* main 
'******************************************************************************************
sub main()
	content()
end sub

'******************************************************************************************
'* content 
'******************************************************************************************
sub content() %>

	<style>
		.keywords {
			margin:10px;
			margin-top:0px;
		}
		.keywords td {
			padding:0.2em;
			padding-right:2em;
		}
	</style>

	<h1>ASP Documentor Reference</h1>

	<p>
		The following things are important in order to create a documentation for your ASP code:
		<ul>
			<li>Works only for ASP written with VBScript</li>
			<li>
				Documents only classes and its public members
				<ul>
					<li>Only one class per .asp file</li>
				</ul>
			</li>
			<li>comments recognized by documentor must start with <code>''</code></li>
			<li>documentor keywords always start with <code>@</code> and end with <code>:</code></li>
			<li>types are surrounded by <code>[]</code> e.g. <code>[bool]</code></li>
			<li>Documentor can be found in the ajaxed console. <code>http://yourhost.com/ajaxed/console/</code> (<a href="http://www.webdevbros.net/ajaxed/">ajaxed</a> must be installed)</li>
		</ul>
	</p>
	
	<strong>Examples of proper comments which will be recognized by documentor:</strong>
	<div class="code">
		<div class="comment">
			'' @DESCRIPTION: some description<br>
			'' @RETURN: [string] returns my name
		</div>
		public function getMe()<br>
		end function<br>
		<br>
		public property get color <span class="comment">''[bool] gets the color</span><br>
		end property
	</div>
	
	<h2>Keywords for class documentation</h2>
	
	<table class="keywords">
	<tr>
		<td><code>@CLASSTITLE</code></td>
		<td>Title of the class</td>
	</tr>
	<tr>
		<td><code>@CREATOR</code></td>
		<td>Name of the class creator. Also email can be added.</td>
	</tr>
	<tr>
		<td><code>@CREATEDON</code></td>
		<td>Date and time of class creation (format up to you)</td>
	</tr>
	<tr>
		<td><code>@CDESCRIPTION</code></td>
		<td>Full class description</td>
	</tr>
	<tr valign="top">
		<td><code>@STATICNAME</code></td>
		<td>
			Name of the variable which holds already an instance of the class.
			This simulates the OO concept of static classes like e.g. <code>StringOperations</code> which is available
			with the staticname <code>str</code>.
		</td>
	</tr>
	<tr>
		<td><code>@POSTFIX</code></td>
		<td>Postfix for SmartSense (intellissense support) plugin in Macromedia Homesite</td>
	</tr>
	<tr>
		<td><code>@VERSION</code></td>
		<td>Class version. e.g. 1.0</td>
	</tr>
	<tr valign="top">
		<td><code>@COMPATIBLE</code></td>
		<td>What browsers are supported by this class. Useful if the class is a control which renders HTML</td>
	</tr>
	<tr>
		<td><code>@REQUIRES</code></td>
		<td>List of other classes which are required. Seperated with ","</td>
	</tr>
	<tr valign="top">
		<td><code>@FRIENDOF</code></td>
		<td>
			Name of the class this class is friend of.
			Useful if your component consists of more classes and they are always used as a bundle.
			e.g. <code>Dropdown</code> and <code>DropdownItem</code>. Both are always used togehter and
			<code>DropdownItem</code> is loaded automatically with <code>Droddown</code>. Thus <code>DropdownItem</code>
			is a friend of <code>Dropdown</code>. 
		</td>
	</tr>
	</table>

	<strong>Example of a proper class documentation:</strong>
	<div class="code">
		<div>
			<pre class="comment">
'' @<span>C</span>LASSTITLE:		Person
'' @CREATOR:		Jack Johnson
'' @CREATEDON:		2000-01-21 20:30
'' @CDESCRIPTION:	Represents a user which is using the application. A user
''                      can be logged in and has different permissions.
'' @STATICNAME:		usr
'' @VERSION:		0.1
			</pre>
		</div>
class User<br>
end class
	</div>


	<h2>Keywords for method documentation</h2>

	<table class="keywords">
	<tr>
		<td><code>@SDESCRIPTION</code></td>
		<td>A one sentence description.</td>
	</tr>
	<tr>
		<td><code>@DESCRIPTION</code></td>
		<td>The full description of the method. Includes all the details which short description does not.</td>
	</tr>
	<tr valign="top">
		<td><code>@PARAM</code></td>
		<td>
			Documentation of a method's parameter.
			The parameter should contain one (or more) type definition(s).
			<code>byRef</code> and <code>byVal</code> are recognized as well.
		</td>
	</tr>
	<tr>
		<td><code>@RETURN</code></td>
		<td>Description about the return value (if <code>function</code>) including its type(s)</td>
	</tr>
	</table>
	
	
	<strong>Example of a proper method documentation:</strong>
	<div class="code">
		<div>
			<pre class="comment">
'' @SDESCRIPTION:	Gets all users created on a given date and a specified role
'' @DESCRIPTION:	Useful if you want to check who signed up when with what role
'' @PARAM:		creationDate [date]: date when the users been created
'' @PARAM:		role [int], [string]: required role. 1 = admin, 0 = common user.
''			if string then "admin" or "common"
'' @RETURN:		[dictionary] list with users. empty dictionary if no users found
			</pre>
		</div>
public function getAllBy(creationDate, role)<br>
end function
	</div>
	
	<h2>Property/public member variables documentation</h2>
	
	<p>
		Properties (<code>public property</code>) and public member variables are documented
		immediately after their definition with a description and it's type.
	</p>
	<strong>Example of a proper property documentation:</strong>
	<div class="code">
		public property get name <span class="comment">''[string] gets the users name</span><br>
		end property<br>
		<br>
		public property set name(val) <span class="comment">''[string] sets the users name</span><br>
		end property
		<br><br>
		<div class="comment">'specifying more types</div>
		public property set role <span class="comment">''[int], [string] sets the role of a user</span><br>
		end property
	</div>
	
	<br>
	<strong>Example of a proper public member documentation:</strong>
	<div class="code">
		public ID <span class="comment">''[int] ID of the user. default = 0</span><br>
		public name <span class="comment">''[string] gets/sets the name of the user</span>
	</div>
	
	<h2>Other documentor keywords (Markers)</h2>

	<table class="keywords">
	<tr valign="top">
		<td><code>STATIC!</code></td>
		<td>Can be used within the description of methods to mark them as static methods (callable without an instance).</td>
	</tr>
	<tr valign="top">
		<td><code>OBSOLETE!</code></td>
		<td>Can bee used within the description of methods, properties or public member variables to indicate that a member is obsolete.</td>
	</tr>
	</table>
	<br>
	<strong>Examples:</strong>
	<div class="code">
		public ID <span class="comment">''[int] OBSOLETE! ID of the user. default = 0</span><br>
		<br>
		<div class="comment">'' @SDESCRIPTION: OBSOLETE! gets an user by its id.</div>
		public function getBy(id)<br>
		end function
		<br><br>
		<div class="comment">'' @SDESCRIPTION: STATIC! gets an user by its id.</div>
		public function getByID(id)<br>
		end function
	</div>

	<h2>Initializations</h2>
	
	Initializations document the default values of class properties. Those values
	are assigned within the constructor using the <code>lib.init</code> method. This method
	tries to get the value from a given variable and if not possible it takes a fallback
	value (default value). Those default values are documented as well because
	its a core concept of the ajaxed library which is used for the config vars of 
	<code>/ajaxedConfig/config.asp</code>. With this it's possible to see the name of the
	config var and its default value.
	<br><br>
	Within custom applications it's possible to create own config variables which are
	initialized with the same concept.
	<br><br>
	<strong>Example:</strong>
	<div class="code">
		class Document
			<div class="indent">
				public location <span class="comment">''[string] gets/sets the documents location</span><br><br>
				public sub class_initialize()
					<div class="indent">location = lib.init(DOCUMENT_LOCATION, "/documents/")</div>
				end sub
			</div>
		end class
	</div>
	
	<h2><a name="parsing"></a>Parsing & other important things</h2>
	
	You should know the following things about the parsing of documentation:
	<ul>
		<li>
			Lines starting with a "<code>-</code>" are treated as list items within your documentation.
			<div class="code">
				<div class="comment">
					''@DESCRIPTION: performs some action. The following needs to be considered:<br>
					''- some list first list item<br>
					''- and another list item. The text must stay in the same line
				</div>
				public sub perform()<br>
				end sub
			</div>
		</li>
		<li>Do not use the <code>public</code> keyword if you don't want to show a method/property in the documentation (no access modifier will result in public member anyway). Can be useful for e.g. protected members, etc</li>
		<li>
			all HTML markup is recognized within the documentation. Exceptions:
			<ul>
				<li>All HTML inside <code>&lt;code&gt;&lt;/code&gt;</code> is escaped and all line breaks are converted to <code>&lt;br/&gt;</code>. This makes the writing of code examples easier and more readable within the code as well (because no HTML markup is needed within code blocks).</li>
			</ul>
		</li>
		<li>
			Use <code>&lt;em&gt;&lt;/em&gt;</code> to emphasize coding keywords e.g. methodname, properties, ...<br>
			The following keywords are emphasized automatically (case sensitive):
			<ul>
				<li>EMPTY, NOTHING, BOOL, INT, STRING, OBJECT, NULL, TRUE, FALSE, RECORDSET, DICTIONARY, BOOLEAN, FLOAT, DOUBLE, ARRAY</li>
			</ul>
		</li>
	</ul>
	
<br>

	<div class="ct">
		<small>Michal Gabrukiewicz, David Rankin. updated July, 2008</small>
	</div>
</body>
</html>
	

<% end sub %>