<?xml version="1.0" encoding="ISO-8859-1" ?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<xsl:output method="html" encoding="ISO-8859-1" />
	<xsl:template match="/">
		<html>
			<head>
				<title>ajaxed API</title>
				<link rel="shortcut icon" href="/ajaxed/console/documentor/icon.png" type="image/ico" />
				<meta name="Generator" content="ASP Documentor" />
				<meta name="description" content="ASP Documentor - documentation for classic ASP (VBScript)" />
				<script src="/ajaxed/prototypejs/prototype.js"></script>
				<style>
					body {
						padding:0px; margin:0px;
					}
					* {
						font-size:14px;
						font-family:tahoma;
					}
					em {
						font-family:courier new;
						font-style:normal;
					}
					.comment {
						color:#008000 !important;
					}
					code, code * {
						font-family:courier new;
						color:#0B7391;
						font-size:9pt !important;
						margin-bottom:0px !important;
					}
					code {
						display:block;
						overflow:auto;
						white-space:nowrap;
						margin-top:3px;
						padding:5px 3px 3px 20px;
					}
					code .ssi-code {
						background:#EFEFEF;
					}
					#menu {
						position:absolute;
						right:10px;
						width:200px;
					}
					body>#menu {
						position:fixed;
					}
					.menuItem {
						margin-bottom:5px;
						font-weight:bold;
					}
					.label {
						font-weight:bold;
						padding-right:2em;
					}
					.cDescription, .cDescription * {
						font-size:16px;
						margin-bottom:10px;
					}
					h1 {
						background-color:#8DB2FF;
						padding:5px;
						margin:15px 0px;
					}
					h2 {
						background-color:#DDD;
						font-size:14px;
						font-weight:bold;
						padding:4px;
						margin-bottom:0;
						border: 1px solid #bbbbbb;
					}
					.method {
						margin:5px 0px;
					}
					.table {
						width:100%;
						border-collapse:collapse;
					}
					.table td, .table th {
						border:1px solid #eee;
						text-align:left;
					}
					.table td {
						padding:8px;
					}
					legend {
						font-size:20px;
						font-weight:bold;
					}
					fieldset {
						padding:10px;
						margin:10px 0;
					}

					.type, .type * {
						font-family:courier;
					}
					.cDescription {
					}
					.mdetails {
						padding:1em;
						margin-top:0.5em;
						border:1px solid #aaa;
						background-color:#FBFBFB;
					}
					.button {
						width:70px;
						padding:2px;
						font-family:tahoma;
					}
					
					a:hover, a:link, a:visited, a:active {
						color:#0000FF;
					}
					.obsolete, .obsolete *, .obsolete a:link, .obsolete a:visited, .obsolete a:active, .obsolete a {
						white-space: nowrap;
						font-weight: normal;
						color:#E10000;
						padding-right:0.5em;
					}
					.default {
						font-style:italic;
					}
					.methods a {
					}
					.param {
						margin-bottom:1em;
						padding-bottom:0.2em;
					}
					.paramName {
						color:#AE008F;
						font-weight:bold;
					}
					.returnName {
						color:#000;
						text-decoration:underline;
						font-weight:bold;
					}
					.return {
						margin-top:0.7em;
					}
					
					/* this is just for Firefox */
					html>body #content {
					}
					
					.class {
						width:680px;
						text-align:left;
					}
					
					.list {
						list-style-position: outside;
						margin:1em 0em 1em 2em;
						padding:0px;
						list-style-type:disc;
					}
				</style>
				
				<style media="print">
					body {
						overflow:auto;
					}
					.notForPrint {
						display:none;
					}
					.alwaysPrint {
						display:block;
					}
					.class {
						width:auto;
						page-break-before:always;
					}
					fieldset {
						border:0px;
					}
					#content {
						overflow:auto;
						height:auto;
					}
					#menu {
						display:none;
					}
					code {
						white-space:auto;
					}
				</style>
				
				<script>
					function togg(eID) {
						$(eID).toggle();
					}
					function doPrint() {
						$$('.alwaysPrint').each(function(el) {
							el.show()
						});
						window.print()
					}
				</script>
			</head>
			
			<body>

			<div id="menu">
			<fieldset style="margin:20px;background-color:#fff;">
				
				
				<xsl:for-each select="/classes/class[not(name = preceding-sibling::class/name)]">
					<xsl:sort select="name"/>
					
					<div>
						
						<xsl:choose>
							<xsl:when test="@obsolete = 1">
								<xsl:attribute name="class">
									<xsl:text>menuItem obsolete</xsl:text>
								</xsl:attribute>
							</xsl:when>
							<xsl:otherwise>
								<xsl:attribute name="class">
									<xsl:text>menuItem</xsl:text>
								</xsl:attribute>
							</xsl:otherwise>
						</xsl:choose>
						
						<xsl:if test="string-length(friendof/types/type) &gt; 0">
							<xsl:attribute name="style">
								<xsl:text>font-weight:lighter;</xsl:text>
							</xsl:attribute>
						</xsl:if>
						
						<a>
							<xsl:attribute name="href">
								<xsl:text>#class_</xsl:text><xsl:value-of select="@id"/>
							</xsl:attribute>
							<xsl:value-of select="name"/>
						</a>
					</div>
					
				</xsl:for-each>
				
				
				<div style="margin-top:20px;">
					<button type="button" class="button" onclick="doPrint()">Print</button>
					
					<div style="font-size:10px;margin-top:8px;">
						documented with<br/>
						<a href="http://www.webdevbros.net/ajaxed">ASP Documentor</a><br/>
						(<xsl:value-of select="//@createdOn" />)
					</div>
				</div>
				
			</fieldset>
			</div>
			
			<table cellpadding="0" cellspacing="0">
			<tr><td style="padding-left:20px;">
			
			
			<div id="content">
				
				<a name="top"></a>	
				<xsl:for-each select="/classes/class">
				<xsl:sort select="name"/>
				<xsl:variable name="currentClass" select="@id" />
				
					<br/><br/>
					<a> 
						<xsl:attribute name="name">
							<xsl:text>class_</xsl:text><xsl:value-of select="@id"/>
						</xsl:attribute>
					</a>
					<fieldset class="class">
					
					<legend>
						Class <xsl:value-of select="name"/>
						<xsl:text> </xsl:text><a class="notForPrint" href="#top" style="text-decoration:none;">^</a>
					</legend>
					
					<div class="cDescription"><xsl:value-of select="description" disable-output-escaping="yes"/></div>
					
					<table cellspacing="0" id="info_{$currentClass}" cellpadding="3" border="0">
					<tr>
						<td class="label">Version:</td>
						<td>
							<xsl:value-of select="version"/>
							<xsl:if test="@obsolete = 1">
								<span class="obsolete">This class is obsolete.</span>
							</xsl:if>
						</td>
					</tr>
					<tr>
						<td class="label">Author:</td>
						<td><xsl:value-of select="author"/> on <xsl:value-of select="created"/></td>
					</tr>
					<tr>
						<td class="label" nowrap="nowrap">Last modified:</td>
						<td><xsl:value-of select="modified"/></td>
					</tr>
					<xsl:if test="string-length(postfix/text()) &gt; 0">
					<tr>
						<td class="label">Postfix:</td>
						<td><xsl:value-of select="postfix"/></td>
					</tr>
					</xsl:if>
					<xsl:if test="string-length(staticname/text()) &gt; 0">
					<tr>
						<td class="label">Staticname:</td>
						<td><em><xsl:value-of select="staticname"/></em></td>
					</tr>					
					</xsl:if>
					<xsl:if test="string-length(requires/types/type/text()) &gt; 0">
					<tr>
						<td class="label">Requires:</td>
						<td>
						<xsl:for-each select="requires/types/type">
							<xsl:call-template name="TheTypes">
								 <xsl:with-param name="aType" select="."/>
								 <xsl:with-param name="aID" select="@id"/>
							</xsl:call-template>
						</xsl:for-each>
						</td>
					</tr>
					</xsl:if>
					<xsl:if test="string-length(compatible/browser/text()) &gt; 0">
					<tr>
						<td class="label">Compatible:</td>
						<td>
							<xsl:for-each select="compatible/browser">
								<xsl:value-of select="."/>
								<xsl:if test="position() &lt; last()">
								    <xsl:text>, </xsl:text>
								</xsl:if>
							</xsl:for-each>
						</td>
					</tr>
					</xsl:if>
					<xsl:if test="string-length(friendof/types/type) &gt; 0">
					<tr>
						<td class="label">Friend of:</td>
						<td>
						<xsl:for-each select="friendof/types/type">
							<xsl:call-template name="TheTypes">
								 <xsl:with-param name="aType" select="."/>
								 <xsl:with-param name="aID" select="@id"/>
							</xsl:call-template>
						</xsl:for-each>
						</td>
					</tr>
					</xsl:if>			
					<xsl:if test="string-length(demo/text()) &gt; 0">
					<tr>
						<td class="label">Demo:</td>
						<td><a href="{demo}" target="_blank"><xsl:value-of select="demo"/></a></td>
					</tr>
					</xsl:if>
					</table>
					
					<h2>
						<span onclick="togg('props_{$currentClass}')" style="cursor:pointer">Properties</span>
						<xsl:text> </xsl:text><a href="#class_{$currentClass}" class="notForPrint" style="text-decoration:none;">^</a>
					</h2>
					
					<table id="props_{$currentClass}" style="display:none; border: 1px solid #888888" class="table alwaysPrint" cellspacing="0" cellpadding="3" border="0">
					<tr>
						<th>Name</th>
						<th>Type</th>
						<th>Description</th>
					</tr>
					<xsl:for-each select="properties/property">
						<xsl:sort select="name"/>
						<tr valign="top">
							<xsl:if test="position() mod 2 != 0">
								<xsl:attribute  name="style">background-color:#EEEEEE</xsl:attribute>
							</xsl:if>
							<td class="property">
							<xsl:choose>
								<xsl:when test="@obsolete = 1">
									<div class="obsolete">This property is obsolete.</div>	
								</xsl:when>
								<xsl:when test="@defaultProperty = 1">
									<xsl:attribute name="style">
										<xsl:text>font-weight:bold;</xsl:text>
									</xsl:attribute>
									<xsl:attribute name="title">
										<xsl:text>The default property</xsl:text>
									</xsl:attribute>
								</xsl:when>
								<xsl:otherwise>
									
								</xsl:otherwise>
							</xsl:choose>
							<xsl:value-of select="name"/>
							</td>
							<td class="type">
								<xsl:for-each select="types/type">
									<xsl:call-template name="TheTypes">
										 <xsl:with-param name="aType" select="."/>
										 <xsl:with-param name="aID" select="@id"/>
									</xsl:call-template>
								</xsl:for-each>
							</td>
							
							<td><xsl:value-of select="description" disable-output-escaping="yes" /><xsl:text disable-output-escaping="yes"><![CDATA[&nbsp;]]></xsl:text></td>
						</tr>	
					</xsl:for-each>
					</table>
					
					<h2>
						<span onclick="togg('meths_{$currentClass}')" style="cursor:pointer">Methods</span>
						<xsl:text> </xsl:text><a href="#class_{$currentClass}" class="notForPrint" style="text-decoration:none;">^</a>
					</h2>
					
					<table id="meths_{$currentClass}" class="table methods alwaysPrint" style="display:none" cellpadding="3" cellspacing="0" border="0">
					<xsl:for-each select="methods/method">
						<xsl:sort select="name"/>
						<tr>
							<td>
								<div class="mName">
									<a href="#method_{generate-id(name)}" >
									<xsl:choose>
									<xsl:when test="string-length(longdescription/text()) &gt; 0">
										<xsl:attribute name="onclick">
											<xsl:text>togg('mdetails_</xsl:text>
												<xsl:value-of select="$currentClass" />
												<xsl:value-of select="generate-id(name)" />
											<xsl:text>')</xsl:text>
										</xsl:attribute>
									</xsl:when>
									<xsl:when test="string-length(return/type/text()) &gt; 0">
										<xsl:attribute name="onclick">
											<xsl:text>togg('mdetails_</xsl:text>
												<xsl:value-of select="$currentClass" />
												<xsl:value-of select="generate-id(name)" />
											<xsl:text>')</xsl:text>
										</xsl:attribute>
									</xsl:when>									
									<xsl:when test="string-length(parameters/parameter/name/text()) &gt; 0">
										<xsl:attribute name="onclick">
											<xsl:text>togg('mdetails_</xsl:text>
												<xsl:value-of select="$currentClass" />
												<xsl:value-of select="generate-id(name)" />
											<xsl:text>')</xsl:text>
										</xsl:attribute>
									</xsl:when>
									</xsl:choose>
										<xsl:if test="@static = 1">
											<xsl:attribute name="style">
												<xsl:text>font-weight:bold;</xsl:text>
											</xsl:attribute>
											<xsl:attribute name="title">
												<xsl:text>This is a static method</xsl:text>
											</xsl:attribute>											
										</xsl:if>
										<xsl:if test="@default = 1">
											<xsl:attribute name="class">
												<xsl:text>default</xsl:text>
											</xsl:attribute>
										</xsl:if>
										<xsl:value-of select="name"/>
										(<xsl:for-each select="parameters/parameter">
											<xsl:value-of select="name" />
											<xsl:if test="position() &lt; last()">
											    <xsl:text>, </xsl:text>
											</xsl:if>
										</xsl:for-each>)
									</a>
								</div>
								<div class="mDesc">
									<xsl:if test="@obsolete = 1">
										<span class="obsolete">This method is obsolete.</span>
									</xsl:if>
									<xsl:value-of select="shortdescription"  disable-output-escaping="yes" />
								</div>
								<div style="display:none" class="mdetails alwaysPrint" id="mdetails_{$currentClass}{generate-id(name)}">
									<xsl:if test="string-length(longdescription/text()) &gt; 0">
										<xsl:value-of select="longdescription" disable-output-escaping="yes" />
										<br /><br/>
									</xsl:if>
									
									<xsl:if test="string-length(parameters/parameter/name/text()) &gt; 0">
										<xsl:for-each select="parameters/parameter">
											<div class="param">
												<span class="paramName"><xsl:value-of select="name/@passed" /><xsl:text> </xsl:text><xsl:value-of select="name" /></span>
												<xsl:if test="string-length(types/type/text()) &gt; 0">
													<span>
														[<xsl:for-each select="types/type">
															<xsl:call-template name="TheTypes">
																 <xsl:with-param name="aType" select="."/>
																 <xsl:with-param name="aID" select="@id"/>
															</xsl:call-template>
														</xsl:for-each>]</span>
												</xsl:if>
												<br/>
												<xsl:value-of select="description" disable-output-escaping="yes" />
											</div>
										</xsl:for-each>
									</xsl:if>
									
									<xsl:if test="string-length(return/type/text()) &gt; 0">
										<div class="return"> <span class="returnName">Returns </span>
											[<xsl:for-each select="return/type">
												<xsl:call-template name="TheTypes">
												 <xsl:with-param name="aType" select="."/>
												 <xsl:with-param name="aID" select="@id"/>
												</xsl:call-template>
											</xsl:for-each>]
											<br/>
											<xsl:value-of select="return/description" disable-output-escaping="yes"  />
										</div>
									</xsl:if> 
									
								</div>
							</td>
						</tr>
						
				</xsl:for-each>
					
				</table>
		
				<xsl:if test="string-length(initializations/init/property/text()) &gt; 0">
					<h2>
						<span onclick="togg('inits_{$currentClass}')" style="cursor:pointer">Initializations</span>
						<xsl:text> </xsl:text><a href="#class_{$currentClass}" class="notForPrint" style="text-decoration:none;">^</a>
					</h2>
			
					<table id="inits_{$currentClass}" class="table alwaysPrint" style="display:none; border: 1px solid #888888" cellpadding="3" cellspacing="0" border="0">
						<tr>
							<th>Config var.</th>
							<th>Property</th>
							<th>Default</th>
						</tr>
						<xsl:for-each select="initializations/init">
						<xsl:sort select="property"/>
							<tr>
								<xsl:if test="position() mod 2 != 0">
									<xsl:attribute  name="style">background-color:#EEEEEE</xsl:attribute>
								</xsl:if>
								<td>
									<div class="iConstant">
										<xsl:value-of select="constant" />
									</div>
								</td>
								<td>
									<div class="iProperty">
										<xsl:value-of select="property" />
									</div>
								</td>
								<td>
									<div class="iDefault">
										<xsl:value-of select="default" />
									</div>
								</td>
							</tr>
						</xsl:for-each>
					</table>
				</xsl:if>
				
				</fieldset>
				
				</xsl:for-each>
				
			</div>
			
			</td>
			</tr>
			</table>
			
			</body>
		</html>
	</xsl:template>
	<!-- 
	/	Function TheTypes(aType, aID)
	/	This behaves like the function call above. Pass a type the class ID(optional) and we draw the type with a link
	/	If the class ID is found we link to the class, otherwise we link to the defined named links below!
	-->
	<xsl:template name="TheTypes">
		<xsl:param name="aType" />
		<xsl:param name="aID" />
		<xsl:variable name="aParam" select="translate($aType,'abcdefghijklmnopqrstuvxyz','ABCDEFGHIJKLMNOPQRSTUVXYZ')" />
			<xsl:choose>
         		<xsl:when test="string-length($aID) &gt; 0">
       				<a class="type">
						<xsl:attribute name="href">
							<xsl:text>#class_</xsl:text>
							<xsl:value-of select="$aID" />
						</xsl:attribute>
						<xsl:value-of select="$aType"/>
					</a>
				</xsl:when>
				<xsl:when test="$aParam = 'DICTIONARY'">
					<a class="type">
						<xsl:attribute name ="href">
							<xsl:text>http://www.w3schools.com/asp/asp_ref_dictionary.asp</xsl:text>
						</xsl:attribute>
						<xsl:attribute name ="target">
							<xsl:text>_blank</xsl:text>
						</xsl:attribute>												
						<xsl:value-of select="$aType"/>
					</a>
				</xsl:when>
				<xsl:when test="$aParam = 'RECORDSET'">
					<a class="type">
						<xsl:attribute name ="href">
							<xsl:text>http://www.w3schools.com/ado/ado_ref_recordset.asp</xsl:text>
						</xsl:attribute>
						<xsl:attribute name ="target">
							<xsl:text>_blank</xsl:text>
						</xsl:attribute>												
						<xsl:value-of select="$aType"/>
					</a>
				</xsl:when>										
         			<xsl:otherwise>
						<span class="type"><xsl:value-of select="$aType"/></span>
         			</xsl:otherwise>
       		</xsl:choose>
			<xsl:if test="position() &lt; last()">
			    <xsl:text>, </xsl:text>
			</xsl:if>
	</xsl:template>
</xsl:stylesheet>
