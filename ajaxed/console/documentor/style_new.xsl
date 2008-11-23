<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<xsl:output method="html" encoding="utf-8" />
	<xsl:template match="/">
				<style>
					h2 {
						background:url(/img/expand.gif) no-repeat 6px;
						margin-top:1em !important;
						padding-left:1.5em;
					}
					.menuItem {
						margin-bottom:0.5em;
					}
					.cDescription, .cDescription * {
						margin-bottom:10px;
					}
					.table th {
						background:#000;
						color:#fff;
					}
					.label {
						font-weight:bold;
					}
					.table td {
						padding:0.2em 0.4em;
					}
					.obsolete, .obsolete *, .obsolete a:link, .obsolete a:visited, .obsolete a:active, .obsolete a {
						white-space: nowrap;
						font-weight: normal;
						color:#f00;
						text-decoration:line-through;
						padding-right:0.5em;
					}
					.default {
						font-style:italic;
					}
					.param {
						margin-bottom:1em;
						padding-bottom:0.2em;
					}
					.paramName {
						font-weight:bold;
					}
					.returnName {
						text-decoration:underline;
						font-weight:bold;
					}
					.return {
						margin-top:0.7em;
					}
					.list {
						margin-top:1em;
					}
					.alias {
						color:#aaa;
						margin-top:6px;
					}
					.code {
						margin:1em 0px;
					}
					.member, .memberAlt {
						padding:1em 0em;
					}
					.member {
					}
					.memberAlt {
						border-bottom:1px solid #aaa;
						background:#F5F5F5;
					}
					.memberName {
						font-weight:bold;
					}
					.memberInfo {
						margin-left:1.9em;
						margin-top:0.1em;
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
					.sidebar {
						display:none;
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

			<div class="sidebar">
				
				<h1>Classes</h1>
				
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
								<xsl:text>margin-left:1em;</xsl:text>
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
					<small>
						generated with
						<a href="http://www.webdevbros.net/ajaxed/">ajaxed ASP Documentor</a>
					</small>
				</div>
				
			</div>
			
				
				<a name="top"></a>	
				<xsl:for-each select="/classes/class">
				<xsl:sort select="name"/>
				<xsl:variable name="currentClass" select="@id" />
				
					<a> 
						<xsl:attribute name="name">
							<xsl:text>class_</xsl:text><xsl:value-of select="@id"/>
						</xsl:attribute>
					</a>
					<div class="contentAlt">
						<div class="content">
							<h1>
								Class <xsl:value-of select="name"/>
								<xsl:text> </xsl:text><a class="notForPrint" href="#top" style="text-decoration:none;">^</a>
							</h1>
						</div>
					</div>
					
			<div class="content text" style="background:#eee;">
				<xsl:for-each select="properties/property">
					<xsl:sort select="name"/>
					<xsl:if test="position() = 1">
						<img src="/img/property.gif" class="icon"/>
					</xsl:if>
					<a href="#property_{generate-id(name)}" onclick="$('props_{$currentClass}').show();">
						<xsl:value-of select="name"/>
					</a>
					<xsl:if test="position() &lt; last()">
					    <xsl:text>, </xsl:text>
					</xsl:if>
				</xsl:for-each>
				<xsl:for-each select="methods/method">
					<xsl:sort select="name"/>
					<xsl:if test="position() = 1">
						<br/><br/>
						<img src="/img/method.gif" class="icon"/>
					</xsl:if>
					<a href="#method_{generate-id(name)}" onclick="$('meths_{$currentClass}').show();">
						<xsl:value-of select="name"/>
					</a>
					<xsl:if test="position() &lt; last()">
					    <xsl:text>, </xsl:text>
					</xsl:if>
				</xsl:for-each>
			</div>
			<div class="content text">
					<div class="cDescription"><xsl:value-of select="description" disable-output-escaping="yes"/></div>
					
					<h2 onclick="togg('meta_{$currentClass}')" style="cursor:pointer">Meta</h2>
					
					<div id="meta_{$currentClass}" style="display:none;" class="alwaysPrint">
						<table cellspacing="0" id="info_{$currentClass}" cellpadding="3" border="0">
						<tr>
							<td class="label">Author:</td>
							<td><xsl:value-of select="author"/> on <xsl:value-of select="created"/></td>
						</tr>
						<xsl:if test="string-length(staticname/text()) &gt; 0">
						<tr>
							<td class="label">Singleton name:</td>
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
						</table>
					</div>
					
					<xsl:if test="count(properties/property) &gt; 0">
					<h2 onclick="togg('props_{$currentClass}')" style="cursor:pointer">Properties</h2>
					<div id="props_{$currentClass}" style="display:none;" class="alwaysPrint">
					<xsl:for-each select="properties/property">
						<xsl:sort select="name"/>
						<a name="property_{generate-id(name)}"></a>
						<div class="member">
							<xsl:if test="position() mod 2 != 0">
								<xsl:attribute name="class">memberAlt</xsl:attribute>
							</xsl:if>
							<div class="memberName">
							<xsl:choose>
								<xsl:when test="@obsolete = 1">
									<xsl:attribute name="class">memberName obsolete</xsl:attribute>	
								</xsl:when>
								<xsl:otherwise>
								</xsl:otherwise>
							</xsl:choose>
							<img src="/img/property.gif" class="icon" alt="Property" />
							<xsl:if test="@defaultProperty = 1">
								<img src="/img/default.png" class="icon" alt="default property"/>
							</xsl:if>
							<xsl:value-of select="name"/>
							<xsl:text disable-output-escaping="yes"><![CDATA[&nbsp;]]></xsl:text>
							<code>
								<xsl:for-each select="types/type">
									<xsl:call-template name="TheTypes">
										 <xsl:with-param name="aType" select="."/>
										 <xsl:with-param name="aID" select="@id"/>
									</xsl:call-template>
								</xsl:for-each>
							</code>
							</div>
							
							<div class="memberInfo"><xsl:value-of select="description" disable-output-escaping="yes" /><xsl:text disable-output-escaping="yes"><![CDATA[&nbsp;]]></xsl:text></div>
						</div>	
					</xsl:for-each>
					</div>
					</xsl:if>
					
					<xsl:if test="count(methods/method) &gt; 0">
					<h2 onclick="togg('meths_{$currentClass}')" style="cursor:pointer">Methods</h2>
					<div id="meths_{$currentClass}" class="methods alwaysPrint" style="display:none">
					<xsl:for-each select="methods/method">
						<xsl:sort select="name"/>
						<div class="member">
						<xsl:if test="position() mod 2 != 0">
							<xsl:attribute name="class">memberAlt</xsl:attribute>
						</xsl:if>
						<div>
							<xsl:choose>
								<xsl:when test="@obsolete = 1">
									<xsl:attribute name="class">
									<xsl:text>memberName obsolete</xsl:text>
								</xsl:attribute>
								</xsl:when>
								<xsl:otherwise>
									<xsl:attribute name="class">
									<xsl:text>memberName</xsl:text>
								</xsl:attribute>
								</xsl:otherwise>
							</xsl:choose>
							<img src="/img/method.gif" class="icon" alt="Method" />
							<xsl:if test="@static = 1">
								<img class="icon" src="/img/static.gif" title="Static method" alt="This method is Static"/>
							</xsl:if>
							<xsl:if test="@default = 1">
								<img src="/img/default.png" class="icon" alt="default method"/>
							</xsl:if>
							<a name="method_{generate-id(name)}" href="javascript:void(0)">
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
								<xsl:value-of select="name"/>
								(<xsl:for-each select="parameters/parameter[@option='0']">
									<xsl:value-of select="name" />
									<xsl:if test="position() &lt; last()">
									    <xsl:text>, </xsl:text>
									</xsl:if>
								</xsl:for-each>)
							</a>
							<xsl:if test="string-length(alias/text()) &gt; 0">
								<xsl:text disable-output-escaping="yes"><![CDATA[&nbsp;&nbsp;]]></xsl:text>
								<span class="alias">alias: <code><xsl:value-of select="alias" disable-output-escaping="yes" /></code></span>
							</xsl:if>
						</div>
						<div class="memberInfo">
							<xsl:value-of select="shortdescription" disable-output-escaping="yes" />
							<div style="display:none" class="alwaysPrint" id="mdetails_{$currentClass}{generate-id(name)}">
								<xsl:if test="string-length(longdescription/text()) &gt; 0">
									<xsl:value-of select="longdescription" disable-output-escaping="yes" />
								</xsl:if>
								<br /><br/>
								
								<xsl:if test="string-length(parameters/parameter/name/text()) &gt; 0">
									<xsl:for-each select="parameters/parameter">
										<div class="param">
											<span class="paramName">
												<xsl:value-of select="name/@passed" />
												<xsl:if test="string-length(name/@passed) &gt; 0">
													<xsl:text disable-output-escaping="yes"><![CDATA[&nbsp;]]></xsl:text>
												</xsl:if>
												<xsl:value-of select="name" />
												<xsl:text disable-output-escaping="yes"><![CDATA[&nbsp;]]></xsl:text>
											</span>
											<xsl:if test="string-length(types/type/text()) &gt; 0">
												<code>
													<xsl:for-each select="types/type">
														<xsl:call-template name="TheTypes">
															 <xsl:with-param name="aType" select="."/>
															 <xsl:with-param name="aID" select="@id"/>
														</xsl:call-template>
													</xsl:for-each></code>
											</xsl:if>
											<br/>
											<xsl:value-of select="description" disable-output-escaping="yes" />
										</div>
									</xsl:for-each>
								</xsl:if>
								
								<xsl:if test="string-length(return/type/text()) &gt; 0">
									<div><span class="returnName">Returns</span>
										<xsl:text disable-output-escaping="yes"><![CDATA[&nbsp;]]></xsl:text>
										<code><xsl:for-each select="return/type">
											<xsl:call-template name="TheTypes">
											 <xsl:with-param name="aType" select="."/>
											 <xsl:with-param name="aID" select="@id"/>
											</xsl:call-template>
										</xsl:for-each></code>
										<br/>
										<xsl:value-of select="return/description" disable-output-escaping="yes"  />
									</div>
								</xsl:if> 
								
							</div>
						</div>
					</div>
						
				</xsl:for-each>
				</div>
				</xsl:if>
		
				<xsl:if test="string-length(initializations/init/property/text()) &gt; 0">
					<h2 onclick="togg('inits_{$currentClass}')" style="cursor:pointer">Initializations</h2>
			
					<table id="inits_{$currentClass}" class="table alwaysPrint" style="display:none; border: 1px solid #888888" cellpadding="3" cellspacing="0" border="0">
						<tr>
							<th>Config variable</th>
							<th>Property</th>
							<th>Default</th>
						</tr>
						<xsl:for-each select="initializations/init">
						<xsl:sort select="property"/>
							<tr>
								<xsl:if test="position() mod 2 != 0">
									<xsl:attribute  name="class">alt</xsl:attribute>
								</xsl:if>
								<td nowrap="nowrap">
									<img src="/img/init.png" class="icon" alt="Initialization variable"/>
									<code>
										<xsl:value-of select="constant" />
									</code>
								</td>
								<td nowrap="nowrap">
									<div class="memberName">
										<img class="icon" src="/img/property.gif" alt="Property"/>
										<xsl:value-of select="property" />
									</div>
								</td>
								<td>
									<code><xsl:value-of select="default" /></code>
								</td>
							</tr>
						</xsl:for-each>
					</table>
				</xsl:if>
				
				<br/>
				</div>
				</xsl:for-each>
				
			
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
