<cfcomponent output="false">
	
	<cfset variables.objFilter = StructNew()/>
	<cfset variables.objFilter.autoFilter = false/>
	<cfset variables.objFilter.localeDates = false/>
	<cfset variables.objFilter.worksheetNames = ""/>
	<cfset variables.objFilter.groupField = ""/>
	<cfset variables.objFilter.groupFieldDisplay = ""/>
	
	
	<cffunction name="init" access="public" returntype="any" hint="I am the initializer." output="false">
		<cfreturn this/>
	</cffunction>
	
	<cffunction name="setAutoFilter" returntype="any">
		<cfargument name="autoFilter" default="0">
		<cfset variables.objFilter.autoFilter = arguments.autoFilter>
		<cfreturn true>
	</cffunction>
	
	<cffunction name="getAutoFilter" returntype="boolean">		
		<cfreturn getObjFilter().autoFilter>
	</cffunction>
	
	<cffunction name="setLocaleDates" returntype="any">
		<cfargument name="localeDates" default="0">
		<cfset variables.objFilter.localeDates = arguments.localeDates>
		<cfreturn true>
	</cffunction>
	
	<cffunction name="setWorksheetNames" returntype="any">
		<cfargument name="worksheetNames" default="0">
		<cfset variables.objFilter.worksheetNames = arguments.worksheetNames>
		<cfreturn true>
	</cffunction>
	<cffunction name="setGroupField" returntype="any">
			<cfargument name="groupField" required="true">
			<cfset variables.objFilter.groupField = arguments.groupField>
		<cfreturn true>
	</cffunction>	
	<cffunction name="setGroupFieldDisplay" returntype="any">
			<cfargument name="groupFieldDisplay" required="false" default="#variables.objFilter.groupField#">
			<cfset variables.objFilter.groupFieldDisplay = arguments.groupFieldDisplay>
		<cfreturn true>
	</cffunction>		
	
	<cffunction name="getWorksheetNames" returntype="string">		
		<cfreturn getObjFilter().worksheetNames>
	</cffunction>
	
	<cffunction name="getLocaleDates" returntype="boolean">
		<cfreturn getObjFilter().localeDates>
	</cffunction>
	
	<cffunction name="getGroupField" returntype="string">		
		<cfreturn getObjFilter().groupField>
	</cffunction>
	<cffunction name="getGroupFieldDisplay" returntype="string">		
		<cfreturn getObjFilter().groupFieldDisplay>
	</cffunction>	

	<cffunction name="getObjFilter" returntype="struct">		
		<cfreturn variables.ObjFilter>
	</cffunction>

	
	
	<cffunction name="processObj" access="public" returntype="any" hint="I take in the data and hand of to the right processor" output="false">
		<cfargument name="rootObj" 		type="any" required="true" default=""/>
		<cfargument name="objFilter" 	type="struct" required="false" default="#StructNew()#">

		<cfset var excelData = ""/>
		<cfif StructKeyExists(arguments.objFilter,"localeDates")>
			<cfset setLocaleDates(objFilter.localeDates) />
		</cfif>
		
		<cfif StructKeyExists(arguments.objFilter,"autoFilter")>
			<cfset setAutoFilter(objFilter.autoFilter)>
		</cfif>
		
		<cfif StructKeyExists(arguments.objFilter,"worksheetNames")>
			<cfset setWorksheetNames(objFilter.worksheetNames)>
		</cfif>
		
		<cfif StructKeyExists(arguments.objFilter,"groupField")>
			<cfset setGroupField(objFilter.groupField)>
			
		</cfif>	
		
		<cfif StructKeyExists(arguments.objFilter,"groupFieldDisplay")>
			<cfset setGroupFieldDisplay(objFilter.groupFieldDisplay)>
		</cfif>			
		<cfif Len(getGroupField())>
			<cfset arguments.rootObj = convertQueryToArrayOfQueries(arguments.rootObj)>
			<cfset setWorksheetNames(variables.tabNames)>
		</cfif> 
		<cfif isQuery(arguments.rootObj)>
			<cfset excelData = processQuery( arguments.rootObj, getWorksheetName( getWorksheetNames(), 1 ) )/>
		<cfelse>
			<cfset excelData = processArrayofQueries( arguments.rootObj, getWorksheetNames() )/>
		</cfif>
		
		<cfreturn getXMLHeader() & startWorkbook() & getStyles() & excelData & endWorkBook()/>
	</cffunction>
	
	<cffunction name="processQuery" access="private" returntype="any" hint="I process a query object into excel" output="false">
		<cfargument name="queryObj" type="any" required="true" default=""/>
		<cfargument name="nodeName" type="any" required="true" default=""/>
		
		<cfset var columnListed = arrayToList( arguments.queryObj.getColumnNames() )/>
		<cfset var columnCount = ListLen( arguments.queryObj.columnList )/>
		<cfset var currVal = ""/>
		<cfset var rendered = ""/>
		<cfsavecontent variable="rendered">
			<cfoutput>
				#startWorksheet( nodeName )#
					#startTable()#
						#startNewRow()#
							<cfloop from="1" to="#columnCount#" index="curr">
								<cfset currVal = ListGetAt( columnListed, curr ) />
								#renderCell( curr, ListGetAt( columnListed, curr ), 'String' )#
							</cfloop>
						#endRow()#
						<cfloop query="arguments.queryObj">
							#startNewRow()#
								<cfloop from="1" to="#columnCount#" index="curr">
									<cfset currVal = ListGetAt( columnListed, curr ) />
									#renderCell( curr, arguments.queryObj[ currVal ][currentrow], determineType( trim( arguments.queryObj[ currVal ][currentrow] ) ) )#
								</cfloop>
							#endRow()#
						</cfloop>
					#endTable()#
				<cfif getAutoFilter()><AutoFilter x:Range="R1C1:R1C#columnCount#" xmlns="urn:schemas-microsoft-com:office:excel"> </AutoFilter></cfif>
		 
				#endWorkSheet()#
			</cfoutput>
		</cfsavecontent>
					
		<cfreturn rendered/>
	</cffunction>
	
	<cffunction name="processArrayOfQueries" access="private" returntype="any" hint="I loop through an array of queries and process each query." output="false">
		<cfargument name="queryArray" type="any" required="true" default=""/>
		<cfargument name="workSheetNames" type="any" required="true" default=""/>
		
		<cfset var excelData = ""/>
		<cfset var finalProduct = ""/>
		<cfloop from="1" to="#ArrayLen(arguments.queryArray)#" index="o">
			<cfset excelData = processQuery( arguments.queryArray[o], getWorksheetName( arguments.workSheetNames, o ) )/>
			<cfset finalProduct = finalProduct & excelData/>
		</cfloop>
		
		<cfreturn finalProduct/>
	</cffunction>
	
	<cffunction name="startWorkBook" access="private" returntype="any" hint="I return the start workbook descriptor" output="false">
		<cfreturn "<Workbook xmlns=""urn:schemas-microsoft-com:office:spreadsheet"" xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"">"/>
	</cffunction>
	
	<cffunction name="EndWorkbook" access="private" returntype="any" hint="I return the end workbook descriptor" output="false">
		<cfreturn "</Workbook>"/>
	</cffunction>
	
	<cffunction name="startWorksheet" access="private" returntype="any" hint="I return the start worksheet descriptor" output="false">
		<cfargument name="node" type="any" required="true" default=""/>
		
		<cfset var wks = ""/>
		<cfsavecontent variable="wks"><cfoutput><Worksheet ss:Name="#arguments.node#"></cfoutput></cfsavecontent>
		
		<cfreturn wks/>
	</cffunction>
	
	<cffunction name="EndWorksheet" access="private" returntype="any" hint="I return the end worksheet descriptor" output="false">
		<cfreturn "</Worksheet>"/>
	</cffunction>
	
	<cffunction name="startTable" access="private" returntype="any" hint="I return the start table descriptor" output="false">
		<cfreturn "<Table>"/>
	</cffunction>
	
	<cffunction name="EndTable" access="private" returntype="any" hint="I return the end table descriptor" output="false">
		<cfreturn "</Table>"/>
	</cffunction>
	
	<cffunction name="startNewRow" access="private" returntype="any" hint="I return the start row descriptor" output="false">
		<cfreturn "<Row>"/>
	</cffunction>
	
	<cffunction name="EndRow" access="private" returntype="any" hint="I return the end row descriptor" output="false">
		<cfreturn "</Row>"/>
	</cffunction>
	
	<cffunction name="renderCell" access="private" returntype="any" hint="I render the cell portion and pass it back" output="false">
		<cfargument name="pos" type="any" required="true" default=""/>
		<cfargument name="data" type="any" required="true" default=""/>
		<cfargument name="bType" type="any" required="true" default=""/>
		
		<cfset var hCell = ""/>
		<cfsavecontent variable="hCell"><cfoutput><Cell ss:StyleID="#getThisStyle(arguments.data)#" ss:Index="#pos#"><Data ss:Type="#arguments.bType#">#scrubData(trim(arguments.data))#</Data></Cell></cfoutput></cfsavecontent>
		
		<cfreturn hCell/>
	</cffunction>
	
	<cffunction name="determineType" access="private" returntype="any" hint="I figure out the type and pass it back" output="true">
		<cfargument name="data" type="any" required="true" default=""/>
		
		<cfif getObjFilter().localeDates>
			<cfif LSisDate( arguments.data )>
				<cfreturn "DateTime"/>
			</cfif>
		<cfelse>
			<cfif isDate( arguments.data )>
				<cfreturn "DateTime"/>
			</cfif>
		</cfif>
		<cfif isNumeric( arguments.data )>
			<cfreturn "Number"/>
		<cfelseif isBoolean( arguments.data )>
			<cfreturn "Boolean"/>
		<cfelse>
			<cfreturn "String"/>
		</cfif>
		
	</cffunction>
	
	<cffunction name="getWorksheetName" access="private" returntype="any" hint="I get the worksheet name from the named list" output="false">
		<cfargument name="workSheetList" type="any" required="true" default=""/>
		<cfargument name="pos" type="any" required="true" default=""/>
		
		<cfif ListLen(arguments.workSheetList) gte pos>
			<cfreturn ListGetAt(arguments.workSheetList, pos)/>
		<cfelse>
			<cfreturn  'Worksheet ' & pos/>
		</cfif>
		
		
	</cffunction>
	
	<cffunction name="getXMLHeader" access="private" returntype="any" hint="I return the xml header" output="false">
		<cfreturn "<?xml version=""1.0""?>"/>
	</cffunction>
	
	<cffunction name="scrubData" access="private" returntype="any" hint="I scrub the data and clean it up for the XML" output="false">	
		<cfargument name="data" type="any" required="true" default=""/>
		
		<cfset var hVal = ""/>
		<cfswitch expression="#determineType(arguments.data)#">
			<cfcase value="DateTime">
				<cfset hVal = #dateformat(arguments.data,'yyyy-mm-dd')# & 'T' & timeformat(arguments.data,'HH:mm')/>
			</cfcase>
			<cfcase value="Number">
				<cfset hVal = #arguments.data#/>
			</cfcase>
			<cfcase value="Boolean">
				<cfset hVal = #iif(arguments.data,DE('1'),DE('0'))#/>				
			</cfcase>			
			<cfdefaultcase>
				<cfif isBinary(arguments.data)>
					<cfset hVal = ""/>
				<cfelse>
					<cfset hVal = #arguments.data#/>
				</cfif>
			</cfdefaultcase>
		</cfswitch>
		<cfreturn xmlFormat(trim(hVal))/>
	</cffunction>
	
	<cffunction name="getThisStyle" access="private" returntype="any" hint="I check the type and determine the proper style for the cell" output="false">
		<cfargument name="data" type="any" required="true" default=""/>
			
		<cfif determineType(arguments.data) eq 'DateTime'>
			<cfreturn 'GenDate'/>
		<cfelse>
			<cfreturn 'Default'/>
		</cfif>		
		
	</cffunction>
	
	<cffunction name="getStyles" access="private" returntype="any" hint="" output="false">
		
		<cfset var sStyles = ""/>
		<cfsavecontent variable="sStyles">
			<Styles>
				<Style ss:ID="Default" ss:Name="Normal">
					<Alignment ss:Vertical="Bottom"/>
					<Borders/>
					<Font/>
					<Interior/>
					<NumberFormat/>
					<Protection/>
				</Style>
				<Style ss:ID="GenDate">
					<NumberFormat ss:Format="General Date"/>
				</Style>
			</Styles>
		</cfsavecontent>
		
		<cfreturn sStyles/>
	</cffunction>
	<cffunction name="convertQueryToArrayOfQueries" returntype="array">
			<cfargument name="qry" required="Yes" >
			
			<!--- 
			*****************************************************************
				groupFieldDisplay is added to allow a more useful option 
				for displaying tab names. 
				For example, an id field might be used to group but the 
				corresponding id description field would be more useful 
				for the display. Not required but defaults to the groupField 
				which is required. 
				
			*****************************************************************
			--->
			
			
			<cfset variables.tabNames = "">
			
			<cfoutput query="arguments.qry">
				<cfif not listFindNoCase(tabNames, arguments.qry[getGroupFieldDisplay()][currentrow])>
					<cfset variables.tabNames = ListAppend(variables.tabNames,arguments.qry[getGroupFieldDisplay()][currentrow])>
				</cfif>
			</cfoutput>
			<!--- use tabnames for the worksheetnames --->
			<cfset setWorksheetNames(variables.tabNames)>
			
			<cfset qryArray = ArrayNew(1)>
			<cfset counter = 1>
			<cfoutput query="arguments.qry"  group="#getGroupField()#">
				<cfset qryArray[counter]= QueryNew(arguments.qry.columnlist)>
					<cfoutput>
						<cfset queryAddrow(qryArray[counter],1)>
						<cfloop list="#arguments.qry.columnlist#" index="listIndex">
							<cfset qryArray[counter][listIndex][qryArray[counter].recordcount ] = arguments.qry[listIndex][currentrow]>
						</cfloop>
					</cfoutput>
				<cfset counter = counter +1>
	
			</cfoutput>

			<cfreturn qryArray>
	
	</cffunction>
</cfcomponent>