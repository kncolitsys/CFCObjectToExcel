<cftry>
	<cfsetting showdebugoutput="false">
	<cfheader name="Content-Disposition" value="attachment; filename=test.xls"> 
	<cfcontent type="application/msexcel" reset="Yes">  
	<!--- <cfcontent type="application/vnd.ms-excel"> --->
	
	<!--- 
		Array of Queries usage
	 --->
	<cfset tmpQry = ArrayNew(1)/>
	<cfset names = "John Doe, Marvin Champ, Alex Money, Lu Sancea, Margot Gumpley"/>
	<cfset ages = "26, 34, 45, 33, 56"/>
	<cfset gender = "Male, Male, Female, Male, Female"/>
	<cfset isAwesome = "yes, no, no, yes, false"/>
	<cfset dob = "7/8/2009|7/8/2009|7/8/2009|7/8/2009|7/8/2009"/>
	
	<cfset myStruct = StructNew()/>
	<cfset myStruct.localeDates = false/>
	<cfset myStruct.autoFilter = false/>
	<cfset myStruct.worksheetNames = "First One,Second"/>
	
	<cfloop from="1" to="5" index="j">
		<cfset tmpQry[j] = QueryNew("Name, Age, Gender, DOB, Awesome")/>
		<cfloop from="1" to="5" index="k">
			<cfset QueryAddRow( tmpQry[j] )/>
			<cfset QuerySetCell( tmpQry[j], "Name", ListGetAt( names,k ) ) />
			<cfset QuerySetCell( tmpQry[j], "Age", ListGetAt( ages,k ) ) />
			<cfset QuerySetCell( tmpQry[j], "Gender", ListGetAt( gender,k ) ) />
			<cfset QuerySetCell( tmpQry[j], "DOB", ListGetAt( dob,k, "|") ) />
			<cfset QuerySetCell( tmpQry[j], "Awesome", ListGetAt( isAwesome,k ) ) />
		</cfloop>
	</cfloop>
		
	<cfset excel = CreateObject( "component", "ObjectToExcel" ).init()/>
	<cfoutput>#excel.processObj( tmpQry, myStruct )#</cfoutput>
	
	
	
	
	<!--- 
		Individual usage
	 --->
	<!---  
	<cfset tmpQry = QueryNew("A_Name, B_Age, C_Gender")/>
	<cfset names = "John Doe, Marvin Champ, Alex Money, Lu Sancea, Margot Gumpley"/>
	<cfset ages = "26, 34, 45, 33, 56"/>
	<cfset gender = "Male, Male, Female, Male, Female"/>
	<cfset wktNames = "People"/>
	<cfloop from="1" to="5" index="j">
		<cfset QueryAddRow( tmpQry )/>
		<cfset QuerySetCell( tmpQry, "A_Name", ListGetAt( names,j ) ) />
		<cfset QuerySetCell( tmpQry, "B_Age", ListGetAt( ages,j ) ) />
		<cfset QuerySetCell( tmpQry, "C_Gender", ListGetAt( gender,j ) ) />
	</cfloop>
		
	<cfset excel = CreateObject( "component", "ObjectToExcel" ).init()/>
	<cfoutput>#excel.processObj( tmpQry, wktNames )#</cfoutput>
	 --->

<cfcatch><cfdump var="#cfcatch#"/><cfabort /></cfcatch>
</cftry>
