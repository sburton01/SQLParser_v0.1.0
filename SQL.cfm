<!---
	Author		: Darwan Leonardo Sitepu
	E-mail		: dlns2001@yahoo.com
	Application	: SQL QUERY SELECT
	Date		: September 7, 2010 7:41 AM
	Summary		: Run SQL SELECT In U Browser And Export Data To Excell (Get All Datasource In U Coldfusion Administrator)
	Revisions	: 	1. Wednesday, June 15, 2011 1:22:33 Change : Tomas (tomamais@gmail.com)
--->
<cfoutput>
<cfscript>
	ServiceFactory = CreateObject("java", "coldfusion.server.ServiceFactory");
</cfscript> 
<title>SQL Runner / Coldfusion Version #ServiceFactory.LicenseService.getMajorVersion()#</title>
<cfparam name="DSN" default="darwanlns">
<cfparam name="show_debugging" default="a">
<cfparam name="convert_TO_Excell" default="NO">
<CfParam name="txtSearch" default="insert,delete,update,drop,alter,create,sp_columns">
<cfif show_debugging neq "dlns2001@yahoo.com">
	<cfsetting showdebugoutput="false">
<Cfelseif convert_TO_Excell EQ "NO">
	<cfsetting showdebugoutput="true">
<Cfelse>
	<cfsetting showdebugoutput="false">
</cfif>
<cffunction name="getdsn" returntype="struct" output="false">
	<cfargument name="dsn" type="string" required="yes">
	<!--- initialize variables --->
	<cfset var dsservice = "">
	<!--- get "factory" --->
	<cfobject action="create"
	         type="java"
	         class="coldfusion.server.ServiceFactory"
	         name="factory">
	<!--- get datasource service --->
	<cfset dsservice = factory.getdatasourceservice()>
	<cfif not structKeyExists(dsservice.datasources,dsn)>
		<cfthrow message="#dsn# is not a registered datasource.">
	</cfif>
	<!--- get dsn --->
	<cfreturn duplicate(dsservice.datasources[dsn])>
</cffunction>
<cfscript>
function getDSNs() {
	var factory = createObject("java","coldfusion.server.ServiceFactory");
	return factory.getDataSourceService().getNames();
}

function dumpVarList(variable) {
    
    // ASSIGN THE delim
    var delim="#Chr(13)##Chr(10)#";
    var var2dump=arguments.variable;
    var label = "";
    var newdump="";
    var keyName="";
    var loopcount=0;
    
    if(arrayLen(arguments) gte 2) delim=arguments[2];
    if(arrayLen(arguments) gte 3) label=arguments[3];
    
    // THE VARIABLE IS A SIMPLE VALUE, SO OUTPUT IT
    if(isStruct(var2dump)) {
      
        for(keyName in var2dump) {
            if(isSimpleValue(var2dump[keyName])) {
			if (keyname contains "DRIVER")writeOutput(var2dump[keyName]);
                
            }
            //else dumpVarList(var2dump[keyName],delim,keyname);
        }
    }
        // THE VARIABLE EXISTS, BUT IS NOT A TYPE WE WISH TO DUMP OUT
   

    return;
}
</cfscript>
<cfform name="frmSQL" action="#CGI.Script_Name#?#CGI.Query_String#" method="post" preservedata="true">
	<cfinput type="hidden" name="show_debugging" value="#show_debugging#">
	<cfinput type="hidden" name="txtSearch" value="#txtSearch#">
<table width="100%">
	<cfinput type="hidden" name="convert_TO_Excell" value="#convert_TO_Excell#">
	<tr>
		<td>Select DataSource</td>
		<td>&nbsp;:&nbsp;</td>
		<td>
			<!---<cfparam name="reqsess_dsn" default="">
			<cfselect name="reqsess_dsn">
				<cfloop array=#getDSNs()# index="name">
					<option value="#name#" <cfif reqsess_dsn Eq name>selected</cfif>>#name# &nbsp;[#dumpVarList(getdsn(name))#]
				</cfloop>
			</cfselect>
            --->
            <!--- Change : Tomas (tomamais@gmail.com) --->
            	<cfset DSNs = getDSNs()>
                <cfparam name="reqsess_dsn" default="">
                <cfselect name="reqsess_dsn">
                    <cfloop index="name" from="1" to="#arrayLen(DSNs)#">
                        <option value="#DSNs[name]#" <cfif reqsess_dsn Eq name>selected</cfif>>#DSNs[name]# &nbsp;[#dumpVarList(getdsn(DSNs[name]))#]
                    </cfloop>
               </cfselect>
           <!--- End Change : Tomas (tomamais@gmail.com) --->
		</td>
	</tr>
	<tr>
		<td>View Output By</td>
		<td>&nbsp;:&nbsp;</td>
		<td>
			<cfparam name="view_output" default="applet">
			<cfselect name="view_output">
				<option value="table" title="All Browser Support(Quick connection)" <cfif view_output Eq "table">selected</cfif>>Table
				<cfif val(ServiceFactory.LicenseService.getMajorVersion()) GT 7>
					<option value="applet" title="Browser Must Support With Java (Quick connection)" <cfif view_output Eq "applet">selected</cfif>>Applet
					<option value="flash" title="Browser Must Support With Flash (Medium connection)" <cfif view_output Eq "flash">selected</cfif>>Flash
					<option value="html" title="All Browser Support (Medium connection)" <cfif view_output Eq "html">selected</cfif>>Html
				</cfif>
			</cfselect>
		</td>
	</tr>
	<tr><td  colspan="5">Query</td></tr>
	<tr><td  colspan="5"><cftextarea name="txtSQL"  rows="20" cols="100%" style="width: 100%; height: 280px;; font-family: courier;" class="codepress sql linenumbers-off" id="sql"></cftextarea></td></tr>
	<tr><td nowrap colspan="5">
		<cfinput type="Button" name="cmdProcess" value="Process Query" onclick="javascript:process_query();">&nbsp;&nbsp;<cfinput type="Button" name="cmdConvert_Excell" value="Convert To Excell" onclick="javascript:convert_excell();">
	</td></tr>
</table>
<Script>
	function convert_excell()
	{
		document.frmSQL.convert_TO_Excell.value ="YES";
		document.frmSQL.submit(); 
	}
	function process_query()
	{
		document.frmSQL.convert_TO_Excell.value ="NO";
		document.frmSQL.submit();
	}
</Script>
</cfform>


<br><br>

<cfform name="Frm_Query" enablecab="yes" format="html">

<cfif isdefined("txtSQL")>
<Cfset TxtRunSql = "#evaluate(DE(preservesinglequotes(txtSQL)))#"> 
	<cftry>
		<cfquery name="qRes" datasource="#reqsess_dsn#">
			#evaluate(DE(preservesinglequotes(TxtRunSql)))#
		</cfquery>

		<cfif isdefined("txtSQL") and (lcase(left(txtSQL,6)) Neq "insert" OR lcase(left(txtSQL,6)) Neq "update" OR lcase(left(txtSQL,6)) Neq "delete")>
			<cfif IsDefined("qRes.recordcount")>
				#qRes.recordcount# found<br><br>
				<cfif qRes.recordcount>
						<cfif convert_TO_Excell NEQ "YES">
							<cfif ServiceFactory.LicenseService.getMajorVersion() GT 7>
								<Cfif UCASE(TRIM(view_output)) NEQ "TABLE">
									<cfinclude template="cfgrid.cfm">									
								<Cfelse>
									<cftable query="qRes" colspacing="2" border="1" htmltable colheaders>
									<cfloop index="iFld" list="#qRes.columnlist#"><cfcol header = "<b>#iFld#</b>" align = "Center" text= "#evaluate("qRes.#iFld#")# &nbsp;"></cfloop>
									</cftable> 
								</Cfif>
							<cfelse>
								<cftable query="qRes" colspacing="2" border="1" htmltable colheaders>
									<cfloop index="iFld" list="#qRes.columnlist#"><cfcol header = "<b>#iFld#</b>" align = "Center" text= "#evaluate("qRes.#iFld#")#"></cfloop>
								</cftable> 
							</cfif>
						<cfelse>
                        	<cfsavecontent variable="strExcelData">
							<!---<cfheader name="content-disposition" value="inline; filename=QUERY_TO_EXCELL[#DateFormat(now(),"ddmmyyyy")#].xls"> 
							<cfcontent type="application/msexcel">--->
                            <style>
                             .XLSNIP
                            	{mso-style-parent:style0;
                            	color:black;
                            	font-size:8.0pt;
                            	font-family:Verdana, sans-serif;
                            	mso-font-charset:0;
                            	mso-number-format:"\@";
                            	vertical-align:top;
                            	border-top:none;
                            	border-right:none;
                            	background:white;
                            	mso-pattern:auto none;
                            	white-space:normal;}
                            </style>
							<table border="1" cellspacing="0" cellpadding="0">
								<tr>
									<cfloop index="iFld" list="#qRes.columnlist#">
										<td bgcolor="##C0C0C0">#iFld#</td>
									</cfloop>
								</tr>
								<cfloop query="qRes">
									<cfset icur=qRes.currentrow>
									<tr>
										<cfloop index="iFld" list="#qRes.columnlist#">
											<cfif IsDate(evaluate("qRes.#iFld#[#icur#]"))>
												<td>#dateformat(evaluate("qRes.#iFld#[#icur#]"),"dd/MMM/yyyy")#</td>
											<cfelse>
												<td class="XLSNIP">#evaluate("qRes.#iFld#[#icur#]")#</td>
											</cfif>
										</cfloop>
									</tr>
								</cfloop>
							</table>
                            </cfsavecontent>
							<cfinput type="hidden" name="convert_TO_Excell" value="NO">
                            <cfheader name="Content-Disposition" value="attachment; filename=QUERY_TO_EXCELL[#DateFormat(now(),'ddmmyyyy')#].xls" />
							<cfif StructKeyExists( URL, "text" )>
                                <cfcontent type="application/msexcel" reset="true" /><cfset WriteOutput(strExcelData.Trim())/>
                                <cfexit />
                            <cfelse>
                                <cfcontent type="application/msexcel" variable="#ToBinary(ToBase64(strExcelData.Trim()))#"/> 
                            </cfif>
						</cfif>
				</cfif>
			</cfif>
		</cfif>
		<cfcatch> 
			<cfrethrow> 
			<cfloop collection = #cfcatch# item = "c"> <br><cfif IsSimpleValue(cfcatch[c])>#c# = #cfcatch[c]#</cfif> </cfloop>
		</cfcatch>
	</cftry>
</cfif>
</Cfform>
</cfoutput>