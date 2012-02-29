<cfcomponent>

	<!--- structure to hold private component settings --->
	<cfset variables.settings = StructNew() />

	<cfset variables.defaultEWSURL = "https://your.server.com/path/to/Envelope.asmx?WSDL" />

	<cffunction name="init" output="false" access="public" hint="Constructor">
		<cfargument name="ewsLocation" type="string" required="false" default="#variables.defaultEWSURL#" hint="WSDL location for the EWS web service" />
		<cfargument name="apiKey" type="string" required="true" hint="API Key for interacting with .NET services" />
		<cfargument name="wsTimeout" type="numeric" required="false" default="10" hint="Timeout duration for accessing the webservice" />
		<cfargument name="errorEmailFrom" type="string" required="true" hint="email address that error reports should come from" />
		<cfargument name="errorEmailTo" type="string" required="true" hint="email address to send error reports to" />
		<cfset variables.settings.ewsLocation = arguments.ewsLocation />
		<cfset variables.settings.apiKey = arguments.apiKey />
		<cfset variables.settings.timeout = arguments.wsTimeout />
		<cfset variables.settings.errorEmailFrom = arguments.errorEmailFrom />
		<cfset variables.settings.errorEmailTo = arguments.errorEmailTo />
		<cfreturn this />
	</cffunction>

	<!--- :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: --->
	<!--- :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: --->
	<!--- EMAIL FUNCTIONS --->
	<!--- :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: --->
	<!--- :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: --->

	<cffunction name="createEmail" access="public" output="false" returnType="boolean">
		<cfargument name="emlUserAddress" type="string" required="true" hint="Format: user@domain.com" />
		<cfargument name="emlTORecipients" type="string" required="true" hint="Comma-delimited list. DO NOT USE SEMICOLONS!" />
		<cfargument name="emlCCRecipients" type="string" required="true" hint="Comma-delimited list. DO NOT USE SEMICOLONS!" />
		<cfargument name="emlBCCRecipients" type="string" required="true" hint="Comma-delimited list. DO NOT USE SEMICOLONS!" />
		<cfargument name="strSubject" type="string" required="true" />
		<cfargument name="strBody" type="string" required="true" />
		<cfargument name="strImportance" type="string" required="false" default="Normal" hint="One of: Normal, High, Low. Not case-sensitive" />
		<cfset var local = {} />
		<cftry>
			<cfset arguments.strImportance = stringWhitelist(arguments.strImportance, "Normal|High|Low", "Normal", false) />
			<cfset local.rspContainer = makeApiRequest("CreateEmail", arguments) />
			<cfcatch>
				<cfset errorEmail(cfcatch, local) />
				<cfreturn false />
			</cfcatch>
		</cftry>
		<cfreturn true />
	</cffunction>

	<cffunction name="getEmailUnreadCount" access="public" output="false" returnType="numeric">
		<cfargument name="emlUserAddress" type="string" required="true" hint="Format: user@domain.com" />
		<cfset var local = {} />
		<cftry>
			<cfset local.rspContainer = makeApiRequest("GetEmailUnreadCount", arguments) />
			<cfcatch>
				<cfset errorEmail(cfcatch, local) />
				<cfreturn -1 />
			</cfcatch>
		</cftry>
		<cfreturn local.rspContainer.SimpleData.XmlText />
	</cffunction>

	<cffunction name="getEmail" output="false" returntype="Query" access="public">
		<cfargument name="emlUserAddress" type="string" required="true" hint="Format: user@domain.com" />
		<cfargument name="intNumMessages" type="numeric" required="false" default="10" />
		<cfset var local = {} />
		<cftry>
			<cfset local.rspContainer = makeApiRequest("GetEmail", arguments) />
			<!--- deserialize the XML into a query object --->
			<cfscript>
				local.itemContainer = local.rspContainer.TableData.diffgram;
				local.qryMessages = QueryNew("subject,from,to,datetimesent,link");
				local.items = 0;
				local.i = 0;

				//check for no emails returned
				if (not structKeyExists(local.itemContainer, "DocumentElement")){ return qryMessages; }

				//get the aray of data
				local.items = local.itemContainer.DocumentElement.XmlChildren;

				//add rows for each message returned from the web service
				QueryAddRow(local.qryMessages, ArrayLen(local.items));

				for (local.i = 1; local.i lte arrayLen(local.items); local.i = local.i + 1){
					//for some reason, sometimes a message is missing the TO attribute, etc.
					//wrapping each statement in a try/catch/continue will allow an individual field to be missing and still pick up the rest of the fields
					local.fieldList = "Subject,From,To,DateTimeSent,Link";
					for (local.j = 1; local.j lte listLen(local.fieldList); local.j = local.j + 1){
						local.key = listGetAt(local.fieldList, local.j);
						if (structKeyExists(local.items[local.i], local.key)){
							QuerySetCell(local.qryMessages, local.key, local.items[local.i][local.key].XmlText, local.i);
						}
					}
				}

				return local.qryMessages;
			</cfscript>
			<cfcatch>
				<cfset errorEmail(cfcatch, local) />
				<cfreturn QueryNew("subject,from,to,datetimesent,link") />
			</cfcatch>
		</cftry>
	</cffunction>

	<!--- :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: --->
	<!--- :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: --->
	<!--- CALENDAR FUNCTIONS --->
	<!--- :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: --->
	<!--- :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: --->

	<cffunction name="createCalendarItem" access="public" output="false" returnType="boolean">
		<cfargument name="emlUserAddress" type="string" required="true" hint="Format: user@domain.com" />
		<cfargument name="lstReqAttendeeEmail" type="string" required="false" default="" hint="Comma-delimited list of email addresses to invite. If all invitee fields are blank, an appointment will be created." />
		<cfargument name="lstOptAttendeeEmail" type="string" required="false" default="" hint="Comma-delimited list of email addresses to invite. If all invitee fields are blank, an appointment will be created." />
		<cfargument name="dtCalItemStart" type="string" required="true" hint="Time values optional. If provided, use 24-hour format. Format: YYYY-MM-DD HH:MM:SS" />
		<cfargument name="dtCalItemEnd" type="string" required="true" hint="Time values optional. If provided, use 24-hour format. Format: YYYY-MM-DD HH:MM:SS" />
		<cfargument name="strCalItemSubject" type="string" required="true" hint="Subject of the appointment/meeting request" />
		<cfargument name="strCalItemLocation" type="string" required="false" default="" hint="Location of appointment/meeting. Optional." />
		<cfargument name="strCalItemBody" type="string" required="false" default="" hint="Body content. HTML allowed (wrap with <html></html> tags)" />
		<cfargument name="blnAllDayFlag" type="boolean" required="false" default="false" hint="whether or not the appointment/meeting is considered 'all-day'" />
		<cfargument name="strCalItemCategories" type="string" required="false" default="" hint="Comma-separated list of category names (strings)" />
		<cfset var local = {} />
		<cftry>
			<cfset local.rspContainer = makeApiRequest("CreateCalendarItem", arguments) />
			<cfcatch>
				<cfset errorEmail(cfcatch, local) />
				<cfreturn false />
			</cfcatch>
		</cftry>
		<cfreturn true />
	</cffunction>

	<cffunction name="getCalendarItemById" access="public" output="false" returntype="query">
		<cfargument name="strUniqueId" type="string" required="true" hint="ID value of the record you want to get" />
		<cfargument name="emlUserAddress" type="string" required="true" hint="Format: user@domain.com" />
		<cfset var local = {} />
		<cftry>
			<cfset local.rspContainer = makeApiRequest("GetCalendarItemById", arguments) />
			<cfset local.rtnCols = "Subject,location,Start,End,Body,Categories,ReqAttendees,OptAttendees,AllDay,Link" />
			<!--- deserialize the XML into a struct object --->
			<cfscript>
				if (isSimpleValue(local.rspContainer) and len(local.rspContainer) eq 0){
					return StructNew();
				}
				local.itemContainer = local.rspContainer.TableData.diffgram;
				local.qryCalendar = queryNew(local.rtnCols);
				local.structCalendar = StructNew();
				local.items = 0;

				//check for no calendar events for specified range, return empty set
				if (not structKeyExists(local.itemContainer, "DocumentElement")){ return local.structCalendar; }

				//get the array of data
				local.items = local.itemContainer.DocumentElement.XmlChildren;

				//popuplate the query
				for (local.i = 1; local.i lte arrayLen(local.items); local.i = local.i + 1){
					local.fieldList = local.rtnCols;
					queryAddRow(local.qryCalendar);
					for (local.j = 1; local.j lte listLen(local.fieldList); local.j = local.j + 1){
						local.key = listGetAt(local.fieldList, local.j);
						if (structKeyExists(local.items[local.i], local.key)){
							if (local.j eq 3 or local.j eq 4){
								//for date fields, convert them to a usable format
								QuerySetCell(local.qryCalendar, local.key, DateConvertISO8601(local.items[local.i][local.key].XmlText, 0 - GetTimeZoneInfo().utcHourOffset), local.i);
							}else{
								QuerySetCell(local.qryCalendar, local.key, local.items[local.i][local.key].XmlText, local.i);
							}
						}
					}
				}

				return local.qryCalendar;
			</cfscript>
			<cfcatch>
				<cfset errorEmail(cfcatch, local) />
				<cfreturn QueryNew("subject,location,start,end,allday,link") />
			</cfcatch>
		</cftry>
	</cffunction>

	<cffunction name="getCalendarItems" output="false" returntype="Query" access="public">
		<cfargument name="emlUserAddress" type="string" required="true" hint="Format: user@domain.com" />
		<cfargument name="dtRangeBegin" type="string" required="true" hint="Time optional, use 24-hour format if specified. Format: YYYY-MM-DD HH:MM" />
		<cfargument name="dtRangeEnd" type="string" required="true" hint="Time optional, use 24-hour format if specified. Format: YYYY-MM-DD HH:MM" />
		<cfset var local = {} />
		<cftry>
			<cfset local.rspContainer = makeApiRequest("GetCalendarItems", arguments) />
			<!--- deserialize the XML into a query object --->
			<cfscript>
				if (isSimpleValue(local.rspContainer) and len(local.rspContainer) eq 0){
					return QueryNew("subject,location,start,end,allday,link");
				}
				local.itemContainer = local.rspContainer.TableData.diffgram;
				local.qryCalendar = QueryNew("subject,location,start,end,allday,link");
				local.items = 0;

				//check for no calendar events for specified range, return empty set
				if (not structKeyExists(local.itemContainer, "DocumentElement")){ return local.qryCalendar; }

				//get the array of data
				local.items = local.itemContainer.DocumentElement.XmlChildren;

				//create enough rows in the query to store the calendar data
				QueryAddRow(local.qryCalendar, arrayLen(local.items));

				//popuplate the query
				for (local.i = 1; local.i lte arrayLen(local.items); local.i = local.i + 1){
					local.fieldList = "Subject,location,Start,End,AllDay,Link";
					for (local.j = 1; local.j lte listLen(local.fieldList); local.j = local.j + 1){
						local.key = listGetAt(local.fieldList, local.j);
						if (structKeyExists(local.items[local.i], local.key)){
							if (local.j eq 3 or local.j eq 4){
								//for date fields, convert them to a usable format
								QuerySetCell(local.qryCalendar, local.key, DateConvertISO8601(local.items[local.i][local.key].XmlText, 0 - GetTimeZoneInfo().utcHourOffset), local.i);
							}else{
								QuerySetCell(local.qryCalendar, local.key, local.items[local.i][local.key].XmlText, local.i);
							}
						}
					}
				}

				return local.qryCalendar;
			</cfscript>
			<cfcatch>
				<cfset errorEmail(cfcatch, local) />
				<cfreturn QueryNew("subject,location,start,end,allday,link") />
			</cfcatch>
		</cftry>
	</cffunction>

	<cffunction name="getCalendarItemsDetailed" output="false" returntype="Query" access="public">
		<cfargument name="emlUserAddress" type="string" required="true" hint="Format: user@domain.com" />
		<cfargument name="dtRangeBegin" type="string" required="true" hint="Time optional, use 24-hour format if specified. Format: YYYY-MM-DD HH:MM" />
		<cfargument name="dtRangeEnd" type="string" required="true" hint="Time optional, use 24-hour format if specified. Format: YYYY-MM-DD HH:MM" />
		<cfset var local = {} />
		<cfset var rtnCols = "subject,location,start,end,allday,link,body,categories,reqattendees,optattendees,id" />
		<cftry>
			<cfset local.rspContainer = makeApiRequest("GetCalendarItemsDetailed", arguments) />
			<!--- deserialize the XML into a query object --->
			<cfscript>
				if (isSimpleValue(local.rspContainer) and len(local.rspContainer) eq 0){
					return QueryNew(rtnCols);
				}
				local.itemContainer = local.rspContainer.TableData.diffgram;
				local.qryCalendar = QueryNew(rtnCols);
				local.items = 0;

				//check for no calendar events for specified range, return empty set
				if (not structKeyExists(local.itemContainer, "DocumentElement")){ return local.qryCalendar; }

				//get the array of data
				local.items = local.itemContainer.DocumentElement.XmlChildren;

				//create enough rows in the query to store the calendar data
				QueryAddRow(local.qryCalendar, arrayLen(local.items));

				//popuplate the query
				for (local.i = 1; local.i lte arrayLen(local.items); local.i = local.i + 1){
					local.fieldList = rtnCols;
					for (local.j = 1; local.j lte listLen(local.fieldList); local.j = local.j + 1){
						local.key = listGetAt(local.fieldList, local.j);
						if (structKeyExists(local.items[local.i], local.key)){
							if (local.j eq 3 or local.j eq 4){
								//for date fields, convert them to a usable format
								QuerySetCell(local.qryCalendar, local.key, DateConvertISO8601(local.items[local.i][local.key].XmlText, 0 - GetTimeZoneInfo().utcHourOffset), local.i);
							}else{
								QuerySetCell(local.qryCalendar, local.key, local.items[local.i][local.key].XmlText, local.i);
							}
						}
					}
				}

				return local.qryCalendar;
			</cfscript>
			<cfcatch>
				<cfset errorEmail(cfcatch, local) />
				<cfreturn QueryNew("subject,location,start,end,allday,link,body,categories,reqattendees,optattendees") />
			</cfcatch>
		</cftry>
	</cffunction>

	<cffunction name="updateCalendarItem" output="false" returntype="boolean" access="public">
		<cfargument name="strUniqueId" type="string" required="true" hint="Unique ID of the calendar record you want to update" />
		<cfargument name="emlUserAddress" type="string" required="true" hint="Format: user@domain.com" />
		<cfargument name="lstReqAttendeeEmail" type="string" required="false" default="" hint="Comma-delimited list of email addresses to invite. If all invitee fields are blank, an appointment will be created." />
		<cfargument name="lstOptAttendeeEmail" type="string" required="false" default="" hint="Comma-delimited list of email addresses to invite. If all invitee fields are blank, an appointment will be created." />
		<cfargument name="dtCalItemStart" type="string" required="true" hint="Time values optional. If provided, use 24-hour format. Format: YYYY-MM-DD HH:MM:SS" />
		<cfargument name="dtCalItemEnd" type="string" required="true" hint="Time values optional. If provided, use 24-hour format. Format: YYYY-MM-DD HH:MM:SS" />
		<cfargument name="strCalItemSubject" type="string" required="true" hint="Subject of the appointment/meeting request" />
		<cfargument name="strCalItemLocation" type="string" required="false" default="" hint="Location of appointment/meeting. Optional." />
		<cfargument name="strCalItemBody" type="string" required="false" default="" hint="Body content. HTML allowed (wrap with <html></html> tags)" />
		<cfargument name="blnAllDayFlag" type="boolean" required="false" default="false" hint="whether or not the appointment/meeting is considered 'all-day'" />
		<cfargument name="strCalItemCategories" type="string" required="false" default="" hint="Comma-separated list of category names (strings)" />
		<cfset var local = {} />
		<cftry>
			<cfset local.rspContainer = makeApiRequest("UpdateCalendarItem", arguments) />
			<cfcatch>
				<cfset errorEmail(cfcatch, local) />
				<cfreturn false />
			</cfcatch>
		</cftry>
		<cfreturn true />
	</cffunction>

	<!--- :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: --->
	<!--- :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: --->
	<!--- CONTACT FUNCTIONS --->
	<!--- :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: --->
	<!--- :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: --->

	<cffunction name="createContact" access="public" output="false" returnType="string">
		<cfargument name="emlUserAddress" type="string" required="true" hint="Format: user@domain.com" />
		<cfargument name="strCtctGivenName" type="string" required="false" default="" hint="First name of new contact" />
		<cfargument name="strCtctSurname" type="string" required="false" default="" hint="Last name of new contact" />
		<cfargument name="strCtctCompanyName" type="string" required="false" default="" hint="Company name of new contact" />
		<cfargument name="strCtctHomePhone" type="string" required="false" default="" hint="Home phone number of new contact" />
		<cfargument name="strCtctWorkPhone" type="string" required="false" default="" hint="Work phone number of new contact" />
		<cfargument name="strCtctMobilePhone" type="string" required="false" default="" hint="Mobile phone number of new contact" />
		<cfargument name="emlCtctEmail1" type="string" required="false" default="" hint="Email1 of new contact" />
		<cfargument name="emlCtctEmail2" type="string" required="false" default="" hint="Email2 of new contact" />
		<cfargument name="emlCtctEmail3" type="string" required="false" default="" hint="Email3 of new contact" />
		<cfargument name="strCtctHomeAddrStreet" type="string" required="false" default="" hint="Home Address: Street & House Number" />
		<cfargument name="strCtctHomeAddrCity" type="string" required="false" default="" hint="Home Address: City" />
		<cfargument name="strCtctHomeAddrStateAbbr" type="string" required="false" default="" hint="Home Address: State Code (2 letters)" />
		<cfargument name="strCtctHomeAddrPostalCode" type="string" required="false" default="" hint="Home Address: Zip" />
		<cfargument name="strCtctHomeAddrCountry" type="string" required="false" default="" hint="Home Address: Country" />
		<cfargument name="strCtctWorkAddrStreet" type="string" required="false" default="" hint="Work Address: Street & House Number" />
		<cfargument name="strCtctWorkAddrCity" type="string" required="false" default="" hint="Work Address: City" />
		<cfargument name="strCtctWorkAddrStateAbbr" type="string" required="false" default="" hint="Work Address: State Code (2 letters)" />
		<cfargument name="strCtctWorkAddrPostalCode" type="string" required="false" default="" hint="Work Address: Zip" />
		<cfargument name="strCtctWorkAddrCountry" type="string" required="false" default="" hint="Work Address: Country" />
		<cfset var local = {} />
		<cfset local.rspContainer = makeApiRequest("CreateContact", arguments) />
		<cfif local.rspContainer.StatusCode.XmlText eq 200>
			<!--- create successful, return the ID --->
			<cfreturn local.rspContainer.SimpleData.XmlText />
		<cfelse>
			<cfreturn "" />
		</cfif>
	</cffunction>

	<cffunction name="updateContact" access="public" output="false" returnType="boolean">
		<cfargument name="strUniqueId" type="string" required="true" hint="The unique id returned when creating or listing contacts" />
		<cfargument name="emlUserAddress" type="string" required="true" hint="Format: user@domain.com" />
		<cfargument name="strCtctGivenName" type="string" required="false" default="" hint="First name of new contact" />
		<cfargument name="strCtctSurname" type="string" required="false" default="" hint="Last name of new contact" />
		<cfargument name="strCtctCompanyName" type="string" required="false" default="" hint="Company name of new contact" />
		<cfargument name="strCtctHomePhone" type="string" required="false" default="" hint="Home phone number of new contact" />
		<cfargument name="strCtctWorkPhone" type="string" required="false" default="" hint="Work phone number of new contact" />
		<cfargument name="strCtctMobilePhone" type="string" required="false" default="" hint="Mobile phone number of new contact" />
		<cfargument name="emlCtctEmail1" type="string" required="false" default="" hint="Email1 of new contact" />
		<cfargument name="emlCtctEmail2" type="string" required="false" default="" hint="Email2 of new contact" />
		<cfargument name="emlCtctEmail3" type="string" required="false" default="" hint="Email3 of new contact" />
		<cfargument name="strCtctHomeAddrStreet" type="string" required="false" default="" hint="Home Address: Street & House Number" />
		<cfargument name="strCtctHomeAddrCity" type="string" required="false" default="" hint="Home Address: City" />
		<cfargument name="strCtctHomeAddrStateAbbr" type="string" required="false" default="" hint="Home Address: State Code (2 letters)" />
		<cfargument name="strCtctHomeAddrPostalCode" type="string" required="false" default="" hint="Home Address: Zip" />
		<cfargument name="strCtctHomeAddrCountry" type="string" required="false" default="" hint="Home Address: Country" />
		<cfargument name="strCtctWorkAddrStreet" type="string" required="false" default="" hint="Work Address: Street & House Number" />
		<cfargument name="strCtctWorkAddrCity" type="string" required="false" default="" hint="Work Address: City" />
		<cfargument name="strCtctWorkAddrStateAbbr" type="string" required="false" default="" hint="Work Address: State Code (2 letters)" />
		<cfargument name="strCtctWorkAddrPostalCode" type="string" required="false" default="" hint="Work Address: Zip" />
		<cfargument name="strCtctWorkAddrCountry" type="string" required="false" default="" hint="Work Address: Country" />
		<cfset var local = {} />
		<cftry>
			<cfset local.rspContainer = makeApiRequest("UpdateContact", arguments) />
			<cfcatch>
				<cfset errorEmail(cfcatch, local.rspContainer)/>
				<cfreturn false />
			</cfcatch>
		</cftry>
		<cfreturn true />
	</cffunction>

	<cffunction name="deleteContact" access="public" output="false" returntype="boolean">
		<cfargument name="strUniqueId" type="string" required="true" hint="The unique id returned when creating or listing contacts" />
		<cfargument name="emlUserAddress" type="string" required="true" hint="The email address of the account containing the contact to be deleted" />
		<cfset var local = {} />
		<cftry>
			<cfset local.rspContainer = makeApiRequest("DeleteContact", arguments) />
			<cfcatch>
				<cfset errorEmail(cfcatch, local.rspContainer)/>
				<cfreturn false />
			</cfcatch>
		</cftry>
		<cfreturn true />
	</cffunction>

	<cffunction name="getContacts" access="public" output="false" returnType="query">
		<cfargument name="emlUserAddress" type="string" required="true" hint="Format: user@domain.com" />
		<cfset var local = {} />
		<cftry>
			<cfset local.rspContainer = makeApiRequest("GetContacts", arguments) />
			<!--- deserialize the XML into a query object --->
			<cfscript>
				local.itemContainer = local.rspContainer.TableData.diffgram;
				local.qryContacts = QueryNew("GivenName,Surname,FileAs,CompanyName,Link,HomePhone,BusinessPhone,MobilePhone,Email1,Email2,Email3,HomeStreet,HomeCity,HomeState,HomeZip,HomeCountry,WorkStreet,WorkCity,WorkState,WorkZip,WorkCountry,UniqueId");
				local.items = 0;

				//check for no calendar events for specified range, return empty set
				if (not structKeyExists(local.itemContainer, "DocumentElement")){ return local.qryContacts; }

				//get the array of data
				local.items = local.itemContainer.DocumentElement.XmlChildren;

				//create enough rows in the query to store the contact data
				QueryAddRow(local.qryContacts, arrayLen(local.items));

				//popuplate the query
				for (local.i = 1; local.i lte arrayLen(local.items); local.i = local.i + 1){
					local.fieldList = "GivenName,Surname,FileAs,CompanyName,Link,HomePhone,WorkPhone,MobilePhone,Email1,Email2,Email3,HomeStreet,HomeCity,HomeState,HomeZip,HomeCountry,WorkStreet,WorkCity,WorkState,WorkZip,WorkCountry,UniqueId";
					for (local.j = 1; local.j lte listLen(local.fieldList); local.j = local.j + 1){
						local.key = listGetAt(local.fieldList, local.j);
						if (structKeyExists(local.items[local.i], local.key)){
							QuerySetCell(local.qryContacts, local.key, local.items[local.i][local.key].XmlText, local.i);
						}
					}
				}

				return local.qryContacts;
			</cfscript>
			<cfcatch>
				<cfset errorEmail(cfcatch, local) />
				<cfreturn QueryNew("GivenName,Surname,FileAs,CompanyName,Link,HomePhone,BusinessPhone,MobilePhone,Email1,Email2,Email3,HomeStreet,HomeCity,HomeState,HomeZip,HomeCountry,WorkStreet,WorkCity,WorkState,WorkZip,WorkCountry,UniqueId") />
			</cfcatch>
		</cftry>
	</cffunction>

	<!--- :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: --->
	<!--- :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: --->
	<!--- TASK FUNCTIONS --->
	<!--- :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: --->
	<!--- :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: --->

	<cffunction name="createTask" access="public" output="false" returnType="boolean">
		<cfargument name="emlUserAddress" type="string" required="true" hint="Format: user@domain.com" />
		<cfargument name="strTaskSubject" type="string" required="true" hint="Subject of new task" />
		<cfargument name="strTaskStartDate" type="string" required="false" default="#dateFormat(now(), 'yyyy-mm-dd')# #timeformat(now(), 'HH:MM:SS')#" hint="date/time the task should be started" />
		<cfargument name="strTaskDueDate" type="string" required="false" default="#dateformat(now(), 'YYYY-MM-DD')# #timeformat(now(), 'HH:MM:SS')#" hint="date/time the task is due" />
		<cfargument name="strTaskImportance" type="string" required="false" default="Normal" hint="Not case-sensitive. One of: High, Low, Normal" />
		<cfargument name="strTaskStatus" type="string" required="false" default="NotStarted" hint="Not case sensitive. One of: NotStarted, WaitingOnOthers, Deferred, InProgress, Completed" />
		<cfargument name="strTaskBody" type="string" required="false" default="" hint="Body of task object. Html allowed (wrap with <html></html> tags)" />
		<cfset var local = {} />
		<cftry>
			<cfset local.rspContainer = makeApiRequest("CreateTask", arguments) />
			<cfcatch>
				<cfset errorEmail(cfcatch, local) />
				<cfreturn false />
			</cfcatch>
		</cftry>
		<cfreturn true />
	</cffunction>

	<cffunction name="getTasks" access="public" output="false" returnType="query">
		<cfargument name="emlUserAddress" type="string" required="true" hint="Format: user@domain.com" />
		<cfset var local = {} />
		<cftry>
			<cfset local.rspContainer = makeApiRequest("GetTasks", arguments) />
			<!--- deserialize the XML into a query object --->
			<cfscript>
				local.itemContainer = local.rspContainer.TableData.diffgram;
				local.qryTasks = QueryNew("Subject,Created,StartDate,DueDate,Importance,Status,Link");
				local.items = 0;

				//check for no calendar events for specified range, return empty set
				if (not structKeyExists(local.itemContainer, "DocumentElement")){ return local.qryTasks; }

				//get the array of data
				local.items = local.itemContainer.DocumentElement.XmlChildren;

				//create enough rows in the query to store the contact data
				QueryAddRow(local.qryTasks, arrayLen(local.items));

				//popuplate the query
				for (local.i = 1; local.i lte arrayLen(local.items); local.i = local.i + 1){
					local.fieldList = "Subject,Created,StartDate,DueDate,Importance,Status,Link";
					for (local.j = 1; local.j lte listLen(local.fieldList); local.j = local.j + 1){
						local.key = listGetAt(local.fieldList, local.j);
						if (structKeyExists(local.items[local.i], local.key)){
							QuerySetCell(local.qryTasks, local.key, local.items[local.i][local.key].XmlText, local.i);
						}
					}
				}

				return local.qryTasks;
			</cfscript>
			<cfcatch>
				<cfset errorEmail(cfcatch, local) />
				<cfreturn  QueryNew("Subject,Created,StartDate,DueDate,Importance,Status,Link") />
			</cfcatch>
		</cftry>
	</cffunction>

	<!--- :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: --->
	<!--- :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: --->
	<!--- PRIVATE FUNCTIONS --->
	<!--- :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: --->
	<!--- :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: --->

	<cffunction name="makeApiRequest" output="false" access="private">
		<cfargument name="method">
		<cfargument name="methodArgs" />
		<cfset var local = StructNew() />
		<cftry>
			<!--- append authentication credentials to every request --->
			<cfset arguments.methodArgs.apiKey = variables.settings.apiKey />
			<!--- make the request --->
			<cfhttp
				url="#replaceNoCase(variables.settings.ewsLocation, "?WSDL", "", "all")#/#arguments.method#"
				result="local.wsResult"
				method="post"
				timeout="#variables.settings.timeout#"
				throwonerror="true"
			>
				<cfloop collection="#arguments.methodArgs#" item="local.key">
					<cfhttpparam type="formfield" name="#local.key#" value="#arguments.methodArgs[local.key]#" />
				</cfloop>
			</cfhttp>
			<cfif trim(local.wsResult.fileContent) eq "Connection Timeout">
				<cfthrow message="Request timeout while connecting to .Net API" detail="#local.wsResult.statusCode#" />
			</cfif>
			<cfif not isXML(local.wsResult.FileContent)>
				<cfthrow message="ASP.NET WS did not return valid XML." detail="#local.wsResult.FileContent#" />
			</cfif>
			<cfset local.wsResponse = xmlParse(local.wsResult.Filecontent) />
			<cfset local.rspContainer = local.wsResponse.WhartonEWSResponse />
			<cfif local.rspContainer.StatusCode.XmlText neq 200>
				<cfthrow message="#local.rspContainer.Msg.XmlText#" errorcode="#local.rspContainer.StatusCode.XmlText#" />
			</cfif>
			<cfcatch>
				<cfset local.arguments = arguments />
				<cfset errorEmail(cfcatch, local) />
				<cfreturn "" />
			</cfcatch>
		</cftry>
		<cfreturn local.rspContainer />
	</cffunction>

	<cffunction name="stringWhitelist" access="private" output="false" returnType="string" hint="makes sure that a string value is on a whitelist; returns the default value if specified value is not white-listed">
		<cfargument name="currentValue" type="string" required="true" />
		<cfargument name="whiteList" type="string" required="true" hint="pipe-delimited list of valid values" />
		<cfargument name="defaultValue" type="string" required="true" hint="The value that should be returned if currentValue is not in the whitelist" />
		<cfargument name="caseSensitive" type="boolean" required="false" default="false" hint="Whether or not the white list and currentValue should be compared CaSe SeNsItIvElY" />
		<cfscript>
			if (arguments.caseSensitive){
				if (listFind(arguments.whiteList, arguments.currentValue, '|')){
					return arguments.currentValue;
				}
			}else{
				if (listFindNoCase(arguments.whiteList, arguments.currentValue, '|')){
					return arguments.currentValue;
				}
			}
			return arguments.defaultValue;
        </cfscript>
	</cffunction>

	<cffunction name="errorEmail" access="private" output="false" returnType="void">
		<cfargument name="dumpException" type="any" required="true" hint="Usually the CFCatch object. this is the exception object that you want to be dumped in the email." />
		<cfargument name="dumpData" type="any" required="false" hint="any extra data that you want dumped in the email" />
		<cfif structKeyExists(dumpData, "apiKey")>
			<cfset dumpData.apiKey = "****************" />
		</cfif>
		<cfif structKeyExists(dumpData, "arguments") and structKeyExists(dumpData.arguments, "apiKey")>
			<cfset dumpData.arguments.apiKey = "****************" />
		</cfif>
		<cfmail from="#variables.settings.errorEmailFrom#" to="#variables.settings.errorEmailTo#" subject="[ENVELOPE] Error trapped in CFC wrapper" type="html">
			<h2>Error Timestamp:</h2>
			<p><cfoutput>#dateFormat(now(), 'yyyy-mm-dd')# #timeFormat(now(), 'HH:MM:SS tt')#</cfoutput></p>
			<h2>Error Info:</h2>
			<cfdump var="#arguments.dumpException#" />
			<cfif structKeyExists(arguments, "dumpData")>
				<h2>Data:</h2>
				<cfdump var="#arguments.dumpData#" />
			</cfif>
		</cfmail>
	</cffunction>

	<cfscript>
	/*
	 * Convert a date in ISO 8601 format to an ODBC datetime.
	 *
	 * @param ISO8601dateString      The ISO8601 date string. (Required)
	 * @param targetZoneOffset      The timezone offset. (Required)
	 * @return Returns a datetime.
	 * @author David Satz (david_satz@hyperion.com)
	 * @version 1, September 28, 2004
	 * @url http://cflib.org/index.cfm?event=page.udfbyid&udfid=1144
	 */
	function DateConvertISO8601(ISO8601dateString, targetZoneOffset) {
	    var rawDatetime = left(ISO8601dateString,10) & " " & mid(ISO8601dateString,12,8);
	    // adjust offset based on offset given in date string
	    if (uCase(mid(ISO8601dateString,20,1)) neq "Z")
	        targetZoneOffset = targetZoneOffset - val(mid(ISO8601dateString,20,3));
	    return DateAdd("h", targetZoneOffset, CreateODBCDateTime(rawDatetime));
	}
	</cfscript>

</cfcomponent>