<!-- #Include file="include\general.asp" -->
<!--#include file="adovbs.inc"--> 

<%Response.Buffer = False%>
<%
	if SecurityCheck(2) = false then ' User must have write access for this department
		Response.Redirect ("restricted.asp")
		Response.end
	end if
%>
<% PageTitle = "Audit Form"%>
	
<!-- #Include file="include\header.asp" -->

<%
	dim con, ActionPlan, Mode, AuditID, audittype
	dim sqlAudit, sqlAudDetails
	dim rsAudit, rsAudDetails
	
	audittype = Request("type")
	
	set con = server.CreateObject("adodb.connection")
	con.Open "DSN=ehs"
	
	Set objCmd  = Server.CreateObject("ADODB.Command")
	objCmd.CommandType = adCmdText
	Set objCmd.ActiveConnection = con
	
	
	Mode = Request("Mode")				' Determines if this is a new AuditForm or an existing one
	AuditID = Request("faID")			' The Facility Audit ID
	ActionPlan = Request("apID")		' The Action Plan ID
	
	if Mode = "Edit" then
		sqlAudDetails = "Select * from FA_Audits where faID = ?"
		objCmd.CommandText = sqlAudDetails
		AuditID = AuditID&""
		objCmd.Parameters.Append objCmd.CreateParameter("faID", adWChar, adParamInput, 50)
		objCmd.Parameters("faID") = AuditID
		
		set rsAudDetails =  server.createobject("adodb.recordset")
		rsAudDetails.Open objCmd
	end if

	' Error detection: If the Lab Name and ActionPlan key is broken raise an error
	if request("error") = "Lab" then
	
	dim errortext
	Select case audittype
		Case "facility"
			errortext= "Workshop/Laboratory Name"
		Case "research"
			errortext= "Research Project"
		Case "curriculum"
			errortext= "Subject Name"
		Case "management"
			errortext= "Dean Director"
	End Select
%>
	<script type="text/javascript">
	<!--
		alert ("You cannot create a <%=audittype%> audit for the same <%=errortext%> in the same year. Please change the <%=errortext%> name.")
	//-->
	</script>
<%
	end if
%>

<form name="audit" action="AuditForm_Process.asp" method="post">
<input type="hidden" name="apID" value="<%=ActionPlan%>">
<input type="hidden" name="faID" value="<%=AuditID%>">
<input type="hidden" name="Mode" value="<%=Mode%>">
<input type="hidden" name="audittype" value="<%=audittype%>">
<input type="hidden" name="action" value="none">



<a href="http://www.uts.edu.au/"><img src="utslogo.gif" width="123" alt="The UTS home page" height="52" style="border:10px solid white" align="left"></a>


<table width="100%" border="0" cellspacing="3">
  <tr> 
	<td><h2>AUDIT FORM - <%=UCase(Left(audittype,1))& Mid(audittype,2) %> Audit</h2></td>
  </tr>
</table>



<table id = "planB">
		<tr> 
		  <td class="label" width="32%">
		  <% 
			Select case audittype
				Case "facility"
					Response.write "Laboratory/Workshop Supervisor:"
				Case "research"
					Response.write "Responsible Investigator:"
				Case "curriculum"
					Response.write "Subject Coordinator:"
				Case "management"
					Response.write "Audit Contact:"
			End Select
		  %>
		  <br><br></td>
		  <td class="label" width="33%">
			<input type="text" name="txt_Sup" size="50" maxlength="150" value="<% if Mode = "Edit" then Response.Write rsAudDetails("faSupervisor")%>"><br><br>
		  </td>
		  <td class="label">Faculty / Unit <br><br></td>
		  <td><%=Session("DepName")%><br><br></td>
		</tr>
		<tr> 
		  <td class="label">
		  <% 
			Select case audittype
				Case "facility"
					Response.write "Workshop/Laboratory Name:"
				Case "research"
					Response.write "Research Project:"
				Case "curriculum"
					Response.write "Subject Name:"
				Case "management"
					Response.write "Dean/Director:"
			End Select
		  %>
		  <br><br></td>
		  <td class="label">
			<input type="text" name="txt_Lab" size="50" maxlength="150" value="<% if Mode = "Edit" then Response.Write rsAudDetails("faLabName")%>"><br><br>
		  </td>
		  <td class="label">
		  <% 
			Select case audittype
				Case "facility"
					Response.write "Location:"
				Case "research"
					Response.write "Research Project Number:"
				Case "curriculum"
					Response.write "Subject Number:"
				Case "management"
					Response.write "Office"
			End Select
		  %>
		  <br><br></td>
		  <td class="label">
			<input type="text" name="txt_Loc" maxlength="150" value="<% if Mode = "Edit" then Response.Write rsAudDetails("faLocation")%>"><br><br>
		  </td>
		</tr>
		<tr> 
		  <td class="label">Name of Auditors<br><br></td>
		  <td class="label">
			<input type="text" name="txt_Assr" size="50" maxlength="150" value="<% if Mode = "Edit" then Response.Write rsAudDetails("faAssesName")%>"><br><br>
		  </td>
		  <td class="label"> Date of Audit<br><br></td>
		  <td class="label">
			<input type="text" name="txt_Date" id="datepicker"  maxlength="150"><br><br>
			

			<script type="text/javascript">
				$( "#datepicker" ).datepicker({
					dateFormat: "dd/mm/yy"
					});	
			</script>


			<!-- script type="text/javascript">
			$( "#datepicker" ).datepicker({
				dateFormat: "yy-mm-dd"
				});			
			</script -->

			<% if Mode = "Edit" then %>

			<!-- script type ="text/javascript">
				$( "#datepicker" ).datepicker( "setDate", "<%=rsAudDetails("faDate")%>" );
			</script -->


			<script type ="text/javascript">
						var parsedDate = $.datepicker.parseDate('dd/mm/yy', '<%=rsAudDetails("faDate")%>');
						$( "#datepicker" ).datepicker( "setDate",parsedDate );
			</script>

			<% end if %>
		  </td>
		</tr>
</table>


<br>
<h2>&nbsp;&nbsp;HOUSEKEEPING & OBSERVATIONS</h2>
<textarea name="txt_hous" rows="9" cols="120"><% if Mode = "Edit" then Response.Write rsAudDetails("faHouseKeeping")%></textarea>

<br>
<h2>&nbsp;&nbsp;ENTER COMPLIANCE FINDINGS</h2>
NOTE: Highlight non-compliances with capitalised "NOT".

	
	<%
	Function ShowStep(StepID)
		dim sqlStep, sqlReq
		dim rsStep, rsReq
		
		sqlStep = "Select stShortName from IN_Steps where stID = " & stepID
		set rsStep = con.Execute (sqlStep)
%>

<table id = "planB">
		 <tr> 
		  <th colspan="3" class="StepMenu"><%=rsStep("stShortName")%></th>
		</tr>
		<tr> 
		  <td class="label" width="30%">COMPLIANCE REQUIREMENTS</td>
		  <td class="label" width="35%">PROCEDURE</td>

		  <td class="label" width="35%">EVIDENCE OF COMPLIANCE</td>
		  <!--<td class="label" width="50%">EVIDENCE OF COMPLIANCE</td> -->
		</tr>
		
<%		
		' Show the requirements
		Set objCmd  = Server.CreateObject("ADODB.Command")
		objCmd.CommandType = adCmdText
		Set objCmd.ActiveConnection = con
		
		
		if Mode = "New" then
			sqlReq =	"SELECT irID " & _
						"FROM IN_Requirements INNER JOIN AP_Requirements ON IN_Requirements.irId = AP_Requirements.arRequirement " & _
						"WHERE IN_Requirements.irStep = ? AND AP_Requirements.arActionPlan = ? and arSelected = Yes" &_
						" AND ir"&audittype&" = Yes order by irDisplayOrder"
						
			objCmd.CommandText = sqlReq
			objCmd.Parameters.Append objCmd.CreateParameter("irStep", adInteger, adParamInput, 50)
			objCmd.Parameters("irStep") = cint(StepID)
			objCmd.Parameters.Append objCmd.CreateParameter("arActionPlan", adInteger, adParamInput, 50)
			objCmd.Parameters("arActionPlan") = cint(ActionPlan)
			
		else
			sqlReq =	"SELECT irID " & _
						"FROM (FA_AuditDetails INNER JOIN IN_Requirements ON FA_AuditDetails.fdRequirement = IN_Requirements.irId) INNER JOIN AP_Requirements ON IN_Requirements.irId = AP_Requirements.arRequirement " & _
						"WHERE fdAudit = ? and irStep = ? and arSelected = Yes AND arActionPlan = ?  AND ir"&audittype&" = Yes order by irDisplayOrder"
			objCmd.CommandText = sqlReq
			objCmd.Parameters.Append objCmd.CreateParameter("fdAudit", adInteger, adParamInput, 50)
			objCmd.Parameters("fdAudit") = cint(AuditID)
			objCmd.Parameters.Append objCmd.CreateParameter("irStep", adInteger, adParamInput, 50)
			objCmd.Parameters("irStep") = cint(StepID)
			objCmd.Parameters.Append objCmd.CreateParameter("arActionPlan", adInteger, adParamInput, 50)
			objCmd.Parameters("arActionPlan") = cint(ActionPlan)

		end if
 
		
		set rsReq =  server.createobject("adodb.recordset")
		rsReq.Open objCmd
		
		while not rsReq.EOF
			ShowRequirement (rsReq("irID"))
		
			rsReq.movenext
		wend
%>
		</table>
<%
	end function
	
	function ShowRequirement(ReqID)
		dim sqlReq
		dim rsReq
		dim data
		
		' Show the requirements
		Set objCmd  = Server.CreateObject("ADODB.Command")
		objCmd.CommandType = adCmdText
		Set objCmd.ActiveConnection = con
		
		
		if Mode = "New" then
			sqlReq =	"SELECT IN_Requirements.*, 0 as Rating " & _
						"FROM IN_Requirements INNER JOIN AP_Requirements ON IN_Requirements.irId = AP_Requirements.arRequirement " & _
						"WHERE AP_Requirements.arActionPlan = ? AND IN_Requirements.irId = ?"&_
						" AND ir"&audittype&" = Yes order by irDisplayOrder"



			objCmd.CommandText = sqlReq
			objCmd.Parameters.Append objCmd.CreateParameter("arActionPlan", adInteger, adParamInput, 50)
			objCmd.Parameters("arActionPlan") = cint(ActionPlan)
			objCmd.Parameters.Append objCmd.CreateParameter("irID", adInteger, adParamInput, 50)
			objCmd.Parameters("irID") = cint(ReqID)
			
		else

			sqlReq =	"SELECT IN_Requirements.*, fdRating as Rating " & _
						"FROM FA_Audits INNER JOIN (FA_AuditDetails INNER JOIN IN_Requirements ON FA_AuditDetails.fdRequirement = IN_Requirements.irId) ON FA_Audits.faID = FA_AuditDetails.fdAudit " & _
						"WHERE faID = ? AND fdRequirement = ? AND ir"&audittype&" = Yes order by irDisplayOrder"
			objCmd.CommandText = sqlReq
			objCmd.Parameters.Append objCmd.CreateParameter("faID", adInteger, adParamInput, 50)
			objCmd.Parameters("faID") = cint(AuditID)
			objCmd.Parameters.Append objCmd.CreateParameter("fdRequirement", adInteger, adParamInput, 50)
			objCmd.Parameters("fdRequirement") = cint(ReqID)
		end if
		
		set rsReq =  server.createobject("adodb.recordset")
		rsReq.Open objCmd
		data = ShowProcedures(ReqID)
		
		Dim row
		For i = 0 to ubound(data)
			set row = data(i)
			dim readLines
			'readlines = row.evidenceCompliance.split(vbCrLf)
			readLines = Split(row.evidenceCompliance, "|")
            'Automatic migration for old data split by comma.  Split on the old character, when we save it will split on the new
            if UBound(readLines) < 1 then
                readLines = Split(row.evidenceCompliance, ",")
            end if
			
			
			
%>
		<tr>
			<% if i=0 then %>
			<td <% if i=0 then %> rowspan=<%=ubound(data)+1%> <% end if %>><span class="label"><%=rsReq("irName")%></span><br><%=rsReq("irdescription")%><br/><br/>
				<span class="label">Compliance Rating</span>
				
				<SELECT name="rate_<%=rsReq("irID")%>">
					<OPTION value=-1 <%if rsReq("Rating") = -1 then Response.Write "selected"%>>N/A</OPTION>
					<OPTION value=0 <%if rsReq("Rating") = 0 then Response.Write " selected"%>>0</OPTION>
					<OPTION value=1 <%if rsReq("Rating") = 1 then Response.Write " selected"%>>1</OPTION>
					<OPTION value=2 <%if rsReq("Rating") = 2 then Response.Write " selected"%>>2</OPTION>
					<OPTION value=3 <%if rsReq("Rating") = 3 then Response.Write " selected"%>>3</OPTION>
					
				</SELECT>
				
			</td>
			<% end if %>
			<td> <%=row.procedures%></td>
			<% if Mode = "Edit" then %>
			
			<td>
				<TEXTAREA rows=3 cols=100 name="text_<%=ReqID%>"><% 
                    if i <= ubound(readLines) then 
                        Response.write( replace(Trim(readlines(i)), ";",",")) 
                    end if %></TEXTAREA>
			</td>
			<% else %>
			<td>
				<TEXTAREA rows=3 cols=100 name="text_<%=ReqID%>"><%=row.evidenceCompliance%></TEXTAREA>
			</td>
			<% end if %>
		</tr>
<%
	next
	end function
	
	
	class RowData
		public procID
		public procedures
		public evidenceCompliance
	End class

	function ShowProcedures(ReqID)
		dim sqlPro, sqlNumOpt
		dim rsPro,rsNumOpt
		dim var,  procedures, procedureID
		dim  rv, dcw
		
		redim dcw(-1)
		
		Set objCmd  = Server.CreateObject("ADODB.Command")
		objCmd.CommandType = adCmdText
		Set objCmd.ActiveConnection = con
	
		if Mode = "New" then
		
			sqlPro =	"SELECT IN_Procedures.*, prID, prChecked, prResponsibilities, prTimeframe, prTextBox " & _
						"FROM IN_Procedures INNER JOIN AP_Procedures ON IN_Procedures.ipID = AP_Procedures.prProcedure " & _
						"WHERE ipRequirement = ? and prActionPlan = ? and prChecked = Yes order by ipDisplayOrder"
'						"WHERE IN_Procedures.ipRequirement = ? and prActionPlan = ? and prChecked = Yes order by prID"		

			objCmd.CommandText = sqlPro
			
			objCmd.Parameters.Append objCmd.CreateParameter("irID", adInteger, adParamInput, 50)
			objCmd.Parameters("irID") = cint(ReqID)
			objCmd.Parameters.Append objCmd.CreateParameter("arActionPlan", adInteger, adParamInput, 50)
			objCmd.Parameters("arActionPlan") = cint(ActionPlan)
			
			'set rsPro = con.Execute (sqlPro) 
			set rsPro =  server.createobject("adodb.recordset")
			rsPro.Open objCmd
					
			if not rsPro.BOF then
				while not rsPro.EOF
					set rv = new RowData
					var = rsPro("ipDefaultAuditText")
					
					if not isnull(var) then 
						rv.evidenceCompliance = replace(var, vbCrLf, vbCrLf & "- ") & chr(13) & chr(10)
						
					end if
					
					procedures = rsPro("ipName")
					procedureId = rsPro("prID")
					if not isnull(procedures) then 
						rv.procedures = replace(procedures, vbCrLf, vbCrLf & "- ") & chr(13) & chr(10)
						rv.procID = procedureID
						'response.write(rv.procedures+"<br/>")
					end if
					
					'set results(rsPro.getRow) = rv
					push dcw, rv
					Set objCmd1  = Server.CreateObject("ADODB.Command")
					objCmd1.CommandType = adCmdText
					Set objCmd1.ActiveConnection = con
					
					sqlNumOpt = "SELECT count(*) as NumOptions " & _
								"FROM IN_Options INNER JOIN (AP_Procedures INNER JOIN AP_Options ON AP_Procedures.prID = AP_Options.aoPro) ON IN_Options.ioID = AP_Options.aoOption " & _
								"WHERE prActionPlan = ? AND prProcedure = ? AND ioActive = Yes"
					
					objCmd1.CommandText = sqlNumOpt
					objCmd1.Parameters.Append objCmd.CreateParameter("prActionPlan", adInteger, adParamInput, 50)
					objCmd1.Parameters("prActionPlan") = cint(ActionPlan)
					objCmd1.Parameters.Append objCmd.CreateParameter("prProcedure", adInteger, adParamInput, 50)
					objCmd1.Parameters("prProcedure") = cint(rsPro("ipID"))
					'set rsNumOpt = con.Execute (sqlNumOpt)
					set rsNumOpt =  server.createobject("adodb.recordset")
					rsNumOpt.Open objCmd1

					if rsNumOpt("NumOptions") > 0 then
						ShowOptions(rsPro("prID"))
					end if
					
					rsPro.movenext
				wend
			end if
		else
		
			'sqlPro =	"SELECT IN_Procedures.*, prID, prChecked, prResponsibilities, prTimeframe, prTextBox " & _
			'			"FROM IN_Procedures INNER JOIN AP_Procedures ON IN_Procedures.ipID = AP_Procedures.prProcedure " & _
			'			"WHERE IN_Procedures.ipRequirement = ? and prActionPlan = ? and prChecked = Yes order by prID"
			
			
			'sqlReq =	"SELECT irID " & _
			'		"FROM (FA_AuditDetails INNER JOIN IN_Requirements ON FA_AuditDetails.fdRequirement = IN_Requirements.irId) INNER JOIN AP_Requirements ON IN_Requirements.irId = AP_Requirements.arRequirement " & _
			'		"WHERE fdAudit = " & AuditID & " and irStep = " &  StepID & " and arSelected = Yes AND arActionPlan = " & ActionPlan
		
			' ?? IN_Procedures.iprequirement
	
			
			
sqlPro =	"SELECT FA_AuditDetails.fdEvidence, IN_Procedures.ipName, prID " & _
						"FROM (FA_AuditDetails INNER JOIN IN_Procedures ON IN_Procedures.ipRequirement = FA_AuditDetails.fdRequirement) " & _
						"INNER JOIN AP_Procedures ON IN_Procedures.ipID = AP_Procedures.prProcedure "&_
   "WHERE fdAudit = ? AND fdRequirement = ? and prActionPlan = ? and prChecked = Yes order by ipDisplayOrder"
   '     "WHERE fdAudit = ? AND fdRequirement = ? and prActionPlan = ? and prChecked = Yes order by prID"
					'	"WHERE fdAudit = ? AND fdRequirement = ? and prActionPlan = ? and prChecked = Yes order by ipDisplayOrder"
			
			

			'	Response.write(sqlPro&" | "&prID)		
			'set rsPro = con.Execute (sqlPro) 
			'Response.write(sqlPro&" | "&AuditID&" | "&ReqID)
			
			objCmd.CommandText = sqlPro
			
			objCmd.Parameters.Append objCmd.CreateParameter("fdAudit", adInteger, adParamInput, 50)
			objCmd.Parameters("fdAudit") = cint(AuditID)
			objCmd.Parameters.Append objCmd.CreateParameter("fdRequirement", adInteger, adParamInput, 50)
			objCmd.Parameters("fdRequirement") = cint(ReqID)
			objCmd.Parameters.Append objCmd.CreateParameter("prActionPlan", adInteger, adParamInput, 50)
			objCmd.Parameters("prActionPlan") = cint(ActionPlan)
			
			'set rsPro = con.Execute (sqlPro) 
			set rsPro =  server.createobject("adodb.recordset")
			rsPro.Open objCmd
		
			if not rsPro.BOF then
				while not rsPro.EOF
					set rv = new RowData
					
					rv.evidenceCompliance = rsPro("fdEvidence") 
					rv.procedures = rsPro("ipName")
					rv.procID = rsPro("prID")
					rsPro.movenext
					
					push dcw, rv
				wend
			end if
		end if
		
		 ShowProcedures = dcw
	end function
	
	function ShowOptions(OptID)
		dim sqlOpt
		dim rsOpt
		dim var

		Set objCmd  = Server.CreateObject("ADODB.Command")
		objCmd.CommandType = adCmdText
		Set objCmd.ActiveConnection = con


		sqlOpt =	"SELECT * FROM IN_Options INNER JOIN AP_Options ON IN_Options.ioID = AP_Options.aoOption " & _
					"WHERE AP_Options.aoPro = ? and aoChecked = Yes"
			objCmd.CommandText = sqlOpt	
			objCmd.Parameters.Append objCmd.CreateParameter("aoPro", adInteger, adParamInput, 50)
			objCmd.Parameters("aoPro") = cint(OptID)
			
		'set rsOpt = con.Execute (sqlOpt)
		set rsOpt =  server.createobject("adodb.recordset")
		rsOpt.Open objCmd
		
		
		while not rsOpt.EOF	
			var = rsOpt("ioDefaultAuditText")
		
			if not isnull(var) then 
				Response.Write "- " & var & vbcrlf
			end if
			
			rsOpt.movenext
		wend
	end function
	
	
	 
	
 
%>
<% ShowStep(3) %>
<BR><BR>
<% ShowStep(2) %>
<BR><BR>
<% ShowStep(1) %>
<BR><BR>
	
	</td>
  </tr>
  <tr><Td colspan="2"><% Response.write("Current Timestamp: " &Hour(Now)&":" &Minute(Now) &":"&Second(Now)&" "&Day(Now) & "/"& Month(Now)&"/"&Year(Now)) %></td></tr>
   <tr> 
	<td colspan="2"> 
	  <input type="button" value="    Save as Draft    " onclick="javascript:audit.action.value='draft';DoSubmit()">
			&nbsp;&nbsp;&nbsp;&nbsp;
	  <input type="button" value="    Save as Final    " onclick="javascript:audit.action.value='final';DoSubmit()">
	</td>
  </tr>
</table>

</form>



<script type="text/javascript">
<!--
	function DoSubmit() {
	
		//Replace any commas on the form as they can mess with form reloading
		

		var x = document.getElementsByTagName("textarea");
		
		for (var i=0; i<x.length; i++) 
		{
			//console.log(x[i]);
			//x[i].value = x[i].value.replace(/,/g,";");
		}


		var message
		
		message = ""
	
		if (document.audit.action.value == "draft") {
			if(!isDate($("#datepicker").val())){
				message = message + " - Date must be DD/MM/YYYY\n"
			}
			if (document.audit.txt_Lab.value == "")
				message = message + " - You must enter the Lab/Workshop name before you can save as a Draft\n"
		}
		else {
			if(!isDate($("#datepicker").val())){
				message = message + " - Date must be DD/MM/YYYY\n" 
			}
		
			if (document.audit.txt_Sup.value == "") 
				message = message + " - You must enter the Supervisor name before you can save as a final\n"
			
			if (document.audit.txt_Lab.value == "") 
				message = message + " - You must enter the Lab/Workshop name before you can save as a final\n"
				
			if (document.audit.txt_Loc.value == "")
				message = message + " - You must enter the Location before you can save as a final\n"
				
			if (document.audit.txt_Assr.value == "")
				message = message + " - You must enter the Assessor name before you can save as a final\n"
				
			if (document.audit.txt_Date.value == "")
				message = message + " - You must enter the Date before you can save as a final\n"
		}
		
		// required checking that does not include the first three steps
		if (message == "")
			document.audit.submit()
		else 
			alert("The following error(s) have been detected:\n\n" + message)
	}
	
	function isDate(txtDate)
	{
		var currVal = txtDate;
		if(currVal == '')
			return false;

		// var rxDatePattern = /^(\d{4})(\/|-)(\d{1,2})(\/|-)(\d{1,2})$/; //Declare Regex
		//var rxDatePattern = /^(\d{2})(\/|-)(\d{2})(\/|-)(\d{4})$/; //Declare Regex
		var rxDatePattern = /^(0?[1-9]|[12][0-9]|3[01])[\/\-](0?[1-9]|1[012])[\/\-]\d{4}$/;

		var dtArray = currVal.match(rxDatePattern); // is format OK?



		if (dtArray == null) 
			return false;

		//Checks for mm/dd/yyyy format.
		dtMonth = dtArray[3];
		dtDay= dtArray[5];
		dtYear = dtArray[1];        

		if (dtMonth < 1 || dtMonth > 12) 
			return false;
		else if (dtDay < 1 || dtDay> 31) 
			return false;
		else if ((dtMonth==4 || dtMonth==6 || dtMonth==9 || dtMonth==11) && dtDay ==31) 
			return false;
		else if (dtMonth == 2) 
		{
			var isleap = (dtYear % 4 == 0 && (dtYear % 100 != 0 || dtYear % 400 == 0));
			if (dtDay> 29 || (dtDay ==29 && !isleap)) 
					return false;
		}
		return true;
	}
//-->
</script>

<!-- #Include file="include\footer.asp" -->