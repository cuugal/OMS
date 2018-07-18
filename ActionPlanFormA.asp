<!-- #Include file="include\general.asp" -->
<%
	if SecurityCheck(1) = false then ' User must have write access for this department
		Response.Redirect ("restricted.asp")
		Response.end
	end if
%>

<% PageTitle = "Action Plan Form A"%>
	
<!-- #Include file="include\header.asp" -->

<%
	dim ActionPlan, con, Existing, apDeanDirector, apDeanOrDirector
	dim sqlFormA
	dim rsFormA
	
	
	class RowData
		public prevdate
		public rating
		public checked
		public description
	End class
	
	ActionPlan = Request("apID")
	Existing = Request("exists")
	
	set con	= server.createobject ("adodb.connection")
	con.open "DSN=ehs"
	
	' If there is already an action plan for this year we need to delete it
	if Existing <> "" then
		dim sqlExisting, rsExisting
		dim sqlDelete
	
		' Get the Action Plan ID for this year
		sqlExisting = "SELECT AP_ActionPlans.apID " & _
			  "FROM AP_ActionPlans " & _
			  "WHERE apStartYear = " & Existing & _
			  " AND apFaculty = " & Session("DepID")
		set rsExisting = con.Execute(sqlExisting)

		' If an Action Plan is found the delete it
		if not rsExisting.BOF then
		
			' WARNING: This delete does a cascadeing delete so related Audits, Compliance Checks and Service Agreements will be deleted
			sqlDelete = "DELETE AP_ActionPlans.apID " & _
						"FROM AP_ActionPlans " & _
						"WHERE apID = " & rsExisting("apID")
			con.Execute(sqlDelete)
			
			' Refresh the parent menu because we delete the AP
			RefreshParent()
		end if
	end if
	
	' Get the requirements 
	if ActionPlan = "" then
		sqlFormA = "SELECT irId, irStep, irName, irFormADescription, -100 as arRating, No as arSelected " & _
				   "FROM IN_Requirements " & _
				   "WHERE irActive = Yes " & _
				   "ORDER BY irDisplayOrder"
	else
		sqlFormA = "SELECT irId, irStep, irName, irFormADescription, arRating, arSelected " & _
				   "FROM IN_Requirements INNER JOIN AP_Requirements ON IN_Requirements.irId = AP_Requirements.arRequirement " & _
				   "WHERE irActive = Yes AND arActionPlan = " & ActionPlan & " " & _
				   "ORDER BY irDisplayOrder"
	end if
	
	set rsFormA = con.Execute(sqlFormA)
	
	sqlPrevious = "select faDate, FA_AuditDetails.* from FA_AuditDetails inner join FA_Audits on FA_Audits.faID = FA_AuditDetails.fdaudit "&_
				"where FA_Audits.faAuditType = 'management' and FA_AuditDetails.fdaudit in (select top 1 faID from FA_Audits inner join ap_ActionPlans on AP_ActionPlans.apID = FA_Audits.faActionPlan where apFaculty = "&Session("DepID")&" order by faDate desc)"
	
	'Response.write(sqlPrevious)
	set rsprevious = con.Execute(sqlPrevious)
	

	
	dim previousData, row, requirement
	set previousData = Server.CreateObject("Scripting.Dictionary")
	
	while not rsprevious.EOF
		set row = new RowData
		row.prevdate = rsprevious("faDate")
		row.rating = rsPrevious("fdRating")
		row.description = rsPrevious("fdEvidence")
		
		requirement = rsPrevious("fdRequirement")
		'Response.write(requirement&" "&row.rating&"<br/>")
		'Only add if key doesn't exist: primitive dictionary in vbscript should do this automatically, but alas it does not.
		If not previousData.Exists(requirement) Then
			previousData.Add requirement , row
		end if
		rsprevious.movenext
	wend
	
    'output for troubleshooting
	'dim allKeys,allItems, myKey, myItem
	'allKeys = previousData.Keys   'Get all the keys into an array
	'allItems = previousData.Items 'Get all the items into an array

	'For i = 0 To previousData.Count - 1 'Iterate through the array
	'  myKey = allKeys(i)   'This is the key value
	'  set myItem = allItems(i) 'This is the item value
	'  Response.Write("The " & i & " value in the Dictionary is " & myKey & " "&myItem.rating&"<br />")
	'Next

	
	if ActionPlan <> "" then
		' get the details of this action plan
		sqlAP = "select * from AP_ActionPlans where apID = " & ActionPlan
		set rsAP = con.execute (sqlAP)
		
		apDeanDirector = rsAP("apDDName")
		apDeanOrDirector = rsAP("apDDOption")
	end if
	
%>

<table width="100%" border="0" cellspacing="3">
<tr> 
	<td><div align="left"><img src="utslogo.gif" width="102" alt="The UTS home page" height="44" style="border:10px solid white" align="right"></div></td>
	<td></td>
</tr>
<tr> 
    <td colspan="2">
		<div align="center"><h2>HEALTH AND SAFETY PLAN</h2></div>
		<h2 align="center">INITIAL COMPLIANCE ASSESSMENT AND RISK IDENTIFICATION PROCESS</h2>
    </td>
</tr>
<tr> 
    <td colspan="2">
		The information you provide here will be used in developing your Health and Safety Plan. It allows you to make a first estimate of compliance ratings for your Plan, and takes you through a process to identify programs for specific hazards. The compliance ratings you provide here can be changed in your draft Plan.
	</td>
</tr>
<tr> 
    <td colspan="2">
		<form action="ActionPlanFormA_Process.asp" name="formA" method="post">
		<input type="hidden" name="action" value="">
		<input type="hidden" name="apID" value="<%=ActionPlan%>">
		<input type="hidden" name="department" value="<%=Session("DepID")%>">
    



	<table width="90%" id = "planA">
        <tr> 
			<td class="label"  width="15%">Faculty/Unit:</td>
			<td><%=Session("DepName")%></td>
        </tr>
        <tr> 
			<td class="label">Dean/Director:</td>
			<td> 
				<input type="text" name="apDeanDirector" value="<%=apDeanDirector%>">
				<i>&nbsp;&nbsp;&nbsp;Select:</i> 
				<input type="radio" name="apDeanOrDirector" value="Dean" <%if apDeanOrDirector = "Dean" then Response.Write " checked"%>>
				Dean &nbsp;&nbsp; 
				<input type="radio" name="apDeanOrDirector" value="Director" <%if apDeanOrDirector = "Director" then Response.Write " checked"%>>
				Director 
			</td>
        </tr>
        <tr> 
			<td class="label">Date:</td>
			<td><%=date()%></td>
        </tr>

        <!--tr>
			<td colspan="5">
				<br>
				<font size="3"><b>Health and Safety Procedures</b></font><br>
				<b>Rate your current level of compliance against each compliance requirement.</b><br>
				<i>Compliance Ratings: 0 = Non-Compliant, 1 = Non-Compliant - some action evident but not fully compliant, 2 = Compliant - just requires maintenance, 3 = Best practice evident</i>
			</td>
        </tr-->
        
	</table>


		<%
		dim date1
		dim req, prevData
		if(Ubound(previousData.keys) >0) then
			set prevData = previousdata.Item(previousData.keys()(0))
			date1 = prevData.prevdate
		end if
		
		
	
		
		%>
        <!--table border="1" width="90%">
        <tr>
			<td width="75%">Compliance Requirement</td>
			<td width="25%">Previous Compliance Rating (0, 1, 2, 3)</td>
			< td width="25%">Compliance Rating (0, 1, 2, 3) at date: <%=date1%></td>
        </tr>
<%
		dim rating
		while not rsFormA.EOF
			
             rating = Empty
			if rsFormA("irStep") = 1 then
			req = rsFormA("irID")
			if previousdata.Exists(req) then
				set prevData =  previousdata.Item(req)
				rating = prevData.rating
			end if
			
			' overwrite with the previous draft if previous draft exists
			if rsFormA("arRating") <> -100 then
				rating = rsFormA("arRating")
			end if
		
			if(rating = "" or rating <0) then rating = "00" end If
			' would be better if 00 was replaced with a dash
%>
			<!tr>
				<td><%=rsFormA("irFormADescription")%> </td>
			
				<td><% =rating %>
                    <input type="hidden" name="rate_<%=rsFormA("irID")%>" value="<% =rating %>" />              
				</td>
			</tr>
<%		
			end if

			rsFormA.movenext
		wend		
%>
        </table-->
        
		<p>
        <br>
		</p>
		<font size="3"><b>Specific Hazard Programs</b></font> - Select hazards identified in your Faculty/Unit<BR>
		<!--i>Compliance Ratings: 0 = Non-Compliant, 1 = Non-Compliant - some action evident but not fully compliant, 2 = Compliant - just requires maintenance, 3 = Best practice evident</i-->
		
        <table id = "planA" width="90%">
        <tr>
			<th width="15%"><strong>Specific Hazards</strong></th>
			<th><strong>For example - staff/student undertake work which involves:</strong></th>
			<!-- td width="25%">Compliance Rating (0, 1, 2, 3) at date: <%=date1%></td -->
			<!--td width="25%">Previous Compliance Rating (0, 1, 2, 3)</td-->
        </tr>
<%
		rsFormA.movefirst

		dim checked
		checked = ""
		
		while not rsFormA.EOF
            rating = Empty
			if rsFormA("irStep") = 2 then
			req = rsFormA("irID")
			if previousdata.Exists(req) then
				set prevData =  previousdata.Item(req)
				rating = prevData.rating
				checked = " checked"
			end if
			
			'over ride if loaded from draft form
			if(rsFormA("arSelected")) then
				checked = " checked"
			end if
			if(rsFormA("arSelected") = false) then
				checked = ""
			end if
			
			if(rsFormA("arRating") <> -100) then
				rating = rsFormA("arRating")
			end if
			if(rating = "" or rating <0) then rating = "-" end if
			
			
%>
			<tr> 
				<td class="label"> 
				  <input type="checkbox" name="req_<%=rsFormA("irID")%>" <%=checked%>>
				  <%=rsFormA("irName")%>
				</td>
				<td><%=rsFormA("irFormADescription")%> </td>
				<!--td><% =rating%></td-->
			</tr>
<%		
			end if

			rsFormA.movenext
		wend
%>
		</table>
		
		<br>

		<!--font size="3"><b>Health and Safety Management</b></font><br>
		<b>Rate your current level of compliance against each compliance requirement.</b><BR>
        <i>Compliance Ratings: 0 = Non-Compliant, 1 = Non-Compliant - some action evident but not fully compliant, 2 = Compliant - just requires maintenance, 3 = Best practice evident</i-->
        
		<!--table border="1" width="90%">
        <tr>
			<td width="75%">Compliance Requirement</td>
			<td width="25%">Compliance Rating (0, 1, 2, 3) at date: <%=date1%> </td>
			<td width="25%">Previous Compliance Rating (0, 1, 2, 3)</td>
        </tr>
<%
		rsFormA.movefirst

		while not rsFormA.EOF
            rating = Empty
			if rsFormA("irStep") = 3 then
			
			    req = rsFormA("irID")
			    if previousdata.Exists(req) then
				    set prevData =  previousdata.Item(req)
				    rating = prevData.rating
			    end if
			
			    ' Override previous if coming from draft
			    if(rsFormA("arRating") <> -100) then
				    rating = rsFormA("arRating")
			    end if
			    if(rating = "" or rating <0) then rating = "-" end if
			
    %>
			    <tr>
				    <td><%=rsFormA("irFormADescription")%></td>
				
				    <td><% =rating %> </td>
			    </tr>
<%		
			end if

			rsFormA.movenext
		wend
%>
        </table-->
<%
	if SecurityCheck(2) then ' User must have write access for this department
%>
		<p>Select either 
			<input type="submit" value="    Save as Draft    " onclick="javascript:formA.action.value='draft'"> 
		or check all compliance requirements are completed and 
			<input type="button" value="    Save as Final    " onclick="javascript:formA.action.value='final';DoSubmit();"></p>
<%
	else
%>
		<p align="center"><input type="submit" value="    Close Window    " onclick="window.close();"></p>
<%
	end if
%>

		</form>
		
	</td>
</tr>
</table>
<script type="text/javascript">
<!--

	function DoSubmit()
	{
		var message
		
		message = ""
		

	
		if (document.formA.apDeanDirector.value == "")
			message = message + " - You must enter the name of the Dean/Director\n"
		
		if (document.formA.apDeanOrDirector[0].checked == false && document.formA.apDeanOrDirector[1].checked == false)
			message = message + " - You must indicate if Dean or Director\n"
			
		if (message == "")
			document.formA.submit()
		else
			alert("The following error(s) have been detected:\n\n" + message)
	}

-->
</script>

<!-- #Include file="include\footer.asp" -->

