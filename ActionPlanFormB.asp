<!-- #Include file="include\general.asp" -->

<%
	if SecurityCheck(1) = false then ' User must have write access for this department
		Response.Redirect ("restricted.asp")
		Response.end
	end if
%>

<% PageTitle = "Action Plan Form B"%>
	
<!-- #Include file="include\header.asp" -->

<%
	dim con, rs, rsRequirement, sql_getap, sql_getrequirements, ActionPlan, checked, anchor
					

	class RowData
		public prevdate
		public rating
		public checked
		public description
	End class
	
	class ProceduresData
		public responsibility
		public checked
		public timeframe
		Public textbox
	End class
	
	ActionPlan = Request("apID")
	
	' Refresh the parent now so that we have a link to the form B
	RefreshParent()
		
	set rs				= server.CreateObject ("adodb.recordset")
	set rsRequirement	= server.CreateObject ("adodb.recordset")
	
	set con				= server.createobject ("adodb.connection")
	con.open "DSN=ehs"
		
	sql_getap = "Select * from AP_ActionPlans where apID = " & ActionPlan
	set rs = con.Execute (sql_getap)
	
	sql_getrequirements =	"SELECT IN_Requirements.* " & _
							"FROM IN_Requirements INNER JOIN AP_Requirements ON IN_Requirements.irId = AP_Requirements.arRequirement " & _
							"WHERE arActionPlan = " & ActionPlan & " and irStep = 2 and arSelected = Yes order by irid"
	set rsRequirement = con.Execute (sql_getrequirements)
	
	
	sqlPrevious = "select faDate, FA_AuditDetails.* from FA_AuditDetails inner join FA_Audits on FA_Audits.faID = FA_AuditDetails.fdaudit "&_
				"where FA_Audits.faAuditType = 'management' and FA_AuditDetails.fdaudit in (select top 1 faID from FA_Audits inner join ap_ActionPlans on AP_ActionPlans.apID = FA_Audits.faActionPlan where apFaculty = "&Session("DepID")&" order by faDate desc)"
	
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
		If not previousData.Exists(requirement) Then
			previousData.Add requirement , row
		end if
		rsprevious.movenext
	wend
	
	sqlProcedures = "select top 1 * from AP_Procedures inner join AP_ActionPlans on AP_Procedures.prActionPlan = AP_ActionPlans.apID where prActionPlan <> "& ActionPlan &" and apFaculty = "& rs("apFaculty") &" order by prActionPlan desc"
	'Response.write(sqlProcedures)
	set procedures = con.Execute(sqlProcedures)
	
	dim procedureData, row1, procedure
	set procedureData = Server.CreateObject("Scripting.Dictionary")
	
	while not procedures.EOF
		set row1 = new ProceduresData

		row1.responsibility = procedures("prResponsibilities")
		row1.checked = procedures("prChecked")
		if row1.checked then
			'Response.write("HELLOW "&procedures("prProcedure")&" "&row1.checked&" "&procedures("prResponsibilities")&"<br/>")
		end if
		row1.timeframe = procedures("prTimeframe")
		procedure = procedures("prProcedure")
		procedureData.Add procedure , row1
		procedures.movenext
	wend
	
	
%>

<% 
	anchor = request("anchor") 
	if anchor <> "" then 
%> 
<script type="text/javascript">
 document.location.hash="#<%=anchor%>" 
</script> 
<% 
	end if 
%> 


 <!-- These are the steps for form B -->

<form action="ActionPlanFormB_Process.asp" method="post" name="formB">
<input type="hidden" name="action" value="none">
<input type="hidden" name="apID" value=<%=ActionPlan%>>
<input type="hidden" name="pointID" value="">
<input type="hidden" name="pointText" value="">

<table width="100%" border="0" cellspacing="3">
<tr>
	<td><!-- commented out the old EHS Branch logo <img src="ehslogo2.gif" width="142" height="111" alt="EHS logo" border="0">-->&nbsp;</td>
		<td><!--div align="right"><img src="utslogo.gif" width="135" height="30"></div-->
	<a href="http://www.uts.edu.au/"><img src="utslogo.gif" width="123" alt="The UTS home page" height="52" style="border:10px solid white" align="right"></a></td>
</tr>
<tr>
 <td colspan="2"><span class="label"><b>STATUS:</b></span>
		<%if rs("apcompleted")=0 then Response.Write "Draft Version "%>
		<%=rs("apVersion")%> 
	</td>
</tr>
<tr>
  <td colspan="2">
	  <div align="center"> 
		<h2><%=Session("DepName")%><br>Health and Safety Plan <%=rs("apStartYear")%> - <%=rs("apEndYear")%><BR>Date created: <%=rs("apCompletionDate")%></h2>
	  </div>
	</td>
  </tr>
 <tr>
	<td colspan="2">The <%=Session("DepName")%> is committed to providing a safe and healthy workplace for students, staff, contractors and visitors. Promoting a safe and healthy workplace is the responsibility of all staff.  This is consistent with the University's Health and Safety Policy, the UTS Health and Safety Plan, the NSW Work Health and Safety Act and associated legislation and with legislation exercised by the NSW Environment Protection Authority.</p>

	<p>This Plan was developed by:</p>

	<p>Names of participants (250 characters max)<BR>
	<textarea rows="5" cols="100" name="developedBy"><%=rs("apDevelopedBy")%></textarea></p>

	<p>This group:</p>
	<ul>
		<li>identified the hazards that may be encountered by staff, students, contractors and visitors,</li>
		<li>assessed the level of compliance of the faculty/unit with each of the compliance requirements,</li>
		<li>agreed on practical procedures to achieve and maintain compliance, and</li>
		<li>designated responsibilities and timeframes for implementation of these procedures.</li>
	</ul>
	
	<p>This Plan outlines the procedures, responsibilities and timeframes for:</p>
	  <ul>
		<li>Health and Safety Management</li>
		<li>Health and Safety Procedures</li>
		<li>Specific Hazard Programs</li>
		<ul>
		<li>Specific Hazards
		  <ul>
<%
	if not rsRequirement.BOF then 
		while not rsRequirement.EOF
			if rsRequirement("irid") <> 16 and rsRequirement("irid") <> 17 and rsRequirement("irid") <> 18 and rsRequirement("irid") <> 19 then 
				Response.Write "<li>" & rsRequirement("irName") & "</li>"
			end if
		
			rsRequirement.MoveNext
		wend
		
		rsRequirement.MoveFirst
%>   
		</ul> <!--dlj 1Apr5 removed an </ul> to put next list items down a level -->
<%
		while not rsRequirement.EOF
			if rsRequirement("irid") = 16 or rsRequirement("irid") = 17 or rsRequirement("irid") = 18 or rsRequirement("irid") = 19 then 
				Response.Write "<li>" & rsRequirement("irName") & "</li>"
			end if
		
			rsRequirement.MoveNext
		wend
	end if
%>		
		</ul>
	</td>
  </tr>
  <!--<tr>
	<td colspan=2 width=100%>
		<table border=1 width=100%>
		<tr>
			<td width=100%>
				<table width=100%>
				<tr>
					<td width=40%><br><br><b>Dean/Director</b></td>
					<td><br><b>Signature -</b><br><br><b>Date <%=date()%></b><br><br></td>
				</tr>
				</table> 
			</td>
		</tr>
		</table>
		<br>
	</td>
  </tr>-->
  <tr> 
	<td colspan="2"> 



<%
	Function ShowStep(StepID)
		dim sqlStep, sqlReq
		dim rsStep, rsReq
		
		sqlStep = "Select stName from IN_Steps where stID = " & stepID
		set rsStep = con.Execute (sqlStep)
%>
		<!--i>Compliance Ratings: 0 = Non-Compliant, 1 = Non-Compliant - some action evident but not fully compliant, 2 = Compliant - just requires maintenance, 3 = Best practice evident</i-->
		
		
		
		
		
		
		<table id = "planB">
		<tr> 
			<th colspan="4"> <%=rsStep("stName")%> </th>
		</tr>
		<%
		dim date1
		dim req, prevData

		if(Ubound(previousData.keys) >0) then
			set prevData = previousdata.Item(previousData.keys()(0))
			date1 = prevData.prevdate
		end if
		%>
		<tr> 
			<td class="label" width ="20%">COMPLIANCE REQUIREMENT</td>
			<!--td>
					<!--note that the date in the following line is month and year NOW. <%=monthname(month(Date()),true)%>&nbsp;<%=year(Date())%>
					It should be month and
					date report is saved i.e =rs("apCompletionDate")
					DLJ made this change 10 August 2004
					-->
				<!--span class="label">1. Compliance rating<br>
					(0, 1, 2, 3) at date: <%=date %></span--> 
			</td-->
			<td>
				<span class="label">Select what is required to comply</span><br>
				&nbsp;&nbsp;<font color="#FF0000"><strong>M</strong> - mandatory procedures to comply 
				with legislative requirements</font><br>
				<input type="checkbox" disabled>
				- optional activities that might be undertaken to achieve compliance
			</td>
			<td class="label" width="30%">Allocate responsibilities</td>
			<td class="label" width="10%">Allocate timeframe to complete by</td>
		</tr>
<%		
		' Show the requirements (only the ones that have been selected) 
		sqlReq =	"SELECT * FROM IN_Requirements INNER JOIN AP_Requirements ON IN_Requirements.irId = AP_Requirements.arRequirement " & _
					"WHERE irStep = " & StepID & " AND arActionPlan = " & ActionPlan & " and arSelected = Yes order by irDisplayOrder"
		set rsReq = con.Execute (sqlReq)
		
		while not rsReq.EOF
			ShowRequirement (rsReq("irID"))
		
			rsReq.movenext
		wend
%>
		</table>
<%
	end function
	
	function ShowRequirement(ReqID)
        rating = Empty
		dim sqlReq, sqlPro, sqlNumPro
		dim rsReq, rsPro, rsNumPro
		
		sqlReq =	"SELECT IN_Requirements.*, AP_Requirements.arRating " & _
					"FROM IN_Requirements INNER JOIN AP_Requirements ON IN_Requirements.irId = AP_Requirements.arRequirement " & _
					"WHERE arActionPlan = " & ActionPlan & " AND irId = " & ReqID
		set rsReq = con.Execute (sqlReq)
		
		sqlNumPro = "SELECT Count(*) AS NumProcedures " & _
					"FROM IN_Procedures " & _
					"WHERE ipRequirement = " & ReqID & " AND ipActive = Yes"
		set rsNumPro = con.Execute (sqlNumPro)
		
		dim req, prevData, rating
		req = rsReq("irID")
		if previousdata.Exists(req) then
			set prevData =  previousdata.Item(req)
			rating = prevData.rating	
		end if
		
		' overwrite with the previous draft if previous draft exists
		if rsReq("arRating") <> "" then
			rating = rsReq("arRating")
		end if
			
		if(rating = "" or rating <0) then rating = "-" end if
	 
%>
		<tr>
			<td rowspan="<%=rsNumPro("NumProcedures")%>"><span class="label"><%=rsReq("irName")%></span><br><%=rsReq("irdescription")%></td>
			
			<!--td rowspan="<%=rsNumPro("NumProcedures")%>"><% =rating %>
                <input type="hidden" name="rate_<%=rsReq("irID")%>" value="<% =rating %>" /> 
			</td-->
<%		

		ShowProcedures ReqID
	end function
	
	function ShowProcedures(ReqID)
		dim sqlPro, sqlNumOpt
		dim rsPro, rsNumOpt
		dim rowNum, checked
		
		sqlPro =	"SELECT IN_Procedures.*, prProcedure, prChecked, prResponsibilities, prTimeframe, prTextBox " & _
					"FROM IN_Procedures INNER JOIN AP_Procedures ON IN_Procedures.ipID = AP_Procedures.prProcedure " & _
					"WHERE ipRequirement = " & ReqID & " and prActionPlan = " & ActionPlan & " and ipActive = Yes order by ipDisplayOrder"
		set rsPro = con.Execute (sqlPro) 
		
		' The first row is treated differently so keep track of which row we are up to
		rowNum = 1

		while not rsPro.EOF
			if rowNum > 1 then
				Response.write "<tr>"
			end if	
			
			'retrieve the procedure info from the previous AP
			dim lastProcedure, prod, responsibility, timeframe, textbox
			prod = rsPro("prProcedure")
			if procedureData.Exists(prod) then
				set lastProcedure =  procedureData.Item(prod)
				responsibility = lastProcedure.responsibility
				timeframe = lastProcedure.timeframe
				textbox = lastProcedure.textbox

				if lastProcedure.checked then
					checked = " checked"
				else
					checked = ""
				end if
			else
				responsibility = ""
				checked = ""
                timeframe = ""
			end if
			
			'Over-ride this if we have come from a draft plan
			if rsPro("prChecked") <> False then
				checked = "checked"
			end if
			if rsPro("prChecked") <> True then
				checked = ""
			end if
			
			if rsPro("prResponsibilities") <> "" then
				responsibility = rsPro("prResponsibilities")
			end if
			
			if rsPro("prTimeframe") <> "" then 
				timeframe = rsPro("prTimeframe")
			end If
			
			if rsPro("prTextBox") <> "" then 'DLJ created the textbox option 17feb15
				textbox = rsPro("prTextBox")
			end If
	

			' Mandatory fields are shown in Red and non-mandatory fields are given a check box
			if rsPro("ipMandatory") = true then
				Response.Write "<td><font color='#FF0000'><B>M</B> " & rsPro("ipName") & "</font> " & rsPro("ipDescription")
			else
				Response.Write "<td><input type='checkbox' name='pro_" & rsPro("prProcedure") & "' " & checked & ">" & rsPro("ipName") & " " & rsPro("ipDescription")		
			end if
			
			'Save the procedure against the plan for future reference in case the procedure changes
			dim text 
			text = rsPro("ipName")&""
			text = Server.HTMLEncode(text)
			%>
				<input type="hidden" name="origProc_<%=rsPro("prProcedure")%>" value="<%=text%>"/>
			<%
		
			' If the procedure has a text box display the text box
			if rsPro("ipIsTextBox") = true then
				'Response.Write "<input type='text' name='proTxt_" & rsPro("prProcedure") & "' value='" & rsPro("prTextBox") & "'>"
				'Response.Write "<input type='text' name='proTxt_" & rsPro("prProcedure") & "' value='" & responsibility & "'>"  NOT RESPONSIBILITY, created textbox line below - DLJ 17feb15
				Response.Write "<input type='text' size='70' name='proTxt_" & rsPro("prProcedure") & "' value='" & textbox & "'>"
			end if

			sqlNumOpt = "SELECT count(*) as NumOptions " & _
						"FROM IN_Options " & _
						"WHERE ioProcedure = " & rsPro("ipID") & " AND ioActive = Yes"
			set rsNumOpt = con.Execute (sqlNumOpt)

			if rsNumOpt("NumOptions") > 0 then

				ShowOptions(rsPro("prProcedure"))
			end if
			
%>
				</td>
				<!--  table input boxes below made to fixed size by DLJ on 21Sept2004 also see line 186 -->
				<td><INPUT type="text" name="resp_<%=rsPro("ipID")%>" value="<%=responsibility%>" size=75><%'if rsPro("ipMandatory") = true then Response.Write " <font size=5 color=red><b>*</b></font>"%></td>
				<td><INPUT type="text" name="time_<%=rsPro("ipID")%>" value="<%=timeframe%>" size =15><%'if rsPro("ipMandatory") = true then Response.Write " <font size=5 color=red><b>*</b></font>"%></td>
			</tr>
<%		
			rowNum = rowNum + 1
			rsPro.movenext
		wend

	end function
	
	function ShowOptions(OptID)
		dim sqlOpt
		dim rsOpt
		
		sqlOpt =	"SELECT * FROM IN_Options INNER JOIN AP_Options ON IN_Options.ioID = AP_Options.aoOption " & _
					"WHERE ioProcedure = " & OptID & " and aoActionPlan = " & ActionPlan & " order by ioDisplayOrder" 
		set rsOpt = con.Execute (sqlOpt)
		
		while not rsOpt.EOF
%>
			<BR><input	type="<%=rsOpt("ioOptionType")%>" 
						name="<%=rsOpt("ioFieldName")%>" 
						value="<%=rsOpt("aoID")%>"
						<% if rsOpt("aoChecked")= true then Response.Write " checked"%>>
			<%=rsOpt("ioDescription")%> 
<%		
            timeframe = Empty
			'retrieve the procedure info from the previous AP
			dim lastProcedure, prod, timeframe
			prod = OptID
			if procedureData.Exists(req) then
				set lastProcedure =  procedureData.Item(prod)
				timeframe = lastProcedure.timeframe
			else
				timeframe = ""
			end if
			
			
			if rsOpt("ioTextBox") = true then
				'Response.Write "<input type='text' name='optTxt_" & rsOpt("ioID") & "' value='" & rsOpt("aoText") & "'"
				Response.Write "<input type='text' name='optTxt_" & rsOpt("ioID") & "' value='" & timeframe & "'"
			
			end if
			
			rsOpt.movenext
		wend
	end function
	
	function ShowAdditionalPoints(PointID)
		dim rsPoint, sqlPoint
					
		set rsPoint = server.CreateObject ("adodb.recordset")
		
		sqlPoint = "Select * from AP_Points where arSection = " & PointID & " and arActionPlan = " & ActionPlan			
		set rsPoint = con.Execute (sqlPoint)
					
		while not rsPoint.EOF
			Response.Write "<li>" & rsPoint("arText") & "</li>"
			rsPoint.MoveNext
		wend
	end function
%>

<% ShowStep(3) %>

<BR><BR>

<% ShowStep(2) %>

<BR><BR>

<% ShowStep(1) %>

<BR><BR>


		<table id = "planB" border="0">
<tr> 
	<th>STEP 4 - HEALTH AND SAFETY RESPONSIBILITIES</th>
</tr>

<!--

<tr> 
	<td> 
		To finalise the plan:<br>&nbsp;&nbsp;&nbsp;&nbsp;
		1. Review the list of safety responsibilities below, add any additional responsibilities and change the headings to reflect the Faculty/Unit structure.<br>&nbsp;&nbsp;&nbsp;&nbsp;
		2. Include any specific responsibilities that have been allocated in the Health and Safety Plan, as appropriate.
	</td>
</tr>

<tr>
	<td>
		<span class="label"><b>Staff must:</b></span>
		<ul>
			<li>take reasonable care of, and cooperate with actions taken to protect, the health and safety of both themselves and others</li>
			<li>follow safe work practices as provided by their supervisor, including the proper use of any personal protective equipment supplied</li>
			<li>seek information or advice from a supervisor before performing new or unfamiliar tasks</li>
			<li>report all health and safety accidents, incidents and hazards to their supervisor as soon as is practicable</li>
			<li>follow the emergency evacuation procedures</li>
			<li>support workplace injury management and return-to-work programs in their work areas.</li>
		</ul>

		<span class="label"><b>Other workers must:</b></span>
		<ul>
			<li>take reasonable care of, and cooperate with actions taken to protect, the health and safety of both themselves and others</li>
			<li>follow safe work practices as provided by their supervisor, including the proper use of any personal protective equipment supplied</li>
			<li>seek information or advice from a supervisor before performing new or unfamiliar tasks</li>
			<li>report all health and safety accidents, incidents and hazards to their supervisor as soon as is practicable</li>
			<li>follow the emergency evacuation procedures.</li>
		</ul>

		<span class="label"><b>Students must:</b></span>
		<ul>
			<li>take reasonable care of, and cooperate with actions taken to protect, the health and safety of both themselves and others</li>
			<li>follow safe work practices, including the proper use of any personal protective equipment supplied</li>
			<li>seek information or advice from a staff member before performing new or unfamiliar tasks</li>
			<li>report all health and safety accidents, incidents and hazards to a staff member as soon as is practicable</li>
			<li>follow the emergency evacuation procedures.</li>
		</ul>

		<span class="label"><b>Visitors to UTS must:</b></span>
		<ul>
			<li>take reasonable care of, and cooperate with actions taken to protect, the health and safety of both themselves and others</li>
			<li>report all health and safety accidents, incidents and hazards to a staff member as soon as is practicable</li>
			<li>follow the emergency evacuation procedures.</li>
		</ul>

	</td>
</tr>


-->

<!-- It would be good to be able to comment out the whole of Step 4 without breaking the Save as Final function. This section is not needed. -->

<tr>
	<td>

		<span class="label"><b>Supervisors and managers</b> must do whatever is reasonably practicable to ensure that both the workplace and the work itself are safe. This includes:</span>

		<ul>
			<li>ensuring that staff are appropriately trained and supervised</li>
			<li>identifying, assessing and managing health and safety risks</li>
			<li>consulting with workers (including staff, affiliates and contractors):
				<ul>
					<li>about issues or changes that affect their health or safety</li>
					<li>during health and safety risk assessments</li>
					<li>when decisions are made about the measures to be taken to eliminate or control these risks</li>
					<li>when reviewing health and safety risk assessments</li>
				</ul>
			</li>
			<li>implementing health and safety risk management programs relevant to their operations, teaching, research and consulting functions and work environment</li>
			<li>reporting (to the Human Resources Unit), investigating and responding to all hazards, accidents, incidents and taking action to control the risk</li>
			<li>assisting with the development, implementation and maintenance of a return to work program for injured staff.</li>


			<a name="section1">
<%
			ShowAdditionalPoints(1)
			
			if SecurityCheck(2) then 
%>
			<li>(Other)&nbsp;&nbsp;<input type="text" name="point_1" size="150">&nbsp;&nbsp;<input type="submit" value="Save Dot Point" onclick="javascript:formB.pointID.value=1;formB.pointText.value=formB.point_1.value;formB.action.value='point'"></li>
<%
			end if
%>
		</ul>
		<span class="label">In addition to the staff responsibilities above, <b>academic staff</b> will:</span>
		<ul>
			<li>provide relevant and practical health and safety information to students (through inclusion in curricula and course notes)</li>
			<li>take steps to ensure students adopt safe work practices</li>
			<li>conduct and document risk assessments on research and consulting programs/projects, and ensuring that risks are eliminated or controlled</li>
			<li>consult with staff who may be affected by health and safety risks during risk assessments, when decisions are made about the measures to be taken to eliminate or control these risks, and when these risk assessments are reviewed.</li>
						 
<%
			ShowAdditionalPoints(2)
			
			if SecurityCheck(2) then
%>
			<li>(Other)&nbsp;&nbsp;<input type="text" name="point_2" size="150">&nbsp;&nbsp;<input type="submit" value="Save Dot Point" onclick="javascript:formB.pointID.value=2;formB.pointText.value=formB.point_2.value;formB.action.value='point'"></li>
<%
			end if
%>
		</ul>





		<span class="label"><b>The <%=rs("apDDOption")%> is also responsible for:</b></span>
		<ul>
			<li>ensure that the Health and Safety Policy and related health and safety risk management programs are effectively implemented in their areas of control</li>
			<li>integrate health and safety risk management into their operations, teaching, research and consulting functions and work environments</li>
			<li>support supervisors and managers in providing appropriate resources for the effective implementation of their faculty/unit health and safety plan</li>
			<li>ensure that managers, supervisors and staff are aware of their responsibilities under the Health and Safety Policy and faculty/unit health and safety plan through effective delegation, training and promotion of the Policy and health and safety procedures</li>
			<li>hold supervisors and managers accountable for their specific responsibilities</li>
			<li>authorise appropriate action to remedy non-compliance with the Health and Safety Policy or health and safety procedures</li>
			<li>ensure that a faculty/unit health and safety plan is developed, implemented and monitored in consultation with staff</li>
			<li>conduct a self-assessment of their faculty or unit's compliance against their faculty/unit health and safety plan at regular intervals and report on progress to the Human Resources Unit.</li>
		</ul>

			<b>and, as part of their academic leadership role <input type="text" name="academic_heads" value="<%if rs("apHeadOfUnit") <> "" then Response.Write rs("apHeadOfUnit") else Response.Write ", heads of academic units"%>" size="25"> are also required to:</b>
		<ul>
			<li>ensure all staff undertake appropriate health and safety risk assessments for curriculum, research and consulting activities</li>
			<li>encourage the incorporation of health and safety risk management into curriculum and research.</li>

<%
			ShowAdditionalPoints(3)
			
			if SecurityCheck(2) then
%>
			<li>(Other)&nbsp;&nbsp;<input type="text" name="point_3" size="150">&nbsp;&nbsp;<input type="submit" value="Save Dot Point" onclick="javascript:formB.pointID.value=3;formB.pointText.value=formB.point_3.value;formB.action.value='point'"></li>
<%
			end if
%>
		</ul>
		<a name="section2">
		Are there any additional responsibilities allocated in the Plan? - <input type="radio" name="add_resp" value="1" <%if rs("apAddResp")=true then Response.Write " checked"%>> Yes <input type="radio" name="add_resp" value="0" <%if rs("apAddResp")=false then Response.Write " checked"%>> No <font color="#FF0000">&lt;Mandatory&gt;</font><br>       
		<BR>
		(Note: You can create new Section 4 headings here.). <BR>
		&nbsp;&nbsp;1 - enter the Position or Title of the person who is responsible and click the Save Additional Responsibility button. The new responsibility will appear.<BR>
		&nbsp;&nbsp;2 - type in the new dot point action and click the save button. The new action will appear under.
<%
		dim rsRespHead, sqlRespHead, rsResp, sqlResp
				
		set rsRespHead = server.CreateObject ("adodb.recordset")
		set rsResp = server.CreateObject ("adodb.recordset")
				
		sqlRespHead = "Select * from AP_ResponsibilityHeadings where rhActionPlan = " & ActionPlan
		set rsRespHead = con.Execute (sqlRespHead)
				
		while not rsRespHead.EOF
			Response.Write "<p><span class='label'>" & rsRespHead("rhTitle") & "</span></p>"
					
			sqlResp = "Select * from AP_Points where arSection = " & rsRespHead("rhID") & " and arActionPlan = " & ActionPlan
			set rsResp = con.Execute(sqlResp)
					
			Response.Write "<ul>"
					
			if not rsResp.BOF then

				while not rsResp.EOF
					Response.Write "<li>" & rsResp("arText") & "</li>"
						
					rsResp.MoveNext
				wend
		
			end if
					
			Response.Write "<li> (Other)&nbsp;&nbsp;<input type='text' name='point_" & rsRespHead("rhID") & "' size=150>&nbsp;&nbsp;<input type=submit value='Save Dot Point' onclick='javascript:formB.pointID.value=" & rsRespHead("rhID") & ";formB.pointText.value=formB.point_" & rsRespHead("rhID") & ".value;formB.action.value=""point""' id=submit1 name=submit1></li>"
			Response.Write "</ul>"
					
			rsRespHead.MoveNext
		wend
		
		if SecurityCheck(2) then
%> 
		<p><input type="text" name="heading" size="50" value="- Position/Title - are also responsible for:">&nbsp;&nbsp;
		   <input type="submit" value="Save Additional Responsibility" onclick="javascript:formB.action.value='addresp'"></p>
		<br><br>
		
<%
		end if 
		

	if SecurityCheck(2) then ' User must have write access for this department
%>
		 <p>Select either	<input type="submit" value="    Save as Draft    " onclick="javascript:formB.action.value='draft'" ID="Submit1" NAME="Submit1"> 
		or Check all compliance requirements are completed and 
					<input type="button" value="    Save as Final    " onclick="javascript:formB.action.value='final';DoSubmit()" ID="Button1" NAME="Button1"></p>
<%
	else
%>
		<p align="center"><input type="submit" value="    Close Window    " onclick="window.close();"></p>
<%
	end if
%>
	</td>
</tr>
</table>

</form>





<%
	Function ScriptStep(StepID)
		dim sqlStep, sqlReq
		dim rsStep, rsReq
		
		sqlStep = "Select stTableHeader from IN_Steps where stID = " & stepID
		set rsStep = con.Execute (sqlStep)
		
		' Show the requirements
		sqlReq =	"SELECT * FROM IN_Requirements INNER JOIN AP_Requirements ON IN_Requirements.irId = AP_Requirements.arRequirement " & _
					"WHERE irStep = " & StepID & " AND arActionPlan = " & ActionPlan & " and arSelected = Yes order by irDisplayOrder"
		set rsReq = con.Execute (sqlReq)
		
		while not rsReq.EOF
			ScriptRequirement (rsReq("irID"))
		
			rsReq.movenext
		wend
	end function
	


	function ScriptRequirement(ReqID)
		dim sqlReq, sqlPro
		dim rsReq, rsPro
		dim jsBuffer, reqName
		
		jsBuffer = ""
		
		sqlReq =	"SELECT  irName, ipID " & _
					"FROM IN_Requirements INNER JOIN IN_Procedures ON IN_Requirements.irId = IN_Procedures.ipRequirement " & _
					"WHERE irId = " & ReqID & " AND ipActive = Yes AND " & _
						"( SELECT Min(ipMandatory) AS Expr1 " & _	
						"FROM IN_Requirements INNER JOIN IN_Procedures ON IN_Requirements.irId = IN_Procedures.ipRequirement " & _
						"WHERE irId = " & ReqID & "  " & _
						"GROUP BY IN_Requirements.irId ) = 0"
						
						'****** COMMENT FOR THE INNER QUERY
						' This part just checks to see if any of the procedures are mandatory - if at least one procedure is mandatory there is no need to do this check
						' This implements the rule that you must have at least one procedure selected for each requirement
						'******
						
		set rsReq = con.Execute (sqlReq)
		
		' if there are no mandatory procedures make sure at least one procedure is selected
		if not rsReq.BOF then
			jsBuffer = " document.formB.pro_" & rsReq("ipID") & ".checked == false "
			reqName = rsReq("irName")
			
			rsReq.movenext
			
			while not rsReq.eof
				jsBuffer = jsBuffer & " &&  document.formB.pro_" & rsReq("ipID") & ".checked == false "
	
				rsReq.movenext
			wend
		end if
		
		if jsBuffer <> "" then
%>
		if ( <%=jsBuffer%> ) {
			message = message + " - You must select at least one procedure for the '<%=reqName%>' Requirement\n"
		}
<%
		end if 

		ScriptProcedures ReqID
	end function
	



	function ScriptProcedures(ReqID)
		dim sqlPro, sqlNumOpt
		dim rsPro, rsNumOpt
		
		sqlPro =	"SELECT IN_Procedures.*, prProcedure, prChecked, prResponsibilities, prTimeframe, prTextBox " & _
					"FROM IN_Procedures INNER JOIN AP_Procedures ON IN_Procedures.ipID = AP_Procedures.prProcedure " & _
					"WHERE IN_Procedures.ipRequirement = " & ReqID & " and prActionPlan = " & ActionPlan & " and ipActive = Yes order by prProcedure"
		set rsPro = con.Execute (sqlPro) 

		while not rsPro.EOF
			' Mandatory fields are shown in Red and non-mandatory fields are given a check box
			if rsPro("ipMandatory") = true then
%>
				if ( document.formB.resp_<%=rsPro("ipID")%>.value == "" ) 
					message = message + " - You must allocate responsibilities for the '<%=rsPro("ipName")%>' procedure\n"
					
				if ( document.formB.time_<%=rsPro("ipID")%>.value == "" ) 
					message = message + " - You must allocate a time frame for the '<%=rsPro("ipName")%>' procedure\n"
<%
			else
%>
				if ( document.formB.pro_<%=rsPro("prProcedure")%>.checked == true ) {
					if ( document.formB.resp_<%=rsPro("ipID")%>.value == "" ) 
						message = message + " - You must allocate responsibilities for the '<%=rsPro("ipName")%>' procedure\n"
						
					if ( document.formB.time_<%=rsPro("ipID")%>.value == "" ) 
						message = message + " - You must allocate a time frame for the '<%=rsPro("ipName")%>' procedure\n"
				}
<%
			end if
			
			sqlNumOpt = "SELECT count(*) as NumOptions " & _
						"FROM IN_Options " & _
						"WHERE ioProcedure = " & rsPro("ipID") & " AND ioActive = Yes"
			set rsNumOpt = con.Execute (sqlNumOpt)

			if rsNumOpt("NumOptions") > 0 then
				ScriptOptions(rsPro("prProcedure"))
			end if

			rsPro.movenext
		wend

	end function
	



	function ScriptOptions(OptID)
		dim sqlOpt
		dim rsOpt
		
		sqlOpt =	"SELECT ioJSValidation " & _
					"FROM IN_Options " & _
					"WHERE ioProcedure = " & OptID & " and ioJSValidation is not null" 
		set rsOpt = con.Execute (sqlOpt)
		
		response.write rsOpt("ioJSValidation")

	end function
%>




<script type="text/javascript">
<!--

	function DoSubmit() {
		var message
		
		message = ""
		
		// required checking that does not include the first three steps
		if ( document.formB.developedBy.value == "" ) 
			message = message + " - You must enter the name, title and position of the people present at the planning workshop\n"
		
// DLJ 1 May 2018 - what are the three functions below for? Check mandatory fields are selected?//
 <!-- % ScriptStep(1) % -->

 <!--% ScriptStep(2) % -->

 <!--% ScriptStep(3) % -->
		
	
		if (document.formB.academic_heads.value == "" || document.formB.academic_heads.value == "")
			message = message + " - You must enter the names of the Heads of Academic Units who are responsible\n"
			
		if (document.formB.add_resp[0].checked == false && document.formB.add_resp[1].checked == false)
			message = message + " - You must indicate if ther are any additional responsibilities for this Action Plan\n"
			
		if (message == "")
			document.formB.submit()
		else 
			alert("The following error(s) have been detected:\n\n" + message)
	}

//-->
</script>

<!-- #Include file="include\footer.asp" -->