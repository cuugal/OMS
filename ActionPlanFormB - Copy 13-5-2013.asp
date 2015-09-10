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
	<td><div align="right"><img src="utslogo.gif" width="135" height="30"></div></td>
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
		<i>Compliance Ratings: 0 = Non-Compliant, 1 = Non-Compliant - some action evident but not fully compliant, 2 = Compliant - just requires maintenance, 3 = Best practice evident</i>
		<table border="1">
		<tr> 
			<td colspan="5" class="StepMenu" bgcolor="#cccccc"><font size="3pt"><b> <%=rsStep("stName")%> </b></font></td>
		</tr>
		<tr> 
			<td class="label">COMPLIANCE REQUIREMENT</td>
			<td>
					<!--note that the date in the following line is month and year NOW. <%=monthname(month(Date()),true)%>&nbsp;<%=year(Date())%>
					It should be month and
					date report is saved i.e =rs("apCompletionDate")
					DLJ made this change 10 August 2004
					-->
				<span class="label">1. Compliance rating<br>
		            (0, 1, 2, or 3)</span> 
		    </td>
			<td>
				<span class="label">2. Note what is required to comply</span><br>
		        &nbsp;&nbsp;<font color="#FF0000">M - mandatory procedures to comply 
		        with legislative requirements</font><br>
		        <input type="checkbox" disabled>
		        - optional activities that might be undertaken to achieve compliance
		    </td>
			<td class="label" width="200">3. Allocate responsibilities</td>
			<td class="label" width="80">4. Allocate timeframe to complete by</td>
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
%>
		<tr>
			<td rowspan="<%=rsNumPro("NumProcedures")%>"><span class="label"><%=rsReq("irName")%></span><br><%=rsReq("irdescription")%></td>
			<td rowspan="<%=rsNumPro("NumProcedures")%>">
				<SELECT name="rate_<%=rsReq("irID")%>">					<OPTION value=0 <%if rsReq("arRating") = 0 then Response.Write " selected"%>>0</OPTION>
					<OPTION value=1 <%if rsReq("arRating") = 1 then Response.Write " selected"%>>1</OPTION>
					<OPTION value=2 <%if rsReq("arRating") = 2 then Response.Write " selected"%>>2</OPTION>
					<OPTION value=3 <%if rsReq("arRating") = 3 then Response.Write " selected"%>>3</OPTION>
				</SELECT>
			</td>
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
			
			' Determine if the option is checked
			if rsPro("prChecked") = true then
				checked = " checked"
			else
				checked = ""
			end if	

			' Mandatory fields are shown in Red and non-mandatory fields are given a check box
			if rsPro("ipMandatory") = true then
				Response.Write "<td><font color='#FF0000'><B>M</B> " & rsPro("ipName") & "</font> " & rsPro("ipDescription")
			else
				Response.Write "<td><input type='checkbox' name='pro_" & rsPro("prProcedure") & "' " & checked & ">" & rsPro("ipName") & " " & rsPro("ipDescription")
			end if
			
			' If the procedure has a text box display the text box
			if rsPro("ipIsTextBox") = true then
				Response.Write "<input type='text' name='proTxt_" & rsPro("prProcedure") & "' value='" & rsPro("prTextBox") & "'>"
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
				<td><INPUT type="text" name="resp_<%=rsPro("ipID")%>" value="<%=rsPro("prResponsibilities")%>" size=50><%'if rsPro("ipMandatory") = true then Response.Write " <font size=5 color=red><b>*</b></font>"%></td>
				<td><INPUT type="text" name="time_<%=rsPro("ipID")%>" value="<%=rsPro("prTimeframe")%>" size =15><%'if rsPro("ipMandatory") = true then Response.Write " <font size=5 color=red><b>*</b></font>"%></td>
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
			if rsOpt("ioTextBox") = true then
				Response.Write "<input type='text' name='optTxt_" & rsOpt("ioID") & "' value='" & rsOpt("aoText") & "'"
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

<% ShowStep(1) %>

<BR><BR>

<% ShowStep(2) %>

<BR><BR>

<% ShowStep(3) %>

<BR><BR>

<table border="0">
<tr> 
	<td class="StepMenu" bgcolor='#CCCCCC'><font size=3pt><b> STEP 4 - HEALTH AND SAFETY RESPONSIBILITIES</b></font></td>
</tr>
<tr> 
	<td> 
		To finalise the plan:<br>&nbsp;&nbsp;&nbsp;&nbsp;
		1. Review the list of safety responsibilities below, add any additional responsibilities and change the headings to reflect the Faculty/Unit structure.<br>&nbsp;&nbsp;&nbsp;&nbsp;
		2. Include any specific responsibilities that have been allocated in the Health and Safety Plan, as appropriate.
	</td>
</tr>
<tr>
	<td>
		<span class="label"><b>All staff, students, contractors and visitors are responsible for:</b></span>
		<ul>
			<li>looking out for hazards, reporting them to the supervisor of the work area and helping to fix hazards</li>
            <li>taking action to avoid, eliminate or minimise risks</li>
            <li>following safe work methods and using personal protective equipment as required</li>
            <li>seeking information or advice as necessary - particularly before carrying out new or unfamiliar work</li>
            <li>reporting accidents and incidents to the supervisor of the work area</li>
           <li>contacting Security (by dialling 6 from internal telephones, or 1800 249 559) to report emergencies that occur on campus</li>
			<li>safely disposing of any hazardous waste produced</li>
            <li>not wilfully placing at risk the health, safety or wellbeing of others.</li>
        </ul>
        <span class="label"><b>Supervisors and managers are also responsible for:</b></span>
		<ul>
			<li>establishing safe work methods</li>
			<li>setting up practical procedures to find and fix hazards</li>
			<li>taking action to fix hazards</li>
			<li>ensuring that management and staff are aware of their responsibilities under the Health and Safety Policy, through effective delegation, training and promotion of the Policy and health and safety procedures</li>
			<li>authorising appropriate action to remedy non-compliance with the Health and Safety Policy or health and safety procedures<br><br>
			
			and, <b>in the case of an accident:</b></li>

			<li>ensuring that the person involved receives first aid and completes an incident report using the HIRO system</li>
			<li>calling the Safety &amp; Wellbeing Branch immediately if the accident has resulted in a serious injury or if there is the risk of a serious injury or illness, such as in the event of a significant chemical spill; damage to equipment; faulty equipment; needlestick injury</li>
			<li>investigating the accident as soon as possible after it occurs (and no later than 48 hours after the accident)</li>
			<li>taking action to prevent a recurrence of the accident<br><br>
			
			and, <b>if hazardous waste is produced:</b><br></li>
			
			<li>ensuring that all hazardous waste is labelled and safely disposed (as per the University's hazardous waste procedures).<br><br></li>
            <a name="section1">
<%
			ShowAdditionalPoints(1)
			
			if SecurityCheck(2) then 
%>
            <li>(Other)&nbsp;&nbsp;<input type="text" name="point_1" size="50">&nbsp;&nbsp;<input type="submit" value="Save Dot Point" onclick="javascript:formB.pointID.value=1;formB.pointText.value=formB.point_1.value;formB.action.value='point'"></li>
<%
			end if
%>
		</ul>
        <span class="label"><b>Academics are also responsible for:</b></span></p>
        <ul>
			<li>providing relevant and practical health and safety information to students</li>
            <li>taking steps to ensure students adopt safe work practices.</li>     
<%
			ShowAdditionalPoints(2)
			
			if SecurityCheck(2) then
%>
			<li>(Other)&nbsp;&nbsp;<input type="text" name="point_2" size="50">&nbsp;&nbsp;<input type="submit" value="Save Dot Point" onclick="javascript:formB.pointID.value=2;formB.pointText.value=formB.point_2.value;formB.action.value='point'"></li>
<%
			end if
%>
		</ul>
        <span class="label"><b>The <%=rs("apDDOption")%> is also responsible for:</b></span>
		<ul>
            <li>developing and maintaining a Faculty/Unit Health and Safety Plan</li>
            <li>ensuring that the Health and Safety Policy and hazard management programs are effectively implemented in their areas of control</li>
            <li>supporting supervisors and holding them accountable for their specific responsibilities</li>
            <li>providing appropriate resources for the effective implementation of this Health and Safety Plan</li>
            <li>ensuring that management and staff are aware of their responsibilities under the Health and Safety Policy through effective delegation, training and promotion of the Health and Safety Policy and procedures<br></li>
            <li>authorising appropriate action to remedy non-compliance with the Health and Safety Policy or with safety procedures<BR></li>
        </ul>
            <b>and, as part of their academic leadership role <input type="text" name="academic_heads" value="<%if rs("apHeadOfUnit") <> "" then Response.Write rs("apHeadOfUnit") else Response.Write ", heads of academic units"%>" size="25"> are also required to:</b>
        <ul>
            <li>ensure appropriate safety issues are included in curriculum and research projects</li>
            <li>encourage the incorporation of health and safety risk management into curriculum and research.</li>
<%
			ShowAdditionalPoints(3)
			
			if SecurityCheck(2) then
%>
			<li>(Other)&nbsp;&nbsp;<input type="text" name="point_3" size="50">&nbsp;&nbsp;<input type="submit" value="Save Dot Point" onclick="javascript:formB.pointID.value=3;formB.pointText.value=formB.point_3.value;formB.action.value='point'"></li>
<%
			end if
%>
		</ul>
		<a name="section2">
        Are there any additional responsibilities allocated in the Plan? - <input type="radio" name="add_resp" value="1" <%if rs("apAddResp")=true then Response.Write " checked"%>> Yes <input type="radio" name="add_resp" value="0" <%if rs("apAddResp")=false then Response.Write " checked"%>> No <font color="#FF0000">&lt;Mandatory&gt;</font><br>       
        <BR>
        ( Note: You can create new Section 4 headings here and then create new dot points under these headings. <BR>
        &nbsp;&nbsp;-First enter the Position or Title of the person who is responsible and click the save button. <BR>
        &nbsp;&nbsp;-Second the new responsibility will appear an then you can add new dot points. <BR>
        &nbsp;&nbsp;-Third type in the new dot point and click the save button and the dot point will appear under the new responsibility heading )
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
					
			Response.Write "<li> (Other)&nbsp;&nbsp;<input type='text' name='point_" & rsRespHead("rhID") & "' size=50>&nbsp;&nbsp;<input type=submit value='Save Dot Point' onclick='javascript:formB.pointID.value=" & rsRespHead("rhID") & ";formB.pointText.value=formB.point_" & rsRespHead("rhID") & ".value;formB.action.value=""point""' id=submit1 name=submit1></li>"
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
		

<% ScriptStep(1) %>

<% ScriptStep(2) %>

<% ScriptStep(3) %>
		
	
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