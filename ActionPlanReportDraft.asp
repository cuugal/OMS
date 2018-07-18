<!-- #Include file="include\general.asp" -->
<!-- This is a print-friendly version of the draft EHS Plan available for the Admin user by clicking on the printer icon from the audit menu (AuditMenu.asp)- CL 02/07/2008 -->

<%
	if SecurityCheck(1) = false then ' User can access with read only
		Response.Redirect ("restricted.asp")
		Response.end
	end if
%>

<% PageTitle = "Action Plan Report"%>
	
<!-- #Include file="include\header.asp" -->
<%	dim con, ActionPlan
		dim sqlAp, sqlReq
		dim rsAp, rsReq
			set con = server.CreateObject("adodb.connection")
			con.Open "DSN=ehs"
			
				ActionPlan = Request("apID")
	sqlAP = "SELECT * FROM AP_ActionPlans INNER JOIN AD_Departments ON AP_ActionPlans.apFaculty = AD_Departments.dpID WHERE apID = " & ActionPlan
	set rsAP = con.Execute (sqlAP)
	
	sqlReq = "SELECT IN_Requirements.* " & _
			 "FROM IN_Requirements INNER JOIN AP_Requirements ON IN_Requirements.irId = AP_Requirements.arRequirement " & _
			 "WHERE arActionPlan = " & ActionPlan & " and irStep = 2 and arSelected = Yes order by irid"
			 set rsReq = con.Execute (sqlReq)
			 
		function ShowLists(ListType)
		dim rslist, sqlList
		
		sqlList =	"SELECT DISTINCTROW IN_Lists.lsPoints " & _
					"FROM AP_Requirements INNER JOIN (IN_Lists INNER JOIN (IN_MapProcLists INNER JOIN AP_Procedures " & _
					"ON IN_MapProcLists.mpINProcedures = AP_Procedures.prProcedure) " & _
					"ON IN_Lists.lsId = IN_MapProcLists.mpList) " & _
					"ON AP_Requirements.arID = AP_Procedures.prReq " & _
					"WHERE lsType = '" & ListType & "' AND prActionPlan = " & ActionPlan & " AND prChecked = Yes AND AP_Requirements.arSelected = Yes"
		
		set rsList = con.Execute(sqllist)
		
		' If threre are no options then 
		if rsList.EOF then
			Response.write "<li>(There are no options available)</li>"
		end if
			
		while not rsList.EOF
			Response.Write "<li>" & rsList("lsPoints") & "</li>"
			
			rslist.MoveNext
		wend
	end function 
%>

<img src="utslogo.gif" width="123" alt="The UTS home page" height="52" style="border:10px solid white" align="left">


	<table width="100%" border="0" cellspacing="3">
	<!--tr> 
		<td><!-- commented out the old EHS Branch logo <img src="ehslogo2.gif" width="142" height="111" alt="EHS logo" border="0">>&nbsp;</td>
		<td><div align="right"><img src="utslogo.gif" alt="UTS logo" width="123" height="52"></div></td>
	</tr-->
	<tr> 
		<td colspan="2"><div align="center"><font size="+3">*** DRAFT PLAN ONLY ***<br></font><h2><%=rsAP("dpName")%><br>Health and Safety Plan <%=rsAP("apStartYear")%> - <%=rsAP("apEndYear")%></h2></div></td>
	</tr>


	<tr> 
		<td>
		<p><b><%=rsAP("apDDOption")%>:&nbsp;<%=rsAP("apDDName")%></b></p>

		<!--p class="label">Date Plan Finalised: <%=rsAP("apCompletionDate")%>	&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;Valid until: <%=valid%></p-->
		</td>
	</tr>



	<!--tr> 
		<td><p><%=rsAP("apDDName")%><br><b><%=rsAP("apDDOption")%></b></p></td>
		<td><!--<p class="label">&nbsp;<br>Date Plan Drafted: <%=rsAP("apCompletionDate")%></p>></td>
	</tr-->


	<tr> 
		<td colspan="2">
		<p>The <%=rsAP("dpName")%> is committed to providing a safe and healthy workplace for students, staff, contractors and visitors. Promoting a safe and healthy workplace is the responsibility of all staff.  This is consistent with the University's Health and Safety Policy, the UTS Health and Safety Plan, the NSW Work Health and Safety Act and associated legislation and with legislation exercised by the NSW Environment Protection Authority.</p>

		<p>This Plan was developed by:<br><br><%=rsAp("apDevelopedBy")%></p>

		<p>This group:</p>
			<ul>
				<li>identified the hazards that may be encountered by staff, students, contractors and visitors,</li>
				<li>assessed the level of compliance of the faculty/unit with each of the compliance requirements,</li>
				<li>agreed on practical procedures to achieve and maintain compliance, and</li>
				<li>designated responsibilities and timeframes for implementation of these procedures.</li>
			</ul>
	</tr>
	<!--

			<p>This Plan includes the procedures, responsibilities and timeframes for:</p>
			<ul>
			<li>Health and Safety Management</li>
			<li>Health and Safety Procedures</li>
			<li>Specific Hazard Programs</li>
			<ul>
			<li>Specific Hazards
				<ul>
	<%
		if not rsReq.BOF then
			while not rsReq.EOF
				if rsReq("irid") <> 16 and rsReq("irid") <> 17 and rsReq("irid") <> 18 and rsReq("irid") <> 19 then 
					Response.Write "<li>" & rsReq("irName") & "</li>"
				end if
			
				rsReq.MoveNext
			wend
			
			rsReq.MoveFirst
	%>   
				</ul>
				</ul>
	<%
			while not rsReq.EOF
				if rsReq("irid") = 16 or rsReq("irid") = 17 or rsReq("irid") = 18 or rsReq("irid") = 19 then 
					Response.Write "<li>" & rsReq("irName") & "</li>"
				end if
			
				rsReq.MoveNext
			wend
		end if
	%>		
			</ul>
		</td>
		</tr>
		<tr>
		<td colspan="2" width="100%">

			<!table border="1" width="100%">
			<tr>
				<td width="100%">
					<table width="100%">
					<tr>
						<td width="40%"><br>
			  <b>Dean/Director:</b></td>
						<td><br>
			  <b>Signature:</b><br><br><br>
			  <b>Date: <<%=date()%>></b><br><br></td>
					</tr>
					</table>
				</td>
			</tr>


	-->


	</table>
	








<!-- Form Outline removed<tr>
	<td colspan=2>--><%
	function ShowFormBOutline()		dim sqlSteps		dim rsSteps	
		sqlSteps = "select * from IN_Steps"
		set rsSteps = con.Execute (sqlSteps)
		
		while not rsSteps.EOF
			Response.Write "<div class='label'>" & rsSteps("stShortName") & "</div>"						ShowReqOutline(rsSteps("stID"))
						rsSteps.movenext
		wend
	end function
	
	function ShowReqOutline(StepID)		dim sqlReq		dim rsReq				sqlReq =	"SELECT IN_Requirements.*, AP_Requirements.arID " & _
					"FROM IN_Requirements INNER JOIN AP_Requirements ON IN_Requirements.irId = AP_Requirements.arRequirement " & _
					"WHERE AP_Requirements.arActionPlan = " & ActionPlan & " AND IN_Requirements.irStep = " & StepID & " and arSelected = Yes"		set rsReq = con.Execute (sqlReq)		
		Response.Write "<ul>"
				while not rsReq.EOF			Response.Write "<li>" & rsReq("irName") & "</li>"						ShowProOutline(rsReq("cpID"))			
			rsReq.movenext		wend				Response.Write "</ul>"
	end function
	
	function ShowProOutline(ReqID)		dim sqlPro
		dim rsPro
		
		sqlPro =	"SELECT IN_Procedures.ipName, AP_Procedures.prTextBox, IN_Procedures.ipNumOptions, AP_Procedures.prID " & _
					"FROM IN_Procedures INNER JOIN AP_Procedures ON IN_Procedures.ipID = AP_Procedures.prProcedure " & _
					"WHERE AP_Procedures.prRequirement = " & ReqID & " and AP_Procedures.prChecked = Yes order by ipDisplayOrder"
		set rsPro = con.Execute (sqlPro)
				Response.Write "<ul>"		
		while not rsPro.EOF
			if rsPro("ipName") <> "" then				Response.Write "<li>" & rsPro("ipName") & "</li>"			else				Response.Write "<li>" & rsPro("prTextBox") & "</li>"			end if
						if rsPro("prID") > 0 then
				ShowOptOutline(rsPro("prID"))			end if
						rsPro.movenext
		wend			Response.Write "</ul>"
	end function			function ShowOptOutline(ProID)
		dim sqlOpt
		dim rsOpt
		
		sqlOpt =	"SELECT IN_Options.ioDescription, AP_Options.aoText, IN_Options.ioTextBox " & _
					"FROM IN_Options INNER JOIN AP_Options ON IN_Options.ioID = AP_Options.aoOption " & _
					"WHERE AP_Options.aoProcedure = " & ProID & " AND AP_Options.aoChecked = Yes order by ioDisplayOrder"
		set rsOpt = con.Execute (sqlOpt)
		
		if not rsOpt.BOF then			Response.Write "<ul>"
			
			while not rsOpt.EOF				if rsOpt("ioTextBox") = true then
					Response.Write "<li>" & rsOpt("aoText") & "</li>"
				else
					Response.Write "<li>" & rsOpt("ioDescription") & "</li>"
				end if				
				rsOpt.movenext			wend						Response.Write "</ul>"
		end if	end function
	
	'ShowFormBOutline()%>

					
					

<%	function ShowAdditionalPoints(PointID)
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
	



<%
	Function ShowStep(StepID)
		dim sqlStep, sqlReq
		dim rsStep, rsReq
		
		sqlStep = "Select stShortName, stName from IN_Steps where stID = " & stepID
		set rsStep = con.Execute (sqlStep)
%>


<!--div class="page-break"></div-->

		<table id = "plan" cellpadding="2" width="100%">
			<tr> 
			  <th colspan="5" class="StepMenu"><%=rsStep("stName")%></th>
			</tr>

			<tr> 
			  <td class="label" width ="20%">COMPLIANCE REQUIREMENT</td>
			<!--note that the date in the following line is month and year NOW. It should be month and
			date report is saved i.e <%=rsAP("apCompletionDate")%>
		DLJ made this change 10 August 2004
			-->
			  <!--td><span class="label">Compliance at <%=rsAP("apCompletionDate")%> <!--<%=monthname(month(Date()),true)%>&nbsp;<%=year(Date())%> >*</span></td-->
			  <td class="label" width ="40%">Health and Safety Procedures</td> <!-- dlj took out open span tag -->
			  <td class="label" width = "20%">Responsibilities</td>
			  <td class="label" width = "10%">Complete by</td>
			</tr>
<%
		'Response.Write rsStep("stBasicHeader")
		
		' Show the requirements
		sqlReq =	"SELECT * FROM IN_Requirements INNER JOIN AP_Requirements ON IN_Requirements.irId = AP_Requirements.arRequirement " & _
					"WHERE IN_Requirements.irStep = " & StepID & " AND AP_Requirements.arActionPlan = " & ActionPlan & " and arSelected = Yes order by irDisplayOrder"
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
		dim sqlReq
		dim rsReq
		
		sqlReq =	"SELECT IN_Requirements.*, AP_Requirements.arRating " & _
					"FROM IN_Requirements INNER JOIN AP_Requirements ON IN_Requirements.irId = AP_Requirements.arRequirement " & _
					"WHERE AP_Requirements.arActionPlan = " & ActionPlan & " AND IN_Requirements.irId = " & ReqID
		set rsReq = con.Execute (sqlReq)
		
		sqlCount =	"SELECT Count(*) AS Expr1 " & _
					"FROM AP_Requirements INNER JOIN AP_Procedures ON AP_Requirements.arID = AP_Procedures.prReq " & _
					"WHERE AP_Procedures.prChecked=Yes AND AP_Requirements.arRequirement = " & ReqID & " AND AP_Requirements.arActionPlan = " & ActionPlan
		set rsCount = con.Execute (sqlCount)
%>
		<tr>
			<td rowspan="<%=rsCount("Expr1")%>"><span class="label"><%=rsReq("irName")%></span><br><%=rsReq("irdescription")%></td>
			<!--td rowspan="<%=rsCount("Expr1")%>"><%=rsReq("arRating")%></td-->
		
<%		

		ShowProcedures ReqID
	end function
	
	function ShowProcedures(ReqID)
		dim sqlPro, sqlNumOpt
		dim rsPro, rsNumOpt
		dim rowNum, checked
		
		sqlPro =	"SELECT IN_Procedures.*, AP_Procedures.prID, AP_Procedures.prChecked, AP_Procedures.prResponsibilities, AP_Procedures.prTimeframe, AP_Procedures.prTextBox " & _
					"FROM IN_Procedures INNER JOIN AP_Procedures ON IN_Procedures.ipID = AP_Procedures.prProcedure " & _
					"WHERE IN_Procedures.ipRequirement = " & ReqID & " and prActionPlan = " & ActionPlan & " and prChecked = Yes order by ipDisplayOrder"
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

			Response.Write "<td>" & rsPro("ipName")
			
			' If the procedure has a text box display the text box
			if rsPro("ipIsTextBox") = true then
				Response.Write rsPro("prTextBox")
			end if

			sqlNumOpt = "SELECT count(*) as NumOptions " & _
						"FROM IN_Options INNER JOIN (AP_Procedures INNER JOIN AP_Options ON AP_Procedures.prID = AP_Options.aoPro) ON IN_Options.ioID = AP_Options.aoOption " & _
						"WHERE prActionPlan = " & ActionPlan & " AND prProcedure = " & rsPro("ipID") & " AND ioActive = Yes"
			set rsNumOpt = con.Execute (sqlNumOpt)

			if rsNumOpt("NumOptions") > 0 then
				ShowOptions(rsPro("prID"))
			end if
%>
				&nbsp;</td>
				<td><%=rsPro("prResponsibilities")%>&nbsp;</td>
				<td><%=rsPro("prTimeframe")%>&nbsp;</td>
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
					"WHERE AP_Options.aoPro = " & OptID & " and aoChecked = Yes order by ioDisplayOrder"
		set rsOpt = con.Execute (sqlOpt)
		
		while not rsOpt.EOF
%>
			<br>
			- <%=rsOpt("ioDescription")%> 
<%		
			if rsOpt("ioTextBox") = true then
				Response.Write "- " & rsOpt("aoText")
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

<br><br>

<% ShowStep(2) %>

<br><br>

<% ShowStep(1) %>

<br><br>

<hr>


		<table width="100%" border="0">
		<tr>
			<td class="label">HEALTH AND SAFETY RESPONSIBILITIES<br><br></td>
		</tr>
		<!-- DLJ 14Jun2017 deleted responsibilties list -->
		<tr>
			<td>Refer to the UTS <i>Health and Safety Responsibilities Vice-Chancellor's Directive</i> on the Health and Safety website for detailed responsibilities of discrete groups within the University community.</td>
			
		</tr>
		<tr><td>&nbsp;</td>
		</tr>
		</table>
<hr>



<table width="100%" border="0">
<tr> 
	<td class="label">RESOURCE LIST</td>
</tr>


<tr>
  <td colspan="2">
		<!--
		XXXXXXXXXXXXXXXXXXXXXXXX
		DLJ EDIT - temporarily commented out until list generation function is repaired 5/5/4
		
		Stephen - Readded this section and tested  to be working correctly according to new specification
		-->
		<p class="label">The following records and documentation are available to prove compliance</p>
		<ul>
<%
		ShowLists("rec")
%>    
		</ul>
		<p class="label">The following signage and posters will be displayed</p>
		<ul>
<%
		ShowLists("sig")
%> 
		</ul>
		<p class="label">Information/training sessions will be conducted</p>
		<ul>
<%
		ShowLists("tra")
%>
		</ul><p class="label">The following is a checklist of procedures to meet compliance in high-risk facilities and work areas</p>
		<ul>
<%
		ShowLists("ccl")
%>

		</ul>
		</td>
	</tr>


















<%
	dim sqlResource, rsResource
	
	sqlResource =	"SELECT * FROM IN_Requirements INNER JOIN AP_Requirements ON IN_Requirements.irId = AP_Requirements.arRequirement " & _
					"where arActionPlan = " & ActionPlan & " and arSelected = Yes order by irid"
	set rsResource = con.Execute (sqlResource)
%>
<!--tr> 
	<td>
		<ul>
<%
		while not rsResource.EOF
			Response.Write "<br><li><b>" & rsResource("irName") & "</b></li>"
			Response.Write "<ul>" & rsResource("irResourceList") & "</ul>"
			
	
			rsResource.MoveNext
		wend
%>
	</ul></td>
</tr-->



</table>
<!-- #Include file="include\footer.asp" -->