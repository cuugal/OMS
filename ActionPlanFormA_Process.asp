<%@ Language=VBScript %>

<!-- #Include file="include\general.asp" -->

<%
	if SecurityCheck(2) = false then ' User must have write access for this department
		Response.Redirect ("restricted.asp")
		Response.end
	end if
%>

<%
	dim action, ActionPlan, con, sqlVersion
	
	set con	= server.createobject ("adodb.connection")	con.open "DSN=ehs"
	
	action = request("action")
	ActionPlan = request("apID")

	if ActionPlan = "" then ' This is a new Action Plan so create a new one

		dim sqlGetDuration, sqlInsertReq, sqlInsertPro, sqlInsertOpt
		dim rsGetDuration, rsIdentity
		
		set rsGetDuration = server.CreateObject("adodb.recordset")
		set rsIdentity = server.CreateObject("adodb.recordset")

		' Get the duration of this action plan (depends on department)		sqlGetDuration = "Select dpActionPlanDuration from AD_Departments where dpID = " & Request("Department")
		set rsGetDuration = con.Execute (sqlGetDuration)		
		' Make sure that we do not allow duplicate Action Plans		on error resume next		
		rsIdentity.Open "AP_ActionPlans", con, 1, 2

		rsIdentity.AddNew
		rsIdentity.fields("apFaculty").Value		= Request("Department")
		rsIdentity.fields("apDDName").Value			= FilterSQL(Request("apDeanDirector"))
		rsIdentity.fields("apCompletionDate").Value = date()
		rsIdentity.fields("apCompleted").Value		= "0"
		rsIdentity.fields("apStartYear").Value		= year(date())
		rsIdentity.fields("apEndYear").Value		= year(date()) + rsGetDuration("dpActionPlanDuration")
		rsIdentity.fields("apDDoption").Value		= FilterSQL(Request("apDeanOrDirector"))
		rsIdentity.fields("apHeadOfUnit").Value		= ""
		rsIdentity.fields("apVersion").Value		= "1"
		rsIdentity.Update

		ActionPlan = rsIdentity("apID")
		rsIdentity.Close
		set rsIdentity = nothing		
		' If there is an error inserting an Action Plan stop execution and call the Error Handler		If Err.Number <> 0 Then
			ErrorHandle()
		end if
		
		sqlInsertReq = "Insert into AP_Requirements (arActionPlan, arRequirement, arRating, arSelected) " &_
					   "SELECT " & ActionPlan & ", irId, 0, no " & _
					   "FROM IN_Requirements"
		con.Execute (sqlInsertReq)
		
		sqlInsertPro = "Insert into AP_Procedures (prReq, prActionPlan, prProcedure, prChecked) " &_
					   "SELECT arID, arActionPlan, ipID, No " & _
					   "FROM IN_Procedures INNER JOIN AP_Requirements ON IN_Procedures.ipRequirement = AP_Requirements.arRequirement " & _
					   "WHERE arActionPlan = " & ActionPlan
		con.Execute (sqlInsertPro)
		
		sqlInsertOpt = "Insert into AP_Options (aoPro, aoActionPlan, aoOption, aoText, aoChecked) " &_
					   "SELECT prID, prActionPlan, ioId, '', no " & _
					   "FROM IN_Options INNER JOIN AP_Procedures ON IN_Options.ioProcedure = AP_Procedures.prProcedure " & _
					   "WHERE prActionPlan = " & ActionPlan
		con.Execute (sqlInsertOpt)
	end if

	select case action
		case "draft"	
			Save()
			
			CloseWindow()
		case "final"	
			Save()
			
			sqlVersion = "Update AP_ActionPlans set apFormACompleted = Yes where apID = " & ActionPlan
			con.Execute (sqlVersion)
			
			Response.Redirect "ActionPlanFormB.asp?apID=" & ActionPlan
			Response.End
	end select

	function Save()
		dim sqlReq, sqlUpdateReq, sqlUpdateAP
		dim rsReq
		dim rating
		
		sqlReq = "select irid, irmandatory, irActive from IN_Requirements"
		set rsReq = con.Execute (sqlReq)
		
		while not rsReq.eof
				if request("rate_" & rsReq("irid")) = "" then 
						rating = "null"
					else 
						rating = request("rate_" & rsReq("irid"))
				end if 
				if rsReq("irMandatory") = true and rsReq("irActive") = true then 
						sqlUpdateReq = "update AP_Requirements set arRating = " & rating & ", arSelected = Yes where arActionPlan = " & ActionPlan & " and arRequirement = " & rsReq("irid")
						con.Execute (sqlUpdateReq)
					else
						if request("req_" & rsReq("irid")) = "on" then
							sqlUpdateReq = "update AP_Requirements set arRating = " & rating & ", arSelected = Yes where arActionPlan = " & ActionPlan & " and arRequirement = " & rsReq("irid")
							con.Execute (sqlUpdateReq)
						else 
							sqlUpdateReq = "update AP_Requirements set arRating = " & rating & ", arSelected = No where arActionPlan = " & ActionPlan & " and arRequirement = " & rsReq("irid")
							con.Execute (sqlUpdateReq)
						end if
				end if
			rsReq.movenext
		wend
					
		sqlUpdateAP = "update AP_ActionPlans set apDDName = '" & FilterSQL(Request("apDeanDirector")) & "', apDDOption = '" & FilterSQL(Request("apDeanOrDirector")) & "' where apID =" & ActionPlan
		'DLJ made change to above line 9Mar4 to ad where apID =" & ActionPlan
		con.Execute (sqlUpdateAP)
	end function
						
	function ErrorHandle()			
		%>
		<script type="text/javascript">
		alert ("An EHS Plan was already found for the current year. The most common cause of this is that the EHS Plan was not refreshed properly.\n\nYour EHS Plan could not be saved please return to the EHS Plan Menu and refresh the page before trying again.")
		</script>
		<%
		CloseWindow()			
	end function
%>
