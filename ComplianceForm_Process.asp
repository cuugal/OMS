<%@ Language=VBScript %>

<!-- #Include file="include\general.asp" -->

<%
	if SecurityCheck(2) = false then ' User must have write access for this department
		Response.Redirect ("restricted.asp")
		Response.end
	end if
%>

<%
	dim action, ActionPlan, Compliance, con, sqlComplete
	
	set con	= server.createobject ("adodb.connection")	con.open "DSN=ehs"

	action = request("action")
	ActionPlan = request("apID")
	Compliance = request("ccID")
	
	if Compliance = "" then
		dim sqlCompInsert
		dim rsIdentity
		
		set rsIdentity = server.CreateObject("adodb.recordset")
		
		rsIdentity.Open "cc_compliance", con, 1, 2

		rsIdentity.AddNew
		rsIdentity.fields("ccActionPlan").Value = ActionPlan
		rsIdentity.fields("ccComplete").Value = 0
		rsIdentity.Update

		Compliance = rsIdentity("ccID")
		rsIdentity.Close
		
		sqlCompInsert = "INSERT INTO CC_ComplianceDetails ( cdCompliance, cdRequirement ) " & _
						"SELECT " & Compliance & ", IN_Requirements.irID " & _
					    "FROM IN_Requirements INNER JOIN (AP_ActionPlans INNER JOIN AP_Requirements ON AP_ActionPlans.apID = AP_Requirements.arActionPlan) ON IN_Requirements.irId = AP_Requirements.arRequirement " & _
					    "WHERE apID = " & ActionPlan '& " AND arSelected = Yes"
		con.Execute(sqlCompInsert)
	end if

	select case action
		case "draft"	
			Save()
			
			CloseWindow()
		case "final"	
			Save()
			
			sqlComplete = "Update CC_Compliance set ccComplete = Yes where ccID = " & Compliance
			con.Execute (sqlComplete)
			
			CloseWindow()
	end select
	function Save()
		dim sqlComp, sqlUpdateComp		dim rsComp				sqlComp = "SELECT cdRequirement " & _
				  "FROM CC_ComplianceDetails " & _
				  "WHERE cdCompliance = " & Compliance		set rsComp = con.Execute (sqlComp)				while not rsComp.eof
			dim NewRate
			
			NewRate = request("txt_" & rsComp("cdRequirement"))						if NewRate = "" then
				NewRate = "null"			end if
					sqlUpdateComp = "UPDATE CC_ComplianceDetails " & _
						    "SET cdNewRating = " & NewRate & _
						    " WHERE cdRequirement = " & rsComp("cdRequirement") & " AND cdCompliance = " & Compliance
			'Response.Write sqlUpdateComp
			'Response.End			con.Execute(sqlUpdateComp)
			
			rsComp.movenext		wend
		
		'Update the following fields:
		sqlUpdateComp = "UPDATE CC_Compliance SET ccAssessor = '" & FilterSQL(request("txtAsses")) & "', ccDate = '" & FilterSQL(request("txtDate")) & _
					   "' WHERE ccID = " & Compliance
		con.Execute(sqlUpdateComp)	end function	
%>