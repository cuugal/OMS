<%@ Language=VBScript %>

<!-- #Include file="include\general.asp" -->

<%
	if SecurityCheck(3) = false then ' User must have write access for this department
		Response.Redirect ("restricted.asp")
		Response.end
	end if
%>

<%
	dim action, ActionPlan, ServiceAgreement, con, sqlComplete
	
	set con	= server.createobject ("adodb.connection")	con.open "DSN=ehs"
	
	action = request("action")
	ActionPlan = request("apID")	ServiceAgreement = request("saID")
	' If this is a new Service Agreement then create it first
	if ServiceAgreement = "" then
	
		dim sqlInsertSA
		dim rsIdentity
		
		set rsIdentity = server.CreateObject("adodb.recordset")
		
		rsIdentity.Open "SA_ServiceAgreement", con, 1, 2

		rsIdentity.AddNew
		rsIdentity.fields("saActionPlan").Value = ActionPlan
		rsIdentity.fields("saComplete").Value = 0
		rsIdentity.Update

		ServiceAgreement = rsIdentity("saID")
		rsIdentity.Close
		
		sqlInsertSA = "insert into SA_ServiceAgreementDetails (sdServiceAgreement, sdRequirement, sdEHSServices, sdContact, sdTimeFrame) " & _
					  "SELECT " & ServiceAgreement & ", irID, null, null, null " & _
					  "FROM IN_Requirements INNER JOIN AP_Requirements ON IN_Requirements.irId = AP_Requirements.arRequirement " & _
					  "WHERE arActionPlan = " & ActionPlan
		con.Execute (sqlInsertSA)
	end if
	

	select case action
		case "draft"	
			Save()
			
			CloseWindow()
		case "final"	
			Save()
			
			sqlComplete = "Update SA_ServiceAgreement set saComplete = Yes where saActionPlan = " & ActionPlan
			con.Execute (sqlComplete)
			
			CloseWindow()
	end select
	function Save()
		dim sqlSA, sqlUpdateSA		dim rsSA
		dim rating				sqlSA = "SELECT sdRequirement " & _
				"FROM SA_ServiceAgreement INNER JOIN SA_ServiceAgreementDetails ON SA_ServiceAgreement.saID = SA_ServiceAgreementDetails.sdServiceAgreement " & _
				"WHERE saID = " & ServiceAgreement		set rsSA = con.Execute (sqlSA)		
		while not rsSA.EOF			sqlUpdateSA = "UPDATE SA_ServiceAgreementDetails SET sdEHSServices = '" & FilterSQL(request("serv_" & rsSA("sdRequirement"))) & "', " & _ 						  "sdContact = '" & FilterSQL(request("cont_" & rsSA("sdRequirement"))) & "', " & _ 						  "sdTimeFrame = '" & FilterSQL(request("time_" & rsSA("sdRequirement"))) & "' " & _ 						  "WHERE sdServiceAgreement = " & ServiceAgreement & " AND sdRequirement = " & rsSA("sdRequirement")
			con.Execute (sqlUpdateSA)
			
			rsSA.movenext
		wend
		
		sqlUpdateSA = "Update SA_ServiceAgreement SET saAddEHSServices = '" & FilterSQL(request("serv_ADD")) & "', " & _
					  "saAddContact = '" & FilterSQL(request("cont_ADD")) & "', " & _
					  "saAddTimeFrame = '" & FilterSQL(request("time_ADD")) & "' " & _
					  "WHERE saActionPlan = " & ActionPlan
		con.Execute (sqlUpdateSA)	end function
%>
