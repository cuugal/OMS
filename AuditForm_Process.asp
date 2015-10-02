<%@ Language=VBScript %>

<!-- #Include file="include\general.asp" -->

<%
	if SecurityCheck(2) = false then ' User must have write access for this department
		Response.Redirect ("restricted.asp")
		Response.end
	end if
%>

<%

    




	dim con, Action, ActionPlan, Mode, AuditID
	dim sqlVersion, sqlLab, sqlAuditInsert, sqlDetailInsert
	dim rsLab, rsIdentity, audit
	
	set con	= server.createobject ("adodb.connection")	con.open "DSN=ehs"
	
	Mode = Request("Mode")				' Determines if this is a new AuditForm or an existing one
	AuditID = Request("faID")			' The Facility Audit ID
	ActionPlan = Request("apID")		' The Action Plan ID
	Action = Request("action")			' Save as Draft or as Final
	audit = request("audittype")
	
	' Check for duplicate LabNames for the same ActionPlan
	if Mode = "New" then
		' this is a new Audit so make sure the Lab Name is unique to this Action Plan
		if LabNameInUse = true then
			Response.Redirect "AuditForm.asp?apID=" & ActionPlan & "&Mode=" & Mode & "&faID=" & AuditID & "&error=Lab" &"&type="&audit
			Response.End
		end if
		
		' Insert the basic requirements into a new Facility Audit
		set rsIdentity = server.CreateObject("adodb.recordset")
		
		rsIdentity.Open "FA_Audits", con, 1, 2

		rsIdentity.AddNew
		rsIdentity.fields("faActionPlan").Value = ActionPlan
		rsIdentity.fields("faComplete").Value = 0
		rsIdentity.Update

		AuditID = rsIdentity("faID")
		rsIdentity.Close
		
		' We put in all of the Requirements even those that are not selected (also inactive ones that may still be selected)
		sqlDetailInsert = "INSERT INTO FA_AuditDetails ( fdAudit, fdRequirement, fdRating ) " & _
						  "SELECT " & AuditID & ", IN_Requirements.irId, 0 " & _
						  "FROM IN_Requirements INNER JOIN AP_Requirements ON IN_Requirements.irId = AP_Requirements.arRequirement " & _
						  "WHERE arActionPlan = " & ActionPlan
		con.Execute(sqlDetailInsert)
	else
		' This is an existing audit so if Lab Name is different then make sure it is unique to this Action Plan
		sqlLab = "select faLabName from FA_Audits where faID = " & AuditID
		set rsLab = con.Execute(sqlLab)
		
		if request("txt_lab") <> rsLab("faLabName") then
		
			if LabNameInUse = true then
				Response.Redirect "AuditForm.asp?apID=" & ActionPlan & "&Mode=" & Mode & "&faID=" & AuditID & "&error=Lab"
				Response.End
			end if
		end if
	end if

	select case action
		case "draft"	
			Save()
			
			CloseWindow()
		case "final"	
			Save()
			
			sqlVersion = "Update FA_Audits set faComplete = Yes where faID = " & AuditID
			con.Execute (sqlVersion)
			
			CloseWindow()
	end select
	
	function LabNameInUse()
		dim sqlLab
		dim rsLab
		
		sqlLab = "Select count(*) as Found from FA_Audits where faActionPlan = " & ActionPlan & " and faLabName = '" & FilterSQL(request("txt_lab")) & "'"
		set rsLab = con.Execute(sqlLab)
		
		if rsLab("Found") > 0 then 
			LabNameInUse = true
		else
			LabNameInUse = false
		end if
	end function
	function Save()
		dim sqlAud, sqlUpdateAud		dim rsAud				sqlAud = "SELECT fdRequirement " & _
				 "FROM FA_AuditDetails INNER JOIN AP_Requirements ON FA_AuditDetails.fdRequirement = AP_Requirements.arRequirement " & _
				 "WHERE fdAudit = " & AuditID & " and arSelected = Yes and arActionPlan = " & ActionPlan		set rsAud = con.Execute (sqlAud)				while not rsAud.eof
			if request("rate_" & rsAud("fdRequirement")) <> "" then				rating = request("rate_" & rsAud("fdRequirement"))			else				rating = 0			end if
		
			'Response.Write "Normal: " & request("text_" & rsAud("fdRequirement")) & vbcrlf			'Response.Write "Changed: " & FilterTrailingCrLf(FilterSQL(request("text_" & rsAud("fdRequirement")))) & vbcrlf
			            'VBScript concatenates by comma when sending fields with same name, not trivial to change this behaviour.            'Obviously we will have commas in our text fields, need to split on a different character than ','            dim tmp            tmp = ""                        for i=0 to request.form("text_" & rsAud("fdRequirement")).count-1
                'split by pipe
                if(i>0) then
                    tmp = tmp & "|"
                end if
                tmp = tmp & request.form("text_" & rsAud("fdRequirement"))(i+1)
            next    '"SET		fdEvidence	= '" & FilterTrailingCrLf(FilterSQL(request("text_" & rsAud("fdRequirement"))))	& "', " & _
	                       			sqlUpdateAud = "UPDATE	FA_AuditDetails " & _
						    "SET		fdEvidence	= '" & FilterTrailingCrLf(FilterSQL(tmp))	& "', " & _					   
                            "		fdRating	= " & rating		& " " & _ 
						   "WHERE	fdRequirement = " & rsAud("fdRequirement") & " AND fdAudit = " & AuditID			con.Execute(sqlUpdateAud)						'Response.Write sqlUpdateAud			'Response.End
			
			rsAud.movenext		wend
		
		'Update the following fields:
		dim audittype, auditUpper
		audittype = request("audittype")
		auditUpper =  UCase(Left(audittype,1))& Mid(audittype,2) 
		sqlUpdateAud = "UPDATE FA_Audits " & _
					   "SET		faSupervisor	= '" & FilterSQL(request("txt_Sup"))	& "', " & _
					   "		faLabName		= '" & FilterSQL(request("txt_lab"))	& "', " & _
					   "		faAssesName		= '" & FilterSQL(request("txt_Assr"))	& "', " & _
					   "		faLocation		= '" & FilterSQL(request("txt_Loc"))	& "', " & _
					   "		faDate			= '" & FilterSQL(request("txt_Date"))	& "', " & _
					   "		faHouseKeeping	= '" & FilterSQL(request("txt_hous"))	& "', " & _
					   "		faAuditType		= '" & FilterSQL(auditupper)	& "' " & _
					   "WHERE	faID = " & AuditID
		con.Execute(sqlUpdateAud)	end function
%>