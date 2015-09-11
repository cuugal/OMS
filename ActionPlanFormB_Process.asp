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
	
	con.CommandTimeout = 1000
	
	action = request("action")
	ActionPlan = request("apID")

	select case action
		case "draft"	
			ShowProcessingMessage()
			SaveDraft()
			
			sqlVersion = "Update AP_ActionPlans set apVersion = apVersion + 1 where apID = " & ActionPlan
			con.Execute (sqlVersion)
			
			CloseWindow()
		case "final"	
			ShowProcessingMessage()
			SaveFinal()
			
			sqlVersion = "Update AP_ActionPlans set apVersion = apVersion + 1 where apID = " & ActionPlan
			con.Execute (sqlVersion)
			
			CloseWindow()
		case "point"	
			SaveDraft() 'Call this to save anything we have done so far
			SavePoint()
			
			Response.Redirect "ActionPlanFormB.asp?apID=" & ActionPlan & "&anchor=section1"
			Response.End
		case "addresp"
			SaveDraft()	'Call this to save anything we have done so far
			SaveResponsibility()
			
			Response.Redirect "ActionPlanFormB.asp?apID=" & ActionPlan & "&anchor=section2"
			Response.End
	end select
	
	function SaveDraft()
		dim rsReq, sqlReq, rsPro, sqlPro, rsOpt, sqlOpt 
		dim sqlReqInsert, sqlProInsert, sqlOptInsert, sqlAPInsert
		
		set rsReq = server.CreateObject("adodb.recordset")
		set rsPro = server.CreateObject("adodb.recordset")
		set rsOpt = server.CreateObject("adodb.recordset")		
		
		' Step 1 - Save the requirements
		sqlReq = "Select arRequirement from AP_Requirements where arActionPlan = " & ActionPlan & " order by arRequirement"
		set rsReq = con.Execute (sqlReq)
		
		while not rsReq.EOF
			' Only need to get the rating			
			if Request ("rate_" & rsReq("arRequirement")) <> "" and Request ("rate_" & rsReq("arRequirement")) <> "-" then
				sqlReqInsert = "Update AP_Requirements set arRating = " & Request ("rate_" & rsReq("arRequirement")) & _
							   " where arActionPlan = " & ActionPlan & " and arRequirement = " & rsReq("arRequirement")
				'response.write sqlReqInsert
				con.Execute (sqlReqInsert)
			end if
		
			rsReq.MoveNext
		wend
'***
'* Stephen - 09/06/2004 
'*
'* When the procedures are saved we need to make sure that the requirement is selected for so that mandatory
'* procedures are not checked when the requirement is not even selected
'***
		' Step 2 - Save the procedures for the requirements that we have selected
		sqlPro =	"SELECT prProcedure, prID, ipMandatory, ipActive " & _
					"FROM AP_Requirements INNER JOIN (IN_Procedures INNER JOIN AP_Procedures ON IN_Procedures.ipID = AP_Procedures.prProcedure) ON AP_Requirements.arID = AP_Procedures.prReq " & _
					"WHERE AP_Procedures.prActionPlan = " & ActionPlan & " AND AP_Requirements.arSelected = Yes"
		set rsPro = con.Execute (sqlPro)
		
		while not rsPro.EOF
			' For a procedure we need the responsibilities and timeframe
			if Request("resp_" & rsPro("prProcedure")) <> "" or Request("time_" & rsPro("prProcedure")) <> "" then
				sqlProInsert = "Update AP_Procedures set prResponsibilities = '" & FilterSQL(Request("resp_" & rsPro("prProcedure"))) & "', prTimeframe = '" & FilterSQL(Request("time_" & rsPro("prProcedure"))) & _
								"' , prOrigProc = '"& FilterSQL(Request("origProc_" & rsPro("prProcedure"))) &_ 
							   "' where prID = " & rsPro("prID")
				response.write(sqlProInsert)			   
				con.Execute (sqlProInsert)
			end if
			
			' If the procedure is optional then save its state
			if rsPro("ipMandatory") = false then
				dim checked
			
				if request("pro_" & rsPro("prProcedure")) = "on" then
					checked = 1
				else
					checked = 0
				end if
			
				sqlProInsert =	"Update AP_Procedures set prChecked = " & checked & _
								" where prID = " & rsPro("prID")
				con.Execute (sqlProInsert)
			else
				if rsPro("ipActive") = true then
					'The procedure is mandatory and active so select it
					sqlProInsert =	"Update AP_Procedures set prChecked = Yes "  & _
									" where prID = " & rsPro("prID")
					con.Execute (sqlProInsert)
				else
					'The procedure is not active so de-select it
					sqlProInsert =	"Update AP_Procedures set prChecked = No "  & _
									" where prID = " & rsPro("prID")
					con.Execute (sqlProInsert)
				end if
			end if
			
			' If the procedure has a text box save the text
			if request("proTxt_" & rsPro("prProcedure")) <> "" then
				sqlProInsert =	"Update AP_Procedures set prTextBox = '" & FilterSQL(request("proTxt_" & rsPro("prProcedure"))) & _
								"' where prID = " & rsPro("prID")
				con.Execute (sqlProInsert)
			end if
		
			rsPro.MoveNext
		wend		
		
		' Step 3 - Save the options
		sqlOpt =	"SELECT aoID, aoOption,ioFieldName,ioOptionType, ioTextBox, ioID " & _
					"FROM IN_Options INNER JOIN AP_Options ON IN_Options.ioID = AP_Options.aoOption " & _
					"WHERE AP_Options.aoActionPlan = " & ActionPlan
		set rsOpt = con.Execute (sqlOpt)
		
		while not rsOpt.EOF
			dim aoOption, aoText
				
			' If option type is radio
			if rsOpt("ioOptionType") = "radio" then
				' This turns off each option
				sqlOptInsert = "update Ap_Options set aoChecked = 0 where aoID = " & rsOpt("aoID")
				con.Execute (sqlOptInsert)
				
				if request(rsOpt("ioFieldName")) <> "" then
					' This turns on the selected option (even if it does it 4 times)
					sqlOptInsert = "update Ap_Options set aoChecked = 1 where aoID = " & request(rsOpt("ioFieldName"))
					con.Execute (sqlOptInsert)
				end if
			end if
				
			' If option type is checkbox
			if rsOpt("ioOptionType") = "checkbox" then
				if request(rsOpt("ioFieldName")) <> "" then
					checked = 1
				else
					checked = 0
				end if

				sqlOptInsert = "update AP_Options set aoChecked = " & checked & " where aoID = " & rsOpt("aoID")
				con.Execute (sqlOptInsert)
			end if
				
			' If the option has a text box
			if rsOpt("ioTextBox") = true then
				sqlOptInsert = "update AP_Options set aoText = '" & FilterSQL(request("optTxt_" & rsOpt("ioID"))) & "' where aoID = " & rsOpt("aoID")
				con.Execute (sqlOptInsert)
			end if
				
			rsOpt.MoveNext
		wend
		
		' Update the last details of form B
		sqlAPInsert =	"Update AP_ActionPlans set apHeadOfUnit = '" & FilterSQL(request("academic_heads")) & _
						"', apAddResp = '" & FilterSQL(request("add_resp")) & _
						"', apDevelopedBy = '" & FilterSQL(request("developedBy")) & _
						"' where apID = " & ActionPlan
		con.Execute (sqlAPInsert)
		
	end function
	
	function SaveFinal()
		SaveDraft()
		
		dim sqlFinal
		
		sqlFinal = "Update AP_ActionPlans set apCompleted = 1, apCompletionDate = '" & Date() & "' where apID = " & ActionPlan
		con.Execute (sqlFinal)
	end function
	
	function SavePoint()
		dim rsPoint, sqlPoint
		
		if request("pointtext") <> "" then
			set rsPoint = server.CreateObject("adodb.recordset")
			sqlPoint = "Insert into AP_Points (arActionPlan,arSection,arText) values (" & ActionPlan & "," & request("pointid") & ",'" & FilterSQL(request("pointtext")) & "')"

			set rsPoint = con.Execute (sqlPoint)
		End If
	end function
	
	function SaveResponsibility()
		dim sqlResp
		
		if request("heading") <> "" then
			set rsResp = server.CreateObject("adodb.recordset")
			sqlResp = "Insert into AP_ResponsibilityHeadings (rhActionPlan,rhTitle) values (" & ActionPlan & ",'" & FilterSQL(request("heading")) & "')"

			con.Execute (sqlResp)
		end if
	end function
	
	function ShowProcessingMessage()
		Response.Write "<!-- #Include file='include\header.asp' -->"
		Response.Write "<table align=center valign=middle><tr><td>"
		Response.Write "Please wait while Form B is being saved to the database. This window will close automatically and "
		Response.Write "the Action Plan menu will be refreshed."
		Response.Write "</td></tr><table>"
		Response.Write "<!-- #Include file='include\footer.asp' -->"
	end function
%>
