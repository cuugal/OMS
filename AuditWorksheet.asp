<!-- #Include file="include\general.asp" -->
<!--#include file="adovbs.inc"--> 

<%
	if SecurityCheck(1) = false then ' User must have write access for this department
		Response.Redirect ("restricted.asp")
		Response.end
	end if
%>

<% PageTitle = "Audit Worksheet"%>
	
<!-- #Include file="include\header.asp" -->

<%
	dim con, ActionPlan, audittype
	dim sqlAp, sqlHaz, sqlMan, sqlProc
	dim rsAp, rsHaz, rsMan, rsProc
	
	set con = server.CreateObject("adodb.connection")
	con.Open "DSN=ehs"
	
	ActionPlan = Request("apID")
	audittype = Request("type")

	Set objCmd  = Server.CreateObject("ADODB.Command")
	objCmd.CommandType = adCmdText
	Set objCmd.ActiveConnection = con
	
	sqlHaz = "SELECT IN_Requirements.irName, AP_Requirements.arRating " & _
			 "FROM IN_Requirements INNER JOIN AP_Requirements ON IN_Requirements.irId = AP_Requirements.arRequirement " & _
			 "WHERE arSelected = Yes AND arActionPlan = ? AND irStep = 2 " & _
			 " AND ir"&audittype&" = Yes "&_
			 "ORDER BY IN_Requirements.irDisplayOrder"
			 ' DLJ 1July2016 - these query types needs to be ordered by IN_Procedures.irDisplayOrder, not IN_Requirements.irDisplayOrder


	objCmd.CommandText = sqlHaz

	objCmd.Parameters.Append objCmd.CreateParameter("arActionPlan", adWChar, adParamInput, 50)
	objCmd.Parameters("arActionPlan") = ActionPlan
	
	set rsHaz =  server.createobject("adodb.recordset")
	rsHaz.Open objCmd
	'set rsHaz = con.Execute (sqlHaz)
	''''''''''''''''''''''''''''''''''''''''''''''
	
	Set objCmd  = Server.CreateObject("ADODB.Command")
	objCmd.CommandType = adCmdText
	Set objCmd.ActiveConnection = con
	
	sqlMan = "SELECT IN_Requirements.irName, AP_Requirements.arRating " & _
			 "FROM IN_Requirements INNER JOIN AP_Requirements ON IN_Requirements.irId = AP_Requirements.arRequirement " & _
			 "WHERE arSelected = Yes AND arActionPlan = ? AND irStep = 3 " & _
			 " AND ir"&audittype&" = Yes "&_
			 "ORDER BY IN_Requirements.irDisplayOrder"
	'set rsMan = con.Execute (sqlMan)
	objCmd.CommandText = sqlMan

	objCmd.Parameters.Append objCmd.CreateParameter("arActionPlan", adWChar, adParamInput, 50)
	objCmd.Parameters("arActionPlan") = ActionPlan
	
	set rsMan =  server.createobject("adodb.recordset")
	rsMan.Open objCmd
	'''''''''''''''''''''''''''''''
	
	Set objCmd  = Server.CreateObject("ADODB.Command")
	objCmd.CommandType = adCmdText
	Set objCmd.ActiveConnection = con
	
	sqlProc = "SELECT IN_Requirements.irName, AP_Requirements.arRating " & _
			 "FROM IN_Requirements INNER JOIN AP_Requirements ON IN_Requirements.irId = AP_Requirements.arRequirement " & _
			 "WHERE arSelected = Yes AND arActionPlan = ? AND irStep = 1 " & _
			 " AND ir"&audittype&" = Yes "&_
			 "ORDER BY IN_Requirements.irDisplayOrder"
	'set rsProc = con.Execute (sqlProc)
	objCmd.CommandText = sqlProc

	objCmd.Parameters.Append objCmd.CreateParameter("arActionPlan", adWChar, adParamInput, 50)
	objCmd.Parameters("arActionPlan") = ActionPlan
	
	set rsProc =  server.createobject("adodb.recordset")
	rsProc.Open objCmd
%>


<a href="http://www.uts.edu.au/"><img src="utslogo.gif" width="123" alt="The UTS home page" height="52" style="border:10px solid white" align="left"></a>



<table  width="95%" border="0" cellspacing="0">
		<tr> 
			<td colspan="2"><div align="left"><h2>AUDIT WORKSHEET - <%=UCase(Left(audittype,1))& Mid(audittype,2) %> Audit</h2></div></td>
		</tr>
		<tr> 
			<td colspan="2">&nbsp;&nbsp;
		</tr>
</table>



<table id = "compact" width="100%">
        <tr> 
          <td class="label" width="50%">
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
          <td class="label">Faculty/Unit: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=Session("DepName")%><br>
          </td>
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
          <td class="label"><% 
			Select case audittype
				Case "facility"
					Response.write "Location:"
				Case "research"
					Response.write "Research Project Number:"
				Case "curriculum"
					Response.write "Subject Number:"
				Case "management"
					Response.write ""
			End Select
		  %>
		  <br><br></td>
        </tr>
        <tr> 
          <td class="label">Name of Auditors:<br><br></td>
          <td class="label">Date of Audit:<br><br></td>
        </tr>
</table>


	</br>





<table id = "compact" width="100%">
        <tr> 
          <td class="label" width="30%">HEALTH AND SAFETY MANAGEMENT</td>
          <!--<td width="20%">&nbsp;</td>-->
          <td class="label" width="30%">SPECIFIC HAZARD PROGRAMS</td>
         <!-- <td width="20%">&nbsp;</td>-->
        </tr>
        <tr> 
          <td>Element</td>
          <!--<td>Compliance rating 0, 1, 2, 3</td>-->
          <td>Program</td>
         <!-- <td>Compliance rating 0, 1, 2, 3</td>-->
        </tr>
        <tr> 
          <td>
<%
		if not rsMan.BOF then
			while not rsMan.EOF
				Response.Write "&nbsp;" & rsMan("irName") & "<BR>"
				
				rsMan.movenext
			wend
		end if
%>
          </td>
         <!-- <td>
<%
		'if not rsMan.BOF then
		'	rsMan.movefirst
'
'			while not rsMan.EOF
'				Response.Write "&nbsp;" & rsMan("arRating") & "<BR>"
'				
'				rsMan.movenext
'			wend
'		end if
%>
		  </td> -->
          <td rowspan="4">
<%
		if not rsHaz.BOF then
			while not rsHaz.EOF
				Response.Write rsHaz("irName") & "<BR>"
				
				rsHaz.movenext
			wend
		end if
%>
          </td>
         <!-- <td rowspan="4"> 
<%
	'	if not rsHaz.BOF then
	'		rsHaz.movefirst
	'	
	'		while not rsHaz.EOF
	'			Response.Write "&nbsp;" & rsHaz("arRating") & "<BR>"
	'			
	'			rsHaz.movenext
	'		wend
	'	end if
%>
          </td> -->
        </tr>
        <tr> 
          <td class="label">HEALTH AND SAFETY PROCEDURES</td>
         <!-- <td>&nbsp;</td>-->
        </tr>
        <tr> 
          <td>Procedure</td>
          <!--<td>Compliance rating 0, 1, 2, 3</td>-->
        </tr>
        <tr> 
          <td>
<%
		if not rsProc.BOF then
			while not rsProc.EOF
				Response.Write rsProc("irName") & "<BR>"
				
				rsProc.movenext
			wend
		end if
%>
          </td>
         <!-- <td>
<%
	'	if not rsProc.BOF then
	'		rsProc.movefirst
'
'			while not rsProc.EOF
'				Response.Write "&nbsp;" & rsProc("arRating") & "<BR>"
'				
'				rsProc.movenext
'			wend
'		end if
%>
          </td>-->
        </tr>
</table>





	</br>

    	


<table id = "worksheet" cellpadding="2" width="100%">
			<tr> 
			  <th colspan="2" class="StepMenu">HOUSEKEEPING</th>
			</tr>
			<tr><td><br><br><br><br><br><br><br><br><br><br></td></tr>
</table>
		
			
	</br>



<div class="page-break"></div>




	<!--  DLJ 23Mar2018 add management specific questions for non management areas  -->

		<%If audittype <> "management" Then %> 
		<table id = "worksheet" cellpadding="2" width="100%">
			<tr>
				<th colspan="2" class="StepMenu">AWARENESS OF MANAGEMENT STRATEGIES</th>
			</tr>
			<tr>
				<td colspan="2">Enter these findings in the Audit Form under the corresponding Compliance Requirement.</td>
			</tr>
			<tr>
				<td width = "25%"><strong>Consultation:</strong> Evident that staff are aware of faculty/unit mechanisms to allow health and safety issues to be heard and shared with workers.</td><td> &nbsp;</td>
			</tr>
			<tr>
				<td><strong>Planning:</strong> Evident that supervisor is aware of their responsibilities under the faculty/unit plan.</td><td></td>
			</tr>
			<tr>
				<td><strong>Training and competency:</strong> Evident that supervisor is aware of faculty/unit training needs analysis.</td><td></td>
			<tr>
				<td><strong>Information:</strong> Evident that staff are aware of faculty/unit health and safety information sources, such as intranet.</td><td></td>
			</tr>
			<tr>
				<td><strong>Faculty/Unit Audits:</strong> Evident that supervisor is aware of faculty/unit run safety audits.</td><td></td>
			</tr>
		</table>
			<% End If %>





        
    <%
	Function ShowStep(StepID)
		dim sqlStep, sqlReq
		dim rsStep, rsReq
		
		sqlStep = "Select stName from IN_Steps where stID = " & stepID
		set rsStep = con.Execute (sqlStep)
%>
		



		
		<table id = "worksheet" cellpadding="2" width="100%">

			<tr> 
			  <th colspan="5" class="StepMenu"><%=rsStep("stName")%></th>
			</tr>



        <tr> 
          <td class="label" width="25%">COMPLIANCE REQUIREMENTS</td>
          <!--<td width="14%"><span class="label">COMPLIANCE RATING 0,1,2,3</span> </td>-->
          <td width="25%"><span class="label">PROCEDURE</td>
          <td class="label" width="25%">EVIDENCE OF COMPLIANCE</td>
		   <td class="label" width="25%">AUDIT NOTES</td>
        </tr>
		
<%		
	Set objCmd  = Server.CreateObject("ADODB.Command")
	objCmd.CommandType = adCmdText
	Set objCmd.ActiveConnection = con
	
		' Show the requirements
		sqlReq =	"SELECT * FROM IN_Requirements INNER JOIN AP_Requirements ON IN_Requirements.irId = AP_Requirements.arRequirement " & _
					"WHERE IN_Requirements.irStep = ? AND AP_Requirements.arActionPlan = ? and arSelected = Yes" &_
					" AND ir"&audittype&" = Yes "
		
		objCmd.CommandText = sqlReq

		objCmd.Parameters.Append objCmd.CreateParameter("irStep", adInteger, adParamInput, 50)
		objCmd.Parameters("irStep") = cint(StepID)
		objCmd.Parameters.Append objCmd.CreateParameter("arActionPlan", adInteger, adParamInput, 50)
		objCmd.Parameters("arActionPlan") = cint(ActionPlan)
	
		'set rsReq = con.Execute (sqlReq)
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

		Set objCmd  = Server.CreateObject("ADODB.Command")
		objCmd.CommandType = adCmdText
		Set objCmd.ActiveConnection = con
		
		sqlReq =	"SELECT IN_Requirements.*, AP_Requirements.arRating " & _
					"FROM IN_Requirements INNER JOIN AP_Requirements ON IN_Requirements.irId = AP_Requirements.arRequirement " & _
					"WHERE AP_Requirements.arActionPlan = ? AND IN_Requirements.irId = ? "&_
					" AND ir"&audittype&" = Yes "
					
		'set rsReq = con.Execute (sqlReq)
		objCmd.CommandText = sqlReq

		objCmd.Parameters.Append objCmd.CreateParameter("arActionPlan", adInteger, adParamInput, 50)
		objCmd.Parameters("arActionPlan") = cint(ActionPlan)
		objCmd.Parameters.Append objCmd.CreateParameter("irId", adInteger, adParamInput, 50)
		objCmd.Parameters("irId") = cint(ReqID)
		
		set rsReq =  server.createobject("adodb.recordset")
		rsReq.Open objCmd
		'''''''''''''''''''''''''''''
		
		Set objCmd  = Server.CreateObject("ADODB.Command")
		objCmd.CommandType = adCmdText
		Set objCmd.ActiveConnection = con
		
		sqlCount =	"SELECT Count(*) AS Expr1 " & _
					"FROM AP_Requirements INNER JOIN AP_Procedures ON AP_Requirements.arID = AP_Procedures.prReq " & _
					"WHERE AP_Procedures.prChecked=Yes AND AP_Requirements.arRequirement =? AND AP_Requirements.arActionPlan = ?" 
		'set rsCount = con.Execute (sqlCount)
		
		objCmd.CommandText = sqlCount
		objCmd.Parameters.Append objCmd.CreateParameter("irId", adInteger, adParamInput, 50)
		objCmd.Parameters("irId") = cint(ReqID)
		objCmd.Parameters.Append objCmd.CreateParameter("arActionPlan", adInteger, adParamInput, 50)
		objCmd.Parameters("arActionPlan") = cint(ActionPlan)
		
		set rsCount =  server.createobject("adodb.recordset")
		rsCount.Open objCmd
		'''''''''''''''''''''''''''''
		
		data = ShowProcedures(ReqID	)
		Dim row
		For i = 0 to ubound(data)
			set row = data(i)
%>
		<tr>
			<% if i=0 then %>
			<td rowspan="<%if rsCount("Expr1") = 0 then response.write "1" else Response.write rsCount("Expr1") %>"><span class="label"><%=rsReq("irName")%></span><br><%=rsReq("irdescription")%>
			<br/>
			<br/>
			<span class="label">Compliance Rating [&nbsp;&nbsp;&nbsp;&nbsp;]</span><br/>
			(0,1,2,N/A)
			</td>
			<% end if %>
				<td><%=row.procedures%></td>
				<td><%=row.evidenceCompliance%></td>
				<!--td> <%If audittype = "facility" Then Response.write "check" Else Response.write " " %></td-->

			<% if i=0 then %>
				<td rowspan="<%if rsCount("Expr1") = 0 then response.write "1" else Response.write rsCount("Expr1") %>">&nbsp;</td>
			
			<% end if %>
<%		
		next

	end function
	
	class RowData
		public procedures
		public evidenceCompliance
	End class
	


	function ShowProcedures(ReqID)
		dim sqlPro, sqlNumOpt
		dim rsPro, rsNumOpt
		dim rowNum, checked
		dim var, rv, dcw
		redim dcw(-1)
		
		var = ""
		
		Set objCmd  = Server.CreateObject("ADODB.Command")
		objCmd.CommandType = adCmdText
		Set objCmd.ActiveConnection = con
		
		'sqlPro =	"SELECT IN_Procedures.*, prID, prChecked, prResponsibilities, prTimeframe, prTextBox " & _
		'			"FROM IN_Procedures INNER JOIN AP_Procedures ON IN_Procedures.ipID = AP_Procedures.prProcedure " & _
		'			"WHERE ipRequirement = ? and prActionPlan = ? and prChecked = Yes AND ip"&audittype&" = Yes " & _
		'			"order by ipDisplayOrder"
		' DLJ Mar2018 - would be great if the record set was selected by ipAudittype = Yes

		sqlPro =	"SELECT IN_Procedures.*, prID, prChecked, prResponsibilities, prTimeframe, prTextBox " & _
					"FROM IN_Procedures INNER JOIN AP_Procedures ON IN_Procedures.ipID = AP_Procedures.prProcedure " & _
					"WHERE ipRequirement = ? and prActionPlan = ? and prChecked = Yes " & _
					"order by ipDisplayOrder"


					'					"WHERE ipRequirement = ? and prActionPlan = ? and prChecked = Yes order by prID"
		'set rsPro = con.Execute (sqlPro) 
		objCmd.CommandText = sqlPro
		objCmd.Parameters.Append objCmd.CreateParameter("irId", adInteger, adParamInput, 50)
		objCmd.Parameters("irId") = cint(ReqID)
		objCmd.Parameters.Append objCmd.CreateParameter("arActionPlan", adInteger, adParamInput, 50)
		objCmd.Parameters("arActionPlan") = cint(ActionPlan)
		
		set rsPro =  server.createobject("adodb.recordset")
		rsPro.Open objCmd
		'''''''''''''''''''''''''''''
		
		' The first row is treated differently so keep track of which row we are up to
		rowNum = 1
		if not rsPro.BOF then
			while not rsPro.EOF
				
				set rv = new RowData
				
				rv.procedures = rsPro("ipName")
				
				' If the procedure has a text box display the text box
				if rsPro("ipIsTextBox") = true then
					'Response.Write rsPro("prTextBox")
					rv.procedures = rv.procedures & " " &rsPro("prTextBox")
				end if

				Set objCmd1  = Server.CreateObject("ADODB.Command")
				objCmd1.CommandType = adCmdText
				Set objCmd1.ActiveConnection = con
		
				sqlNumOpt = "SELECT count(*) as NumOptions " & _
							"FROM IN_Options INNER JOIN (AP_Procedures INNER JOIN AP_Options ON AP_Procedures.prID = AP_Options.aoPro) ON IN_Options.ioID = AP_Options.aoOption " & _
							"WHERE prActionPlan = ? AND prProcedure = ? AND ioActive = Yes"
				'set rsNumOpt = con.Execute (sqlNumOpt)
				objCmd1.CommandText = sqlNumOpt
				objCmd1.Parameters.Append objCmd.CreateParameter("prActionPlan", adInteger, adParamInput, 50)
				objCmd1.Parameters("prActionPlan") = cint(ActionPlan)
				objCmd1.Parameters.Append objCmd.CreateParameter("prProcedure", adInteger, adParamInput, 50)
				objCmd1.Parameters("prProcedure") = cint(rsPro("ipID"))
				
				set rsNumOpt =  server.createobject("adodb.recordset")
				rsNumOpt.Open objCmd1
				'''''''''''''''''''''''''''''

				if rsNumOpt("NumOptions") > 0 then
					'??
					'ShowOptions(rsPro("prID"))
				end if

				rv.evidencecompliance = rsPro("ipDefaultAuditText")
				push dcw, rv
				
				rowNum = rowNum + 1
				rsPro.movenext
			wend
		end if
		ShowProcedures = dcw
	end function
	


	function ShowOptions(OptID)
		dim sqlOpt
		dim rsOpt
		
		Set objCmd  = Server.CreateObject("ADODB.Command")
		objCmd.CommandType = adCmdText
		Set objCmd.ActiveConnection = con
		
		sqlOpt =	"SELECT * FROM IN_Options INNER JOIN AP_Options ON IN_Options.ioID = AP_Options.aoOption " & _
					"WHERE AP_Options.aoPro = " & OptID & " and aoChecked = Yes"
		'set rsOpt = con.Execute (sqlOpt)
		objCmd.CommandText = sqlOpt
		objCmd.Parameters.Append objCmd.CreateParameter("aoPro", adInteger, adParamInput, 50)
		objCmd.Parameters("aoPro") = cint(OptID)

		set rsOpt =  server.createobject("adodb.recordset")
		rsOpt.Open objCmd
		'''''''''''''''''''''''''''''
		
		while not rsOpt.EOF
%>
			<BR>
			- <%=rsOpt("ioDescription")%> 
<%		
			if rsOpt("ioTextBox") = true then
				Response.Write "- " & rsOpt("aoText")
			end if
			
			rsOpt.movenext
		wend
	end function
%>

<div class="page-break"></div></br>
	<% ShowStep(3) %>
<div class="page-break"></div></br>
	<% ShowStep(2) %>
<div class="page-break"></div></br>
	<% ShowStep(1) %>

<BR><BR>


    </td>

  </tr>
</table>



<!-- #Include file="include\footer.asp" -->
