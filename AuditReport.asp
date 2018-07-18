<!-- #Include file="include\general.asp" -->
<%
	if SecurityCheck(2) = false then ' User must have write access for this department
		Response.Redirect ("restricted.asp")
		Response.end
	end if
	
	audittype = Request("type")
%>
<% PageTitle = audittype&" Audit Report"%>
	
<!-- #Include file="include\header.asp" -->

<%
	dim con, ActionPlan, AuditCount, AuditID
	dim sqlAudit, sqlAudDetails, sqlHaz, sqlMan, sqlProc
	dim rsAudit, rsAudDetails, rsHaz, rsMan, rsProc
	
	set con = server.CreateObject("adodb.connection")
	con.Open "DSN=ehs"
	
	ActionPlan = Request("apID")
	AuditID = request("faID")
	
	sqlAudDetails = "Select * from FA_Audits where faID = " & AuditID
	rsAudDetails = con.Execute(sqlAudDetails)


		' DLJ 19March2015 added to all three sql below: AND ir"&audittype&" = Yes 


	sqlHaz = "SELECT IN_Requirements.*, fdRating as Rating " & _
			 "FROM (FA_AuditDetails INNER JOIN IN_Requirements ON FA_AuditDetails.fdRequirement = IN_Requirements.irId) INNER JOIN AP_Requirements ON IN_Requirements.irId = AP_Requirements.arRequirement " & _
			 "WHERE fdAudit = " & AuditID & " AND irStep = 2 and arSelected = Yes AND arActionPlan = " & ActionPlan & " AND ir"&audittype&" = Yes " & _
			 "ORDER BY IN_Requirements.irDisplayOrder"
	set rsHaz = con.Execute (sqlHaz)
	
	sqlMan = "SELECT IN_Requirements.*, fdRating as Rating " & _
			 "FROM (FA_AuditDetails INNER JOIN IN_Requirements ON FA_AuditDetails.fdRequirement = IN_Requirements.irId) INNER JOIN AP_Requirements ON IN_Requirements.irId = AP_Requirements.arRequirement " & _
			 "WHERE fdAudit = " & AuditID & " AND irStep = 3 and arSelected = Yes AND arActionPlan = " & ActionPlan & " AND ir"&audittype&" = Yes " & _
			 "ORDER BY IN_Requirements.irDisplayOrder"
	set rsMan = con.Execute (sqlMan)
	
	sqlProc = "SELECT IN_Requirements.*, fdRating as Rating " & _
			 "FROM (FA_AuditDetails INNER JOIN IN_Requirements ON FA_AuditDetails.fdRequirement = IN_Requirements.irId) INNER JOIN AP_Requirements ON IN_Requirements.irId = AP_Requirements.arRequirement " & _
			 "WHERE fdAudit = " & AuditID & " AND irStep = 1 and arSelected = Yes AND arActionPlan = " & ActionPlan & " AND ir"&audittype&" = Yes " & _
			 "ORDER BY IN_Requirements.irDisplayOrder"
	set rsProc = con.Execute (sqlProc)
%>



<img src="utslogo.gif" width="123" alt="The UTS home page" height="52" style="border:10px solid white" align="left">

<table width="100%" border="0" cellspacing="3">
	<tr> 
		<% if rsAudDetails("faComplete") then  %>
			<td span="2"><h2>HEALTH AND SAFETY AUDIT REPORT</h2></td>
		<% Else %>
			<td span="2"><h2>HEALTH AND SAFETY AUDIT REPORT - DRAFT</h2></td>
		<% End if %>
	</tr>
</table>



<table id = "planA" cellpadding="1" width="100%">
        <tr> 
          <td class="label" width="32%">
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
          <td width="33%">
            <%=rsAudDetails("faSupervisor")%><br><br>
          </td>
          <td class="label">Faculty/Unit
		   
		  <br><br></td>
          <td><%=Session("DepName")%><br><br></td>
        </tr>
        <tr> 
          <td class="label"><% 
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
          <td>
            <%=rsAudDetails("faLabName")%><br><br>
          </td>
          <td class="label">
		  <% 
			Select case audittype
				Case "facility"
					Response.write "Location:"
				Case "research"
					Response.write "Research Project Number:"
				Case "curriculum"
					Response.write "Subject Number:"
				Case "management"
					Response.write "Office"
			End Select
		  %><br><br></td>
          <td>
            <%=rsAudDetails("faLocation")%><br><br>
          </td>
        </tr>
        <tr> 
          <td class="label">Name of Assessor<br><br></td>
          <td>
            <%=rsAudDetails("faAssesName")%><br><br>
          </td>
          <td class="label"> Date<br><br></td>
          <td>
            <%=rsAudDetails("faDate")%><br><br>
          </td>
        </tr>
</table>

<p></br></p>

      




<table id = "planA" cellpadding="1" width="100%">
		<tr> 
		  <th colspan="4" class="StepMenu">SUMMARY OF RESULTS</th>
		</tr>

        <tr> 
          <td class="label">HEALTH AND SAFETY MANAGEMENT</td>
          <td width="15%">&nbsp;</td>
          <td class="label">SPECIFIC HAZARD PROGRAMS</td>
          <td width="15%">&nbsp;</td>
        </tr>

        <tr> 
          <td>Element</td>
          <td>Compliance rating (0, 1, 2)</td>
          <td>Program</td>
          <td>Compliance rating (0, 1, 2)</td>
        </tr>

        <tr> 
          <td>
<%
		while not rsMan.EOF
			Response.Write "&nbsp;" & rsMan("irName") & "<BR>"
			
			rsMan.movenext
		wend
%>
          </td>
          <td>
<%
        If Not (rsMan.EOF And rsMan.BOF) then
		    rsMan.movefirst
        end if

		while not rsMan.EOF
			'Response.Write "&nbsp;" & rsMan("Rating") & "<BR>"
			'If statement below by DLJ 9March2015
			If rsMan("Rating") = -1 Then 
			Response.Write "&nbsp;" & "-" & "<BR>"
			else
			Response.Write "&nbsp;" & rsMan("Rating") & "<BR>"
			End if

			rsMan.movenext
		wend
%>
		  </td>
          <td rowspan="4">
<%
		while not rsHaz.EOF
			Response.Write rsHaz("irName") & "<BR>"
			
			rsHaz.movenext
		wend
%>
          </td>
          <td rowspan="4">
<%
		If Not (rsHaz.EOF And rsHaz.BOF) then
		    rsHaz.movefirst
        end if
		
		while not rsHaz.EOF
			'Response.Write "&nbsp;" & rsHaz("Rating") & "<BR>"
			'If statement below by DLJ 9March2015
			If rsHaz("Rating") = -1 Then 
			Response.Write "&nbsp;" & "-" & "<BR>"
			else
			Response.Write "&nbsp;" & rsHaz("Rating") & "<BR>"
			End if

			
			rsHaz.movenext
		wend
%>
          </td>
        </tr>

        <tr> 
          <td class="label">HEALTH AND SAFETY PROCEDURES</td>
          <td>&nbsp;</td>
        </tr>

        <tr> 
          <td>Procedure</td>
          <td>Compliance rating (0, 1, 2)</td>
        </tr>

        <tr> 
          <td>
<%
		while not rsProc.EOF
			Response.Write rsProc("irName") & "<BR>"
			
			rsProc.movenext
		wend
%>
          </td>
          <td>
<%
		
		If Not (rsProc.EOF And rsProc.BOF) then
' DLJ added above IF condition because causing recordset error where there are no PROCEDURES records i.e. in Research and Curriculum audits.
		rsProc.movefirst
		End If
				while not rsProc.EOF
					'Response.Write "&nbsp;" & rsProc("Rating") & "<BR>"

					'If statement below by DLJ 9March2015
					If rsProc("Rating") = -1 Then 
					Response.Write "&nbsp;" & "-" & "<BR>"
					else
					Response.Write "&nbsp;" & rsProc("Rating") & "<BR>"
					End if
					
					rsProc.movenext
				Wend
		

%>
          </td>
        </tr>
</table>


<p><font size="1">Compliance Ratings:<br>
0 = Non-compliant<br>
1 = Non-compliant - some action evident but not yet compliant<br>
2 = Compliant - just requires maintenance
</font></p>



<table id = "planA" cellpadding="1" width="100%">
		<tr> 
			<th colspan="2" class="StepMenu">HOUSEKEEPING &amp; NOTES</th>
		</tr>
		<tr>
			<td><%=replace(rsAudDetails("faHouseKeeping"), vbCrLf, "<BR>")%></td>
		</tr>
</table>




    
<%
	Function ShowStep(StepID)
		dim sqlStep, sqlReq
		dim rsStep, rsReq
		
		sqlStep = "Select stName from IN_Steps where stID = " & stepID
		set rsStep = con.Execute (sqlStep)
%>



<table id = "planA" cellpadding="1" width="100%">
		 <tr> 
          <th colspan="6" class="StepMenu"><%=rsStep("stName")%></th>
        </tr>

        <tr> 
          <td width="10%"><strong>Compliance Requirements</strong></td>
          <td  width="7%"><strong>Compliance Rating (0,1,2</strong>)</td>
          <td><strong>Evidence of Compliance</strong><br><font size="1">Note that non-compliances are highlighted with capitalised "NOT".</font></td>
            <% if rsAudDetails("faComplete") then  %>
             <td width="7%"><strong>Person to Action</strong></td>
             <td width="7%"><strong>Target Completion Date</strong></td>
             <td width="7%"><strong>Date Completed</strong></td>
		    <% end if  %>
        </tr>
		
<%		
		' Show the requirements
		sqlReq =	"SELECT irID " & _
					"FROM (FA_AuditDetails INNER JOIN IN_Requirements ON FA_AuditDetails.fdRequirement = IN_Requirements.irId) INNER JOIN AP_Requirements ON IN_Requirements.irId = AP_Requirements.arRequirement " & _
					"WHERE fdAudit = " & AuditID & " and irStep = " &  StepID & " and arSelected = Yes AND arActionPlan = " & ActionPlan &  " AND ir"&audittype&" = Yes "
					
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
		
		sqlReq =	"SELECT IN_Requirements.*, fdRating as Rating " & _
					"FROM (FA_Audits INNER JOIN FA_AuditDetails ON FA_Audits.faID = FA_AuditDetails.fdAudit) INNER JOIN IN_Requirements ON FA_AuditDetails.fdRequirement = IN_Requirements.irId " & _
					"WHERE faID = " & AuditID & " AND fdRequirement = " & ReqID
		set rsReq = con.Execute (sqlReq)
%>

		<tr>
			<td><span class="label"><%=rsReq("irName")%></span><br><%=rsReq("irdescription")%></td>
			
			
			<!--td><center><%=rsReq("Rating")%></center></td-->
			<td><center><%
			'If statement  by DLJ 9March2015 to replace above commented out line
			If rsReq("Rating") = -1 Then 
			Response.Write "&nbsp;" & "n/a"
			else
			Response.Write rsReq("Rating")
			End if
			%></center></td>


			<td>
				<%ShowProcedures ReqID%>
			&nbsp;</td>
            <% if rsAudDetails("faComplete") then  %>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <% end if  %>
		</tr>
<%
	end function
	
	function ShowProcedures(ReqID)
		dim sqlPro
		dim rsPro
		
		sqlPro =	"SELECT fdEvidence " & _
					"FROM FA_AuditDetails " & _
					"WHERE fdAudit = " & AuditID & " AND fdRequirement = " & ReqID
		set rsPro = con.Execute (sqlPro) 
        'response.write(sqlPro)
		
		if not rsPro.BOF then
			while not rsPro.EOF
				' Response.Write replace(rsPro("fdEvidence"), vbcrlf, "<BR>")
				Response.Write Replace((replace(rsPro("fdEvidence"), vbcrlf, "<BR/>")), "|", "<br/>")

				rsPro.movenext
			wend
		end if
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