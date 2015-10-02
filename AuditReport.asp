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

<table width="100%" border="0" cellspacing="3">
  <tr> 
    <td></td>
    <td> 
      <div align="right"><img src="utslogo.gif" width="135" height="30"></div>
    </td>
  </tr>
  <tr> 
    <td colspan="2"><h2>HEALTH AND SAFETY AUDIT REPORT</h2>
      &nbsp;&nbsp;
      <table border="1" width="100%">
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
    </td>
  </tr>
    </td>
  </tr>
  <tr> 
    <td colspan="2"> 
      <h2><br>SUMMARY OF RESULTS</h2>
		<p><font size="1">Compliance Ratings:<br>
		0 = Non-compliant<br>
		1 = Non-compliant - some action evident but not yet compliant<br>
		2 = Compliant - just requires maintenance<br>
		3 = Best practice evident</font></p>
    </td>
  </tr>
  <tr> 
    <td colspan="2"> 
      <table width="100%" border="1">
        <tr> 
          <td class="label">HEALTH AND SAFETY MANAGEMENT</td>
          <td width="15%">&nbsp;</td>
          <td class="label">SPECIFIC HAZARD PROGRAMS</td>
          <td width="15%">&nbsp;</td>
        </tr>
        <tr> 
          <td>Element</td>
          <td>Compliance rating<br>[0, 1, 2, 3]</td>
          <td>Program</td>
          <td>Compliance rating<br>[0, 1, 2, 3]</td>
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
		rsMan.movefirst

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
		rsHaz.movefirst
		
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
          <td>Compliance rating<br>[0, 1, 2, 3]</td>
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
    </td>
  </tr>
  <tr> 
    <td colspan="2"> 
		<h2>HOUSEKEEPING &amp; NOTES</h2>
    		<table border="1" width="100%">
		<tr><td><br><%=replace(rsAudDetails("faHouseKeeping"), vbCrLf, "<BR>")%><br><br></td></tr>
		</table>
		<br>
    </td>
  </tr>
  <tr> 
    <td colspan="2"> 
      <h2>STATUS OF COMPLIANCE - RESULTS IN DETAIL<br>   </h2>
    </td>
  </tr>
  <tr> 
    <td colspan="2">
    
    <%
	Function ShowStep(StepID)
		dim sqlStep, sqlReq
		dim rsStep, rsReq
		
		sqlStep = "Select stShortName from IN_Steps where stID = " & stepID
		set rsStep = con.Execute (sqlStep)
%>
		<table border="1" cellpadding="2" width="100%">
		 <tr> 
          <td colspan="5" class="StepMenu"><%=rsStep("stShortName")%></td>
        </tr>
        <tr> 
          <td class="label" width="15%">COMPLIANCE REQUIREMENTS</td>
          <td width="10%"><span class="label"><center>COMPLIANCE RATING 0,1,2,3</center></span> </td>
          <td class="label">EVIDENCE OF COMPLIANCE</td>
		  <!-- <td class="label" width="50%">EVIDENCE OF COMPLIANCE</td> -->
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
			Response.Write "&nbsp;" & "-"
			else
			Response.Write rsReq("Rating")
			End if
			%></center></td>


			<td>
				<%ShowProcedures ReqID%>
			&nbsp;</td>
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
				Response.Write Replace((replace(rsPro("fdEvidence"), vbcrlf, "<BR/>")), "|", "<br/>-")

				rsPro.movenext
			wend
		end if
	end function
%>

	<% ShowStep(3) %>
	<BR><BR>
	<% ShowStep(1) %>
	<BR><BR>
	<% ShowStep(2) %>
	<BR><BR>
    
    </td>
  </tr>
</table>

<!-- #Include file="include\footer.asp" -->