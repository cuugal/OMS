<!-- #Include file="include\general.asp" -->
<%
	if SecurityCheck(1) = false then
		Response.Redirect ("restricted.asp")
		Response.end
	end if
	
	PageTitle = "Health and Safety Online Management System"
	PageName  = "AuditMenu.asp"%>
	
<!-- #Include file="include\header_menu.asp" -->
<!--#include file="adovbs.inc"--> 
<%
	dim con, ActionPlan	dim sqlDraft, sqlFinal, sqlAP	dim rsDraft, rsFinal, rsAP
	set con	= server.createobject ("adodb.connection")
		con.open "DSN=ehs"
			' Get the ID of the ActionPlan this Service Agreement will be attached to	
	 Set objCmd  = Server.CreateObject("ADODB.Command")
  objCmd.CommandType = adCmdText
  Set objCmd.ActiveConnection = con
  
  ' Get the ID of the ActionPlan that the FA and CC worksheets will be attached to
  sqlAPID = "select apID " & _
        "from AP_ActionPlans " & _
        "where apFaculty = ? and apStartYear = ( " & _
        "   SELECT max(apStartYear) " & _
        "   FROM AP_ActionPlans " & _
        "   WHERE apFaculty = ? and apCompleted = Yes )"
  objCmd.CommandText = sqlAPID
  objCmd.Parameters.Append objCmd.CreateParameter("dpID", adWChar, adParamInput, 50)
  objCmd.Parameters("dpID") = Session("DepID") 
  objCmd.Parameters.Append objCmd.CreateParameter("dpID1", adWChar, adParamInput, 50)
  objCmd.Parameters("dpID1") = Session("DepID")
  'set rsAPID = con.Execute(sqlAPID) 
  set rsAPID =  server.createobject("adodb.recordset")
  rsAPID.Open objCmd	
	if not rsAPID.BOF then		ActionPlan = rsAPID("apID")	else		ActionPlan = ""	end if
	
	'Set this parameter up her as it is re-used many times
	Set DepID = objCmd.CreateParameter("dpID", adWChar, adParamInput, 50)
	DepID.Value = Session("DepID")
	
	

	Set objCmd  = Server.CreateObject("ADODB.Command")
	objCmd.CommandType = adCmdText
	Set objCmd.ActiveConnection = con
  
	sqlAP = "SELECT AP_ActionPlans.apStartYear, AP_ActionPlans.apID " & _
			"FROM AP_ActionPlans " & _
			"WHERE AP_ActionPlans.apFaculty = ? " & _
			"ORDER BY AP_ActionPlans.apStartYear desc"
	objCmd.CommandText = sqlAP
	objCmd.Parameters.Append DepID
	'set rsAP = con.Execute (sqlAP)
	set rsAP =  server.createobject("adodb.recordset")
	rsAP.Open objCmd
	
	
	Set objCmd  = Server.CreateObject("ADODB.Command")
	objCmd.CommandType = adCmdText
	Set objCmd.ActiveConnection = con
	sqlDraft = "SELECT apStartYear, apID, faComplete, faLabName, faID, faDate " & _
			   "FROM AP_ActionPlans INNER JOIN FA_Audits ON AP_ActionPlans.apID = FA_Audits.faActionPlan " & _
			   "WHERE  faComplete = No  AND apFaculty = ? " & _
			   "ORDER BY AP_ActionPlans.apStartYear desc"
	objCmd.CommandText = sqlDraft
	objCmd.Parameters.Append DepID
	'set rsDraft = con.execute(sqlDraft)
	set rsDraft =  server.createobject("adodb.recordset")
	rsDraft.Open objCmd

	Set objCmd  = Server.CreateObject("ADODB.Command")
	objCmd.CommandType = adCmdText
	Set objCmd.ActiveConnection = con
	'sqlAll = "SELECT max(apID) as ap_ID, max(faID) as fa_ID, max(faComplete) as complete, max(faDate) as auditdate, max(faAuditType) as audittype , max(faLabName) as labname, max(faAssesName) as assessor, "&_
	'		"count(FA_AuditDetails.fdAudit) as total, count(IIf(FA_AuditDetails.fdrating = 2, 1, Null)) as conformance, max(faID) as ID "&_
	'		"FROM ((AP_ActionPlans) "&_
	'		"INNER JOIN FA_Audits ON AP_ActionPlans.apID = FA_Audits.faActionPlan) "&_
	'		"LEFT OUTER JOIN FA_AuditDetails on FA_AuditDetails.fdAudit = FA_Audits.faid  "&_
	'		   "WHERE AP_ActionPlans.apFaculty = ? " & _
	'		   "group by FA_AuditDetails.fdAudit ORDER BY max(faDate) desc"
	
'ORIGINAL BELOW
	'sqlAll = "SELECT apID as ap_ID,  faComplete as complete, faDate as auditdate, faAuditType as audittype , faLabName as labname, faLocation as location, faAssesName as assessor, "&_
	'		"count(FA_AuditDetails.fdAudit) as total, sum(IIf(FA_AuditDetails.fdRating = 2, 1, 0)) as conformance, faID as fa_ID "&_
	'		"FROM (((AP_ActionPlans) INNER JOIN FA_Audits ON AP_ActionPlans.apID = FA_Audits.faActionPlan) INNER JOIN FA_AuditDetails on FA_AuditDetails.fdAudit = FA_Audits.faID) "&_
	'		" INNER JOIN IN_Requirements ON FA_AuditDetails.fdRequirement = IN_Requirements.irID "&_
	'		"WHERE AP_ActionPlans.apFaculty = ? and IN_Requirements.irManagement AND "&_
	'		" ((FA_Audits.faAuditType = 'facility' AND IN_Requirements.irFacility ) OR (FA_Audits.faAuditType = 'management' AND IN_Requirements.irManagement ) "&_
	'		" OR (FA_Audits.faAuditType = 'research' AND IN_Requirements.irResearch ) OR (FA_Audits.faAuditType = 'curriculum' AND IN_Requirements.irCurriculum )) "&_
'
'				"group by FA_AuditDetails.fdAudit, faAuditType, faLabName, faLocation, faAssesName, faID , faDate, faComplete, faID, apID ORDER BY max(faDate) desc"


'DLJ 5may15 changed below using iif to count total as being NOT -1 and conformance as being 2 or 3
'DLJ 6 May removed IN_Requirements.irManagement AND
'	sqlAll = "SELECT apID as ap_ID,  faComplete as complete, faDate as auditdate, faAuditType as audittype , faLabName as labname, faLocation as location, faAssesName as assessor, "&_
'			"sum(IIf(FA_AuditDetails.fdRating <> -1, 1, 0)) as total, sum(IIf(FA_AuditDetails.fdRating = 2 OR FA_AuditDetails.fdRating = 3, 1, 0)) as conformance, faID as fa_ID "&_
'			"FROM (((AP_ActionPlans) INNER JOIN FA_Audits ON AP_ActionPlans.apID = FA_Audits.faActionPlan) INNER JOIN FA_AuditDetails on FA_AuditDetails.fdAudit = FA_Audits.faID) "&_
'			" INNER JOIN IN_Requirements ON FA_AuditDetails.fdRequirement = IN_Requirements.irID "&_
'			"WHERE AP_ActionPlans.apFaculty = ? and "&_
'			" ((FA_Audits.faAuditType = 'facility' AND IN_Requirements.irFacility ) OR (FA_Audits.faAuditType = 'management' AND IN_Requirements.irManagement ) "&_
'			" OR (FA_Audits.faAuditType = 'research' AND IN_Requirements.irResearch ) OR (FA_Audits.faAuditType = 'curriculum' AND IN_Requirements.irCurriculum )) "&_
'
'				"group by FA_AuditDetails.fdAudit, faAuditType, faLabName, faLocation, faAssesName, faID , faDate, faComplete, faID, apID ORDER BY max(faDate) desc"

'to help diagnose inaccuracy in counting changed to count conformance and nonconformance, rather than total
	sqlAll = "SELECT apID as ap_ID, fdAudit as auditno, faComplete as complete, faDate as auditdate, faAuditType as audittype , faLabName as labname, faLocation as location, faAssesName as assessor, "&_
			"sum(IIf(FA_AuditDetails.fdRating = 0 OR FA_AuditDetails.fdRating = 1, 1, 0)) as nonconformance, sum(IIf(FA_AuditDetails.fdRating = 2 OR FA_AuditDetails.fdRating = 3, 1, 0)) as conformance, faID as fa_ID "&_
			"FROM (((AP_ActionPlans) INNER JOIN FA_Audits ON AP_ActionPlans.apID = FA_Audits.faActionPlan) INNER JOIN FA_AuditDetails on FA_AuditDetails.fdAudit = FA_Audits.faID) "&_
			" INNER JOIN IN_Requirements ON FA_AuditDetails.fdRequirement = IN_Requirements.irID "&_
			"WHERE AP_ActionPlans.apFaculty = ? and "&_
			" ((FA_Audits.faAuditType = 'facility' AND IN_Requirements.irFacility ) OR (FA_Audits.faAuditType = 'management' AND IN_Requirements.irManagement ) "&_
			" OR (FA_Audits.faAuditType = 'research' AND IN_Requirements.irResearch ) OR (FA_Audits.faAuditType = 'curriculum' AND IN_Requirements.irCurriculum )) "&_

				"group by FA_AuditDetails.fdAudit, faAuditType, faLabName, faLocation, faAssesName, faID , faDate, faComplete, faID, apID ORDER BY max(faDate) desc"
	
	objCmd.CommandText = sqlAll
	objCmd.Parameters.Append DepID
	'set rsFinal = con.execute(sqlFinal)
	set rsFinal =  server.createobject("adodb.recordset")
	rsFinal.Open objCmd
	
	

%>

<table style="width:100%" cellspacing="0" border="0" cellpadding="4" align="center">
<tr  bgcolor="#6699cc" style="height:40px">
	<td colspan="2" style="text-align:left">
		<font size="+1" face="arial" color="white">
			<b>&nbsp; Audits for the <% =Session("DepName") %></b>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<font size="-1" face="arial" color="white"><%=Session("AccessLevel")%></font>
        </font>
	</td>
</tr>
<tr>
	<td>
		<div style="border: solid 4px black; border-bottom: none;padding:5px">
			<strong>Create Audit Worksheet</strong><br/>
			<% if ActionPlan <> "" then
			%>
				<a href="javascript:void(0)" onclick="javascript:OpenWindow('AuditWorksheet.asp?apID=<%=ActionPlan%>&type=facility');">Facility Audit</a>
				<a href="javascript:void(0)" onclick="javascript:OpenWindow('AuditWorksheet.asp?apID=<%=ActionPlan%>&type=management');">Management Audit</a>
				<a href="javascript:void(0)" onclick="javascript:OpenWindow('AuditWorksheet.asp?apID=<%=ActionPlan%>&type=research');">Research Audit</a>
				<a href="javascript:void(0)" onclick="javascript:OpenWindow('AuditWorksheet.asp?apID=<%=ActionPlan%>&type=curriculum');">Curriculum Audit</a>
			<%
			else
			%>
					You need to complete a Plan before you can create a Facility Audit Worksheet
			<%
			end if %>
			
		</div>
		<div style="border: solid 4px black; padding:5px; text-align:left">
			<strong>Audit Worksheet</strong><BR>
			<ul>
			<li>a template to help conduct an audit on a facility or work area in your faculty/unit based on your Plan</li>
			<li>used by the Safety &amp; Wellbeing Branch, but can be used by anyone at any point in time</li>
			<li>can be printed off and used to note audit findings using audit criteria derived from your Plan</li>
			</ul>
			
		</div>		
	</td>
	<td>
		<div style="border: solid 4px black; border-bottom: none;padding:5px">	
			<strong>Enter Audit Findings</strong><br/>
			<%
			if SecurityCheck(2) = true then ' User must have write access for this department
				if ActionPlan <> "" then
			%>
					<a href="javascript:void(0)" onclick="javascript:OpenWindow('AuditForm.asp?apID=<%=ActionPlan%>&Mode=New&type=facility');">Facility Audit</a>
					<a href="javascript:void(0)" onclick="javascript:OpenWindow('AuditForm.asp?apID=<%=ActionPlan%>&Mode=New&type=management');">Management Audit</a>
					<a href="javascript:void(0)" onclick="javascript:OpenWindow('AuditForm.asp?apID=<%=ActionPlan%>&Mode=New&type=research');">Research Audit</a>
					<a href="javascript:void(0)" onclick="javascript:OpenWindow('AuditForm.asp?apID=<%=ActionPlan%>&Mode=New&type=curriculum');">Curriculum Audit</a>

			<%
				else
			%>
					<p>You need to complete a Plan before you can create a Facility Audit Form</p>
			<%	end If
			else		
			%>
			<br>

			<%
			end if %>

		</div>
		<div style="border: solid 4px black; padding:5px; text-align:left">
			<strong>Audit Form</strong><BR>
			<ul>
			<li> online form used to enter audit results previously recorded on Audit Worksheet</li>
			<li>used by the Safety &amp; Wellbeing Branch, but can be used by Faculty/Unit management</li>
			<li>can be saved as 'draft' and returned to at any time</li>
			</ul>
			<br/>
		</div>
	</td>
	</tr>
	<tr>
		<td colspan="2">
			<div style="border: solid 4px black; padding:5px">
				<strong>View Audit Reports</strong> (Facility/Management/Research/Curriculum)
			</div>
		</td>
	</tr>
</table>
<style type = "text/css">
	.header{
		color:white;
		background-color:#6699cc;
		font-family: "Arial",Arial,sans-serif;
		font-size: 12pt;
	}
	
	body.DTTT_Print .header{
		font-size: 10pt;
		background-color: white;
		color:#000;
	}

	#audits{
		width:100%;
		padding-top:5px;
	}
	ul
	{
		list-style: square inside url('data:image/gif;base64,R0lGODlhBQAKAIABAAAAAP///yH5BAEAAAEALAAAAAAFAAoAAAIIjI+ZwKwPUQEAOw==');
	}
	
</style>
<br/>
<table id="audits" class="display"  cellspacing="0">
	<thead>
		<tr class="header" >
            <!-- can potentially lower this security setting, depending on who needs to see these menu items.  Also change value @ line 277 -->
		    <% if SecurityCheck(4) = true then %>
			    <th style="width:150px">Action</th>
               <% end if %>
			<th>Date of Audit</th>
			<th>Audit Type</th>
			<th>Name</th>
			<th>Non-Conformances</th>
			<th>Auditors</th>
			
		</tr>
	</thead>
		<tbody>
		<% DIM recordCount
			recordCount = 0
			if not rsFinal.BOF then
			while not rsFinal.EOF 
			
			dim conformance ,  total, nonconformance
			' why is cint used here?
            ' AA - VBScript doesn't play nicely with some numbers, using cint ensures we get a math function rather than a string operation
			' DLJ 9Nov2015 changed to count conformances and nonconformances for simplicity in debugging counting error
			conformance= cint(rsFinal("conformance"))
			'total = cint(rsFinal("total"))
			nonconformance= cint(rsFinal("nonconformance"))
			'nonconformance = total-conformance
			total = conformance + nonconformance
			
			dim name
			select case rsFinal("audittype")
				Case "management"
					name = "Dean/Director"
				Case "facility"
					name = rsFinal("labname")&" "&rsFinal("location")
				Case "research"
					name = "Research Project"
				Case "curriculum"
					name = rsFinal("labname")&" "&rsFinal("location")
				Case else
					name = rsFinal("labname")&" "&rsFinal("location")
			end select
			
			%>
			
		<tr>

            
            <% if SecurityCheck(4) = true then %>
			    <td> <a href="javascript:void(0)" onclick="javascript:OpenWindow('AuditReport.asp?apID=<%=rsFinal("ap_ID")%>&faID=<%=rsFinal("fa_ID")%>&type=<%=lcase(rsFinal("audittype"))%>');">View</a>&nbsp No.<%=rsFinal("auditno")%>
				    <% if not rsFinal("complete") then %>
			          /<a href="javascript:void(0)" onclick="javascript:OpenWindow('AuditForm.asp?apID=<%=rsFinal("ap_ID")%>&faID=<%=rsFinal("fa_ID")%>&type=<%=lcase(rsFinal("audittype"))%>&Mode=Edit');">Edit</a>
			          /<a href="javascript:void(0)" onclick="checkDelete(<%=rsFinal("fa_ID")%>)">Delete</a>
				    <% end if %>
			    </td>
            <% end if %>
			<td><%=rsFinal("auditdate")%></td>
			<td><%=rsFinal("audittype")%></td>
			<td><%=name%><%if not rsFinal("complete") then %> -[DRAFT]<% end if %></td>
			<!--td><%=nonconformance%> / <%=total%></td-->
			<td><%=nonconformance%> / <%=total%></td>
			<td><%=rsFinal("assessor")%></td>

		</tr>
		<% recordCount = recordCount + 1
			rsFinal.movenext
			wend
		end if %>

		
	</tbody>
</table>

<script type="text/javascript">
   

$(document).ready(function () {
    $.fn.dataTable.moment('D/MM/YYYY');
    var table = $('#audits').dataTable({
        "order": [[ 1, "desc" ]],
        "dom": 'T<"clear">lfrtip',
        "tableTools": {
            "sSwfPath": "/TableTools/swf/copy_csv_xls_pdf.swf",
			  "aButtons": [ "copy", "print" ]
        }
    } );
});

   

     function checkDelete(abc) {
         var r = confirm("Are you sure you wish to delete this draft?");
         if (r == true) {
             // Fire off the request to /form.php
             request = $.ajax({
                 url: "AJAXDeleteDraft.asp",
                 type: "post",
                 data: "faid=" + abc,
                 async: false,
                 success: function (data) {

                     location.reload();
                 }
             });
         } else {
             //Don't need to do anything.
         }
     }
        </script>

<!-- #Include file="include\footer.asp" -->