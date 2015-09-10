<!-- #Include file="include\general.asp" -->

<%
	if SecurityCheck(4) = false then
		Response.Redirect ("restricted.asp")
		Response.end
	end if
	
	PageTitle = "Health and Safety Online Management System"
	PageName  = "AuditMenu.asp"
	
	
	set con	= server.createobject ("adodb.connection")
		con.open "DSN=ehs"
		
	Set objCmd  = Server.CreateObject("ADODB.Command")
	objCmd.CommandType = adCmdText
	Set objCmd.ActiveConnection = con

	' Get the ID of the ActionPlan that the FA and CC worksheets will be attached to
	sqlAPID = "select * from AD_Users inner join AD_Departments on AD_Departments.dpID = AD_Users.lgDepartment where lgAuditor = yes order by lgID asc"

	objCmd.CommandText = sqlAPID
	set rsUsers =  server.createobject("adodb.recordset")
	rsUsers.Open objCmd
	  
	  
	Set objCmd  = Server.CreateObject("ADODB.Command")
	objCmd.CommandType = adCmdText
	Set objCmd.ActiveConnection = con

	' Get the ID of the ActionPlan that the FA and CC worksheets will be attached to
	sqlDepartment = "select * from AD_Departments"

	objCmd.CommandText = sqlDepartment
	set rsDepartments =  server.createobject("adodb.recordset")
	rsDepartments.Open objCmd
  
%>
	
<!-- #Include file="include\header_menu.asp" -->
<!--#include file="adovbs.inc"--> 

<table style="width:100%" cellspacing="0" border="0" cellpadding="4" align="center">
<tr  bgcolor="#6699cc" style="height:40px">
	<td colspan="2" style="text-align:left">
		<font size="+1" face="arial" color="white">
			<b>&nbsp;Manage Auditors</b>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<font size="-1" face="arial" color="white"><%=Session("AccessLevel")%></font>
        </font>
	</td>
</tr>
<tr>
	<td></td>
</tr>
</table>
<input type="button" value="Add New Auditor" onclick="javascript:createDialog('','','','', 'new');"/>

<table id="auditors" class="display"  cellspacing="0">
	<thead>
		<tr class="header" >
		
			<th>Action</th>
			<th>User Name</th>
			<th>Password</th>
			<th>Department</th>
		</tr>
	</thead>
		<tbody>
		<%
			if not rsUsers.BOF then
			while not rsUsers.EOF 
			
			%>
			
		<tr>
			<td> <a href="javascript:void(0)" onclick="javascript:createDialog('<%=rsUsers("lgID")%>','<%=rsUsers("lgName")%>','<%=rsUsers("lgPassword")%>','<%=rsUsers("lgDepartment")%>' ,'edit');">Edit</a>
				<a href="javascript:void(0)" onclick="javascript:deleteDialog('<%=rsUsers("lgID")%>');">Delete</a>
			
			</td>
			<td><%=rsUsers("lgName")%></td>
			<td><%=rsUsers("lgPassword")%></td>
			<td><%=rsUsers("dpName")%></td>
		</tr>
		<% 
			rsUsers.movenext
			wend
		end if %>

		
	</tbody>
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
	fieldset { padding:0; border:0; margin-top:25px; }

	input.text { margin-bottom:12px; width:95%; padding: .4em; }
	
	.overflow {
		height: 200px;
	}

	.ui-dialog {
		overflow: visible !important;
	}
}
 
 
</style>

<div id="dialog-form" title="Edit Auditor">
<p class="validateTips">All fields are required.</p>
	<form id="edit">
		<fieldset>
			<input type="hidden" name="id" id="id"/>
			<input type="hidden" name="mode" id="mode"/>
			<label for="name">Name</label>
			<input type="text" name="name" id="name" value="" class="text ui-widget-content ui-corner-all">
			<label for="password">Password</label>
			<input type="text" name="password" id="password" value="" class="text ui-widget-content ui-corner-all">
			<label for="department">Department</label><br/>
			<select style="width: 545px;" name="department" id="department"  class="text ui-widget-content ui-corner-all">
			<%
		
			if not rsDepartments.BOF then
				while not rsDepartments.EOF 
					
				%>
					<option value="<%=rsDepartments("dpID") %>"><%=rsDepartments("dpName") %></option>
				<% 
				rsDepartments.movenext
				wend
			end if %>
			</select>
			<!-- Allow form submission with keyboard without duplicating the dialog button -->
			<input type="submit" tabindex="-1" style="position:absolute; top:-1000px">
		</fieldset>
	</form>
</div>

<div id="dialog-confirm" title="Delete this Auditor?">
<form id="delete">
	<input type="hidden" name="id" id="delete_id"/>
	<input type="hidden" name="mode" id="mode" value="delete"/>
</form>
<p><span class="ui-icon ui-icon-alert" style="float:left; margin:0 7px 20px 0;"></span>This auditor will be permanently deleted and cannot be recovered. Are you sure?</p>
</div>

<script type="text/javascript">

function createDialog(id, name, password, department, mode){
	$("#id").val(id);
	$("#name").val(name);
	$("#password").val(password);
	$("#mode").val(mode);
	//console.log("boom"+department);
	dialog.dialog('open');
	$("#department").val(department); 
	$( "#department" ).selectmenu();
	
}




function deleteDialog(id){
	$("#delete_id").val(id);
	delete_Dialog.dialog('open');	
}


$(document).ready( function () {
    $('#auditors').dataTable( {
        "dom": 'T<"clear">lfrtip',
        "tableTools": {
            "sSwfPath": "/TableTools/swf/copy_csv_xls_pdf.swf",
			  "aButtons": [ "copy", "print" ]
        }
    } );
} );
var dialog;
 dialog = $("#dialog-form").dialog({
		autoOpen: false,
		height: 400,
		width: 700,
		modal: true,
		buttons: {
			"Save": saveAuditor,
			Cancel: function() {
				dialog.dialog( "close" );
			}
		}
	});

	var delete_Dialog;
 delete_Dialog = $("#dialog-confirm").dialog({
		autoOpen: false,
		height: 180,
		width: 350,
		modal: true,
		buttons: {
			"Delete": deleteAuditor,
			Cancel: function() {
			$( this ).dialog( "close" );
			}
		}
	});
		
	function saveAuditor() {
		var serializedData = $("#edit").serialize();
		// Fire off the request to /form.php
		request = $.ajax({
			url: "AJAXSaveUser.asp",
			type: "post",
			data: serializedData,
			async: false,
			success: function(data) {
                 dialog.dialog( "close" );
				 location.reload();
              }
		});
		return true;
	}
	
	
	function deleteAuditor() {
		var serializedData = $("#delete").serialize();
		// Fire off the request to /form.php
		request = $.ajax({
			url: "AJAXSaveUser.asp",
			type: "post",
			data: serializedData,
			async: false,
			success: function(data) {
                 dialog.dialog( "close" );
				 location.reload();
              }
		});
		return true;
	}
</script>
