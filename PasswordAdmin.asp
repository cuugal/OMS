<!-- #Include file="include\general.asp" -->

<%
	if SecurityCheck(3) = false then
		Response.Redirect ("restricted.asp")
		Response.end
	end if

	PageTitle = "EHS Online Management System!"
%>
	
<!-- #Include file="include\header_menu.asp" -->

<%
	dim con
	dim viewSql, editSql, passSql, supSql
	dim viewRS, editRS, passRS, supRS
	
	set con	= server.createobject ("adodb.connection")
		con.open "DSN=ehs"
		
	viewSql = "SELECT AD_Users.lgID, AD_Users.lgName " & _
			  "FROM AD_Users " & _
			  "WHERE lgDepartment = " & Session("DepID") & " AND lgView = Yes AND lgEdit = No AND lgChangePassword = No AND lgSuperUser = No"
	set viewRS = con.Execute(viewSql)
				
	editSql = "SELECT AD_Users.lgID, AD_Users.lgName " & _
			  "FROM AD_Users " & _
			  "WHERE lgDepartment = " & Session("DepID") & " AND lgView = Yes AND lgEdit = Yes AND lgChangePassword = No AND lgSuperUser = No"
	set editRS = con.Execute(editSql)

	passSql = "SELECT AD_Users.lgID, AD_Users.lgName " & _
			  "FROM AD_Users " & _
			  "WHERE lgDepartment = " & Session("DepID") & " AND lgView = Yes AND lgEdit = Yes AND lgChangePassword = Yes AND lgSuperUser = No"
	set passRS = con.Execute(passSql)
				 
	supSql = "SELECT AD_Users.lgID, AD_Users.lgName " & _
			 "FROM AD_Users " & _
			 "WHERE lgView = Yes AND lgEdit = Yes AND lgChangePassword = Yes AND lgSuperUser = Yes"		  
	set supRS = con.Execute(supSql)
	
	' Update the passwords
	dim PassID
	dim Password
	
	PassID = request("PassID")
	Password = request("Password1")
	
	if PassID <> "" then
		dim updateSql, userSql, userRs
	
		'Determine the ID of the current user (in case we need to change the current users password
		userSql = "SELECT lgID FROM AD_Users WHERE lgName = '" & FilterSQL(Session("Login")) & "' AND lgPassword = '" & FilterSQL(Session("Pass")) & "'"
		set userRs = con.Execute(userSql)

		updateSql = "Update AD_USers set lgPassword = '" & FilterSQL(Password) & "' where lgID = " & PassID
		con.Execute(updateSql)

		' If we just changed the current users password then update the session password
		if PassID = cstr(userRs("lgID")) then
			Session("Pass") = Password
		end if
		
		' Alert popup to confirm the password was changed
%>
		<script type="text/javascript">
		<!--
			alert ("The new password has been saved.")
		//-->
		</script>
<%
	end if
%>  
<font size=2 face=Arial>

<TABLE width=640 cellspacing="0" border="0" cellpadding=4 align="center">
<%
	if SecurityCheck(4) = true then
		sqlDepList = "Select dpID, dpName from AD_Departments"
		set rsDepList = con.Execute (sqlDepList)
%>		
	<tr>
		<td colspan=5>
		
		<table cellpadding=0 cellspacing="0" border="0">
		
		<form action="PasswordAdmin.asp" method="post" name="passform">
		
		<tr>             
			<td>
				<font size=2 face=Arial><b>Change passwords for</b> &nbsp;&nbsp;</font>
				<font size=-1 face=arial>
				<SELECT name="department" onchange="javascript:this.form.submit();">
					<OPTION value=-1>- Please Select a Department/Faculty -</OPTION>
<%
			while not rsDepList.EOF
%>
					<option value=<%=rsDepList("dpid")%> <%if cint(rsDepList("dpid")) = cint(Session("DepID")) then Response.Write " selected"%>><%=rsDepList("dpname")%></option>
<%
				rsDepList.MoveNext
			wend
%>
				</SELECT>
				</font>
			</td>	
		</tr>
		
		</form>
		
		</table>
		
		</td>
	</tr>
	<tr>
		<td colspan=3><br><br></td>
	</tr> 
<%
	end if
%>     
<TR>
	<TD><font size=2 face=Arial><b>Password Access Level</b></font></TD>
	<TD><font size=2 face=Arial><b>Login Name</b></font></TD>
	<TD><font size=2 face=Arial><b>New Password</b></font></TD>
	<TD><font size=2 face=Arial><b>New Password (again)</b></font></TD>
	<TD><font size=2 face=Arial></font></TD>
</TR>

<form name="viewForm" action="PasswordAdmin.asp" method=get>

<TR>
	<TD><font size=2 face=Arial>View Only</font></TD>
	<TD><font size=2 face=Arial><INPUT type="hidden" name=PassID value=<%=viewRS("lgID")%>><%=viewRS("lgName")%></font></TD>
	<TD><font size=2 face=Arial><INPUT type="password" name=password1></font></TD>
	<TD><font size=2 face=Arial><INPUT type="password" name=password2></font></TD>
	<TD><nobr><INPUT type="reset" value="Reset">&nbsp;&nbsp;<INPUT type="button" value="Change Password" onclick="DoSubmit_View();"></nobr></TD>
</TR>

</form>
<form name="editForm" action="PasswordAdmin.asp" method=post>

<TR>
	<TD><font size=2 face=Arial>Create / Edit / View</font></TD>
	<TD><font size=2 face=Arial><INPUT type="hidden" name=PassID value=<%=editRS("lgID")%>><%=editRS("lgName")%></font></TD>
	<TD><font size=2 face=Arial><INPUT type="password" name=password1></font></TD>
	<TD><font size=2 face=Arial><INPUT type="password" name=password2></font></TD>
	<TD><nobr><INPUT type="reset" value="Reset">&nbsp;&nbsp;<INPUT type="button" value="Change Password" onclick="DoSubmit_Edit();"></nobr></TD>
</TR>

</form>
<form name="passForm" action="PasswordAdmin.asp" method=post>

<TR>
	<TD><font size=2 face=Arial>Create / Edit / View / Change Password</font></TD>
	<TD><font size=2 face=Arial><INPUT type="hidden" name=PassID value=<%=passRS("lgID")%>><%=passRS("lgName")%></font></TD>
	<TD><font size=2 face=Arial><INPUT type="password" name=password1></font></TD>
	<TD><font size=2 face=Arial><INPUT type="password" name=password2></font></TD>
	<TD><nobr><INPUT type="reset" value="Reset">&nbsp;&nbsp;<INPUT type="button" value="Change Password" onclick="DoSubmit_Pass();"></nobr></TD>
</TR>

</form>


<%
	if SecurityCheck(4) = true then
%>
	<form name="supForm" action="PasswordAdmin.asp" method=post>

	<TR>
		<TD><font size=2 face=Arial>EHS Administrator</font></TD>
		<TD><font size=2 face=Arial><INPUT type="hidden" name=PassID value=<%=supRS("lgID")%>><%=supRS("lgName")%></font></TD>
		<TD><font size=2 face=Arial><INPUT type="password" name=password1></font></TD>
		<TD><font size=2 face=Arial><INPUT type="password" name=password2></font></TD>
		<TD><nobr><INPUT type="reset" value="Reset">&nbsp;&nbsp;<INPUT type="button" value="Change Password" onclick="DoSubmit_Sup();"></nobr></TD>
	</TR>
	
	</form>
<%
	end if
%>
</TABLE>

<script type="text/javascript">

	function DoSubmit_View() {
		var message
		
		message = ""
	
		if (document.viewForm.password1.value != document.viewForm.password2.value) {
			message = message + "The passwords are different. Please re-enter the new password."
		}
		
		if (document.viewForm.password1.value.length < 6) {
			message = message + "The password must be a least six characters long."
		}
		
		if (message == "") {
			document.viewForm.submit()
		}
		else {
			alert (message)
			document.viewForm.reset()
		}
	}
	
	function DoSubmit_Edit() {
		var message
		
		message = ""
	
		if (document.editForm.password1.value != document.editForm.password2.value) {
			message = message + "The passwords are different. Please re-enter the new password."
		}
		
		if (document.editForm.password1.value.length < 6) {
			message = message + "The password must be a least six characters long."
		}
		
		if (message == "") {
			document.editForm.submit()
		}
		else {
			alert (message)
			document.editForm.reset()
		}
	}
	
	function DoSubmit_Pass() {
		var message
		
		message = ""
	
		if (document.passForm.password1.value != document.passForm.password2.value) {
			message = message + "The passwords are different. Please re-enter the new password."
		}
		
		if (document.passForm.password1.value.length < 6) {
			message = message + "The password must be a least six characters long."
		}
		
		if (message == "") {
			document.passForm.submit()
		}
		else {
			alert (message)
			document.passForm.reset()
		}
	}
	
	function DoSubmit_Sup() {
		var message
		
		message = ""
	
		if (document.supForm.password1.value != document.supForm.password2.value) {
			message = message + "The passwords are different. Please re-enter the new password."
		}
		
		if (document.supForm.password1.value.length < 6) {
			message = message + "The password must be a least six characters long."
		}
		
		if (message == "") {
			document.supForm.submit()
		}
		else {
			alert (message)
			document.supForm.reset()
		}
	}

</script>

<!-- #Include file="include\footer.asp" -->