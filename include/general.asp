<%
const SECURITYLVL1 = -101238
const SECURITYLVL2 = -105439
const SECURITYLVL3 = -106547
const SECURITYLVL4 = -102523

' Security check and redirection
if LoggedIn() = false then
	Response.Redirect "index.asp" 
end if

dim sqlDeptID
dim rsDeptID

' Allow the department to be changed after login or by re-login
if request("Department") <> "" then
	SetDepartment(request("Department"))
end if

' Set the department session variable if it is not already set
if session("DepID") = "" then
	set con	= server.createobject ("adodb.connection")	con.open "DSN=ehs;uid=;pwd=;"

	sqlDeptID = "SELECT dpID " & _
				"FROM AD_Departments INNER JOIN AD_Users ON AD_Departments.dpID = AD_Users.lgDepartment " & _
				"WHERE AD_Users.lgName= '" & FilterSQL(Session("Login")) & "' and lgPassword = '" & FilterSQL(Session("Pass")) & "'"
	set rsDeptID = con.Execute (sqlDeptID)
	
	SetDepartment(rsDeptID("dpID"))
end if

function SetDepartment(DepID)
	set con	= server.createobject ("adodb.connection")	con.open "DSN=ehs;uid=;pwd=;"
	
	sqlDeptID = "SELECT dpID, dpName FROM AD_Departments WHERE dpID = " & DepID
	set rsDeptID = con.Execute (sqlDeptID)

	Session("DepID") = rsDeptID("dpID")
	Session("DepName") = rsDeptID("dpName")
end function

function LoggedIn()
	if Session("Login") <> "" then
		dim sqlPass
		dim rsPass
		dim con
		
		set con	= server.createobject ("adodb.connection")		con.open "DSN=ehs;uid=;pwd=;"
		
		' Compare the passwords
		sqlPass = "SELECT * FROM AD_Users WHERE lgName = '" & FilterSQL(Session("Login")) & "' AND lgPassword = '" & FilterSQL(Session("Pass")) & "'"
		set rsPass = con.Execute (sqlPass)
		
		if not rsPass.BOF then
			LoggedIn = true
		else
			LoggedIn = false
		end if
	else
		LoggedIn = false
	end if
end function

function SecurityCheck(Level)
	' Determine if the user has the right access for the current page
	dim sqlUser
	dim rsUser
	dim con, passed
	
	passed = true
		
	set con	= server.createobject ("adodb.connection")	con.open "DSN=ehs"
	
	sqlUser = "SELECT * FROM AD_Users WHERE lgName = '" & FilterSQL(Session("Login")) & "' AND lgPassword = '" & FilterSQL(Session("Pass")) & "'"
	set rsUser = con.execute(sqlUser)
	
	'Response.Write rsUser("lgSuperUser") & " Sup<BR>"
	'Response.Write rsUser("lgEdit") & " Edit<BR>"
	'Response.Write rsUser("lgChangePassword") & " Pass<BR>"
	'Response.Write rsUser("lgDepartment") & " deptid<BR>"
	'Response.Write Department & " depname<BR>"
	'Response.Write Session("login") & " login<BR>"
	'Response.Write level & " level<BR>"
	
	
	'Level 1 = read, 2 = write, 3 = password, 4 = super , 5 = auditor
	select case(Level)
		case 1
			if rsUser("lgSuperUser") <> true then
				if cint(rsUser("lgDepartment")) <> cint(Session("DepID")) then
					passed = false
				end if
			end if
		case 2
			if rsUser("lgSuperUser") <> true then
				if rsUser("lgEdit") <> true or cint(rsUser("lgDepartment")) <> cint(Session("DepID")) then
					passed = false
				end if
			end if
		case 3
			if rsUser("lgSuperUser") <> true then
				if rsUser("lgChangePassword") <> true or cint(rsUser("lgDepartment")) <> cint(Session("DepID")) then
					passed = false
				end if
			end if
		case 4
			if rsUser("lgSuperUser") <> true then
				passed = false
			end if
		case 5
			if rsUser("lgAuditor") <> true then
				passed = false
			end if
	end select
	
	'Response.Write passed
	'Response.End
	
	SecurityCheck = passed
end Function

function CloseWindow
	' This function writes the code to close a window using javascript
	RefreshParent()
%>
	<script language=javascript>
		window.close()
	</script>
<%
end function

function RefreshParent
	' This function writes the code to close a window using javascript
%>
	<script language=javascript>
		opener.location.reload()
	</script>
<%
end function

function FilterSQL(sqlString)
	FilterSQL = replace(sqlString,"'","''")
end function

function FilterTrailingCrLf(sqlString)
	dim RegEx, result
	
	set RegEx = new RegExp
	regEx.Pattern = "^(?:\r\n)+|(?:\r\n)+$|((\r\n)\2)\2+"
	regEx.MultiLine = False  
	regEx.Global = True   

	FilterTrailingCrLf = regEx.Replace (sqlString, "$1")

end function

function Department_List
	' declare local variables
	dim con, rs, sql_department

	sql_department = "Select dpID, dpName from AD_Departments"

	' set recordset and connection objects
	set rs = server.CreateObject ("adodb.recordset")
	set con = server.createobject ("adodb.connection")

	con.open "DSN=ehs"

	set rs = con.Execute (sql_department)
%>
	<SELECT name=department>		<OPTION value=-1>- Please Select a Faculty/Unit -</OPTION>
<%		while not rs.EOF
%>			<option value=<%=rs("dpid")%>><%=rs("dpname")%></option>
<%
			rs.MoveNext		wend
%>	</SELECT>
<%
end function

sub push(byref arr, var) 
   dim uba 
   uba = UBound(arr) 
   redim preserve arr(uba+1) 
   set arr(uba+1) = var 
   
 end sub

'Function InsertRecord( tblName, ArrFlds, ArrValues )
	' This function recieves a tablename and an Array of 
	' Fieldnames and an Array of field Values. 
	' It returns the ID of the record that has been inserted.


	' Turn error handling on.
	'On Error Resume Next

	'dim cnnInsert, rstInsert, thisID

	' Object instantiation.	
	'Set cnnInsert = Server.CreateObject ("ADODB.Connection")
	'Set rstInsert = Server.CreateObject ("ADODB.Recordset")

	' Open our connection to the database.
	'cnnInsert.open ("DSN=ehs")
			
	' Open our Table (using the tblName that we passed in).
	'rstInsert.Open tblName, cnnInsert, adOpenKeyset, _
	'                   adLockOptimistic, adCmdTable

	' Call the AddNew method.	
	'rstInsert.AddNew  ArrFlds, ArrValues

	' Commit the changes.		
	'rstInsert.Update				

	' Retrieve the ID.
	'thisID = rstInsert("ID")

	'If Err.number = 0 Then
	' If the Err.number = 0 then close everything and 
	' return the ID.
	    'rstInsert.Close
	    'Set rstInsert = Nothing
	    'cnnInsert.close
	    'Set cnnInsert = Nothing
				
	    'InsertRecord = thisID
			
	'Else
	' Woops, an error occurred, close everything and display an 
	' error message to the user.
	    'rstInsert.Close
	    'Set rstInsert = Nothing
	    'cnnInsert.close
	    'Set cnnInsert = Nothing
			
	    ' Call our re-useable Error handler function.	
	    'InsertRecord = Err.number & ", " & Err.Description & ", " & Err.Source
			
	'End If
		
'End Function

%>

