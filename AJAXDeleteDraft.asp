<!--#include file="adovbs.inc"--> 
<%


	apid = request("apid")
    faid = request("faid")

  
   
    if(apid <> "") then

    Response.write(apid)
	    set con	= server.createobject ("adodb.connection")
		    con.open "DSN=ehs"
		
	    Set objCmd  = Server.CreateObject("ADODB.Command")
	    objCmd.CommandType = adCmdText
	    Set objCmd.ActiveConnection = con
		  

	    sqlDel = "delete from AP_ActionPlans where apid = ?"
		
	    objCmd.CommandText = sqlDel
		 
	    objCmd.Parameters.Append objCmd.CreateParameter("apid", adInteger, adParamInput, 50)
	    objCmd.Parameters("apid") = cint(apid)


	    set rsFinal =  server.createobject("adodb.recordset")
	    rsFinal.Open objCmd

	    Response.write(1)
    end if

    
    if(faid <> "") then
      
        Response.write("faid"&faid)
	    set con	= server.createobject ("adodb.connection")
		    con.open "DSN=ehs"
		
	    'Set objCmd  = Server.CreateObject("ADODB.Command")
	    'objCmd.CommandType = adCmdText
	    'Set objCmd.ActiveConnection = con
		  

	    'sqlDel = "delete from FA_AuditDetails where fdaudit = ?"
		
	    'objCmd.CommandText = sqlDel
		 
	    'objCmd.Parameters.Append objCmd.CreateParameter("faid", adInteger, adParamInput, 50)
	    'objCmd.Parameters("faid") = cint(faid)


	    'set rsFinal =  server.createobject("adodb.recordset")
	    'rsFinal.Open objCmd



        Set objCmd  = Server.CreateObject("ADODB.Command")
	    objCmd.CommandType = adCmdText
	    Set objCmd.ActiveConnection = con

        sqlDel = "delete from FA_Audits where faid = ?"
		
	    objCmd.CommandText = sqlDel
		 
	    objCmd.Parameters.Append objCmd.CreateParameter("faid", adInteger, adParamInput, 50)
	    objCmd.Parameters("faid") = cint(faid)


	    set rsFinal =  server.createobject("adodb.recordset")
	    rsFinal.Open objCmd

	    Response.write(1)
    end if

	
%>