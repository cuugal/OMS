<%@Language = VBscript%>

<!-- #Include file="include\general.asp" -->

<%
	if SecurityCheck(1) = false then ' User must have write access for this department
		Response.Redirect ("restricted.asp")
		Response.end
	end if
%>

<% PageTitle = "Service Agreement Report"%>

<!-- #Include file="include\header.asp" -->

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Service Agreement Management Report</title>
</head>
<body>


<form method="POST" action="SAMgmtReport.asp">

<%
'********************Gathering information from the exsisting form*****************************
numRecords = Request.Form("hdnRecordCount") 
   'SA = Request.Form("chkSA")
   'C = Request.Form("txtASComments")

dim strMonth
dim strMonthVal
dim strYear
dim con
dim strSQL
dim rsQueryA
dim rsQueryB
dim strDate
dim strD
dim numRecordCounter
dim numRecords
dim sdRequirement

strMonth = Request.Form("cboMonth")
strYear = Request.Form("txtYear")
if strMonth ="" or strYear ="" then %>
			<script type="text/javascript">
				alert("Either the month or year field is blank. Please check your month or year and try again.")
				window.close();
			</script>
<%else
strDate = "1/"+strMonth+"/"+strYear
strDate = cdate(strDate)
'Response.Write(strDate)

select case strMonth
   case "1":strMonthVal = "January" 
   case "2":strMonthVal = "February" 
   case "3":strMonthVal = "March"
   case "4":strMonthVal = "April"
   case "5":strMonthVal = "May"
   case "6":strMonthVal = "June"
   case "7":strMonthVal = "July"
   case "8":strMonthVal = "August"
   case "9":strMonthVal = "September"
   case "10":strMonthVal = "October"
   case "11":strMonthVal = "November"
   case "12":strMonthVal = "December"
end select   
strd = strMonthVal + " " +strYear
%>
<p><b><font size="4">EHS Service Agreements from <%=strd%>  To <%=date()%> 
</font></b></p> 

<b> 

<%


'*******************************Database connectivity code************************************************    
	set con	= server.createobject ("adodb.connection")
	con.open "DSN=ehs"
'**************************************query section starts here*******************************************************
    strSQl = "SELECT * FROM AP_ActionPlans INNER JOIN AD_Departments ON AP_ActionPlans.apFaculty = AD_Departments.dpID where apCompletiondate > #"& strDate &"# order by apFaculty"
    
    set rsQueryA = con.Execute (strSQL)
    
    strSQL = "SELECT irID, irName AS Requirement, arRating AS Rating, sdID, saSA, saC, saActionPlan, sdServiceActioned, sdRequirement, sdComments, sdEHSServices AS EHSServices, sdContact AS Contact, sdTimeFrame AS TimeFrame, saAddEHSServices AS AddEHSServices, saAddContact AS AddContact, saAddTimeFrame AS AddTimeFrame "_ 
    &"FROM SA_ServiceAgreement INNER JOIN (IN_Requirements INNER JOIN (SA_ServiceAgreementDetails INNER JOIN AP_Requirements ON SA_ServiceAgreementDetails.sdRequirement = AP_Requirements.arRequirement) ON IN_Requirements.irId = AP_Requirements.arRequirement) ON (SA_ServiceAgreement.saActionPlan = AP_Requirements.arActionPlan) "_ 
    &" AND (SA_ServiceAgreement.saID = SA_ServiceAgreementDetails.sdServiceAgreement)"_
    &" WHERE sdServiceActioned = 0"
    
    set rsQueryB = con.Execute (strSQL) %>
    
    
        
<%'***********************************************section ends here******************************************************
%>
<%'****************************************Reporting Section starts here***********************************************
  dim flg
  numRecordCounter = 0    
  while not rsQueryA.EOF
  dim cMonth
  dim cYear
  dim SaActionPlan
  cMonth = month(rsQueryA("apCompletionDate"))
  cYear = Year(rsQueryA("apCompletionDate"))
  select case cMonth
   case "1":strMonthVal = "January"
   case "2":strMonthVal = "February" 
   case "3":strMonthVal = "March"
   case "4":strMonthVal = "April"
   case "5":strMonthVal = "May"
   case "6":strMonthVal = "June"
   case "7":strMonthVal = "July"
   case "8":strMonthVal = "August"
   case "9":strMonthVal = "September"
   case "10":strMonthVal = "October"
   case "11":strMonthVal = "November"
   case "12":strMonthVal = "December"
end select  
  
   %></b><p><b><font size="4"><%=rsQueryA("dpName")%></font></b></p>
           <table border="1" cellpadding=3 id="table1">
             <tr> 
             <td class="label">COMPLIANCE REQUIREMENT</td>
             <td class="label"><center>Compliance<br>rating at<br><%=strMonthVal%>, <%=cYear%></center></td>
             <td class="label">EHS SERVICES</td>
             <td class="label">FACULTY/UNIT CONTACT</td>
             <td class="label">TIMEFRAME</td>
          
             </tr>
      <%  while not rsQueryB.EOF 
          if rsQueryA("ApId").value = rsQueryB("saActionPlan").value then 
             
          if rsQueryB("EHSServices")<> "" then 
            numRecordCounter = numRecordCounter + 1  
            sdRequirement = rsQueryB("sdRequirement")  
            
                  %>
          
          <input type ="hidden" name=hdnSdRequirement<%=numRecordCounter%>value =<%=sdRequirement %>>  
          
          <tr> 
          <td> 
			<%=rsQueryB("Requirement")%>&nbsp;
          </td>
          <td> 
            <center><%=rsQueryB("Rating")%>&nbsp;</center>
          </td>
          <td> 
                <%=rsQueryB("EHSServices")%>&nbsp;
          </td>
          <td> 
            <%=rsQueryB("Contact")%>&nbsp;
          </td>
          <td> 
            <%=rsQueryB("TimeFrame")%>&nbsp;
          </td>
                     
                  
          <%                       
             end if
                   
            dim EHSVal
            dim AddContact
            dim AddTimeFrame
            dim saSA
            dim Comments 
            
            saActionPlan = rsQueryB("saActionPlan").value
            EHSVal = rsQueryB("AddEHSServices").value 
            AddContact = rsQueryB("AddContact").value 
            AddTimeFrame = rsQueryB("AddTimeFrame").value
            saSA = rsQueryB("saSA").value
            Comments = rsQueryB("sac").value  
            
          end if  
     
'***********************************************************************************************            
            rsQueryB.movenext
       wend
'********************************************code for additional EHS Services*******************
     
     if EHSVal <> "" then %>
		<tr> 
          <td> 
			Additional EHS Services
          </td>
          <td> 
            <center>--</center>
          </td>
          <td> 
            
            <%=EHSVal%>
          </td>
          <td> 
            <%=AddContact%>
          </td>
          <td> 
           <%=AddTimeFrame%></td>
         
            <%end if
'**********************************************code ends here***********************************    
         %></table><%
           rsQueryB.moveFirst
        rsQueryA.movenext  
  wend      
%>
    <input type ="hidden" name=hdnRecordCount value =<%=numRecordCounter %>>
<%end if%>    
    
</form>
</body>
</html>