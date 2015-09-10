<!-- #Include file="include\general.asp" -->
<%
  if SecurityCheck(2) = false then ' User must have write access for this department
    Response.Redirect ("restricted.asp")
    Response.end
  end if
%>
<% PageTitle = "Compliance Checking Form"%>
  
<!-- #Include file="include\header.asp" -->

<%
  dim con, ComplianceID
  dim sqlDate
  dim rsDate
  
  ComplianceID = request("ccID")
  ActionPlan = request("apID")
  
  set con = server.CreateObject("adodb.connection")
  con.Open "DSN=ehs"
  
  ' DELETE ANY DRAFT OR EXISTING CC FOR THE CURRENT DEPARTMENT
  if request("exists") = "yes" then
    ' This query will do a cascade delete and also delete the ServiceAgreementDetails
    sqlDel = "DELETE * " & _
         "FROM CC_Compliance " & _
         "WHERE ccActionPlan = " & ActionPlan
    con.Execute (sqlDel) 
    
    RefreshParent()
  end if
  
  if ComplianceID = "" then   
    sqlDate = "SELECT apID as ActionPlan, apCompletionDate " & _
          "FROM AP_ActionPlans " & _
          "WHERE apID = " & ActionPlan
  else
    sqlDate = "SELECT ccActionPlan AS ActionPlan, ccDate AS CompDate, ccAssessor, apCompletionDate " & _
          "FROM CC_Compliance INNER JOIN AP_ActionPlans ON CC_Compliance.ccActionPlan = AP_ActionPlans.apID " & _
          "WHERE ccID = " & ComplianceID
  end if
  
  set rsDate = con.Execute(sqlDate)
%>
<form name="compForm" action="ComplianceForm_Process.asp" method="post">
<input type="hidden" name="action" value="none">
<input type="hidden" name="apID" value="<%=rsDate("ActionPlan")%>">
<input type="hidden" name="ccID" value="<%=ComplianceID%>">

<table width="100%" border="0" cellspacing="3">
<tr> 
<td><!-- removed old EHS Branch logo CL 14/4/09 <img src="ehslogo2.gif" width="142" height="111" alt="EHS logo" border="0">-->&nbsp;</td>
<td> 
<div align="right"><img src="utslogo.gif" width="135" height="30" alt="UTS Logo"></div>
    </td>
  </tr>
  <tr> 
    <td colspan="2"> &nbsp; 
      <table width="100%" border="1" cellspacing="1" cellpadding="0">
        <tr> 
          <td> 
            <table border="0" width="100%">
              <tr> 
                <td class="label" width="15%">Faculty/Unit:</td>
                <td><%=Session("DepName")%></td>
              </tr>
              <tr> 
                <td class="label">Name of Assessor:</td>
                <td><input type="text" name="txtAsses" size="50" maxlength=150 value="<%if ComplianceID <> "" then Response.Write rsDate("ccAssessor")%>"></td>
              </tr>
              <tr> 
                <td class="label">Date:</td>
                <td><input type="text" name="txtDate" size="50" value="<%if ComplianceID <> "" then Response.Write rsDate("CompDate")%>"> (dd/mm/yyyy)</td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td colspan="2"> 
      <h2><br>
        &nbsp;&nbsp;STATUS OF COMPLIANCE WITH EHS PLAN</h2>
    </td>
  </tr>

  <tr>
    <td colspan="2">&nbsp;&nbsp;*Compliance Ratings:<br>&nbsp;&nbsp;
    <b>0</b> = Non-compliant;&nbsp;&nbsp;&nbsp;<b>1</b> = Non-compliant - Some action evident but not yet compliant;&nbsp;&nbsp;&nbsp;<b>2</b> = Compliant - just requires maintenance;&nbsp;&nbsp;&nbsp;<b>3</b> = Best practice evident</td>
  </tr>


  <tr>
    <td colspan="2"> 
      <table width="100%" border="1">
        <tr> 
          <td class="label" width="33%">COMPLIANCE REQUIREMENTS</td>
          <td class="label" width="15%">RATING FROM LAST<br>
            EHS PLAN (<%=rsDate("apCompletionDate")%>)</td>
          <td class="label">NEW COMPLIANCE <br>
            RATING (0,1,2,3)</td>
        </tr>
<%
  function ShowComplianceChecking()
    dim sqlSteps, sqlComp
    dim rsSteps, rsComp
    
    sqlSteps = "SELECT IN_Steps.stShortName, IN_Steps.stID " & _
           "FROM IN_Steps " & _
           "ORDER BY IN_Steps.stReportOrder"
    set rsSteps = con.Execute(sqlSteps)
%>  
<%    
    while not rsSteps.eof
      if ComplianceID = "" then
        sqlComp = "SELECT irName, irID, arRating, 0 as cdNewRating, irFormADescription, irStep " & _
              "FROM IN_Requirements INNER JOIN (AP_ActionPlans INNER JOIN AP_Requirements ON AP_ActionPlans.apID = AP_Requirements.arActionPlan) ON IN_Requirements.irId = AP_Requirements.arRequirement " & _
              "WHERE apID = " & ActionPlan & "AND arSelected = Yes AND irStep = " & rsSteps("stID")
      else
        sqlComp = "SELECT irName, irId, arRating, cdNewRating, irFormADescription, irStep " & _
              "FROM IN_Requirements INNER JOIN ((CC_ComplianceDetails INNER JOIN CC_Compliance ON CC_ComplianceDetails.cdCompliance = CC_Compliance.ccID) INNER JOIN AP_Requirements ON (CC_Compliance.ccActionPlan = AP_Requirements.arActionPlan) AND (CC_ComplianceDetails.cdRequirement = AP_Requirements.arRequirement)) ON (CC_ComplianceDetails.cdRequirement = IN_Requirements.irId) AND (IN_Requirements.irId = AP_Requirements.arRequirement) " & _
              "WHERE ccID = " & ComplianceID & " AND irStep = " & rsSteps("stID") & " AND arSelected = Yes " & _
              "ORDER BY irDisplayOrder"
      end if
      
      set rsComp = con.Execute(sqlComp)     
%>
      <tr><td colspan="3"><br><b><%=rsSteps("stShortName")%></b><br><br></td></tr>
<!--DLJ hack of 28 June6 starts here Purpose is to add CA descriptors to form like FormA-->
<!-- 3 if cases for each step. Steps 1 and 3 are the same. Step 2 must show irName since this info is not included in -->
<!-- the field irFormADescription for Step 2, Specific/High Risk Hazards -->
<%
      while not rsComp.eof
            if rsComp("irStep") = 1 then
%>
        <tr>
          <td><%=rsComp("irFormADescription")%></td> 
          <td>&nbsp;&nbsp;&nbsp;&nbsp;<%=rsComp("arRating")%></td>
          <td>
            &nbsp;&nbsp;&nbsp;&nbsp;
            <select name="txt_<%=rsComp("irID")%>">
              <option value="0" <%if rsComp("cdNewRating") = 0 then Response.Write " selected"%>>0</option>
              <option value="1" <%if rsComp("cdNewRating") = 1 then Response.Write " selected"%>>1</option>
              <option value="2" <%if rsComp("cdNewRating") = 2 then Response.Write " selected"%>>2</option>
              <option value="3" <%if rsComp("cdNewRating") = 3 then Response.Write " selected"%>>3</option>
            </select>
          </td>
        </tr>
<%
        end if
        rsComp.movenext
      wend
%>



<%
      rsComp.movefirst
      while not rsComp.eof
            if rsComp("irStep") = 2 then
%>
        <tr>
          <td><b><%=rsComp("irName")%></B> <%=rsComp("irFormADescription")%></td>
          <td>&nbsp;&nbsp;&nbsp;&nbsp;<%=rsComp("arRating")%></td>
          <td>
            &nbsp;&nbsp;&nbsp;&nbsp;
            <select name="txt_<%=rsComp("irID")%>">
              <option value="0" <%if rsComp("cdNewRating") = 0 then Response.Write " selected"%>>0</option>
              <option value="1" <%if rsComp("cdNewRating") = 1 then Response.Write " selected"%>>1</option>
              <option value="2" <%if rsComp("cdNewRating") = 2 then Response.Write " selected"%>>2</option>
              <option value="3" <%if rsComp("cdNewRating") = 3 then Response.Write " selected"%>>3</option>
            </select>
          </td>
        </tr>
<%
        end if
        rsComp.movenext
      wend
%>
    <%
        rsComp.movefirst

        while not rsComp.EOF
          if rsComp("irStep") = 3 then
    %>
            <tr>
              <td><%=rsComp("irFormADescription")%></td>
              <td>&nbsp;&nbsp;&nbsp;&nbsp;<%=rsComp("arRating")%></td>
              <td>
                &nbsp;&nbsp;&nbsp;&nbsp;
                <SELECT name="txt_<%=rsComp("irID")%>">
                  <option value="0" <%if rsComp("cdNewRating") = 0 then Response.Write " selected"%>>0</option>
                  <option value="1" <%if rsComp("cdNewRating") = 1 then Response.Write " selected"%>>1</option>
                  <option value="2" <%if rsComp("cdNewRating") = 2 then Response.Write " selected"%>>2</option>
                  <option value="3" <%if rsComp("cdNewRating") = 3 then Response.Write " selected"%>>3</option>
                </select>
            </td>
          </tr>
    <%    
          end if

          rsComp.movenext
        wend

          rsSteps.movenext
    wend
%>

<!--DLJ hack of 28 June ends here-->



<%
  end function
  
  ShowComplianceChecking
%>        
      </table>
    </td>
  </tr>
  <tr>
  <td><br></td>
  </tr>
  <tr> 
    <td colspan="2"> 
      <input type="button" value="    Save as Draft    " onclick="javascript:compForm.action.value='draft';DoSubmit();">
      &nbsp;&nbsp;&nbsp;&nbsp;
      <input type="button" value="    Save as Final    " onclick="javascript:compForm.action.value='final';DoSubmit();">
    </td>
  </tr>
</table>

</form>

<script type="text/javascript">
<!--
  function DoSubmit() {
    var dt=document.compForm.txtDate
    
    if (isDate(dt.value)==false){
      dt.focus()
    }
    else {
      if (document.compForm.action.value == "final")
        if (document.compForm.txtAsses.value == "")
          alert ("Please enter the Name of the Assessor")
        else
          document.compForm.submit()
      else
        document.compForm.submit()
    }
  }
//-->
</script>

<!-- #Include file="include\footer.asp" -->

<script type="text/javascript">
/**
 * DHTML date validation script. Courtesy of SmartWebby.com (http://www.smartwebby.com/dhtml/)
 */
// Declaring valid date character, minimum year and maximum year
var dtCh= "/";
var minYear=1900;
var maxYear=2100;

function isInteger(s){
  var i;
    for (i = 0; i < s.length; i++){   
        // Check that current character is number.
        var c = s.charAt(i);
        if (((c < "0") || (c > "9"))) return false;
    }
    // All characters are numbers.
    return true;
}

function stripCharsInBag(s, bag){
  var i;
    var returnString = "";
    // Search through string's characters one by one.
    // If character is not in bag, append to returnString.
    for (i = 0; i < s.length; i++){   
        var c = s.charAt(i);
        if (bag.indexOf(c) == -1) returnString += c;
    }
    return returnString;
}

function daysInFebruary (year){
  // February has 29 days in any year evenly divisible by four,
    // EXCEPT for centurial years which are not also divisible by 400.
    return (((year % 4 == 0) && ( (!(year % 100 == 0)) || (year % 400 == 0))) ? 29 : 28 );
}
function DaysArray(n) {
  for (var i = 1; i <= n; i++) {
    this[i] = 31
    if (i==4 || i==6 || i==9 || i==11) {this[i] = 30}
    if (i==2) {this[i] = 29}
   } 
   return this
}

function isDate(dtStr){
  var daysInMonth = DaysArray(12)
  var pos1=dtStr.indexOf(dtCh)
  var pos2=dtStr.indexOf(dtCh,pos1+1)
  var strDay=dtStr.substring(0,pos1)
  var strMonth=dtStr.substring(pos1+1,pos2)
  var strYear=dtStr.substring(pos2+1)
  strYr=strYear
  if (strDay.charAt(0)=="0" && strDay.length>1) strDay=strDay.substring(1)
  if (strMonth.charAt(0)=="0" && strMonth.length>1) strMonth=strMonth.substring(1)
  for (var i = 1; i <= 3; i++) {
    if (strYr.charAt(0)=="0" && strYr.length>1) strYr=strYr.substring(1)
  }
  month=parseInt(strMonth)
  day=parseInt(strDay)
  year=parseInt(strYr)
  if (pos1==-1 || pos2==-1){
    alert("The date format should be : dd/mm/yyyy")
    return false
  }
  if (strMonth.length<1 || month<1 || month>12){
    alert("Please enter a valid month")
    return false
  }
  if (strDay.length<1 || day<1 || day>31 || (month==2 && day>daysInFebruary(year)) || day > daysInMonth[month]){
    alert("Please enter a valid day")
    return false
  }
  if (strYear.length != 4 || year==0 || year<minYear || year>maxYear){
    alert("Please enter a valid 4 digit year between "+minYear+" and "+maxYear)
    return false
  }
  if (dtStr.indexOf(dtCh,pos2+1)!=-1 || isInteger(stripCharsInBag(dtStr, dtCh))==false){
    alert("Please enter a valid date")
    return false
  }
return true
}

function ValidateForm(){
  var dt=document.frmSample.txtDate
  if (isDate(dt.value)==false){
    dt.focus()
    return false
  }
    return true
 }

</script>