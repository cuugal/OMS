<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
   "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
	<title><% = PageTitle %></title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="oms.css" type="text/css">
	<link rel="stylesheet" href="datatables/css/jquery.ui.css" type="text/css">
	<script type="text/javascript" src="datatables/jquery.min.js"></script>
	<script type="text/javascript" src="datatables/jquery-ui.js"></script>


<style>
#plan {
    font-family: "Trebuchet MS", Arial, Helvetica, sans-serif;
    border-collapse: collapse;
    width: 95%;
	page-break-after : avoid;
}

#plan td, #plan th {
    border: 1px solid #323232;
    padding: 8px;
}

/* #plan tr:nth-child(even){background-color: #f2f2f2;} */

#plan tr:hover {background-color: #ddd;}

#plan th {
    padding-top: 12px;
    padding-bottom: 12px;
    text-align: left;
    background-color: #0f4beb;
    color: white;
}




#planA {
    font-family: "Trebuchet MS", Arial, Helvetica, sans-serif;
    border-collapse: collapse;
    width: 95%;
}

#planA td, #planA th {
    border: 1px solid #b2b2b2;
    padding: 8px;
}

/* #planA tr:nth-child(even){background-color: #f2f2f2;} */

/*#planA tr:hover {background-color: #ddd;} */

#planA th {
    padding-top: 12px;
    padding-bottom: 12px;
    text-align: left;
    background-color: #0f4beb;
    color: #ffffff;
}



#planB {
    font-family: "Trebuchet MS", Arial, Helvetica, sans-serif;
    border-collapse: collapse;
    width: 95%;
}

#planB td, #planB th {
    border: 1px solid #b2b2b2;
    padding: 8px;
}

/* #planB tr:nth-child(even){background-color: #f2f2f2;} */

/*#planB tr:hover {background-color: #ddd;} */

#planB th {
    padding-top: 12px;
    padding-bottom: 12px;
    text-align: left;
    background-color: #0f4beb;
    color: #ffffff;
}




#worksheet {
    font-family: "Trebuchet MS", Arial, Helvetica, sans-serif;
	font-size: 12pt;
    border-collapse: collapse;
    width: 95%;
}

#worksheet td, #worksheet th {
    border: 1px solid #b2b2b2;
    padding: 4px;
}

/* #planA tr:nth-child(even){background-color: #f2f2f2;} */

/*#planA tr:hover {background-color: #ddd;} */

#worksheet th {
    padding-top: 12px;
    padding-bottom: 12px;
    text-align: left;
    background-color: #0f4beb;
    color: #ffffff;
}




#compact {
    font-family: "Trebuchet MS", Arial, Helvetica, sans-serif;
    border-collapse: collapse;
    width: 95%;
}

#compact td, #compact th {
    border: 1px solid #b2b2b2;
    padding: 3px;
}


#compact th {
    padding-top: 12px;
    padding-bottom: 12px;
    text-align: left;
    background-color: #0f4beb;
    color: #ffffff;
}


@media all {
	.page-break	{ display: none; }
}

@media print {
	.page-break	{ display: block; page-break-before: always; }
}

</style>


</head>
<body bgcolor="#ffffff" text="#000000">

<!-- Standard Page Header -->