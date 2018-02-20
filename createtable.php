<!DOCTYPE html>
<html>
<head>
<style>

table {
	font-family: "Trebuchet MS", Arial, Helvetica, sans-serif;
    width: 100%;
    border-collapse: collapse;
}

table, td, th {
    border: 1px solid #ddd;
    padding: 4px;
}

table tr:nth-child(even){background-color: #f2f2f2;}
table tr:hover {background-color: #ddd;}
table th {
    padding-top: 12px;
    padding-bottom: 12px;
    text-align: left;
    background-color: #4CAF50;
    color: white;
}

</style>
</head>
<body>
<?php

	$mysqli = mysqli_connect("localhost", "root", "", "x2crm");
	$res = $mysqli->query("

SELECT
	a.assignedTo,
	a.associationName,
	a.dueDate, 
	a.completeDate,
	(a.completeDate - a.dueDate) as total,
	b.text,
	c.eventSubtype,
	c.eventStatus
FROM x2_actions a
	INNER JOIN x2_action_text b
	  ON a.id = b.actionId	
	INNER JOIN x2_action_meta_data c
	  ON a.id = c.actionId
WHERE a.type = 'event' and 
		NOT a.associationName = 'Contacts' and 
		NOT a.associationName = 'Calendar' and  
		WEEKOFYEAR(FROM_UNIXTIME(a.completeDate)) >= WEEKOFYEAR(NOW())-3 and
		WEEKOFYEAR(FROM_UNIXTIME(a.completeDate)) <= WEEKOFYEAR(NOW())+1 and
		YEAR(FROM_UNIXTIME(a.completeDate)) = YEAR(NOW())
ORDER BY a.assignedTo, a.completeDate ");


echo "<table>
<tr>
<th>Provider</th>
<th>Client</th>
<th>Start</th>
<th>End</th>
<th>Hours</th>
<th>Description</th>
<th>Type</th>
<th>Status</th>
</tr>";

while($row = mysqli_fetch_array($res)) {
    echo "<tr>";
    echo "<td>" . $row['assignedTo'] . "</td>";
    echo "<td>" . $row['associationName'] . "</td>";
    echo "<td>" . $row['dueDate'] . "</td>";
    echo "<td>" . $row['completeDate'] . "</td>";
    echo "<td>" . $row['total'] . "</td>";
	echo "<td>" . $row['text'] . "</td>";
    echo "<td>" . $row['eventSubtype'] . "</td>";
    echo "<td>" . $row['eventStatus'] . "</td>";
    echo "</tr>";
}


mysqli_close($mysqli);

?>
</body>
</html>





