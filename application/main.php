<?php session_start(); ?>
<!DOCTYPE html>
<html>
<head>
	<title>SIMS: Student Information Management System</title>
	<link rel="stylesheet" type="text/css" href="login.css">
</head>
<body>
<center>
<h1>Welcome, <?php $_SESSION['username']; ?></h1>
<h2>Student Data</h2>
<section>
	<button type="button" onclick="location.href='home.php'">Home</button>
	<input type="text" id="svalue">
	<select>
		<option value="Name">Name</option>
		<option value="Grade">Grade</option>
		<option value="Race">Race</option>
		<option value="Gender">Gender</option>
		<option value="School">School</option>
		<option value="option">More to come</option>
	  </select>
	  <button type="button" onclick="alert('Feature Not Avaliable')">Search</button>
		<button type="button" onclick="alert('Feature Not Avaliable')">Import Data</button>
	  <button type="button" onclick="alert('Feature Not Avaliable')">Export Data</button>
	<br>
	<br>
  <img src="cmsLogo.png" alt="">
</center>
</body>
</html>
