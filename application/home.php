<?php
      if(isset($_POST['file'])){
      session_start();
      $inputFileName = $_POST['file'];
      $_SESSION['file'] = $inputFileName;
    }
?>


<!DOCTYPE html>
<html>
<head>
	<title>SIMS-Home</title>
	<link rel="stylesheet" type="text/css" href="login.css">
</head>

<body>
<center>
<br>
<br>
<h1>Admin Homepage</h1>
<br>
<form method="post" action="load.php">
	<button type="button" onclick="location.href='file_view.php'">File Manager</button>
	<!-- <input type="text" id="svalue"> -->
  <button type="button" onclick="location.href='main.php'">Dashboard</button>

  <input type="file" value="Select File" name="file">

  <!-- <button type="submit" onclick="location.href='load.php'">Data Upload</button> -->
  <button type="submit" name="upload" >Data Upload</button>

  <button type="button" onclick="alert('Feature Not Avaliable')">Account Manager</button>
</form>
  <br>
  <br>
  <img src="cmsLogo.png" alt="">

</center>
</body>

</html>
