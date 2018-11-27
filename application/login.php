<?php
//
// $host="localhost";
// $user="root";
// $password="";
// $db="StudentTest"
//
// mysqli_connect($host,$user,$password);
// mysqli_select_db($db);

$con = new mysqli('localhost', 'root', '');
mysqli_select_db($con, 'StudentTest');

if(isset($_POST['username'])){
    $uname=$_POST['username'];
    $password=$_POST['password'];

    $sql="select * from loginform where User='".$uname."'AND Pass='".$password."' limit 1";

    $result=mysqli_query($con,$sql);

    if(mysqli_num_rows($result) == 1){
        // echo "Successful Login";
        // exit();
        session_start();
        $_SESSION['username'] = $_POST['username'];
        header('Location: home.php');
    }
    else{

        echo "Unsuccessful Login";
        exit();
    }

}
?>

<!DOCTYPE html>
    <head>
            <title>Login</title>
            <link rel="stylesheet" type="text/css" href="login.css">
    </head>


    <body>
        <div class ="loginbox">
        <img src="avatar.png" class = "avatar">
            <h1>Login Here:</h1>
            <form method="POST" action="#">
                <p>Username</p>
                <input type="text" name="username" placeholder="Enter Username">
                <p>Password</p>
                <input type="password" name="password" placeholder="Enter Password">
                <input type="submit" name ="" value="Login">
                <!-- <br>

                Create Account Placeholder
                <br>
                <br>
                Reset Password Placeholder -->

            </form>
        </div>
        <!-- <img src="cmsLogo.png" class = "logo" >
       <p class="quote">
           It's relationships, not programs, that change children.
           <br>
           -Bill Milliken
       </p> -->
    </body>



</html>
