<?php
       
        echo "<table border='1'>";
        echo "<tr>"
            ."<th>Last Name</th>"
            ."<th>First Name</th>"
            ."<th>DCSP/Student ID</th>"
            ."<th>Birthdate</th>"
            ."<th>Match</th>"
            ."<th>School Number</th>"
            ."<th>School Name</th>"
            ."<th>Grade Level</th>"
            ."<th>Race Description</th>"
            ."<th>Gender Code</th>"
            ."<th>FRL</th>"
            ."<th>Cumulative GPA</th>"
            ."<th>Days Enrolled</th>"
            ."<th>Days Absent</th>"
            ."<th>Referrals</th>"
            ."<th>ISSP</th>"
            ."<th>OSSP</th>"
            ."<th>iReady Math BY</th>"
            ."<th>iReady Math MY</th>"
            ."<th>iReady Math EY</th>"
            ."<th>iReady ELA BY</th>"
            ."<th>iReady ELA MY</th>"
            ."<th>iReady ELA EY</th>"
            ."<th>Achieve_BY</th>"
            ."<th>Achieve_MY</th>"
            ."<th>Achieve_EY</th>"
            ."<th>Math_Course_1</th>"
            ."<th>Math_Grade_1</th>"
            ."<th>Math_Course_2</th>"
            ."<th>Math_Grade_2</th>"
            ."<th>Math_Course_3</th>"
            ."<th>Math_Grade_3</th>"
            ."<th>Math_Course_4</th>"
            ."<th>Math_Grade_4</th>"
            ."<th>English_Course_1</th>"
            ."<th>English_Grade_1</th>"
            ."<th>English_Course_2</th>"
            ."<th>English_Grade_2</th>"
            ."<th>English_Course_3</th>"
            ."<th>English_Grade_3</th>"
            ."<th>English_Course_4</th>"
            ."<th>English_Grade_4</th>"
            ."<th>Science_Course_1</th>"
            ."<th>Science_Grade_1</th>"
            ."<th>Science_Course_2</th>"
            ."<th>Science_Grade_2</th>"
            ."<th>Science_Course_3</th>"
            ."<th>Science_Grade_3</th>"
            ."<th>AFL</th>"
            ."<th>After School</th>"
            ."<th>AmeriCorps</th>"
            ."<th>BTS</th>"
            ."<th>Gear Up</th>"
            ."<th>SEP</th>"
            ."</tr>";
        
        class TableRows extends RecursiveIteratorIterator
        {
           function __construct($it) 
           {
               parent::__construct($it, self::LEAVES_ONLY);
           } 
           
           function current()
           {
               return "<td style='width: 150px; border: 1px solid black>". 
                       parent::current(). "</td>";
           }
           
           function beginChildren() 
           {
               echo "<tr>";
           }
           
           function endChildren() {
               echo "</tr>"."\n";
           }
        }
        
        $servername = "localhost";
        $username = "root";
        $password = "Database";
        $dbname = "csv_db";
        
        try
        {
        $connection = new PDO("mysql:host=$servername;dbname=$dbname", $username, 
                              $password);
        $connection->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
        $stmt = $connection->query("SELECT * FROM tbl_name");
        $stmt->execute();
        
                

        foreach(new TableRows(new RecursiveArrayIterator($stmt->fetchAll())) as $k=>$v)
        {
         echo $v;
        }
        
        }
        catch(PDOException $e)
        {
            echo "Error: " . $e->getMessage();
        }
        $connection = null;
        echo "</table>";
        
        
    ?>

