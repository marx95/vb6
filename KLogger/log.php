<?php
	if($_GET['f'] == 1)
	{
		$cnn_mysql	= mysql_connect("dbmy0058.whservidor.com", "pgcontrol_1", "667d2d7");
		$sdb_mysql	= mysql_select_db("pgcontrol_1", $cnn_mysql);
		
		$maquina 	= base64_encode($_POST['maquina']);
		$logs 		= base64_encode($_POST['logs']);
		$hora 		= base64_encode(date("d/m/Y h:i:s"));
		
		mysql_query("INSERT INTO logs (maquina, logs, hora) VALUES('$maquina', '$logs', '$hora')");
		die("#SUCESSO#");
	}
	
	if($_GET['f'] == 3)
	{
		$cnn_mysql	= mysql_connect("dbmy0058.whservidor.com", "pgcontrol_1", "667d2d7");
		$sdb_mysql	= mysql_select_db("pgcontrol_1", $cnn_mysql);
		$maquina	= $_GET['maquina'];
		
		$q_mysql	= mysql_query("SELECT hora, logs, maquina FROM logs WHERE maquina='$maquina' ORDER BY id");
		while($rows	= mysql_fetch_array($q_mysql))
		{
			echo base64_decode($rows['hora']) . "<br>";
		}
		die("</table>");
	}
?>
<form method="post" action="?f=1" name="FormKL">
<input name="maquina" id="maquina" value="" />
<input name="logs" id="logs" value="" />
<input type="submit" name="enviar" id="enviar" />
</form>