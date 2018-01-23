<!--<!doctype html public "-//w3c//dtd xhtml 1.0 transitional//en" "http://www.w3.org/tr/xhtml1/dtd/xhtml1-transitional.dtd">
<html>
<head>
<title>.::DIRCON | Sistema web de gestiÃ³n de cobranzas::.</title>
<script src="scripts/jquery.min_lb.js"></script>
<script type="text/javascript" src="scripts/jquery_lb.js"></script>
<script type="text/javascript" src="scripts/lb.jquery.min.js"></script>
<link rel="stylesheet" type=text/css href="scripts/lb.css" media=all>
<script type="text/javascript">
    $(document).ready(function(){
      sexylightbox.initialize({color:'black', dir: 'imagenes'});
    });
</script>
</head>
<script language="javascript">
function global_popup_iwtsystem(ventana,ruta,nombre,propiedades)
{
	if(ventana==null)
	{
	ventana=window.open(ruta,nombre,propiedades);	
	}
	else
	{
		if(ventana.closed)			
		{
		ventana=window.open(ruta,nombre,propiedades);	
		}
	}
	ventana.focus();	
	return ventana;			
}
function ingresar()
{
	if(trim(formula.usr.value)==""){alert("debe ingresar un usuario.");return;}
	if(trim(formula.pwd.value)==""){alert("debe ingresar una contraseÃ±a.");return;}
	document.formula.submit();
}
function limpiar()
{
	document.formula.reset();
}
function trim(string)
{
	while(string.substr(0,1)==" ")
	string = string.substring(1,string.length) ;
	while(string.substr(string.length-1,1)==" ")
	string = string.substring(0,string.length-2) ;
	return string;
}	
</script>
<body topmargin=0 leftmargin=0>
<form name="formula" method="post" action="dcs_uservalida.asp" onsubmit="return false;">
	<table width="100%" height="100%" style="position: fixed" border=0 cellpadding=0 cellspacing=0>
	<tr>
		<td align=center>
			<table width="780" height="480" border=0 cellpadding=0 cellspacing=0>
			<tr>
				<td height="400"><!--<td background="imagenes/bienvenida.jpg" height="400">
						<table width="780" height="400" border=0 cellpadding=0 cellspacing=0>
							<tr>
								<td height="170" align=center><img src="imagenes/dcs_logo.jpg" height="140" border=0></td></tr>
							<tr>
								<td>
									<table align=right>
									<tr>
										<td><font face=arial size=2 color="#00529b">usuario:</font></td>
										<td><input name="usr" size=15 maxlength=30 style="width: 100px; height: 14px; font-size: x-small"></td>											
										<td width=200>&nbsp;</td>
									<tr>
									<tr>
										<td><font face=arial size=2 color="#00529b">contraseÃ±a:</font></td>
										<td><input type="password" name="pwd" size=15 maxlength=30 style="width: 100px; height: 14px; font-size: x-small"></td>
										<td width=200>&nbsp;</td></tr>									
									<tr>
										<td colspan=2 align=right><input type="image" src="imagenes/btingresar.png" alt="ingresar" title="ingresar" border=0 onclick="ingresar();"></a>&nbsp;<a style="text-decoration: none" href="javascript:limpiar();" ><img title=limpiar border=0 alt=limpiar src="imagenes/btlimpiar.png"></a></td>
										<td width=200>&nbsp;</td></tr>	
									<tr>
										<td colspan=2 align=right>
											<table>
											<tr>
												<td align=center><a style="text-decoration: none" href="#tb_inline?height=100&amp;width=400&amp;background=#000&amp;inlineid=olvclave" rel=sexylightbox ><font size=2 face="arial" color="#00529b">olvidé&nbsp;mi&nbsp;contraseña&nbsp;</font></a></td>
												<td align=center><a style="text-decoration: none" href="#tb_inline?height=100&amp;width=400&amp;background=#000&amp;inlineid=olvclave" rel=sexylightbox ><img title="olvidé mi contraseña" border =0 alt  ="olvidé mi contraseña" src="imagenes/modclave.png" ></a></td>
											</tr>
											</table>
										<td width=200>&nbsp;</td>
									</tr>
									</table>
								</td>
							</tr>
						</table>
				</td>			
			 </tr>		
    		</table>
		</td>
	</tr>
	</table>
<script language=javascript>
	document.formula.usr.focus();
</script>	
<div id="olvclave" style="display: none"> 
<iframe src="enviarclave.asp" bgcolor="#ffffff" style="position: absolute" align=top height=100 width=400 allowtransparency frameborder=0 name="fr_carpeta1" id="fr_carpeta1" ><blockquote><p>debe utilizar iexplorer 5.5 o superior.</p></blockquote></iframe></div></form>
</body>
</html>-->

<!DOCTYPE html>
<html lang="es">
  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
    <!-- Meta, title, CSS, favicons, etc. -->
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">

    <title>.::DIRCON | Sistema web de gestiÃ³n de cobranzas::.</title>

    <!-- Bootstrap -->
    <link href="vendors/bootstrap/dist/css/bootstrap.min.css" rel="stylesheet">
    <!-- Font Awesome -->
    <link href="vendors/font-awesome/css/font-awesome.min.css" rel="stylesheet">
    <!-- NProgress -->
    <link href="vendors/nprogress/nprogress.css" rel="stylesheet">
    <!-- Animate.css -->
    <link href="vendors/animate.css/animate.min.css" rel="stylesheet">

    <!-- Custom Theme Style -->
    <link href="build/css/custom.min.css" rel="stylesheet">
  </head>

  <body class="login">
    <div>
      <a class="hiddenanchor" id="signup"></a>
      <a class="hiddenanchor" id="signin"></a>

      <div class="login_wrapper">
        <div class="animate form login_form">
          <section class="login_content">
            <form>
              <h1>Acceder</h1>
              <div>
                <input type="text" class="form-control" placeholder="Username" required="" />
              </div>
              <div>
                <input type="password" class="form-control" placeholder="Password" required="" />
              </div>
              <div>
                <a class="btn btn-default submit" href="index.html">Ingresar</a>
                <a class="reset_pass" href="#">Â¿Olvido su contraseÃ±a?</a>
              </div>

              <div class="clearfix"></div>

              <div class="separator">
                <div class="clearfix"></div>
                <br />

                <div>
                  <h1> Direct Contact Solutions </h1>
                  <p>2012-2017&copy; Todos los derechos reservados.  Terminos y condiciones</p>
                </div>
              </div>
            </form>
          </section>
        </div>

        <div id="register" class="animate form registration_form">
          <section class="login_content">
            <form>
              <h1>Create Account</h1>
              <div>
                <input type="text" class="form-control" placeholder="Username" required="" />
              </div>
              <div>
                <input type="email" class="form-control" placeholder="Email" required="" />
              </div>
              <div>
                <input type="password" class="form-control" placeholder="Password" required="" />
              </div>
              <div>
                <a class="btn btn-default submit" href="index.html">Submit</a>
              </div>

              <div class="clearfix"></div>

              <div class="separator">
                <p class="change_link">Already a member ?
                  <a href="#signin" class="to_register"> Log in </a>
                </p>

                <div class="clearfix"></div>
                <br />

                <div>
                  <h1><i class="fa fa-paw"></i> Gentelella Alela!</h1>
                  <p>©2016 All Rights Reserved. Gentelella Alela! is a Bootstrap 3 template. Privacy and Terms</p>
                </div>
              </div>
            </form>
          </section>
        </div>
      </div>
    </div>
  </body>
</html>
