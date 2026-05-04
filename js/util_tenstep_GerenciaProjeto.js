<<<<<<< HEAD
// JScript source code

function uncheck(form, field)
{
	for (i=0; i<document.forms[form].elements[field].length; i++)
		document.forms[form].elements[field][i].checked = false;
}

function check_cpf (numcpf)
{
x = 0;
soma = 0;
dig1 = 0;
dig2 = 0;
texto = "";
numcpf1="";
len = numcpf.length; x = len -1;
// var numcpf = "12345678909";
for (var i=0; i <= len - 3; i++) {
y = numcpf.substring(i,i+1);
soma = soma + ( y * x);
x = x - 1;
texto = texto + y;
}
dig1 = 11 - (soma % 11);
if (dig1 == 10) dig1=0 ;
if (dig1 == 11) dig1=0 ;
numcpf1 = numcpf.substring(0,len - 2) + dig1 ;
x = 11; soma=0;
for (var i=0; i <= len - 2; i++) {
soma = soma + (numcpf1.substring(i,i+1) * x);
x = x - 1;
}
dig2= 11 - (soma % 11);
if (dig2 == 10) dig2=0;
if (dig2 == 11) dig2=0;
//alert ("Digito Verificador : " + dig1 + "" + dig2);
if ((dig1 + "" + dig2) == numcpf.substring(len,len-2)) {
return true;
}
return false;
}

function valida_email(email) {

  if ((email.indexOf('@') == email.lastIndexOf('@')) &&
      (email.indexOf('@') > 0) &&
      (email.lastIndexOf('.') > (email.indexOf('@') + 1)) &&
      (email.lastIndexOf('.') + 1 != email.length)) return true
  else return false
}

function consisteEmail (email) {
  //O pattern abaixo reflete as seguintes regras:
  //- no máximo 25 caracteres (alfanumérico, hífen, ponto) possíveis antes da @, sendo q o último
  //  deve ser um caracter alfanumérico
  //- cada parte do domínio pode ter no máximo 30 caracteres (alfanumérico, hífen), último não 
  //  pode ser hífen
  //- domínio pode ter até 4 partes separadas por ponto, sendo que a ultima tem entre 2 e 3
  //  caracteres (alfanumérico), geralmente com, net, br, etc
  pattern  = /(^[a-z0-9.-]{1,24}[a-z0-9]@([a-z0-9-]{1,29}[a-z0-9].){1,3}[a-z0-9]{2,3}$)/;
  
  return pattern.test(email);

}

function formataData(campo){
//Formata o campo de data de acordo com a tecla pressionada
  
  var pos = 2;
  var pos1 = 4;
  var tecla = event.keyCode;
  var tam = 0;
  var vrout = "";

  vr = ""
  for (i=0 ;i <= campo.value.length; i++) {
    c = campo.value.substr(i,1);
  if (c >= "0" && c <= "9") {
    vr += c;
  }
  }
  tam = vr.length ;

  if (tecla == 8 ){ 
  tam = tam - 1 ; 
  }
  
  if ( tecla == 9 || tecla == 8 || tecla >= 48 && tecla <= 57 || tecla >= 96 && tecla <= 105){

  vrout = "" ;
  for (i=0 ;i <= tam ; i++) {
  
  c = vr.substr(i,1);

  switch (i){  
    case pos : 
    case pos1 :      
      vrout += "/";      
      break;  
    } // switch
  vrout += c;
  } // for
  
  campo.value = vrout;
  }
    
  else {
    return(false);
  }

  return(true);
  
}

//testa se caractere é numérico
function isNum(caractere)
{
  var strValidos = "0123456789."
  if ( strValidos.indexOf( caractere ) == -1 )
    return false;
  return true;
}
function isNum2(caractere)
{
  var strValidos = "0123456789"
  if ( strValidos.indexOf( caractere ) == -1 )
    return false;
  return true;
}
  
//retorna verdadeiro se tecla em event é númerica ou backspace
//retorna falso caso contrário
function validaTeclaNumerica(event)
{
  var BACKSPACE=  8;
  var key;
  var tecla;

  if(navigator.appName.indexOf("Netscape")!= -1)  
    tecla= event.which;    
  else
    tecla= event.keyCode;

  key = String.fromCharCode( tecla);

  if ( tecla == 13 )
    return false;
  if ( tecla == BACKSPACE )
    return true;
  return ( isNum(key));

} 
function validaTeclaNumerica2(event)
{
  var BACKSPACE=  8;
  var key;
  var tecla;

  if(navigator.appName.indexOf("Netscape")!= -1)  
    tecla= event.which;    
  else
    tecla= event.keyCode;

  key = String.fromCharCode( tecla);

  if ( tecla == 13 )
    return false;
  if ( tecla == BACKSPACE )
    return true;
  return ( isNum2(key));

} 

function validaData ( data ) {
//Verifica se a data é válida (data em português) 
//no formato DD/MM/YY ou DD/MM/YYYY

  if (data == "" ) {
    return (false) ;
  }  
  else {
  
    var datePat = /^(\d{1,2})(\/|-)(\d{2})\2(\d{4})$/i;

    var matchArray = data.match(datePat); //verifica o formato
    if (matchArray == null) {
      return (false);
    }
  
    day = matchArray[1]; // divide a data em variáveis
    month = matchArray[3];
    year = matchArray[4];

    if (month < 1 || month > 12) { // verifica o intervalo de meses
      return (false);
    }
    if (day < 1 || day > 31) {
      return (false);
    }
    if ((month==4 || month==6 || month==9 || month==11) && day==31) {
      return (false);
    }
    if (month == 2) { //verifica data de 29 de fevereiro
      var isleap = (year % 4 == 0 && (year % 100 != 0 || year % 400 == 0));
    if (day>29 || (day==29 && !isleap)) {
      return (false);
    }
    }  
  }
  return (true);  // data válida
  
}

function validaTecla(event, compare)
{
  var BACKSPACE=  8;
  var key;
  var tecla;

  if(navigator.appName.indexOf("Netscape")!= -1)  
    tecla= event.which;    
  else
    tecla= event.keyCode;

  key = String.fromCharCode( tecla);

  if ( tecla == 13 )
    return false;
  if ( tecla == BACKSPACE )
    return true;
  return ( compare.indexOf(key)!=-1);
}

// Open a pop up window
function openPopup(winurl, winname, winWidth, winHeight) 
{
  // Position
  var x = (window.screen.width - winWidth)/ 2;
  var y = (window.screen.height - winHeight)/2;
  // Object window
  var win = null;
  params = "width="+winWidth+", height="+winHeight+", top="+y+",left="+x+", scrollBars=yes, resizable=yes, " +
           "toolbar=no, menubar=no, location=no, directories=no";
  win = window.open(winurl, winname, params);
  //win.moveTo(x, y);
  //win.focus();
}



=======
// JScript source code

function uncheck(form, field)
{
	for (i=0; i<document.forms[form].elements[field].length; i++)
		document.forms[form].elements[field][i].checked = false;
}

function check_cpf (numcpf)
{
x = 0;
soma = 0;
dig1 = 0;
dig2 = 0;
texto = "";
numcpf1="";
len = numcpf.length; x = len -1;
// var numcpf = "12345678909";
for (var i=0; i <= len - 3; i++) {
y = numcpf.substring(i,i+1);
soma = soma + ( y * x);
x = x - 1;
texto = texto + y;
}
dig1 = 11 - (soma % 11);
if (dig1 == 10) dig1=0 ;
if (dig1 == 11) dig1=0 ;
numcpf1 = numcpf.substring(0,len - 2) + dig1 ;
x = 11; soma=0;
for (var i=0; i <= len - 2; i++) {
soma = soma + (numcpf1.substring(i,i+1) * x);
x = x - 1;
}
dig2= 11 - (soma % 11);
if (dig2 == 10) dig2=0;
if (dig2 == 11) dig2=0;
//alert ("Digito Verificador : " + dig1 + "" + dig2);
if ((dig1 + "" + dig2) == numcpf.substring(len,len-2)) {
return true;
}
return false;
}

function valida_email(email) {

  if ((email.indexOf('@') == email.lastIndexOf('@')) &&
      (email.indexOf('@') > 0) &&
      (email.lastIndexOf('.') > (email.indexOf('@') + 1)) &&
      (email.lastIndexOf('.') + 1 != email.length)) return true
  else return false
}

function consisteEmail (email) {
  //O pattern abaixo reflete as seguintes regras:
  //- no máximo 25 caracteres (alfanumérico, hífen, ponto) possíveis antes da @, sendo q o último
  //  deve ser um caracter alfanumérico
  //- cada parte do domínio pode ter no máximo 30 caracteres (alfanumérico, hífen), último não 
  //  pode ser hífen
  //- domínio pode ter até 4 partes separadas por ponto, sendo que a ultima tem entre 2 e 3
  //  caracteres (alfanumérico), geralmente com, net, br, etc
  pattern  = /(^[a-z0-9.-]{1,24}[a-z0-9]@([a-z0-9-]{1,29}[a-z0-9].){1,3}[a-z0-9]{2,3}$)/;
  
  return pattern.test(email);

}

function formataData(campo){
//Formata o campo de data de acordo com a tecla pressionada
  
  var pos = 2;
  var pos1 = 4;
  var tecla = event.keyCode;
  var tam = 0;
  var vrout = "";

  vr = ""
  for (i=0 ;i <= campo.value.length; i++) {
    c = campo.value.substr(i,1);
  if (c >= "0" && c <= "9") {
    vr += c;
  }
  }
  tam = vr.length ;

  if (tecla == 8 ){ 
  tam = tam - 1 ; 
  }
  
  if ( tecla == 9 || tecla == 8 || tecla >= 48 && tecla <= 57 || tecla >= 96 && tecla <= 105){

  vrout = "" ;
  for (i=0 ;i <= tam ; i++) {
  
  c = vr.substr(i,1);

  switch (i){  
    case pos : 
    case pos1 :      
      vrout += "/";      
      break;  
    } // switch
  vrout += c;
  } // for
  
  campo.value = vrout;
  }
    
  else {
    return(false);
  }

  return(true);
  
}

//testa se caractere é numérico
function isNum(caractere)
{
  var strValidos = "0123456789."
  if ( strValidos.indexOf( caractere ) == -1 )
    return false;
  return true;
}
function isNum2(caractere)
{
  var strValidos = "0123456789"
  if ( strValidos.indexOf( caractere ) == -1 )
    return false;
  return true;
}
  
//retorna verdadeiro se tecla em event é númerica ou backspace
//retorna falso caso contrário
function validaTeclaNumerica(event)
{
  var BACKSPACE=  8;
  var key;
  var tecla;

  if(navigator.appName.indexOf("Netscape")!= -1)  
    tecla= event.which;    
  else
    tecla= event.keyCode;

  key = String.fromCharCode( tecla);

  if ( tecla == 13 )
    return false;
  if ( tecla == BACKSPACE )
    return true;
  return ( isNum(key));

} 
function validaTeclaNumerica2(event)
{
  var BACKSPACE=  8;
  var key;
  var tecla;

  if(navigator.appName.indexOf("Netscape")!= -1)  
    tecla= event.which;    
  else
    tecla= event.keyCode;

  key = String.fromCharCode( tecla);

  if ( tecla == 13 )
    return false;
  if ( tecla == BACKSPACE )
    return true;
  return ( isNum2(key));

} 

function validaData ( data ) {
//Verifica se a data é válida (data em português) 
//no formato DD/MM/YY ou DD/MM/YYYY

  if (data == "" ) {
    return (false) ;
  }  
  else {
  
    var datePat = /^(\d{1,2})(\/|-)(\d{2})\2(\d{4})$/i;

    var matchArray = data.match(datePat); //verifica o formato
    if (matchArray == null) {
      return (false);
    }
  
    day = matchArray[1]; // divide a data em variáveis
    month = matchArray[3];
    year = matchArray[4];

    if (month < 1 || month > 12) { // verifica o intervalo de meses
      return (false);
    }
    if (day < 1 || day > 31) {
      return (false);
    }
    if ((month==4 || month==6 || month==9 || month==11) && day==31) {
      return (false);
    }
    if (month == 2) { //verifica data de 29 de fevereiro
      var isleap = (year % 4 == 0 && (year % 100 != 0 || year % 400 == 0));
    if (day>29 || (day==29 && !isleap)) {
      return (false);
    }
    }  
  }
  return (true);  // data válida
  
}

function validaTecla(event, compare)
{
  var BACKSPACE=  8;
  var key;
  var tecla;

  if(navigator.appName.indexOf("Netscape")!= -1)  
    tecla= event.which;    
  else
    tecla= event.keyCode;

  key = String.fromCharCode( tecla);

  if ( tecla == 13 )
    return false;
  if ( tecla == BACKSPACE )
    return true;
  return ( compare.indexOf(key)!=-1);
}

// Open a pop up window
function openPopup(winurl, winname, winWidth, winHeight) 
{
  // Position
  var x = (window.screen.width - winWidth)/ 2;
  var y = (window.screen.height - winHeight)/2;
  // Object window
  var win = null;
  params = "width="+winWidth+", height="+winHeight+", top="+y+",left="+x+", scrollBars=yes, resizable=yes, " +
           "toolbar=no, menubar=no, location=no, directories=no";
  win = window.open(winurl, winname, params);
  //win.moveTo(x, y);
  //win.focus();
}



>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
