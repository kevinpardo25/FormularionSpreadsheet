Main.gs
============================


//funcion que llama al HTML
function doGet(e){
 return HtmlService.createHtmlOutputFromFile('Index');
  
}

function writeForm(form) {
  try {  
    var distinguishedName = form.distinguishedName; //these match to the named fields in your form mismo nombre del formulario html
    var moreDetails = form.moreDetails;
    var cedula=form.cedula;
 
    var ss = SpreadsheetApp.openById('1gJol802ZIiAxEk7w5WmQ8XhmPlRqQxA9lFVur8UXCAc'); //Id de la tabla de excel
    var sheet = ss.getSheetByName('Datos');
    var newRow = sheet.getLastRow()+1;//va a la ultima fila          
     
    //escribe en el excel como tal
    var range = sheet.getRange(newRow, 1);    
    range.setValue(distinguishedName);
    
    range = sheet.getRange(newRow, 2);
    range.setValue(moreDetails);
    
    range=sheet.getRange(newRow,3);
    range.setValue(cedula);
     
   //Mensaje de confirmacion por html
    var confirmationMessage = ['<h2>Registro exitoso!</h2>', 
                               
                              ];
    //Se envia el correo funcionando y probado  
    mensaje= "se ha registrado el usuario"+distinguishedName
    MailApp.sendEmail("kevinpardo25@gmail.com", "Correo automatizado", mensaje);                              
    var len = confirmationMessage.length-1;
    Logger.log('len= ' + len);
    var i = Math.floor(Math.random() * len);//randomizes from the array
     
    return confirmationMessage[i]; //displays randomized message
  } catch (error) {
     
    return error.toString();
  }
}





// funcion que guarda los datos en la tabla de excel
function procesaFormDatosPersona (e){
var sNombre=e.nombre;
var sApellido=e.apellido;
var sCedula=e.cedula;
var sSexo=e.sexo;
var sCelular=e.celular;
var sDireccion=e.direccion;
  
//se indica en que hoja de calculo se va a almacenar por su id
var hojaCalculo= SpreadsheetApp.openById("1gJol802ZIiAxEk7w5WmQ8XhmPlRqQxA9lFVur8UXCAc");
  //se indica en que hoja de calculo se guardara en este caso en la hoja 1 llamada sheet1
var hojaDatos= hojaCalculo.getSheetByName('Datos');
var ultimaFila= hojaDatos.getLastRow();

hojaDatos.getRange(ultimaFila+1, 1).setValue(sNombre);
hojaDatos.getRange(ultimaFila+1, 2).setValue(sApellido);
hojaDatos.getRange(ultimaFila+1, 3).setValue(sCedula);
hojaDatos.getRange(ultimaFila+1, 4).setValue(sSexo);
hojaDatos.getRange(ultimaFila+1, 5).setValue(sCelular);
hojaDatos.getRange(ultimaFila+1, 6).setValue(sDireccion);
}
//fin de la funcion
// funcion que devuelve los nombres (fila1)
function getNombreSS(e){
  var sId=e.id;
  return buscaEnSheet(sId,1);
}

//funcion que devuelve por apellido(fila2)
function getApellidoSS(e){
 var sId=e.id;
  return buscaEnSheet(sId,2)
  
}



//funcion que busca en sheet como una BD
function buscaEnSheet(sId, comlumna){
  var hojaCalculo=SpreadsheetApp.openById("1gJol802ZIiAxEk7w5WmQ8XhmPlRqQxA9lFVur8UXCAc");
  var hojaDatos= hojaCalculo.getSheetByName('Datos');
  
  var numColumns= hojaDatos.getLastColumn();
  var ultimaFila= hojaDatos.getLastRow();
  var sw=0;
  
  var row= hojaDatos.getRange(1, 1, ultimaFila, numColumns).getValues();
  
  for (var i=1; i<row.length; i++){
    for (var col=0; col<row[i].length; col++){
      var id= row[i][2].toString();
      if(sId==id){
       var indice= i+1;
        sw=1;
      }
  }
  }
  if(sw==1){
   var info= hojaDatos.getRange(indice, columna).getValue();
    return info;
    
  }
}
  
  
  
  

