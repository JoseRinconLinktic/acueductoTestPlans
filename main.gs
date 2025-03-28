//Version 28/03/2025
var personalAccessToken = "AOJJZijWT6k27LMzQk3oZXdJvjfpOV8zHngfMpxy2CAEzQCnFW6IJQQJ99BBACAAAAAL75KlAAASAZDO2F4l";
var sheets_url = "https://docs.google.com/spreadsheets/d/1JvZmEcFZjWtah81L_AHPsGUpYUgBSVq4dh891NN1sl0/edit?gid=0#gid=0";
//var personalAccessToken = "A9pcpSOWQ5G8znF9u80WmJNb38kQeP8y2A4mwQ2TxTL7XwowveIvJQQJ99BCACAAAAAL75KlAAASAZDO4cT8";
var page_name = "CORR Y TRAV"; //Aqui se le pone nombre a la hoja
var data_sheets = SpreadsheetApp.openByUrl(sheets_url);
var page = data_sheets.getSheetByName(page_name);

var suite_consultar;

var suites_consultar = [199036,200056, 200103,199019, 200054,200101]; //Aqui se crean los archivos

function crear_varios(){
  for(i=0;i<suites_consultar.length;i++){
    suite_consultar = suites_consultar[i];
    Logger.log(`Consultando ${suite_consultar}...`);
    obtener_test_plans()
  }
  escribirEnHoja();
}

if (page) {
  // Si la hoja existe, limpia su contenido
  Logger.log(`La hoja "${page_name}" ya existe. Limpiando contenido...`);
  page.clear();
} else {
  // Si la hoja no existe, la crea
  Logger.log(`La hoja "${page_name}" no existe. Creando nueva hoja...`);
  page = data_sheets.insertSheet(page_name);
}


var nuevos_encabezados = [
  "ID Test Plan", 
  "Nombre Test Plan", 
  "ID de Test Suite", 
  "Nombre Test Suite", 
  "ID de Test Case", 
  "Nombre Test Case", 
  "Estado Test Case", 
  "Fecha de modificación Estado Test Case",
  "Caso Corrido por",
  "URL de Test Case"
];
page.getRange(1, 1, 1, nuevos_encabezados.length).setValues([nuevos_encabezados]);

function solicitud_https(url) {
  var maxRetries = 5; // Número máximo de reintentos
  var retryDelay = 100; // Tiempo de espera entre reintentos en milisegundos
  var attempts = 0;
  var response;

  while (attempts < maxRetries) {
    attempts++;
    try {
      var options = {
        "method" : "get",
        "headers" : {
          "Content-Type" : "application/json",
          "Authorization" : "Basic " + Utilities.base64Encode(':' + personalAccessToken)
        },
        "muteHttpExceptions": true  
      };

      response = UrlFetchApp.fetch(url, options);

      if (response.getResponseCode() == 200) {
        var jsonResponse = JSON.parse(response.getContentText());
        return jsonResponse;
      } else {
        Logger.log("Intento " + attempts + " fallido. Código de respuesta: " + response.getResponseCode());
      }
    } catch (e) {
      Logger.log("Error en el intento " + attempts + ": " + e.message);
    }

    // Espera antes de reintentar
    Utilities.sleep(retryDelay);
  }

  // Si se alcanzan los reintentos máximos sin éxito
  Logger.log("No se pudo acceder a la URL después de " + maxRetries + " intentos.");
  return 0;
}

function obtener_test_plans() {
  //var url = "https://dev.azure.com/Linktic/384%20-%20CORE%20FIDUPREVISORA/_apis/test/plans?api-version=5.0";
  var url = "https://dev.azure.com/Linktic/348%20-%20SGDEA%20ACUEDUCTO/_apis/test/plans?api-version=5.0";
  var test_plans_json = solicitud_https(url);
  
  if(test_plans_json != 0){

    var plans_quantity = test_plans_json.count;

    for( var i = 0; i < plans_quantity; i++){
      var plan = test_plans_json.value[i];
      var plan_id = plan.id;
      var plan_title = plan.name;
      var plan_url = plan.url;

      obtener_test_suite(plan_id,plan_title,plan_url)
    }
    
  } else{

    Logger.log("No se encontro nada")

  }
}

function obtener_test_suite(plan_id, plan_title, plan_url){

  //var url = `https://dev.azure.com/Linktic/384%20-%20CORE%20FIDUPREVISORA/_apis/test/plans/${plan_id}/suites?api-version=5.0`;
  var url = `https://dev.azure.com/Linktic/348%20-%20SGDEA%20ACUEDUCTO/_apis/test/plans/${plan_id}/suites?api-version=5.0`;
  var test_suites_json = solicitud_https(url);

  var suite_quantity = test_suites_json.count;

  //Logger.log("Encontro una Suite");

  for( var i = 0; i < suite_quantity; i++){
      var suite = test_suites_json.value[i];
      var suite_id = suite.id;
      var suite_title = suite.name;
      var suite_url = suite.url;

      try{
        var suite_parent = suite.parent["id"];
        if(suite_parent == suite_consultar){
          obtener_test_cases(plan_id, plan_title, plan_url, suite_id,suite_title,suite_url);
        }
      }

      catch (e){
        Logger.log(e);
      }
    }

}

function obtener_test_cases(plan_id, plan_title, plan_url,suite_id,suite_title,suite_url){
  //var url = `https://dev.azure.com/Linktic/384%20-%20CORE%20FIDUPREVISORA/_apis/test/Plans/${plan_id}/Suites/${suite_id}/Points?api-version=5.0`;
  var url = `https://dev.azure.com/Linktic/348%20-%20SGDEA%20ACUEDUCTO/_apis/test/Plans/${plan_id}/Suites/${suite_id}/Points?api-version=5.0`;

  var test_cases_json = solicitud_https(url);

  var case_quantity = test_cases_json.count;

  for (var i = 0; i < case_quantity; i++){
    var cases = test_cases_json.value[i];
    var case_id = cases.testCase.id;
    var case_url = `https://dev.azure.com/Linktic/4fe5ab76-a2ad-4a1f-aeea-6f9ab718b878/_workitems/edit/${case_id}`;
    var case_outcome = cases.outcome;
    var run_id = cases.lastTestRun["id"];
    var run_by = cases.lastResultDetails.runBy["displayName"];

    //var url_wi = `https://dev.azure.com/Linktic/384%20-%20CORE%20FIDUPREVISORA/_apis/wit/workitems/${case_id}?api-version=7.1-preview`;
    var url_wi = `https://dev.azure.com/Linktic/348%20-%20SGDEA%20ACUEDUCTO/_apis/wit/workitems/${case_id}?api-version=5.0`;

    var url_run = `https://dev.azure.com/Linktic/348%20-%20SGDEA%20ACUEDUCTO/_apis/test/runs/${run_id}/results?api-version=5.0`;

    var case_wi = solicitud_https(url_wi);

    if(run_id == 0){
      var case_timestamp = case_wi.fields['Microsoft.VSTS.Common.StateChangeDate'];
    }else{
      var case_run = solicitud_https(url_run);
      var case_timestamp = case_run.value[0].completedDate;
    }
    var case_title =  case_wi.fields['System.Title'];

    var dateObj = new Date(case_timestamp);

    var options = { year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit', second: '2-digit'};
    var formattedDate = dateObj.toLocaleDateString('es-CO', options);

    var datos = {};

    datos.plan_id = plan_id;
    datos.plan_title = plan_title;
    datos.plan_url = plan_url;
    datos.suite_id = suite_id;
    datos.suite_title = suite_title;
    datos.suite_url = suite_url;
    datos.case_id = case_id;
    datos.case_outcome = case_outcome;
    datos.case_url = case_url;
    datos.timestamp = formattedDate;
    datos.case_title = case_title;
    datos.run_by = run_by;
    subir_sheets(datos);
  }
}


function subir_sheets(datos) {
  var maxRetries = 3; // Número máximo de reintentos
  var retryDelay = 2000; // Tiempo de espera entre reintentos en milisegundos
  var attempts = 0;

  // Lógica de reintentos
  while (attempts < maxRetries) {
    try {

      // Define los datos que se agregarán en una nueva fila
      var fila = [
        datos.plan_id,
        datos.plan_title,
        datos.suite_id,
        datos.suite_title,
        datos.case_id,
        datos.case_title,
        datos.case_outcome,
        datos.timestamp,
        datos.run_by,
        datos.case_url
      ];

      // Agrega la fila al final de la hoja
      page.appendRow(fila);

      //Logger.log("Fila agregada correctamente.");
      break; // Si no hubo error, salimos del ciclo de reintentos
    } catch (e) {
      // Si ocurre un error, registrar el mensaje de error y esperar antes del siguiente intento
      attempts++;
      Logger.log('Error al agregar fila: ' + e.message);

      if (attempts >= maxRetries) {
        //Logger.log('Número máximo de reintentos alcanzado. Error: ' + e.message);
        throw e; // Si se alcanzó el máximo de reintentos, lanzar el error
      }

      // Si no hemos alcanzado el máximo de intentos, esperamos antes de intentar de nuevo
      var delay = Math.min(retryDelay * Math.pow(2, attempts), 10000); // Limitar el retraso máximo a 10 segundos
      //Logger.log('Reintentando en ' + delay / 1000 + ' segundos...');
      Utilities.sleep(delay); // Espera antes de reintentar
    }
  }
}

function escribirEnHoja() {
  page.getRange("K3").setValue("Fecha de ultima actualización");
  
  // Obtener la fecha de ejecución actual
  var fechaHora = new Date();
  var options = { year: 'numeric', month: 'numeric', day: 'numeric', hour: '2-digit', minute: '2-digit', second: '2-digit' };
  var formattedDate = fechaHora.toLocaleDateString('es-CO', options); // Formato en español
  
  // Escribir la fecha de ejecución en la celda K4
  page.getRange("K4").setValue(formattedDate);
  
  Logger.log("Fecha de última actualización guardada en la celda K4.");
}
