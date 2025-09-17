# Funciones_excel
Automatizar excel

funcion 2

function main(workbook: ExcelScript.Workbook) {
  // Categorías de errores
  const type_of_errors: string[] = [
    "BCP",
    "Sistema",
    "Error",
    "eRRor"
  ];

  const CELDA = "B2"; // Celda de ingreso y salida de datos
  const hoja = workbook.getActiveWorksheet();
  const rango = hoja.getRange(CELDA);

  // Construir mensaje de ayuda para validación de datos
  const errorMessage = type_of_errors
    .map((err, idx) => `${idx + 1}. ${err}`)
    .join("\n");

  // Configurar validación de datos con mensaje de entrada
  rango.getDataValidation().setPrompt({
    showPrompt: true,
    title: "Escribe los números de errores separados por comas. Ej. 1,2,3",
    message: errorMessage
  });

  // Obtener entrada del usuario
  const input = rango.getValue().toString().trim();
  if (!input) {
    return; // Si está vacío, no hacemos nada
  }

  // Separar entrada en números
  const tokens = input.split(",").map(x => x.trim());
  const validIndices: number[] = [];
  const invalidTokens: string[] = [];

  for (const t of tokens) {
    const idx = parseInt(t, 10) - 1; // Convertir a índice
    if (!isNaN(idx) && idx >= 0 && idx < type_of_errors.length) {
      validIndices.push(idx);
    } else {
      invalidTokens.push(t); // Guardar los inválidos
    }
  }

  // Construir salida
  let output: string;
  if (invalidTokens.length > 0) {
    output = `⚠️ Valores inválidos: ${invalidTokens.join(", ")}`;
  } else {
    output = validIndices.map(i => type_of_errors[i]).join(", ");
  }

  // Sobrescribir celda
  rango.setValue(output);
}


funcion 1

function main(workbook: ExcelScript.Workbook) {
  // Categorias de errores
  let type_of_errors: string[] = [
    "BCP",
    "Sistema",
    "Error",
    "eRRor"
  ]
  let CELDA = "B2"; // Celda de ingreso y salida de datos
  let hoja = workbook.getActiveWorksheet();
  let rango = hoja.getRange(CELDA);
  let arreglo: string[] = [];
  let inputNumbers: number[] = [];
  let input: string;
  let output: string = "";
  let errorMesage: strig = "";
  //Iteradores
  let c: string = 'a';
  let i: number;

  // Creando string para las opciones de errores en las instrucciones
  for(i = 0; i < type_of_errors.length; i++){
    errorMesage += `${i+1}. ${type_of_errors[i]}\n`
  }

  // Configurar validación de datos con mensaje de entrada
  let validacion = rango.getDataValidation();
  validacion.setPrompt({
    showPrompt: true,
    title: "Escribe los números de errores separados por comas. Ej. 1,2,3",
    message: errorMesage
  });

  // Escaneando los datos de entrada
  input = rango.getValue().toString();
  arreglo = input.split(",");

  for(c of arreglo){
    inputNumbers.push(parseInt(c)-1);
  }

  for(i of inputNumbers){
    output += type_of_errors[i] + "  ";
  }

  hoja.getRange(CELDA).setValue(output);
}
