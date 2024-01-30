document.getElementById('generar').addEventListener('click', mostrarDatos);
document.getElementById('generarExcel').addEventListener('click', generarExcel);
document.getElementById('btn_plantilla').addEventListener('click', descargarModeloPlantilla);

var selectedSheet;
var columnData;
var nuevoLibro = XLSX.utils.book_new();


function mostrarDatos() {
    // Obtén el tipo de anteojos seleccionado
    var tipoAnteojo = document.getElementById('tipoAnteojo').value;

    // Obtén la acción seleccionada 
    var accion = document.getElementById('accion').value;

    // Obtén el elemento de entrada de archivo
    var input = document.getElementById('fileInput');

    // Verifica si se seleccionó un archivo
    if (input.files.length > 0) {
        var file = input.files[0];

        // Creo un FileReader para leer el archivo
        var reader = new FileReader();

        reader.onload = function (e) {
            var data = e.target.result;

            // Usa SheetJS para analizar el archivo
            var workbook = XLSX.read(data, { type: 'binary' });

            // Selecciona la hoja según el tipo de anteojos
            var sheetIndex = (tipoAnteojo === 'receta') ? 0 : 1;

            // Obtiene la hoja correspondiente
            var selectedSheetName = workbook.SheetNames[sheetIndex];
            selectedSheet = workbook.Sheets[selectedSheetName];

            // Obtiene los datos según la acción seleccionada
            columnData = [];
            var range = XLSX.utils.decode_range(selectedSheet['!ref']);

            //Manejo la accion de acuerdo con el tipo de anteojo
            if (tipoAnteojo === 'receta') {

                switch (accion) {
                    case ('sku'):
                        columnData = generarSKU(range);
                        break;
                    case ('nombres'):
                        columnData = generarNombresReceta(range);
                        break;
                    case ('descripcion'):
                        var columnData = generarDescripcionReceta(range);                        
                        break;
                    case ('personalizado'):
                        // Lógica personalizada aquí
                        break;
                }

            } else {
                switch (accion) {
                    case ('sku'):
                        columnData = generarSKU(range);
                        break;
                    case ('nombres'):
                        columnData = generarNombresSol(range);
                        break;
                    case ('descripcion'):
                        columnData = generarDescripcionSol(range);
                        break;
                    case ('personalizado'):
                        // Lógica personalizada aquí
                        break;
                }
            }

            // Muestra los datos en el textarea
            document.getElementById('resultTextarea').value = columnData.join('\n');
        };

        // Lee el contenido del archivo como binario
        reader.readAsBinaryString(file);
    } else {
        // Mostrar mensaje de copiado
        var mensajeError = document.getElementById("errorArchivo");
        mensajeError.style.display = "block";

        setTimeout(function () {
            mensajeError.style.display = "none";
        }, 2000); // El mensaje se ocultará después de 2 segundos (puedes ajustar el tiempo según tus necesidades)
    }
}

function generarSKU(range) {
    var resultado = [];
    for (var i = range.s.r + 1; i <= range.e.r; ++i) {
        var cellAddressA = XLSX.utils.encode_cell({ r: i, c: 0 });
        var cellValue = selectedSheet[cellAddressA] ? selectedSheet[cellAddressA].v : undefined;
        resultado.push(cellValue);
    }
    return resultado;

}

function generarNombresReceta(range) {
    var resultado = [];
    for (var i = range.s.r + 1; i <= range.e.r; i++) {
        var cellAddressB = XLSX.utils.encode_cell({ r: i, c: 1 });
        var cellAddressH = XLSX.utils.encode_cell({ r: i, c: 7 });
        var cellAddressI = XLSX.utils.encode_cell({ r: i, c: 8 });
        var cellAddressJ = XLSX.utils.encode_cell({ r: i, c: 9 });
        var cellAddressN = XLSX.utils.encode_cell({ r: i, c: 13 });
        var cellAddressQ = XLSX.utils.encode_cell({ r: i, c: 16 });

        var cellValueB = selectedSheet[cellAddressB] ? selectedSheet[cellAddressB].v : '';
        var cellValueH = selectedSheet[cellAddressH] ? selectedSheet[cellAddressH].v : '';
        var cellValueI = selectedSheet[cellAddressI] ? selectedSheet[cellAddressI].v : '';
        var cellValueJ = selectedSheet[cellAddressJ] ? selectedSheet[cellAddressJ].v : '';
        var cellValueN = selectedSheet[cellAddressN] ? selectedSheet[cellAddressN].v : '';
        var cellValueQ = selectedSheet[cellAddressQ] ? selectedSheet[cellAddressQ].v : '';

        // Verificar cada valor antes de incluirlo en nombreCompleto, si es vacío lo omite
        var nombreCompleto = '';
        if (cellValueB) nombreCompleto += `${cellValueB.trim()} `; // trim() para eliminar espacios al final
        if (cellValueI) nombreCompleto += `${cellValueI.trim()} `;
        if (cellValueH) nombreCompleto += `${cellValueH.trim()} `;
        if (cellValueJ) nombreCompleto += `color ${cellValueJ} `;
        if (cellValueN) nombreCompleto += `cal ${cellValueN} `; //es un num por eso no hace falta cortarlo
        if (cellValueQ) nombreCompleto += `- ${cellValueQ.trim()}`;

        resultado.push(nombreCompleto.trim());
    }
    return resultado;
}

function generarNombresSol(range) {
    var resultado = [];
    for (var i = range.s.r + 1; i <= range.e.r; i++) {
        var cellAddressC = XLSX.utils.encode_cell({ r: i, c: 2 });
        var cellAddressK = XLSX.utils.encode_cell({ r: i, c: 10 });
        var cellAddressJ = XLSX.utils.encode_cell({ r: i, c: 9 });
        var cellAddressL = XLSX.utils.encode_cell({ r: i, c: 11 });
        var cellAddressX = XLSX.utils.encode_cell({ r: i, c: 23 });

        var cellValueC = selectedSheet[cellAddressC] ? selectedSheet[cellAddressC].v : '';
        var cellValueK = selectedSheet[cellAddressK] ? selectedSheet[cellAddressK].v : '';
        var cellValueJ = selectedSheet[cellAddressJ] ? selectedSheet[cellAddressJ].v : '';
        var cellValueL = selectedSheet[cellAddressL] ? selectedSheet[cellAddressL].v : '';
        var cellValueX = selectedSheet[cellAddressX] ? selectedSheet[cellAddressX].v : '';

        // Verificar cada valor antes de incluirlo en nombreCompleto, si es vacío lo omite
        var nombreCompleto = '';
        if (cellValueC) nombreCompleto += `${cellValueC.trim()} `; // trim() para eliminar espacios al final
        if (cellValueK) nombreCompleto += `${cellValueK.trim()} `;
        if (cellValueJ) nombreCompleto += `${cellValueJ.trim()} `;
        if (cellValueL) nombreCompleto += `color ${cellValueL.trim()} `;
        if (cellValueX) nombreCompleto += `- ${cellValueX.trim()}`;

        resultado.push(nombreCompleto.trim());
    }
    return resultado;
}

function generarDescripcionSol(range) {
    var resultadoDescripcionSol = [];
    for (var i = range.s.r + 1; i <= range.e.r; i++) {
        var cellAddressC = XLSX.utils.encode_cell({ r: i, c: 2 });
        var cellAddressD = XLSX.utils.encode_cell({ r: i, c: 3 });
        var cellAddressJ = XLSX.utils.encode_cell({ r: i, c: 9 });
        var cellAddressK = XLSX.utils.encode_cell({ r: i, c: 10 });
        var cellAddressL = XLSX.utils.encode_cell({ r: i, c: 11 });
        var cellAddressM = XLSX.utils.encode_cell({ r: i, c: 12 });
        var cellAddressN = XLSX.utils.encode_cell({ r: i, c: 13 });
        var cellAddressO = XLSX.utils.encode_cell({ r: i, c: 14 });
        var cellAddressU = XLSX.utils.encode_cell({ r: i, c: 20 });
        var cellAddressV = XLSX.utils.encode_cell({ r: i, c: 21 });
        var cellAddressW = XLSX.utils.encode_cell({ r: i, c: 22 });
        var cellAddressX = XLSX.utils.encode_cell({ r: i, c: 23 });
        var cellAddressY = XLSX.utils.encode_cell({ r: i, c: 24 });
        var cellAddressZ = XLSX.utils.encode_cell({ r: i, c: 25 });

        var cellValueC = selectedSheet[cellAddressC] ? selectedSheet[cellAddressC].v : '';
        var cellValueD = selectedSheet[cellAddressD] ? selectedSheet[cellAddressD].v : '';
        var cellValueJ = selectedSheet[cellAddressJ] ? selectedSheet[cellAddressJ].v : '';
        var cellValueK = selectedSheet[cellAddressK] ? selectedSheet[cellAddressK].v : '';
        var cellValueL = selectedSheet[cellAddressL] ? selectedSheet[cellAddressL].v : '';
        var cellValueM = selectedSheet[cellAddressM] ? selectedSheet[cellAddressM].v : '';
        var cellValueN = selectedSheet[cellAddressN] ? selectedSheet[cellAddressN].v : '';
        var cellValueO = selectedSheet[cellAddressO] ? selectedSheet[cellAddressO].v : '';
        var cellValueU = selectedSheet[cellAddressU] ? selectedSheet[cellAddressU].v : '';
        var cellValueV = selectedSheet[cellAddressV] ? selectedSheet[cellAddressV].v : '';
        var cellValueW = selectedSheet[cellAddressW] ? selectedSheet[cellAddressW].v : '';
        var cellValueX = selectedSheet[cellAddressX] ? selectedSheet[cellAddressX].v : '';
        var cellValueY = selectedSheet[cellAddressY] ? selectedSheet[cellAddressY].v : '';
        var cellValueZ = selectedSheet[cellAddressZ] ? selectedSheet[cellAddressZ].v : '';


        // Verificar cada valor antes de incluirlo en nombreCompleto, si es vacío lo omite
        var nombreCompleto = '';
        nombreCompleto += '<p>Anteojos de sol ';
        if (cellValueC) nombreCompleto += `${cellValueC.trim()} `; // trim() para eliminar espacios al final
        if (cellValueK) nombreCompleto += `${cellValueK.trim()} `;
        if (cellValueJ) nombreCompleto += `${cellValueJ.trim()} `;
        if (cellValueL) nombreCompleto += `color ${cellValueL.trim()} `;
        if (cellValueU) nombreCompleto += `cal ${cellValueU}. `;
        nombreCompleto += 'Original, con estuche y garantía oficial. </p>';
        nombreCompleto += '<p>';
        if (cellValueX) nombreCompleto += `Colección: ${cellValueX.trim()}. <br>`;
        if (cellValueM) nombreCompleto += `Material del armazón: ${cellValueM.trim()}. <br>`;
        if (cellValueO) nombreCompleto += `Color de la lente: ${cellValueO.trim()}. <br>`;
        if (cellValueD) nombreCompleto += `Color del frente: ${cellValueD.trim()}. <br>`;
        if (cellValueZ) nombreCompleto += `Color de patilla: ${cellValueZ.trim()}. <br>`;
        if (cellValueN) nombreCompleto += `Tipo de filtro: ${cellValueN.trim()}. <br>`;
        if (cellValueY) nombreCompleto += `Pais de origen: ${cellValueY.trim()}. <br>`;
        nombreCompleto += '</p>';
        nombreCompleto += '<p>';
        nombreCompleto += 'Medidas: <br>';
        if (cellValueU) nombreCompleto += `Diámetro de la lente: ${cellValueU}mm. <br>`;
        if (cellValueV) nombreCompleto += `Largo de puente: ${cellValueV}mm. <br>`;
        if (cellValueW) nombreCompleto += `Largo de patilla: ${cellValueW}mm. <br>`;
        nombreCompleto += '</p>'
        nombreCompleto += '<p> 1 año de Garantía por defectos de fabricación. NO cubre fallas por mal uso del producto. </p>';
        nombreCompleto += '<p>Envío gratis a todo el país. <br></p>';

        resultadoDescripcionSol.push(nombreCompleto.trim());
    }
    return resultadoDescripcionSol;
}

function generarDescripcionReceta(range) {
    var resultadoDescripcionReceta = [];
    for (var i = range.s.r + 1; i <= range.e.r; i++) {
        var cellAddressB = XLSX.utils.encode_cell({ r: i, c: 1 });
        var cellAddressH = XLSX.utils.encode_cell({ r: i, c: 7 });
        var cellAddressI = XLSX.utils.encode_cell({ r: i, c: 8 });
        var cellAddressJ = XLSX.utils.encode_cell({ r: i, c: 9 });
        var cellAddressK = XLSX.utils.encode_cell({ r: i, c: 10 });
        var cellAddressL = XLSX.utils.encode_cell({ r: i, c: 11 });
        var cellAddressM = XLSX.utils.encode_cell({ r: i, c: 12 });
        var cellAddressN = XLSX.utils.encode_cell({ r: i, c: 13 });
        var cellAddressO = XLSX.utils.encode_cell({ r: i, c: 14 });
        var cellAddressP = XLSX.utils.encode_cell({ r: i, c: 15 });
        var cellAddressQ = XLSX.utils.encode_cell({ r: i, c: 16 });
        var cellAddressR = XLSX.utils.encode_cell({ r: i, c: 17 });

        var cellValueB = selectedSheet[cellAddressB] ? selectedSheet[cellAddressB].v : '';
        var cellValueH = selectedSheet[cellAddressH] ? selectedSheet[cellAddressH].v : '';
        var cellValueI = selectedSheet[cellAddressI] ? selectedSheet[cellAddressI].v : '';
        var cellValueJ = selectedSheet[cellAddressJ] ? selectedSheet[cellAddressJ].v : '';
        var cellValueK = selectedSheet[cellAddressK] ? selectedSheet[cellAddressK].v : '';
        var cellValueL = selectedSheet[cellAddressL] ? selectedSheet[cellAddressL].v : '';
        var cellValueM = selectedSheet[cellAddressM] ? selectedSheet[cellAddressM].v : '';
        var cellValueN = selectedSheet[cellAddressN] ? selectedSheet[cellAddressN].v : '';
        var cellValueO = selectedSheet[cellAddressO] ? selectedSheet[cellAddressO].v : '';
        var cellValueP = selectedSheet[cellAddressP] ? selectedSheet[cellAddressP].v : '';
        var cellValueQ = selectedSheet[cellAddressQ] ? selectedSheet[cellAddressQ].v : '';
        var cellValueR = selectedSheet[cellAddressR] ? selectedSheet[cellAddressR].v : '';

        // Verificar cada valor antes de incluirlo en nombreCompleto, si es vacío lo omite
        var nombreCompleto = '';
        nombreCompleto += '<p>Armazón para anteojos ';
        if (cellValueB) nombreCompleto += `${cellValueB.trim()} `; // trim() para eliminar espacios al final
        if (cellValueI) nombreCompleto += `${cellValueI.trim()} `;
        if (cellValueH) nombreCompleto += `${cellValueH.trim()} `;
        if (cellValueJ) nombreCompleto += `Color ${cellValueJ} `;
        nombreCompleto += 'Original, con estuche y garantía oficial. </p>';
        nombreCompleto += '<p>';
        if (cellValueQ) nombreCompleto += `Colección: ${cellValueQ.trim()}. <br>`;
        if (cellValueK) nombreCompleto += `Material del armazón: ${cellValueK.trim()}. <br>`;
        if (cellValueL) nombreCompleto += `Color del frente: ${cellValueL.trim()}. <br>`;
        if (cellValueM) nombreCompleto += `Color de patilla: ${cellValueM.trim()}. <br>`;
        if (cellValueR) nombreCompleto += `Pais de origen: ${cellValueR.trim()}. <br>`;
        nombreCompleto += '</p>';
        nombreCompleto += '<p>';
        nombreCompleto += 'Medidas: <br>';
        if (cellValueN) nombreCompleto += `Diámetro de la lente: ${cellValueN}mm. <br>`;
        if (cellValueO) nombreCompleto += `Largo de puente: ${cellValueO}mm. <br>`;
        if (cellValueP) nombreCompleto += `Largo de patilla: ${cellValueP}mm. <br>`;
        nombreCompleto += '</p>'
        nombreCompleto += '<p> 1 año de Garantía por defectos de fabricación. NO cubre fallas por mal uso del producto. </p>';
        nombreCompleto += '<p>Envío gratis a todo el país. <br></p>';

        resultadoDescripcionReceta.push(nombreCompleto.trim());
    }
    return resultadoDescripcionReceta;
}



function copiarAlPortapapeles() {
    // Seleccionar el elemento de texto
    var textarea = document.getElementById("resultTextarea");

    // Seleccionar el texto del elemento
    textarea.select();
    textarea.setSelectionRange(0, 99999); // Para dispositivos móviles

    // Copiar el texto al portapapeles
    document.execCommand("copy");

    // Mostrar mensaje de copiado
    var mensajeCopiado = document.getElementById("mensajeCopiado");
    mensajeCopiado.style.display = "block";

    setTimeout(function () {
        mensajeCopiado.style.display = "none";
    }, 2000); // El mensaje se ocultará después de 2 segundos (puedes ajustar el tiempo según tus necesidades)
}

function generarExcel() {
    var tipoAnteojo = document.getElementById('tipoAnteojo').value;
    if (tipoAnteojo === 'receta') {
        generarExcelReceta();
    } else {
        generarExcelSol();
    }

}

function generarExcelSol() {
    //Cargo todos los datos que necesito
    var input = document.getElementById('fileInput');
    if (input.files.length > 0) {
        var file = input.files[0];
        var reader = new FileReader();
        reader.onload = function (e) {
            var data = e.target.result;
            var workbook = XLSX.read(data, { type: 'binary' });
            var sheetIndex = 1;
            var selectedSheetName = workbook.SheetNames[sheetIndex];
            selectedSheet = workbook.Sheets[selectedSheetName];
            var range = XLSX.utils.decode_range(selectedSheet['!ref']);

            // Crear un nuevo libro de Excel
            var nuevoLibro = XLSX.utils.book_new();

            // Crear una nueva hoja de cálculo con los encabezados de columnas
            var nuevaHoja = XLSX.utils.aoa_to_sheet([['SKU', 'Nombre de producto', 'Precio', 'Stock', 'Categoría', 'Descripción', 'Marca', 'Color Lente', 'Polarizado', 'Espejado', 'Apto Graduable', 'Color Armazón', 'Forma', 'Cat', 'Calibre']]);


            // ... Código para generar las columnas ...

            //COLUMNA A: SKU
            // Generar SKU y cargar la columna A desde la celda A2
            var columnDataSKU = generarSKU(range);
            columnDataSKU.forEach((value, index) => {
                // Agregar la fila a la hoja de cálculo para la columna A
                XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 0 }, raw: false });
            });


            //COLUMNA B: Nombre
            // Generar nombres de receta y cargar la columna B desde la celda B2
            var columnDataNombre = generarNombresSol(range);
            columnDataNombre.forEach((value, index) => {
                // Agregar la fila a la hoja de cálculo para la columna B
                XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 1 }, raw: false });
            });


            //COLUMNA C Y D VACIAS

            //COLUMNA E: Categoría
            for (var i = 1; i < columnDataSKU.length; i++) {
                // Construir la referencia de la celda en la columna E
                var cellRef = 'E' + (i + 1);

                // Agregar "lentes de sol" en la celda especificada
                nuevaHoja[cellRef] = { t: 's', v: 'lentes de sol' };
            }


            // //COLUMNA F: Descripción
            var columnDataDes = generarDescripcionSol(range);
            columnDataDes.forEach((value, index) => {
                XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 5 }, raw: false });
            });


            //COLUMNA G: Marca
            var datosColumna = agregarUnaColumna(range, 2);
            datosColumna.forEach((value, index) => {
                XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 6 }, raw: false });
            });

            //COLUMNA H: Color lente

            var datosColumna = agregarUnaColumna(range, 15);
            datosColumna.forEach((value, index) => {
                XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 7 }, raw: false });
            });


            //COLUMNA I: Polarizado
            var datosColumna = agregarUnaColumna(range, 16);
            datosColumna.forEach((value, index) => {
                XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 8 }, raw: false });
            });

            //COLUMNA J: Espejado
            var datosColumna = agregarUnaColumna(range, 17);
            datosColumna.forEach((value, index) => {
                XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 9 }, raw: false });
            });

            //COLUMNA K: Apto Graduable
            var datosColumna = agregarUnaColumna(range, 18);
            datosColumna.forEach((value, index) => {
                XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 10 }, raw: false });
            });

            //COLUMNA L: Color Armazón
            var datosColumna = agregarUnaColumna(range, 4);
            datosColumna.forEach((value, index) => {
                XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 11 }, raw: false });
            });

            //COLUMNA M: Forma
            var datosColumna = agregarUnaColumna(range, 6);
            datosColumna.forEach((value, index) => {
                XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 12 }, raw: false });
            });

            //COLUMNA N: Categoría
            var datosColumna = agregarUnaColumna(range, 8);
            datosColumna.forEach((value, index) => {
                XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 13 }, raw: false });
            });

            //COLUMNA O: Calibre
            var datosColumna = agregarUnaColumna(range, 20);
            datosColumna.forEach((value, index) => {
                XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 14 }, raw: false });
            });
            // ... Fin de generación de columnas ...          





            // Añadir la nueva hoja al nuevo libro de Excel
            XLSX.utils.book_append_sheet(nuevoLibro, nuevaHoja, 'de Sol');

            // Crear un blob con el nuevo libro, primero en binary y luego convertido con la funcion s2ab
            var blob = XLSX.write(nuevoLibro, { bookType: 'xlsx', type: 'binary' });


            // Descargar el nuevo archivo Excel
            saveAs(new Blob([s2ab(blob)], { type: "application/octet-stream" }), 'AnteojosDeSol.xlsx');
        }
        // Lee el contenido del archivo como binario
        reader.readAsBinaryString(file);
    } else {
        // Mostrar mensaje de copiado
        var mensajeError = document.getElementById("errorArchivo");
        mensajeError.style.display = "block";

        setTimeout(function () {
            mensajeError.style.display = "none";
        }, 2000); // El mensaje se ocultará después de 2 segundos (puedes ajustar el tiempo según tus necesidades)
    }


}

function generarExcelReceta() {
    //Cargo todos los datos que necesito

    var input = document.getElementById('fileInput');
    if (input.files.length > 0) {
        var file = input.files[0];
        var reader = new FileReader();
        reader.onload = function (e) {
            var data = e.target.result;
            var workbook = XLSX.read(data, { type: 'binary' });
            var sheetIndex = 0;
            var selectedSheetName = workbook.SheetNames[sheetIndex];
            selectedSheet = workbook.Sheets[selectedSheetName];
            var range = XLSX.utils.decode_range(selectedSheet['!ref']);

            // Crear un nuevo libro de Excel
            var nuevoLibro = XLSX.utils.book_new();

            // Crear una nueva hoja de cálculo con los encabezados de columnas
            var nuevaHoja = XLSX.utils.aoa_to_sheet([['SKU', 'Nombre de producto', 'Precio', 'Stock', 'Categoría', 'Descripción', 'Marca', 'Color Armazón', 'Forma', 'Categoría', 'Calibre', 'Link foto']]);


            // ... Código para generar las columnas ...

            //COLUMNA A: SKU
            // Generar SKU y cargar la columna A desde la celda A2
            var columnDataSKU = generarSKU(range);
            columnDataSKU.forEach((value, index) => {
                // Agregar la fila a la hoja de cálculo para la columna A
                XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 0 }, raw: false });
            });


            //COLUMNA B: Nombre
            // Generar nombres de receta y cargar la columna B desde la celda B2
            var columnDataNombre = generarNombresReceta(range);
            columnDataNombre.forEach((value, index) => {
                // Agregar la fila a la hoja de cálculo para la columna B
                XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 1 }, raw: false });
            });


            //COLUMNA C Y D VACIAS

            //COLUMNA E: Categoría
            for (var i = 1; i < columnDataSKU.length; i++) {
                // Construir la referencia de la celda en la columna E
                var cellRef = 'E' + (i + 1);

                // Agregar "lentes de sol" en la celda especificada
                nuevaHoja[cellRef] = { t: 's', v: 'Recetado' };
            }


            // //COLUMNA F: Descripción
            var columnDataDes = generarDescripcionReceta(range);
            columnDataDes.forEach((value, index) => {
                XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 5 }, raw: false });
            });


            //COLUMNA G: Marca
            var datosColumna = agregarUnaColumna(range, 1);
            datosColumna.forEach((value, index) => {
                XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 6 }, raw: false });
            });

            //COLUMNA H: Color armazón

            var datosColumna = agregarUnaColumna(range, 3);
            datosColumna.forEach((value, index) => {
                XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 7 }, raw: false });
            });


            //COLUMNA I: Forma
            var datosColumna = agregarUnaColumna(range, 5);
            datosColumna.forEach((value, index) => {
                XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 8 }, raw: false });
            });

            //COLUMNA J: Categoría
            var datosColumna = agregarUnaColumna(range, 6);
            datosColumna.forEach((value, index) => {
                XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 9 }, raw: false });
            });

            //COLUMNA K: Calibre
            var datosColumna = agregarUnaColumna(range, 13);
            datosColumna.forEach((value, index) => {
                XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 10 }, raw: false });
            });

            //COLUMNA L: Link Drive



            // ... Fin de generación de columnas ...       


            // Añadir la nueva hoja al nuevo libro de Excel
            XLSX.utils.book_append_sheet(nuevoLibro, nuevaHoja, 'de receta');

            // Crear un blob con el nuevo libro, primero en binary y luego convertido con la funcion s2ab
            var blob = XLSX.write(nuevoLibro, { bookType: 'xlsx', type: 'binary' });


            // Descargar el nuevo archivo Excel
            saveAs(new Blob([s2ab(blob)], { type: "application/octet-stream" }), 'AnteojosDeReceta.xlsx');
        }
        // Lee el contenido del archivo como binario
        reader.readAsBinaryString(file);
    } else {
        // Mostrar mensaje de copiado
        var mensajeError = document.getElementById("errorArchivo");
        mensajeError.style.display = "block";

        setTimeout(function () {
            mensajeError.style.display = "none";
        }, 2000); // El mensaje se ocultará después de 2 segundos (puedes ajustar el tiempo según tus necesidades)
    }



}

function s2ab(s) {
    var buf = new ArrayBuffer(s.length); //convert s to arrayBuffer
    var view = new Uint8Array(buf);  //create uint8array as viewer
    for (var i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF; //convert to octet
    return buf;
}

function agregarUnaColumna(range, columna) {
    var resultado = [];
    for (var i = range.s.r + 1; i <= range.e.r; ++i) {
        var cellAddress = XLSX.utils.encode_cell({ r: i, c: columna });
        var cellValue = selectedSheet[cellAddress] ? selectedSheet[cellAddress].v : undefined;
        resultado.push(cellValue);
    }
    return resultado;

}

function descargarModeloPlantilla() {
    



    
}