document.getElementById('generar').addEventListener('click', mostrarDatos);
document.getElementById('generarExcel').addEventListener('click', generarExcel);
document.getElementById('btn_plantilla').addEventListener('click', descargarModeloPlantilla);

function descargarModeloPlantilla() {
    var reader = new FileReader();
    // Crear un nuevo libro de Excel
    var nuevoLibro = XLSX.utils.book_new();

    // Crear una nueva hoja de cálculo con los encabezados de columnas
    var nuevaHojaReceta = XLSX.utils.aoa_to_sheet([['SKU', 'Marca', 'Color ArmazónING', 'Color ArmazónESP', 'FormaESP', 'CategoríaESP', 'Modelo', 'NombreMod', 'CodColor', 'Material', 'ColorPatilla', 'Calibre', 'Puente', 'Patilla', 'Colección', 'País de Origen', 'Links fotos']]);

    var nuevaHojaSol = XLSX.utils.aoa_to_sheet([['SKU', 'Marca', 'ColArmazónING', 'ColArmazónESP', 'FormaESP', 'CategoríaESP', 'Modelo', 'NombreMod', 'CodColor', 'Material', 'Filtro', 'ColorPatilla', 'ColorLenteING', 'ColorLenteESP', 'Polarizado', 'Espejado', 'Apto Graduado', 'Calibre', 'Puente', 'Patilla', 'Colección', 'País de Origen', 'Links fotos']]);

    // Añadir la nueva hoja al nuevo libro de Excel
    XLSX.utils.book_append_sheet(nuevoLibro, nuevaHojaReceta, 'de receta');
    XLSX.utils.book_append_sheet(nuevoLibro, nuevaHojaSol, 'de sol');

    // Crear un blob con el nuevo libro, primero en binary y luego convertido con la funcion s2ab
    var blob = XLSX.write(nuevoLibro, { bookType: 'xlsx', type: 'binary' });

    // Descargar el nuevo archivo Excel
    saveAs(new Blob([s2ab(blob)], { type: "application/octet-stream" }), 'PlantillaAnteojos.xlsx');

    // Lee el contenido del archivo como binario
    reader.readAsBinaryString(file);
}

function mostrarDatos() {
    // Me guardo el tipo de anteojos seleccionado
    var tipoAnteojo = document.getElementById('tipoAnteojo').value;

    // Me guardo la acción seleccionada 
    var accion = document.getElementById('accion').value;

    // Me guardo el elemento de entrada de archivo
    var input = document.getElementById('fileInput');

    // Verifico si se seleccionó un archivo
    if (input.files.length > 0) {
        var file = input.files[0];

        // Creo un FileReader para leer el archivo
        var reader = new FileReader();

        reader.onload = function (e) {
            var data = e.target.result;

            // Uso SheetJS para analizar el archivo
            var workbook = XLSX.read(data, { type: 'binary' });

            // Selecciono la hoja según el tipo de anteojos
            var sheetIndex = (tipoAnteojo === 'receta') ? 0 : 1;

            // Obtengo la hoja correspondiente
            var selectedSheetName = workbook.SheetNames[sheetIndex];
            selectedSheet = workbook.Sheets[selectedSheetName];

            // Obtengo los datos según la acción seleccionada
            columnData = [];
            var range = XLSX.utils.decode_range(selectedSheet['!ref']);

            //Manejo la accion de acuerdo con el tipo de anteojo
            if (tipoAnteojo === 'receta') {
                switch (accion) {
                    case ('sku'):
                        columnData = agregarUnaColumna(range, 0);
                        break;
                    case ('nombres'):
                        columnData = generarNombresReceta(range);
                        break;
                    case ('descripcion'):
                        var columnData = generarDescripcionReceta(range);
                        break;                    
                }
            } else {
                switch (accion) {
                    case ('sku'):
                        columnData = agregarUnaColumna(range, 0);
                        break;
                    case ('nombres'):
                        columnData = generarNombresSol(range);
                        break;
                    case ('descripcion'):
                        columnData = generarDescripcionSol(range);
                        break;                    
                }
            }

            // Muestra los datos en el textarea
            document.getElementById('resultTextarea').value = columnData.join('\n');
        };

        // Lee el contenido del archivo como binario
        reader.readAsBinaryString(file);
    } else {
        error();
    }
}

    function generarExcel() {
        //Cargo todos los datos que necesito
        var input = document.getElementById('fileInput');
        var tipoAnteojo = document.getElementById('tipoAnteojo').value;
        if (input.files.length > 0) {
            var file = input.files[0];
            var reader = new FileReader();
            reader.onload = function (e) {
                var data = e.target.result;
                var workbook = XLSX.read(data, { type: 'binary' });
                var sheetIndex = (tipoAnteojo === 'receta') ? 0 : 1;
                var selectedSheetName = workbook.SheetNames[sheetIndex];
                selectedSheet = workbook.Sheets[selectedSheetName];
                var range = XLSX.utils.decode_range(selectedSheet['!ref']);

                // Crear un nuevo libro de Excel
                var nuevoLibro = XLSX.utils.book_new();

                // Crear una nueva hoja de cálculo con los encabezados de columnas
                var nuevaHoja = XLSX.utils.aoa_to_sheet([['Tipo', 'SKU', 'Nombre', 'Publicado', 'Visibilidad en el catálogo', 'Descripción corta', 'Estado del impuesto', '¿Existencias?', 'Inventario', '¿Permitir reservas de productos agotados?', '¿Vendido individualmente?', '¿Permitir valoraciones de clientes?', 'Precio normal', 'Categorías', 'Imágenes', 'Posición', 'Nombre del atributo 1', 'Valor(es) del atributo 1', 'Atributo visible 1', 'Atributo global 1', 'Nombre del atributo 2', 'Valor(es) del atributo 2', 'Atributo visible 2', 'Atributo global 2', 'Nombre del atributo 3', 'Valor(es) del atributo 3', 'Atributo visible 3', 'Atributo global 3', 'Nombre del atributo 4', 'Valor(es) del atributo 4', 'Atributo visible 4', 'Atributo global 4', 'Nombre del atributo 5', 'Valor(es) del atributo 5', 'Atributo visible 5', 'Atributo global 5', 'Nombre del atributo 6', 'Valor(es) del atributo 6', 'Atributo visible 6', 'Atributo global 6', 'Nombre del atributo 7', 'Valor(es) del atributo 7', 'Atributo visible 7', 'Atributo global 7', 'Nombre del atributo 8', 'Valor(es) del atributo 8', 'Atributo visible 8', 'Atributo global 8', 'Nombre del atributo 9', 'Valor(es) del atributo 9', 'Atributo visible 9', 'Atributo global 9']]);

                if (tipoAnteojo === 'receta') {

                    //Generar columnas de receta 
                    generarColumnasReceta(range, nuevaHoja);

                    // Añadir la nueva hoja al nuevo libro de Excel
                    XLSX.utils.book_append_sheet(nuevoLibro, nuevaHoja, 'de receta');

                    // Crear un blob con el nuevo libro, primero en binary y luego convertido con la funcion s2ab
                    var blob = XLSX.write(nuevoLibro, { bookType: 'xlsx', type: 'binary' });

                    // Descargar el nuevo archivo Excel
                    saveAs(new Blob([s2ab(blob)], { type: "application/octet-stream" }), 'Publicar_Receta.xlsx');

                } else {
                    //Generar columnas de sol 
                    generarColumnasSol(range, nuevaHoja);

                    // Añadir la nueva hoja al nuevo libro de Excel
                    XLSX.utils.book_append_sheet(nuevoLibro, nuevaHoja, 'de Sol');

                    // Crear un blob con el nuevo libro, primero en binary y luego convertido con la funcion s2ab
                    var blob = XLSX.write(nuevoLibro, { bookType: 'xlsx', type: 'binary' });

                    // Descargar el nuevo archivo Excel
                    saveAs(new Blob([s2ab(blob)], { type: "application/octet-stream" }), 'Publicar_Sol.xlsx');
                }

            }
            // Lee el contenido del archivo como binario
            reader.readAsBinaryString(file);

        } else {
            error()
        }

    }


    //MÉTODOS COMUNES

    function error() {
        var mensajeError = document.getElementById("errorArchivo");
        mensajeError.style.display = "block";
        setTimeout(function () {
            mensajeError.style.display = "none";
        }, 2000); // El mensaje se ocultará después de 2 segundos 

    }

    function s2ab(s) {
        var buf = new ArrayBuffer(s.length);
        var view = new Uint8Array(buf);
        for (var i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
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


    //MÉTODOS PARA RECETA

    function generarColumnasReceta(range, nuevaHoja) {
        var columnaSKU;

        //COLUMNAS DE CONTENIDO VARIABLE

        //COLUMNA B: SKU
        columnaSKU = [];
        columnaSKU = agregarUnaColumna(range, 0);
        columnaSKU.forEach((value, index) => {
            XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 1 }, raw: false });
        });

        //COLUMNA C: Nombre
        nuevaColumna = [];
        nuevaColumna = generarNombresReceta(range);
        nuevaColumna.forEach((value, index) => {
            XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 2 }, raw: false });
        });

        //COLUMNA F: Descripción
        nuevaColumna = [];
        nuevaColumna = generarDescripcionReceta(range);
        nuevaColumna.forEach((value, index) => {
            XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 5 }, raw: false });
        });

        //COLUMNA O: Imágenes
        nuevaColumna = [];
        nuevaColumna = agregarUnaColumna(range, 16);
        nuevaColumna.forEach((value, index) => {
            XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 14 }, raw: false });
        });

        //COLUMNA R: Calibre
        nuevaColumna = [];
        nuevaColumna = agregarUnaColumna(range, 11);
        nuevaColumna.forEach((value, index) => {
            XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 17 }, raw: false });
        });

        //COLUMNA V: Categoría
        nuevaColumna = [];
        nuevaColumna = agregarUnaColumna(range, 5);
        nuevaColumna.forEach((value, index) => {
            (value) ? value = value.charAt(0).toUpperCase() + value.slice(1).toLowerCase() : '';
            XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 21 }, raw: false });
        });

        //COLUMNA Z: Color del armazón
        nuevaColumna = [];
        nuevaColumna = agregarUnaColumna(range, 3);
        nuevaColumna.forEach((value, index) => {
            (value) ? value = value.charAt(0).toUpperCase() + value.slice(1).toLowerCase() : '';
            XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 25 }, raw: false });
        });

        //COLUMNA AD: Forma
        nuevaColumna = [];
        nuevaColumna = agregarUnaColumna(range, 4);
        nuevaColumna.forEach((value, index) => {
            (value) ? value = value.charAt(0).toUpperCase() + value.slice(1).toLowerCase() : '';
            XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 29 }, raw: false });
        });

        //COLUMNA AH: Marca
        nuevaColumna = [];
        nuevaColumna = agregarUnaColumna(range, 1);
        nuevaColumna.forEach((value, index) => {
            (value) ? value = value.charAt(0).toUpperCase() + value.slice(1).toLowerCase() : '';
            XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 33 }, raw: false });
        });



        //COLUMNAS DE CONTENIDO FIJO

        //ColumnaSKU tiene el total de registros/filas de la nueva hoja 
        for (var i = 1; i <= columnaSKU.length; i++) {

            var cellRefA = 'A' + (i + 1);
            nuevaHoja[cellRefA] = { t: 's', v: 'simple' };

            //cargo todas las columnas juntas que llevan el mismo valor
            var cellRefs = ['D', 'H', 'L', 'S', 'T', 'W', 'X', 'AA', 'AB', 'AE', 'AF', 'AI', 'AJ'];
            cellRefs.forEach(function (cellRef) {
                nuevaHoja[cellRef + (i + 1)] = { t: 'n', v: '1' };
            });

            //cargo todas las columnas juntas que llevan el mismo valor
            var cellRefs = ['I', 'J', 'K', 'P'];
            cellRefs.forEach(function (cellRef) {
                nuevaHoja[cellRef + (i + 1)] = { t: 'n', v: '0' };
            });

            var columna;
            columna = 'E' + (i + 1);
            nuevaHoja[columna] = { t: 's', v: 'visible' };

            columna = 'G' + (i + 1);
            nuevaHoja[columna] = { t: 's', v: 'taxable' };

            columna = 'N' + (i + 1);
            nuevaHoja[columna] = { t: 's', v: 'Recetados' };

            columna = 'Q' + (i + 1);
            nuevaHoja[columna] = { t: 's', v: 'Calibre' };

            columna = 'U' + (i + 1);
            nuevaHoja[columna] = { t: 's', v: 'Categoría' };

            columna = 'Y' + (i + 1);
            nuevaHoja[columna] = { t: 's', v: 'Color Armazón' };

            columna = 'AC' + (i + 1);
            nuevaHoja[columna] = { t: 's', v: 'Forma' };

            columna = 'AG' + (i + 1);
            nuevaHoja[columna] = { t: 's', v: 'Marca' };

        }




    }

    function generarNombresReceta(range) {
        var resultado = [];
        for (var i = range.s.r + 1; i <= range.e.r; i++) {
            var cellAddressB = XLSX.utils.encode_cell({ r: i, c: 1 });
            var cellAddressG = XLSX.utils.encode_cell({ r: i, c: 6 });
            var cellAddressH = XLSX.utils.encode_cell({ r: i, c: 7 });
            var cellAddressI = XLSX.utils.encode_cell({ r: i, c: 8 });
            var cellAddressL = XLSX.utils.encode_cell({ r: i, c: 11 });
            var cellAddressO = XLSX.utils.encode_cell({ r: i, c: 14 });

            var cellValueB = selectedSheet[cellAddressB] ? selectedSheet[cellAddressB].v.charAt(0).toUpperCase() + selectedSheet[cellAddressB].v.slice(1).toLowerCase() : '';
            var cellValueG = selectedSheet[cellAddressG] ? selectedSheet[cellAddressG].v : '';
            var cellValueH = selectedSheet[cellAddressH] ? selectedSheet[cellAddressH].v.charAt(0).toUpperCase() + selectedSheet[cellAddressH].v.slice(1).toLowerCase() : '';
            var cellValueI = selectedSheet[cellAddressI] ? selectedSheet[cellAddressI].v : '';
            var cellValueL = selectedSheet[cellAddressL] ? selectedSheet[cellAddressL].v : '';
            var cellValueO = '';//La coleccion podría ser un numero y por ende no es aplicable el toUpperCase/toLowerCase
            if (selectedSheet[cellAddressO] && typeof selectedSheet[cellAddressO].v === 'string') {
                cellValueO = selectedSheet[cellAddressO].v.charAt(0).toUpperCase() + selectedSheet[cellAddressO].v.slice(1).toLowerCase();
            }
            // Verificar cada valor antes de incluirlo en nombreCompleto, si es vacío lo omite
            var nombreCompleto = '';
            if (cellValueB) nombreCompleto += `${cellValueB.trim()} `; // trim() para eliminar espacios al final
            if (cellValueH) nombreCompleto += `${cellValueH.trim()} `;
            if (cellValueG) nombreCompleto += `${cellValueG.trim()} `;
            if (cellValueI) nombreCompleto += `color ${cellValueI} `;
            if (cellValueL) nombreCompleto += `cal ${cellValueL} `; //es un num por eso no hace falta cortarlo
            if (cellValueO) nombreCompleto += `- ${cellValueO.trim()}`;

            resultado.push(nombreCompleto.trim());
        }
        return resultado;
    }

    function generarDescripcionReceta(range) {
        var resultadoDescripcionReceta = [];

        for (var i = range.s.r + 1; i <= range.e.r; i++) {
            var cellAddressB = XLSX.utils.encode_cell({ r: i, c: 1 });
            var cellAddressC = XLSX.utils.encode_cell({ r: i, c: 2 });
            var cellAddressG = XLSX.utils.encode_cell({ r: i, c: 6 });
            var cellAddressH = XLSX.utils.encode_cell({ r: i, c: 7 });
            var cellAddressI = XLSX.utils.encode_cell({ r: i, c: 8 });
            var cellAddressJ = XLSX.utils.encode_cell({ r: i, c: 9 });
            var cellAddressK = XLSX.utils.encode_cell({ r: i, c: 10 });
            var cellAddressL = XLSX.utils.encode_cell({ r: i, c: 11 });
            var cellAddressM = XLSX.utils.encode_cell({ r: i, c: 12 });
            var cellAddressN = XLSX.utils.encode_cell({ r: i, c: 13 });
            var cellAddressO = XLSX.utils.encode_cell({ r: i, c: 14 });
            var cellAddressP = XLSX.utils.encode_cell({ r: i, c: 15 });

            var cellValueB = selectedSheet[cellAddressB] ? selectedSheet[cellAddressB].v.charAt(0).toUpperCase() + selectedSheet[cellAddressB].v.slice(1).toLowerCase() : '';          
            var cellValueC = selectedSheet[cellAddressC] ? selectedSheet[cellAddressC].v.charAt(0).toUpperCase() + selectedSheet[cellAddressC].v.slice(1).toLowerCase() : '';
            var cellValueG = selectedSheet[cellAddressG] ? selectedSheet[cellAddressG].v : '';
            var cellValueH = selectedSheet[cellAddressH] ? selectedSheet[cellAddressH].v.charAt(0).toUpperCase() + selectedSheet[cellAddressH].v.slice(1).toLowerCase() : '';
            var cellValueI = selectedSheet[cellAddressI] ? selectedSheet[cellAddressI].v : '';
            var cellValueJ = selectedSheet[cellAddressJ] ? selectedSheet[cellAddressJ].v.charAt(0).toUpperCase() + selectedSheet[cellAddressJ].v.slice(1).toLowerCase() : '';
            var cellValueK = selectedSheet[cellAddressK] ? selectedSheet[cellAddressK].v.charAt(0).toUpperCase() + selectedSheet[cellAddressK].v.slice(1).toLowerCase() : '';
            var cellValueL = selectedSheet[cellAddressL] ? selectedSheet[cellAddressL].v : '';
            var cellValueM = selectedSheet[cellAddressM] ? selectedSheet[cellAddressM].v : '';
            var cellValueN = selectedSheet[cellAddressN] ? selectedSheet[cellAddressN].v : '';
            var cellValueO = ''; //La coleccion podría ser un numero y por ende no es aplicable el toUpperCase/toLowerCase
            if (selectedSheet[cellAddressO] && typeof selectedSheet[cellAddressO].v === 'string') {
                cellValueO = selectedSheet[cellAddressO].v.charAt(0).toUpperCase() + selectedSheet[cellAddressO].v.slice(1).toLowerCase();
            }
            var cellValueP = selectedSheet[cellAddressP] ? selectedSheet[cellAddressP].v : '';

            // Verificar cada valor antes de incluirlo en nombreCompleto, si es vacío lo omite
            var nombreCompleto = '';
            nombreCompleto += '<p>Armazón para anteojos ';
            if (cellValueB) nombreCompleto += `${cellValueB.trim()} `; // trim() para eliminar espacios al final
            if (cellValueH) nombreCompleto += `${cellValueH.trim()} `;
            if (cellValueG) nombreCompleto += `${cellValueG.trim()} `;
            if (cellValueI) nombreCompleto += `Color ${cellValueI} `;
            nombreCompleto += 'Original, con estuche y garantía oficial. </p>';
            nombreCompleto += '<p>';
            if (cellValueO) nombreCompleto += `Colección: ${cellValueO.trim()}. <br>`;
            if (cellValueJ) nombreCompleto += `Material del armazón: ${cellValueJ.trim()}. <br>`;
            if (cellValueC) nombreCompleto += `Color del frente: ${cellValueC.trim()}. <br>`;
            if (cellValueK) nombreCompleto += `Color de patilla: ${cellValueK.trim()}. <br>`;
            if (cellValueP) nombreCompleto += `Pais de origen: ${cellValueP.trim()}. <br>`;
            nombreCompleto += '</p>';
            nombreCompleto += '<p>';
            nombreCompleto += 'Medidas: <br>';
            if (cellValueL) nombreCompleto += `Diámetro de la lente: ${cellValueL}mm. <br>`;
            if (cellValueM) nombreCompleto += `Largo de puente: ${cellValueM}mm. <br>`;
            if (cellValueN) nombreCompleto += `Largo de patilla: ${cellValueN}mm. <br>`;
            nombreCompleto += '</p>'
            nombreCompleto += '<p> 1 año de Garantía por defectos de fabricación. NO cubre fallas por mal uso del producto. </p>';
            nombreCompleto += '<p>Envío gratis a todo el país. <br></p>';

            resultadoDescripcionReceta.push(nombreCompleto.trim());
        }
        return resultadoDescripcionReceta;
    }


    //MÉTODOS PARA SOL

function generarColumnasSol(range, nuevaHoja) {
    var columnaSKU;

    //COLUMNAS DE CONTENIDO VARIABLE

    //COLUMNA B: SKU
    columnaSKU = [];
    columnaSKU = agregarUnaColumna(range, 0);
    columnaSKU.forEach((value, index) => {
        XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 1 }, raw: false });
    });

    //COLUMNA C: Nombre
    nuevaColumna = [];
    nuevaColumna = generarNombresSol(range);
    nuevaColumna.forEach((value, index) => {
        XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 2 }, raw: false });
    });

    //COLUMNA F: Descripción
    nuevaColumna = [];
    nuevaColumna = generarDescripcionSol(range);
    nuevaColumna.forEach((value, index) => {
        XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 5 }, raw: false });
    });

    //COLUMNA O: Imágenes
    nuevaColumna = [];
    nuevaColumna = agregarUnaColumna(range, 22);
    nuevaColumna.forEach((value, index) => {
        XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 14 }, raw: false });
    });

    //COLUMNA R: Apto Graduable
    nuevaColumna = [];
    nuevaColumna = agregarUnaColumna(range, 16);
    nuevaColumna.forEach((value, index) => {
        (value) ? value = value.charAt(0).toUpperCase() + value.slice(1).toLowerCase() : '';
        XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 17 }, raw: false });
    });

    //COLUMNA V: Calibre
    nuevaColumna = [];
    nuevaColumna = agregarUnaColumna(range, 17);
    nuevaColumna.forEach((value, index) => {
        XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 21 }, raw: false });
    });

    //COLUMNA Z: Categoría
    nuevaColumna = [];
    nuevaColumna = agregarUnaColumna(range, 5);
    nuevaColumna.forEach((value, index) => {
        (value) ? value = value.charAt(0).toUpperCase() + value.slice(1).toLowerCase() : '';
        XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 25 }, raw: false });
    });

    //COLUMNA AD: Color Armazón
    nuevaColumna = [];
    nuevaColumna = agregarUnaColumna(range, 3);
    nuevaColumna.forEach((value, index) => {
        (value) ? value = value.charAt(0).toUpperCase() + value.slice(1).toLowerCase() : '';
        XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 29 }, raw: false });
    });

    //COLUMNA AH: Color lente
    nuevaColumna = [];
    nuevaColumna = agregarUnaColumna(range, 13);
    nuevaColumna.forEach((value, index) => {
        (value) ? value = value.charAt(0).toUpperCase() + value.slice(1).toLowerCase() : '';
        XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 33 }, raw: false });
    });

    //COLUMNA AL: Espejado
    nuevaColumna = [];
    nuevaColumna = agregarUnaColumna(range, 15);
    nuevaColumna.forEach((value, index) => {
        (value) ? value = value.charAt(0).toUpperCase() + value.slice(1).toLowerCase() : '';
        XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 37 }, raw: false });
    });

    //COLUMNA AP: Forma
    nuevaColumna = [];
    nuevaColumna = agregarUnaColumna(range, 4);
    nuevaColumna.forEach((value, index) => {
        (value) ? value = value.charAt(0).toUpperCase() + value.slice(1).toLowerCase() : '';
        XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 41 }, raw: false });
    });

    //COLUMNA AT: Marca
    nuevaColumna = [];
    nuevaColumna = agregarUnaColumna(range, 1);
    nuevaColumna.forEach((value, index) => {
        (value) ? value = value.charAt(0).toUpperCase() + value.slice(1).toLowerCase() : '';
        XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 45 }, raw: false });
    });

    //COLUMNA AX: Polarizado
    nuevaColumna = [];
    nuevaColumna = agregarUnaColumna(range, 14);
    nuevaColumna.forEach((value, index) => {
        (value) ? value = value.charAt(0).toUpperCase() + value.slice(1).toLowerCase() : '';
        XLSX.utils.sheet_add_aoa(nuevaHoja, [[value]], { origin: { r: index + 1, c: 49 }, raw: false });
    });

    //COLUMNAS DE CONTENIDO FIJO

    //ColumnaSKU tiene el total de registros/filas de la nueva hoja 
    for (var i = 1; i <= columnaSKU.length; i++) {


        //cargo todas las columnas juntas que llevan el mismo valor
        var cellRefs = ['D', 'H', 'L', 'S', 'T', 'W', 'X', 'AA', 'AB', 'AE', 'AF', 'AI', 'AJ', 'AM', 'AN', 'AQ', 'AR', 'AU', 'AV', 'AY', 'AZ'];
        cellRefs.forEach(function (cellRef) {
            nuevaHoja[cellRef + (i + 1)] = { t: 'n', v: '1' };
        });

        //cargo todas las columnas juntas que llevan el mismo valor
        var cellRefs = ['I', 'J', 'K', 'P'];
        cellRefs.forEach(function (cellRef) {
            nuevaHoja[cellRef + (i + 1)] = { t: 'n', v: '0' };
        });

        var columna;
        
        columna = 'A' + (i + 1);
        nuevaHoja[columna] = { t: 's', v: 'simple' };

        columna = 'E' + (i + 1);
        nuevaHoja[columna] = { t: 's', v: 'visible' };

        columna = 'G' + (i + 1);
        nuevaHoja[columna] = { t: 's', v: 'taxable' };

        columna = 'N' + (i + 1);
        nuevaHoja[columna] = { t: 's', v: 'Lentes de sol' };

        columna = 'Q' + (i + 1);
        nuevaHoja[columna] = { t: 's', v: 'Apto Graduable' };

        columna = 'U' + (i + 1);
        nuevaHoja[columna] = { t: 's', v: 'Calibre' };

        columna = 'Y' + (i + 1);
        nuevaHoja[columna] = { t: 's', v: 'Categoría' };

        columna = 'AC' + (i + 1);
        nuevaHoja[columna] = { t: 's', v: 'Color Armazón' };

        columna = 'AG' + (i + 1);
        nuevaHoja[columna] = { t: 's', v: 'Color Lente' };

        columna = 'AK' + (i + 1);
        nuevaHoja[columna] = { t: 's', v: 'Espejado' };

        columna = 'AO' + (i + 1);
        nuevaHoja[columna] = { t: 's', v: 'Forma' };

        columna = 'AS' + (i + 1);
        nuevaHoja[columna] = { t: 's', v: 'Marca' };

        columna = 'AW' + (i + 1);
        nuevaHoja[columna] = { t: 's', v: 'Polarizado' };
    }
}

function generarNombresSol(range) {
    var resultado = [];
    for (var i = range.s.r + 1; i <= range.e.r; i++) {
        var cellAddressB = XLSX.utils.encode_cell({ r: i, c: 1 });
        var cellAddressG = XLSX.utils.encode_cell({ r: i, c: 6 });
        var cellAddressH = XLSX.utils.encode_cell({ r: i, c: 7 });        
        var cellAddressI = XLSX.utils.encode_cell({ r: i, c: 8 });
        var cellAddressU = XLSX.utils.encode_cell({ r: i, c: 20 });

        var cellValueB = selectedSheet[cellAddressB] ? selectedSheet[cellAddressB].v.charAt(0).toUpperCase() + selectedSheet[cellAddressB].v.slice(1).toLowerCase() : '';
        var cellValueG = selectedSheet[cellAddressG] ? selectedSheet[cellAddressG].v : '';
        var cellValueH = selectedSheet[cellAddressH] ? selectedSheet[cellAddressH].v.charAt(0).toUpperCase() + selectedSheet[cellAddressH].v.slice(1).toLowerCase() : '';
        var cellValueI = selectedSheet[cellAddressI] ? selectedSheet[cellAddressI].v : '';
        var cellValueU = ''; //La coleccion podría ser un numero y por ende no es aplicable el toUpperCase/toLowerCase
            if (selectedSheet[cellAddressU] && typeof selectedSheet[cellAddressU].v === 'string') {
                cellValueU = selectedSheet[cellAddressU].v.charAt(0).toUpperCase() + selectedSheet[cellAddressU].v.slice(1).toLowerCase();
            }

        // Verificar cada valor antes de incluirlo en nombreCompleto, si es vacío lo omite
        var nombreCompleto = '';
        if (cellValueB) nombreCompleto += `${cellValueB.trim()} `; // trim() para eliminar espacios al final
        if (cellValueH) nombreCompleto += `${cellValueH.trim()} `;
        if (cellValueG) nombreCompleto += `${cellValueG.trim()} `;
        if (cellValueI) nombreCompleto += `color ${cellValueI.trim()} `;
        if (cellValueU) nombreCompleto += `- ${cellValueU.trim()}`;

        resultado.push(nombreCompleto.trim());
    }
    return resultado;
}

function generarDescripcionSol(range) {
    var resultadoDescripcionSol = [];
    for (var i = range.s.r + 1; i <= range.e.r; i++) {
        var cellAddressB = XLSX.utils.encode_cell({ r: i, c: 1 });
        var cellAddressC = XLSX.utils.encode_cell({ r: i, c: 2 });
        var cellAddressG = XLSX.utils.encode_cell({ r: i, c: 6 });
        var cellAddressH = XLSX.utils.encode_cell({ r: i, c: 7 });
        var cellAddressI = XLSX.utils.encode_cell({ r: i, c: 8 });
        var cellAddressJ = XLSX.utils.encode_cell({ r: i, c: 9 });
        var cellAddressK = XLSX.utils.encode_cell({ r: i, c: 10 });
        var cellAddressL = XLSX.utils.encode_cell({ r: i, c: 11 });
        var cellAddressM = XLSX.utils.encode_cell({ r: i, c: 12 });
        var cellAddressR = XLSX.utils.encode_cell({ r: i, c: 17 });
        var cellAddressS = XLSX.utils.encode_cell({ r: i, c: 18 });
        var cellAddressT = XLSX.utils.encode_cell({ r: i, c: 19 });
        var cellAddressU = XLSX.utils.encode_cell({ r: i, c: 20 });
        var cellAddressV = XLSX.utils.encode_cell({ r: i, c: 21 });

        //si hay valor: transforma las palabras en la primer letra en mayus y las otras en min, sino un ''
        var cellValueB = selectedSheet[cellAddressB] ? selectedSheet[cellAddressB].v.charAt(0).toUpperCase() + selectedSheet[cellAddressB].v.slice(1).toLowerCase() : '';
        var cellValueC = selectedSheet[cellAddressC] ? selectedSheet[cellAddressC].v.charAt(0).toUpperCase() + selectedSheet[cellAddressB].v.slice(1).toLowerCase() : '';
        var cellValueG = selectedSheet[cellAddressG] ? selectedSheet[cellAddressG].v : '';
        var cellValueH = selectedSheet[cellAddressH] ? selectedSheet[cellAddressH].v.charAt(0).toUpperCase() + selectedSheet[cellAddressH].v.slice(1).toLowerCase() : '';
        var cellValueI = selectedSheet[cellAddressI] ? selectedSheet[cellAddressI].v : '';
        var cellValueJ = selectedSheet[cellAddressJ] ? selectedSheet[cellAddressJ].v.charAt(0).toUpperCase() + selectedSheet[cellAddressJ].v.slice(1).toLowerCase() : '';
        var cellValueK = selectedSheet[cellAddressK] ? selectedSheet[cellAddressK].v : '';
        var cellValueL = selectedSheet[cellAddressL] ? selectedSheet[cellAddressL].v.charAt(0).toUpperCase() + selectedSheet[cellAddressL].v.slice(1).toLowerCase() : '';
        var cellValueM = selectedSheet[cellAddressM] ? selectedSheet[cellAddressM].v.charAt(0).toUpperCase() + selectedSheet[cellAddressM].v.slice(1).toLowerCase() : '';
        var cellValueR = selectedSheet[cellAddressR] ? selectedSheet[cellAddressR].v : '';
        var cellValueS = selectedSheet[cellAddressS] ? selectedSheet[cellAddressS].v : '';
        var cellValueT = selectedSheet[cellAddressT] ? selectedSheet[cellAddressT].v : '';
        var cellValueU = ''; //La coleccion podría ser un numero y por ende no es aplicable el toUpperCase/toLowerCase
        if (selectedSheet[cellAddressU] && typeof selectedSheet[cellAddressU].v === 'string') {
            cellValueU = selectedSheet[cellAddressU].v.charAt(0).toUpperCase() + selectedSheet[cellAddressU].v.slice(1).toLowerCase();
        }
        var cellValueV = selectedSheet[cellAddressV] ? selectedSheet[cellAddressV].v : '';


        // Verificar cada valor antes de incluirlo en nombreCompleto, si es vacío lo omite/ trim() para eliminar espacios vacíos al final

        var nombreCompleto = '';
        nombreCompleto += '<p>Anteojos de sol ';
        if (cellValueB) nombreCompleto += `${cellValueB.trim()} `; 
        if (cellValueH) nombreCompleto += `${cellValueH.trim()} `;
        if (cellValueG) nombreCompleto += `${cellValueG.trim()} `;
        if (cellValueI) nombreCompleto += `color ${cellValueI.trim()} `;
        if (cellValueR) nombreCompleto += `cal ${cellValueR}. `;
        nombreCompleto += 'Original, con estuche y garantía oficial. </p>';
        nombreCompleto += '<p>';
        if (cellValueU) nombreCompleto += `Colección: ${cellValueU.trim()}. <br>`;
        if (cellValueJ) nombreCompleto += `Material del armazón: ${cellValueJ.trim()}. <br>`;
        if (cellValueM) nombreCompleto += `Color de la lente: ${cellValueM.trim()}. <br>`;
        if (cellValueC) nombreCompleto += `Color del frente: ${cellValueC.trim()}. <br>`;
        if (cellValueL) nombreCompleto += `Color de patilla: ${cellValueL.trim()}. <br>`;
        if (cellValueK) nombreCompleto += `Tipo de filtro: ${cellValueK.trim()}. <br>`;
        if (cellValueV) nombreCompleto += `Pais de origen: ${cellValueV.trim()}. <br>`;
        nombreCompleto += '</p>';
        nombreCompleto += '<p>';
        nombreCompleto += 'Medidas: <br>';
        if (cellValueR) nombreCompleto += `Diámetro de la lente: ${cellValueR}mm. <br>`;
        if (cellValueS) nombreCompleto += `Largo de puente: ${cellValueS}mm. <br>`;
        if (cellValueT) nombreCompleto += `Largo de patilla: ${cellValueT}mm. <br>`;
        nombreCompleto += '</p>'
        nombreCompleto += '<p> 1 año de Garantía por defectos de fabricación. NO cubre fallas por mal uso del producto. </p>';
        nombreCompleto += '<p>Envío gratis a todo el país. <br></p>';

        resultadoDescripcionSol.push(nombreCompleto.trim());
    }
    return resultadoDescripcionSol;
}