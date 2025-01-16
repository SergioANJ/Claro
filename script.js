// Lista para almacenar los nuevos datos dinámicos
const nuevosDatos = [];
let datosExcel = {}; // Objeto para almacenar los datos cargados del archivo Excel
let contadorCampoDinamico = 0; // Contador para los campos dinámicos de tipo archivo

// Evento para agregar un nuevo campo dinámico
document.getElementById('agregarCampoBtn').addEventListener('click', function () {
    const tipoCampo = document.getElementById('tipoCampo').value;

    if (tipoCampo === 'texto') {
        const nuevoCampoNombre = String(document.getElementById('nuevoCampoNombre').value).trim();
        const nuevoCampoValor = document.getElementById('nuevoCampoValor').value.trim();

        if (nuevoCampoNombre && nuevoCampoValor) {
            nuevosDatos.push({ nombre: nuevoCampoNombre, valor: nuevoCampoValor });
            document.getElementById('nuevoCampoNombre').value = '';
            document.getElementById('nuevoCampoValor').value = '';
            actualizarCamposDinamicos();
        } else {
            alert('Por favor ingresa tanto el nombre como el valor del nuevo campo.');
        }
    } else if (tipoCampo === 'archivo') {
        const archivo = document.getElementById('nuevoCampoArchivo').files[0];
        if (archivo) {
            const reader = new FileReader();
            reader.onload = function (e) {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const sheet = workbook.Sheets[workbook.SheetNames[0]];
                const sheetData = XLSX.utils.sheet_to_json(sheet, { header: 1 }).slice(1);
    
                // Extraer la primera columna y agregarla a nuevosDatos
                const nuevaColumna = sheetData.map(row => row[0]).filter(value => value); // Filtrar valores no vacíos
                
                // Incrementar el contador y crear el nuevo nombre
                contadorCampoDinamico++;
                const nuevoNombreCampo = `CampoDinamico_${contadorCampoDinamico}`;
                
                nuevosDatos.push({ nombre: nuevoNombreCampo, valor: nuevaColumna }); // Guardar como un solo campo
    
                actualizarCamposDinamicos(); // Actualizar la visualización de campos dinámicos
            };
            reader.readAsArrayBuffer(archivo);
        } else {
            alert('Por favor selecciona un archivo para agregar.');
        }
    }
});
// Función para actualizar los campos dinámicos en el formulario
function actualizarCamposDinamicos() {
    const camposDinamicosDiv = document.getElementById('camposDinamicos');
    camposDinamicosDiv.innerHTML = '';
    nuevosDatos.forEach(dato => {
        const div = document.createElement('div');
        div.classList.add('form-group');
        div.innerHTML = `
            <label for="${dato.nombre}">${dato.nombre}:</label>
            <input type="text" class="form-control" id="${dato.nombre}" name="${dato.nombre}" value="${dato.valor}" readonly>
        `;
        camposDinamicosDiv.appendChild(div);
    });
}
function formatearFecha(fechaHora) {
    const fecha = new Date(fechaHora);

    // Formatear los componentes de la fecha y hora
    const anio = fecha.getFullYear();
    const mes = String(fecha.getMonth() + 1).padStart(2, '0'); // Meses son base 0
    const dia = String(fecha.getDate()).padStart(2, '0');
    const horas = String(fecha.getHours()).padStart(2, '0');
    const minutos = String(fecha.getMinutes()).padStart(2, '0');
    const segundos = String(fecha.getSeconds()).padStart(2, '0');
    const ampm = horas >= 12 ? 'p.m.' : 'a.m.'; // Determinar AM o PM

    // Retornar la fecha en formato YYYY-MM-DD HH:mm:ss (24H)
    return `${anio}-${mes}-${dia} ${horas}:${minutos}:${segundos}`;
}

// Evento para cargar el archivo procesado
document.getElementById('cargar-archivo-btn').addEventListener('click', function () {
    const archivo = document.getElementById('cargarArchivo').files[0];
    if (archivo) {
        const reader = new FileReader();
        reader.onload = function (e) {
            const contenido = e.target.result;

            // Dividir el contenido en filas
            const filas = contenido.split('\n');

            // Procesar cada fila como un arreglo de columnas
            const sheetData = filas.map(fila => fila.split(',').map(celda => celda.trim()));

            // Asignar los datos del archivo a los campos HTML (primera fila son los encabezados)
            document.getElementById('campana').value = sheetData[1]?.[0] || '';
            document.getElementById('tipoDestino').value = sheetData[1]?.[1] || '';
            document.getElementById('plataforma').value = sheetData[1]?.[3] || '';
            document.getElementById('fechaHora').value = sheetData[1]?.[4] || '';

            // Asignar directamente las columnas de interés
            document.getElementById('urlImagen').value = sheetData[1]?.[6] || ''; // Columna urlimagen
            document.getElementById('nombreTemplate').value = sheetData[1]?.[7] || ''; // Columna nombretemplate
            document.getElementById('urlPublica').value = sheetData[1]?.[8] || ''; // Columna urlpublica

            // Guardar los datos del archivo CSV para uso posterior
            datosExcel = sheetData;
        };

        reader.readAsText(archivo, 'UTF-8');
    } else {
        alert('Por favor selecciona un archivo.');
    }
});


// Escuchar el evento de selección del archivo de destino
document.getElementById('destino').addEventListener('change', function () {
    const archivo = document.getElementById('destino').files[0]; // Obtener el archivo seleccionado
    if (archivo) {
        // Obtener el nombre del archivo
        const nombreArchivo = archivo.name;

        // Asignar el nombre del archivo al campo MIN
        document.getElementById('min').value = nombreArchivo;
    }
});

// Evento para procesar el formulario y generar el Excel
document.getElementById('generarExcelBtn').addEventListener('click', function (e) {
    e.preventDefault();

    const campana = document.getElementById('campana').value;
    const tipoDestino = document.getElementById('tipoDestino').value || 'MIN';
    const archivoDestino = document.getElementById('destino').files[0];
    const plataforma = document.getElementById('plataforma').value || 'Whatsapp';
    let fechaHora = document.getElementById('fechaHora').value;

    // Formatear la fecha
    fechaHora = formatearFecha(fechaHora);

    // Obtener campos adicionales
    const min = document.getElementById('min').value;
    const urlImagen = document.getElementById('urlImagen').value;
    const nombreTemplate = document.getElementById('nombreTemplate').value;
    const urlPublica = document.getElementById('urlPublica').value;

    if (!archivoDestino) {
        alert('Por favor, sube un archivo Excel.');
        return;
    }

    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];

        const destinos = XLSX.utils.sheet_to_json(sheet, { header: 1 }).slice(1).flat();

        // Obtener los valores de la columna MIN del archivo de destino
        const sheetData = XLSX.utils.sheet_to_json(sheet, { header: 1 }).slice(1);
        const mins = sheetData.map(row => row[5]);  // Suponiendo que la columna MIN está en la posición 5 (columna 6)

        // Obtener los valores de Campo_Archivo
        const campoArchivoValues = nuevosDatos.filter(dato => dato.nombre.startsWith('CampoDinamico_')).map(dato => dato.valor).flat() || [];

        const nombresColumnasCombinadas = [
            'MIN',
            'urlimagen',
            'nombretemplate',
            'urlpublica',
            ...nuevosDatos.map(dato => String(dato.nombre)) // Convertir los nombres de los campos dinámicos a cadenas
        ].join(',');

        const result = destinos.map((destino, index) => {
            const fila = {
                CAMPANA: campana,
                TIPODESTINO: tipoDestino,
                DESTINO: destino,
                PLATAFORMA: plataforma,
                FECHAHORAENVIO: fechaHora,
                MIN: mins[index] || destino,
                url_imagen: urlImagen,
                nombre_template: nombreTemplate,
                url_publica: urlPublica,
            };
            
            // Agregar todos los campos dinámicos (incluyendo Campo_Archivo si existe)
            nuevosDatos.forEach(dato => {
                if (dato.nombre.startsWith('CampoDinamico_')) {
                    // Si es Campo_Archivo, asignar el valor correspondiente de la fila
                    fila[dato.nombre] = dato.valor[index] || '';
                } else {
                    // Para otros campos dinámicos, asignar el valor directamente
                    fila[dato.nombre] = dato.valor;
                }
            });

            // Construcción de la columna combinada sin coma final innecesaria
            const valoresBase = [
                mins[index] || destino,
                urlImagen,
                nombreTemplate,
                urlPublica
            ];
        
                // Agregar valores de campos dinámicos a la columna combinada si existen
            nuevosDatos.forEach(dato => {
                if (dato.nombre.startsWith('CampoDinamico_')) {
                    valoresBase.push(dato.valor[index] || ''); // Agregar el valor correspondiente de Campo_Archivo
                } else {
                    valoresBase.push(dato.valor); // Agregar otros campos dinámicos
                }
            });
        
            // Combinar y evitar coma final
            fila[nombresColumnasCombinadas] = valoresBase.join(',');
        
            return fila;
        });        

        const columnasOrdenadas = [
            'CAMPANA',
            'TIPODESTINO',
            'DESTINO',
            'PLATAFORMA',
            'FECHAHORAENVIO',
            'MIN',
            'url_imagen',
            'nombre_template',
            'url_publica',
            ...nuevosDatos.map(dato => String(dato.nombre)),
            nombresColumnasCombinadas,
        ];

        // Ordenar datos y convertir números grandes a texto
        const resultOrdenado = result.map(fila =>
            columnasOrdenadas.reduce((acc, columna) => {
                let valor = fila[columna] || '';
                if (typeof valor === 'number' && valor.toString().includes('e')) {
                    valor = valor.toString(); // Convertir a texto
                }
                acc[columna] = valor;
                return acc;
            }, {})
        );

        // Filtrar las columnas a ocultar
        const columnasAExcluir = [
            'MIN',
            'url_imagen',
            'nombre_template',
            'url_publica',
            ...nuevosDatos.map(dato => String(dato.nombre))  // Incluir las columnas dinámicas
        ];

        const resultSinOcultas = resultOrdenado.map(fila => {
            const filteredFila = {};
            for (let key in fila) {
                if (!columnasAExcluir.includes(key)) {
                    filteredFila[key] = fila[key];
                }
            }
            return filteredFila;
        });

        // Crear hoja con los datos filtrados
        const worksheet = XLSX.utils.json_to_sheet(resultSinOcultas, { header: Object.keys(resultSinOcultas[0] || {}) });

        // Asegurarse de que los valores numéricos grandes se guarden como texto
        Object.keys(worksheet)
            .filter(cellAddress => /^[A-Z]+\d+$/.test(cellAddress)) // Solo direcciones de celdas válidas
            .forEach(cellAddress => {
                const cell = worksheet[cellAddress];
                if (cell && typeof cell.v === 'number') {
                    cell.v = cell.v.toString(); // Convertir a texto
                    cell.t = 's'; // Forzar tipo de celda a texto
                }
            });

            const workbookFinal = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbookFinal, worksheet, 'Datos Generados');
    
            const nombreArchivo = document.getElementById('nombreArchivo').value.trim();
            const nombreFinal = nombreArchivo ? `${nombreArchivo}.xlsx` : 'archivo_generado.xlsx';  
            
            // Guardar el archivo Excel generado
            XLSX.writeFile(workbookFinal, nombreFinal);
            alert('¡Archivo generado exitosamente!')
    
            // Limpiar los campos del formulario después de la descarga
            setTimeout(function () {
                document.getElementById('form').reset();
                document.getElementById('camposDinamicos').innerHTML = '';
            }, 100);
        };
    
        reader.readAsArrayBuffer(archivoDestino);
    });