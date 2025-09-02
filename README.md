<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Control de Mantenimiento de Vehículos</title>
    <script src="https://unpkg.com/xlsx/dist/xlsx.full.min.js"></script>
    <style>
        @media print {
            @page {
                size: landscape;
                margin: 0.5cm;
            }
            body, html {
                width: 27.9cm;
                height: 21.6cm;
                margin: 0;
                padding: 0;
            }
            .container {
                width: 100%;
                height: 100%;
                padding: 0.5cm;
                margin: 0;
                box-shadow: none;
                border: none;
            }
            .no-print {
                display: none !important;
            }
            .form-title {
                box-shadow: none;
            }
            table {
                page-break-inside: avoid;
            }
        }
        
        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }
        
        body {
            background-color: #f0f3f5;
            color: #333;
            line-height: 1.4;
            padding: 15px;
            display: flex;
            justify-content: center;
            align-items: center;
            min-height: 100vh;
        }
        
        .container {
            width: 27.9cm;
            min-height: 21.6cm;
            background: white;
            padding: 0.7cm;
            box-shadow: 0 0 15px rgba(0, 0, 0, 0.1);
            margin: 0 auto;
            overflow-x: auto;
        }
        
        .header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 0.5cm;
            padding-bottom: 0.3cm;
            border-bottom: 2px solid #ffcc00;
        }
        
        .logo-container {
            display: flex;
            align-items: center;
        }
        
        .logo-img {
            width: 200px;
            height: 60px;
            object-fit: contain;
            margin-right: 15px;
        }
        
        .company-info {
            text-align: right;
        }
        
        .company-name {
            font-size: 18px;
            font-weight: bold;
            color: #0066a1;
            margin-bottom: 5px;
        }
        
        .form-title {
            text-align: center;
            background-color: #ffcc00;
            padding: 8px;
            margin: 0.4cm 0;
            border-radius: 4px;
            font-size: 16px;
            color: #333;
        }
        
        .info-section {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 10px;
            margin-bottom: 0.4cm;
            font-size: 12px;
        }
        
        .info-box {
            padding: 6px 8px;
            border: 1px solid #ddd;
            border-radius: 3px;
            background-color: #f9f9f9;
        }
        
        .info-label {
            font-weight: bold;
            color: #0066a1;
        }
        
        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 11px;
            table-layout: fixed;
            margin-bottom: 0.5cm;
        }
        
        th {
            background-color: #ffcc00;
            padding: 5px;
            text-align: center;
            border: 1px solid #ddd;
            font-weight: bold;
        }
        
        td {
            padding: 5px;
            border: 1px solid #ddd;
            text-align: center;
            height: 40px;
            vertical-align: middle;
        }
        
        .placa-cell {
            text-align: center;
            font-weight: bold;
            background-color: #f5f5f5;
            width: 8%;
            position: sticky;
            left: 0;
            z-index: 10;
        }
        
        .modelo-cell {
            text-align: left;
            padding-left: 8px;
            background-color: #f5f5f5;
            width: 15%;
            position: sticky;
            left: 8%;
            z-index: 10;
        }
        
        .service-cell {
            width: 10%;
        }
        
        .date-cell {
            width: 8%;
        }
        
        .workshop-cell {
            width: 12%;
        }
        
        .cost-cell {
            width: 8%;
            background-color: #f9f9f9;
        }
        
        .comments-cell {
            width: 15%;
            text-align: left;
            padding: 5px;
        }
        
        .service-number {
            font-weight: bold;
            color: #0066a1;
            background-color: #e6f7ff;
            text-align: center;
        }
        
        .footer {
            text-align: center;
            margin-top: 0.5cm;
            padding-top: 8px;
            border-top: 1px solid #ddd;
            color: #666;
            font-size: 11px;
        }
        
        .buttons-container {
            display: flex;
            justify-content: center;
            gap: 15px;
            margin: 15px auto;
            flex-wrap: wrap;
        }
        
        .action-btn {
            width: 180px;
            padding: 10px;
            color: white;
            border: none;
            border-radius: 4px;
            font-size: 14px;
            cursor: pointer;
            transition: all 0.3s;
        }
        
        .print-btn {
            background: linear-gradient(to bottom, #0066a1, #004d7a);
        }
        
        .excel-btn {
            background: linear-gradient(to bottom, #0d8044, #0a6636);
        }
        
        .save-btn {
            background: linear-gradient(to bottom, #ff9800, #f57c00);
        }
        
        .load-btn {
            background: linear-gradient(to bottom, #9c27b0, #7b1fa2);
        }
        
        .import-btn {
            background: linear-gradient(to bottom, #607d8b, #455a64);
        }
        
        .action-btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.2);
        }
        
        .service-minor {
            background-color: #e6f7ff;
        }
        
        .service-major {
            background-color: #fff2e6;
        }
        
        input[type="text"], input[type="date"], input[type="number"], select {
            border: none;
            border-bottom: 1px dashed #999;
            padding: 3px;
            width: 100%;
            background: transparent;
            font-size: 11px;
            text-align: center;
        }
        
        .section-title {
            color: #0066a1;
            margin: 15px 0 10px 0;
            font-size: 16px;
            border-bottom: 2px solid #0066a1;
            padding-bottom: 5px;
        }
        
        .service-container {
            margin-bottom: 20px;
            overflow-x: auto;
        }
        
        .notification {
            position: fixed;
            top: 20px;
            right: 20px;
            padding: 15px;
            border-radius: 5px;
            color: white;
            font-weight: bold;
            z-index: 1000;
            opacity: 0;
            transition: opacity 0.5s;
        }
        
        .success {
            background-color: #4caf50;
        }
        
        .error {
            background-color: #f44336;
        }
        
        .file-input-container {
            position: relative;
            display: inline-block;
        }
        
        .file-input {
            position: absolute;
            left: 0;
            top: 0;
            opacity: 0;
            width: 100%;
            height: 100%;
            cursor: pointer;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <div class="logo-container">
                <img src="https://cdn.imweb.me/upload/S202402132ab908c48484a/ad9072e67d6f0.png" alt="Logo de la empresa" class="logo-img">
            </div>
            <div class="company-info">
                <div class="company-name">Hansae GSN S.A.</div>
                <p>Control de Mantenimiento de Vehículos - 6 Servicios</p>
            </div>
        </div>
        
        <div class="info-section">
            <div class="info-box">
                <span class="info-label">Fecha de control:</span> <input type="date" id="fecha-control">
            </div>
            <div class="info-box">
                <span class="info-label">Responsable:</span> <input type="text" id="responsable" placeholder="Nombre completo">
            </div>
            <div class="info-box">
                <span class="info-label">Total vehículos:</span> 14
            </div>
        </div>
        
        <div class="form-title">REGISTRO DE MANTENIMIENTO DE VEHÍCULOS - 6 SERVICIOS POR VEHÍCULO</div>
        
        <div class="buttons-container no-print">
            <button class="action-btn save-btn" onclick="guardarDatos()">Guardar Datos</button>
            <button class="action-btn load-btn" onclick="cargarDatos()">Cargar Datos</button>
            <div class="file-input-container">
                <button class="action-btn import-btn">Importar desde Excel</button>
                <input type="file" id="file-input" class="file-input" accept=".xlsx, .xls" onchange="importarExcel(event)">
            </div>
            <button class="action-btn excel-btn" onclick="exportarExcel()">Exportar a Excel</button>
            <button class="action-btn print-btn" onclick="window.print()">Imprimir Formato</button>
        </div>
        
        <!-- Servicio 1 -->
        <div class="service-container">
            <div class="section-title">SERVICIO 1</div>
            <table id="servicio1">
                <thead>
                    <tr>
                        <th class="placa-cell">Placa</th>
                        <th class="modelo-cell">Modelo</th>
                        <th class="service-cell">Tipo Servicio</th>
                        <th class="date-cell">Fecha Servicio</th>
                        <th class="date-cell">Próximo Servicio</th>
                        <th class="workshop-cell">Taller</th>
                        <th class="cost-cell">Costo ($)</th>
                        <th class="comments-cell">Comentarios</th>
                    </tr>
                </thead>
                <tbody>
                    <!-- Los datos se generarán con JavaScript -->
                </tbody>
            </table>
        </div>
        
        <!-- Servicio 2 -->
        <div class="service-container">
            <div class="section-title">SERVICIO 2</div>
            <table id="servicio2">
                <thead>
                    <tr>
                        <th class="placa-cell">Placa</th>
                        <th class="modelo-cell">Modelo</th>
                        <th class="service-cell">Tipo Servicio</th>
                        <th class="date-cell">Fecha Servicio</th>
                        <th class="date-cell">Próximo Servicio</th>
                        <th class="workshop-cell">Taller</th>
                        <th class="cost-cell">Costo ($)</th>
                        <th class="comments-cell">Comentarios</th>
                    </tr>
                </thead>
                <tbody>
                    <!-- Los datos se generarán con JavaScript -->
                </tbody>
            </table>
        </div>
        
        <!-- Servicio 3 -->
        <div class="service-container">
            <div class="section-title">SERVICIO 3</div>
            <table id="servicio3">
                <thead>
                    <tr>
                        <th class="placa-cell">Placa</th>
                        <th class="modelo-cell">Modelo</th>
                        <th class="service-cell">Tipo Servicio</th>
                        <th class="date-cell">Fecha Servicio</th>
                        <th class="date-cell">Próximo Servicio</th>
                        <th class="workshop-cell">Taller</th>
                        <th class="cost-cell">Costo ($)</th>
                        <th class="comments-cell">Comentarios</th>
                    </tr>
                </thead>
                <tbody>
                    <!-- Los datos se generarán con JavaScript -->
                </tbody>
            </table>
        </div>
        
        <!-- Servicio 4 -->
        <div class="service-container">
            <div class="section-title">SERVICIO 4</div>
            <table id="servicio4">
                <thead>
                    <tr>
                        <th class="placa-cell">Placa</th>
                        <th class="modelo-cell">Modelo</th>
                        <th class="service-cell">Tipo Servicio</th>
                        <th class="date-cell">Fecha Servicio</th>
                        <th class="date-cell">Próximo Servicio</th>
                        <th class="workshop-cell">Taller</th>
                        <th class="cost-cell">Costo ($)</th>
                        <th class="comments-cell">Comentarios</th>
                    </tr>
                </thead>
                <tbody>
                    <!-- Los datos se generarán con JavaScript -->
                </tbody>
            </table>
        </div>
        
        <!-- Servicio 5 -->
        <div class="service-container">
            <div class="section-title">SERVICIO 5</div>
            <table id="servicio5">
                <thead>
                    <tr>
                        <th class="placa-cell">Placa</th>
                        <th class="modelo-cell">Modelo</th>
                        <th class="service-cell">Tipo Servicio</th>
                        <th class="date-cell">Fecha Servicio</th>
                        <th class="date-cell">Próximo Servicio</th>
                        <th class="workshop-cell">Taller</th>
                        <th class="cost-cell">Costo ($)</th>
                        <th class="comments-cell">Comentarios</th>
                    </tr>
                </thead>
                <tbody>
                    <!-- Los datos se generarán con JavaScript -->
                </tbody>
            </table>
        </div>
        
        <!-- Servicio 6 -->
        <div class="service-container">
            <div class="section-title">SERVICIO 6</div>
            <table id="servicio6">
                <thead>
                    <tr>
                        <th class="placa-cell">Placa</th>
                        <th class="modelo-cell">Modelo</th>
                        <th class="service-cell">Tipo Servicio</th>
                        <th class="date-cell">Fecha Servicio</th>
                        <th class="date-cell">Próximo Servicio</th>
                        <th class="workshop-cell">Taller</th>
                        <th class="cost-cell">Costo ($)</th>
                        <th class="comments-cell">Comentarios</th>
                    </tr>
                </thead>
                <tbody>
                    <!-- Los datos se generarán con JavaScript -->
                </tbody>
            </table>
        </div>
        
        <div class="footer">
            <p>Este documento es parte del sistema de control de flota - Formato ID: F-VH-2023 - Versión: 2.0</p>
        </div>
    </div>

    <div id="notification" class="notification"></div>

    <script>
        // Datos de los vehículos
        const vehiculos = [
            {placa: "856KRT", modelo: "HYUNDAI STARIA 2024"},
            {placa: "128GKG", modelo: "HONDA HRV 2016"},
            {placa: "704KFR", modelo: "HONDA HRV 2020"},
            {placa: "802JVT", modelo: "HYUNDAI STARIA 2023"},
            {placa: "805KDV", modelo: "TOYOTA PRADO 2023"},
            {placa: "842KQC", modelo: "TOYOTA PRADO 2024"},
            {placa: "502KPB", modelo: "CHEVROLET CAPTIVA 2023"},
            {placa: "210KMC", modelo: "KIA CARNIVAL 2024"},
            {placa: "053KVJ", modelo: "HONDA PILOT 2024"},
            {placa: "638KWY", modelo: "HYUNDAI SANTA FE 2025"},
            {placa: "250KWY", modelo: "HYUNDAI STARIA 2025"},
            {placa: "132KXZ", modelo: "HYUNDAI TUCSON 2025"},
            {placa: "924KXD", modelo: "TOYOTA INNOVA 2025"},
            {placa: "C833BZT", modelo: "HYUNDAI STARIA 2025"}
        ];

        // Inicializar las tablas
        document.addEventListener('DOMContentLoaded', function() {
            for (let i = 1; i <= 6; i++) {
                inicializarTabla('servicio' + i);
            }
            cargarDatos(); // Intentar cargar datos guardados al iniciar
        });

        // Función para inicializar una tabla
        function inicializarTabla(idTabla) {
            const tabla = document.getElementById(idTabla);
            const tbody = tabla.querySelector('tbody');
            tbody.innerHTML = '';
            
            vehiculos.forEach(vehiculo => {
                const fila = document.createElement('tr');
                fila.innerHTML = `
                    <td class="placa-cell">${vehiculo.placa}</td>
                    <td class="modelo-cell">${vehiculo.modelo}</td>
                    <td class="service-cell">
                        <select name="tipo">
                            <option value="">Seleccionar</option>
                            <option value="minor">Menor</option>
                            <option value="major">Mayor</option>
                        </select>
                    </td>
                    <td class="date-cell"><input type="date" name="fecha"></td>
                    <td class="date-cell"><input type="date" name="proximo"></td>
                    <td class="workshop-cell"><input type="text" name="taller"></td>
                    <td class="cost-cell"><input type="number" name="costo" min="0"></td>
                    <td class="comments-cell"><input type="text" name="comentarios"></td>
                `;
                tbody.appendChild(fila);
            });
        }

        // Función para guardar datos en localStorage
        function guardarDatos() {
            try {
                const datos = {
                    fecha: document.getElementById('fecha-control').value,
                    responsable: document.getElementById('responsable').value,
                    servicios: {}
                };
                
                for (let i = 1; i <= 6; i++) {
                    const servicioId = 'servicio' + i;
                    datos.servicios[servicioId] = [];
                    
                    const tabla = document.getElementById(servicioId);
                    const filas = tabla.querySelectorAll('tbody tr');
                    
                    filas.forEach((fila, index) => {
                        datos.servicios[servicioId][index] = {
                            placa: vehiculos[index].placa,
                            modelo: vehiculos[index].modelo,
                            tipo: fila.querySelector('[name="tipo"]').value,
                            fecha: fila.querySelector('[name="fecha"]').value,
                            proximo: fila.querySelector('[name="proximo"]').value,
                            taller: fila.querySelector('[name="taller"]').value,
                            costo: fila.querySelector('[name="costo"]').value,
                            comentarios: fila.querySelector('[name="comentarios"]').value
                        };
                    });
                }
                
                localStorage.setItem('mantenimientoVehiculos', JSON.stringify(datos));
                mostrarNotificacion('Datos guardados correctamente', 'success');
            } catch (error) {
                mostrarNotificacion('Error al guardar los datos: ' + error.message, 'error');
            }
        }

        // Función para cargar datos desde localStorage
        function cargarDatos() {
            try {
                const datosGuardados = localStorage.getItem('mantenimientoVehiculos');
                if (!datosGuardados) {
                    mostrarNotificacion('No hay datos guardados', 'error');
                    return;
                }
                
                const datos = JSON.parse(datosGuardados);
                
                document.getElementById('fecha-control').value = datos.fecha || '';
                document.getElementById('responsable').value = datos.responsable || '';
                
                for (let i = 1; i <= 6; i++) {
                    const servicioId = 'servicio' + i;
                    if (datos.servicios && datos.servicios[servicioId]) {
                        const tabla = document.getElementById(servicioId);
                        const filas = tabla.querySelectorAll('tbody tr');
                        
                        datos.servicios[servicioId].forEach((dato, index) => {
                            if (filas[index]) {
                                filas[index].querySelector('[name="tipo"]').value = dato.tipo || '';
                                filas[index].querySelector('[name="fecha"]').value = dato.fecha || '';
                                filas[index].querySelector('[name="proximo"]').value = dato.proximo || '';
                                filas[index].querySelector('[name="taller"]').value = dato.taller || '';
                                filas[index].querySelector('[name="costo"]').value = dato.costo || '';
                                filas[index].querySelector('[name="comentarios"]').value = dato.comentarios || '';
                                
                                // Aplicar color según tipo de servicio
                                const row = filas[index];
                                if (dato.tipo === 'minor') {
                                    row.classList.add('service-minor');
                                    row.classList.remove('service-major');
                                } else if (dato.tipo === 'major') {
                                    row.classList.add('service-major');
                                    row.classList.remove('service-minor');
                                } else {
                                    row.classList.remove('service-minor', 'service-major');
                                }
                            }
                        });
                    }
                }
                
                mostrarNotificacion('Datos cargados correctamente', 'success');
            } catch (error) {
                mostrarNotificacion('Error al cargar los datos: ' + error.message, 'error');
            }
        }

        // Función para importar datos desde Excel
        function importarExcel(event) {
            const file = event.target.files[0];
            if (!file) return;
            
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    
                    // Limpiar todos los campos primero
                    limpiarFormulario();
                    
                    // Procesar cada hoja del Excel
                    for (let i = 1; i <= 6; i++) {
                        const sheetName = `Servicio ${i}`;
                        if (workbook.SheetNames.includes(sheetName)) {
                            const worksheet = workbook.Sheets[sheetName];
                            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                            
                            // Saltar la fila de encabezados
                            for (let j = 1; j < jsonData.length; j++) {
                                const row = jsonData[j];
                                if (row.length >= 8) {
                                    const placa = row[0];
                                    const tipo = row[2];
                                    const fecha = row[3];
                                    const proximo = row[4];
                                    const taller = row[5];
                                    const costo = row[6];
                                    const comentarios = row[7];
                                    
                                    // Buscar el vehículo por placa y servicio
                                    const tabla = document.getElementById(`servicio${i}`);
                                    const filas = tabla.querySelectorAll('tbody tr');
                                    
                                    for (let k = 0; k < filas.length; k++) {
                                        const filaPlaca = filas[k].querySelector('.placa-cell').textContent;
                                        if (filaPlaca === placa) {
                                            filas[k].querySelector('[name="tipo"]').value = tipo || '';
                                            filas[k].querySelector('[name="fecha"]').value = convertirFechaExcel(fecha) || '';
                                            filas[k].querySelector('[name="proximo"]').value = convertirFechaExcel(proximo) || '';
                                            filas[k].querySelector('[name="taller"]').value = taller || '';
                                            filas[k].querySelector('[name="costo"]').value = costo || '';
                                            filas[k].querySelector('[name="comentarios"]').value = comentarios || '';
                                            
                                            // Aplicar color según tipo de servicio
                                            if (tipo === 'minor') {
                                                filas[k].classList.add('service-minor');
                                                filas[k].classList.remove('service-major');
                                            } else if (tipo === 'major') {
                                                filas[k].classList.add('service-major');
                                                filas[k].classList.remove('service-minor');
                                            }
                                            
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    
                    // Procesar hoja de resumen si existe
                    if (workbook.SheetNames.includes('Resumen')) {
                        const worksheet = workbook.Sheets['Resumen'];
                        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                        
                        // Buscar fecha y responsable en el resumen
                        if (jsonData.length > 0 && jsonData[0].length >= 2) {
                            document.getElementById('fecha-control').value = convertirFechaExcel(jsonData[0][1]) || '';
                        }
                        if (jsonData.length > 1 && jsonData[1].length >= 2) {
                            document.getElementById('responsable').value = jsonData[1][1] || '';
                        }
                    }
                    
                    mostrarNotificacion('Datos importados desde Excel correctamente', 'success');
                    
                    // Limpiar el input de archivo
                    event.target.value = '';
                    
                } catch (error) {
                    mostrarNotificacion('Error al importar desde Excel: ' + error.message, 'error');
                    console.error(error);
                }
            };
            reader.readAsArrayBuffer(file);
        }

        // Función para convertir fecha de Excel a formato YYYY-MM-DD
        function convertirFechaExcel(fechaExcel) {
            if (!fechaExcel) return '';
            
            // Si es un número (formato de fecha de Excel)
            if (typeof fechaExcel === 'number') {
                // Las fechas en Excel son días desde el 1 de enero de 1900
                const fecha = new Date((fechaExcel - 25569) * 86400 * 1000);
                return fecha.toISOString().split('T')[0];
            }
            
            // Si ya es una cadena en formato de fecha
            if (typeof fechaExcel === 'string') {
                // Intentar parsear diferentes formatos de fecha
                const fecha = new Date(fechaExcel);
                if (!isNaN(fecha.getTime())) {
                    return fecha.toISOString().split('T')[0];
                }
            }
            
            return '';
        }

        // Función para limpiar el formulario
        function limpiarFormulario() {
            document.getElementById('fecha-control').value = '';
            document.getElementById('responsable').value = '';
            
            for (let i = 1; i <= 6; i++) {
                const tabla = document.getElementById(`servicio${i}`);
                const filas = tabla.querySelectorAll('tbody tr');
                
                filas.forEach(fila => {
                    fila.querySelector('[name="tipo"]').value = '';
                    fila.querySelector('[name="fecha"]').value = '';
                    fila.querySelector('[name="proximo"]').value = '';
                    fila.querySelector('[name="taller"]').value = '';
                    fila.querySelector('[name="costo"]').value = '';
                    fila.querySelector('[name="comentarios"]').value = '';
                    fila.classList.remove('service-minor', 'service-major');
                });
            }
        }

        // Función para exportar a Excel
        function exportarExcel() {
            try {
                const wb = XLSX.utils.book_new();
                const today = new Date();
                const dateString = today.toISOString().split('T')[0];
                
                // Crear una hoja por cada servicio
                for (let i = 1; i <= 6; i++) {
                    const servicioId = 'servicio' + i;
                    const tabla = document.getElementById(servicioId);
                    const ws = XLSX.utils.table_to_sheet(tabla);
                    XLSX.utils.book_append_sheet(wb, ws, `Servicio ${i}`);
                }
                
                // Crear hoja resumen
                const datos = [];
                
                // Encabezados
                datos.push(['Placa', 'Modelo', 'Servicio', 'Tipo', 'Fecha', 'Próximo', 'Taller', 'Costo', 'Comentarios']);
                
                // Datos de cada servicio
                for (let i = 1; i <= 6; i++) {
                    const servicioId = 'servicio' + i;
                    const tabla = document.getElementById(servicioId);
                    const filas = tabla.querySelectorAll('tbody tr');
                    
                    filas.forEach(fila => {
                        const placa = fila.querySelector('.placa-cell').textContent;
                        const modelo = fila.querySelector('.modelo-cell').textContent;
                        const tipo = fila.querySelector('[name="tipo"]').value;
                        const fecha = fila.querySelector('[name="fecha"]').value;
                        const proximo = fila.querySelector('[name="proximo"]').value;
                        const taller = fila.querySelector('[name="taller"]').value;
                        const costo = fila.querySelector('[name="costo"]').value;
                        const comentarios = fila.querySelector('[name="comentarios"]').value;
                        
                        datos.push([placa, modelo, `Servicio ${i}`, tipo, fecha, proximo, taller, costo, comentarios]);
                    });
                }
                
                const wsResumen = XLSX.utils.aoa_to_sheet(datos);
                XLSX.utils.book_append_sheet(wb, wsResumen, 'Resumen');
                
                XLSX.writeFile(wb, `Mantenimiento_Vehiculos_${dateString}.xlsx`);
                mostrarNotificacion('Datos exportados a Excel correctamente', 'success');
            } catch (error) {
                mostrarNotificacion('Error al exportar a Excel: ' + error.message, 'error');
            }
        }

        // Mostrar notificación
        function mostrarNotificacion(mensaje, tipo) {
            const notification = document.getElementById('notification');
            notification.textContent = mensaje;
            notification.className = 'notification ' + tipo;
            notification.style.opacity = '1';
            
            setTimeout(() => {
                notification.style.opacity = '0';
            }, 3000);
        }

        // Cambiar color de fondo según tipo de servicio
        document.addEventListener('change', function(e) {
            if (e.target.tagName === 'SELECT' && e.target.name === 'tipo') {
                const row = e.target.closest('tr');
                if (e.target.value === 'minor') {
                    row.classList.add('service-minor');
                    row.classList.remove('service-major');
                } else if (e.target.value === 'major') {
                    row.classList.add('service-major');
                    row.classList.remove('service-minor');
                } else {
                    row.classList.remove('service-minor', 'service-major');
                }
            }
        });
    </script>
</body>
</html>
