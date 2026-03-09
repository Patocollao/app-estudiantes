<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Interfaz de la Aplicación</title>
    <style>
        /* 1. Variables de Color (Paleta del Portal) */
        :root {
            --app-fondo: #1B222C;        /* Fondo principal azul noche */
            --app-borde: #F6C128;        /* Amarillo oro para detalles */
            --app-texto-titulos: #FFFFFF;/* Blanco para máxima lectura */
            --app-texto-datos: #CBD5E1;  /* Gris claro para contenido */
            --fondo-sitio: #11151c;      /* Fondo exterior para simular tu web */
        }

        /* 2. Reseteo y Estilo Base del Body */
        body {
            /* Simula el fondo de la página donde incrustarás la app */
            background-color: var(--fondo-sitio); 
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            display: flex;
            justify-content: center;
            align-items: center;
            min-height: 100vh;
            margin: 0;
            padding: 20px;
            box-sizing: border-box;
        }

        /* 3. Contenedor Principal de tu App */
        .app-contenedor {
            background-color: var(--app-fondo);
            border: 1px solid var(--app-borde);
            border-radius: 10px;
            padding: 30px;
            width: 100%;
            max-width: 600px; /* Ancho máximo de tu app */
            box-shadow: 0 4px 15px rgba(0, 0, 0, 0.3);
            color: var(--app-texto-titulos);
        }

        /* 4. Tipografía de la App */
        .app-titulo {
            margin-top: 0;
            color: var(--app-borde);
            font-size: 1.5rem;
            margin-bottom: 10px;
            font-weight: 600;
        }

        .app-texto-secundario {
            color: var(--app-texto-datos);
            font-size: 0.95rem;
            margin-bottom: 25px;
            line-height: 1.5;
        }

        /* 5. Elementos de Interfaz (Ejemplo de formulario) */
        .app-formulario {
            display: flex;
            gap: 12px;
            flex-wrap: wrap;
        }

        .app-input {
            flex: 1;
            min-width: 200px;
            padding: 12px;
            border-radius: 6px;
            border: 1px solid rgba(203, 213, 225, 0.3); /* Borde gris sutil */
            background-color: rgba(0, 0, 0, 0.2); /* Fondo ligeramente más oscuro */
            color: var(--app-texto-titulos);
            font-size: 1rem;
            transition: border-color 0.3s;
        }

        .app-input::placeholder {
            color: rgba(203, 213, 225, 0.6);
        }

        .app-input:focus {
            outline: none;
            border-color: var(--app-borde);
        }

        /* 6. Estilo de los Botones */
        .app-boton {
            background-color: transparent;
            border: 1px solid var(--app-borde);
            color: var(--app-borde);
            border-radius: 6px;
            padding: 12px 24px;
            cursor: pointer;
            font-weight: 600;
            font-size: 1rem;
            transition: all 0.3s ease;
        }

        .app-boton:hover {
            background-color: var(--app-borde);
            color: var(--app-fondo); /* El texto se vuelve oscuro al pasar el mouse */
        }
    </style>
</head>
<body>

    <div class="app-contenedor">
        <h2 class="app-titulo">Panel de la Aplicación</h2>
        <p class="app-texto-secundario">Este es el espacio de tu app incrustada. El fondo, los bordes y la tipografía ya coinciden exactamente con el estilo de tu portal.</p>
        
        <div class="app-formulario">
            <input type="text" class="app-input" placeholder="Ingresa un dato...">
            <button class="app-boton">Ejecutar Acción</button>
        </div>
    </div>
    </body>
</html>
