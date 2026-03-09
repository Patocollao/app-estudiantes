/* Paleta de colores para la App */
:root {
    --app-fondo: #1B222C;        /* Fondo principal azul noche */
    --app-borde: #F6C128;        /* Amarillo oro para detalles */
    --app-texto-titulos: #FFFFFF;/* Blanco para máxima lectura */
    --app-texto-datos: #CBD5E1;  /* Gris claro para el resto del contenido */
}

/* Contenedor principal de tu App incrustada */
.app-contenedor {
    background-color: var(--app-fondo);
    border: 1px solid var(--app-borde);
    border-radius: 10px;
    color: var(--app-texto-titulos);
    padding: 24px;
    /* Sombra sutil para que resalte un poco del fondo de la página */
    box-shadow: 0 4px 15px rgba(0, 0, 0, 0.3); 
}

/* Estilo para los botones o elementos interactivos dentro de tu App */
.app-boton {
    background-color: transparent;
    border: 1px solid var(--app-borde);
    color: var(--app-borde);
    border-radius: 8px;
    padding: 10px 20px;
    cursor: pointer;
    transition: background-color 0.3s;
}

.app-boton:hover {
    background-color: rgba(246, 193, 40, 0.1); /* Fondo amarillo transparente al pasar el mouse */
}

/* Textos secundarios dentro de la app */
.app-texto-secundario {
    color: var(--app-texto-datos);
    font-size: 0.95rem;
}
