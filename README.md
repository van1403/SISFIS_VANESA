# SISFIS_VANESA

Sistema de informes de fiscalización. Proyecto académico que usa solo datos ficticios.

Permite crear, editar, listar y eliminar informes, con vista previa en tiempo real. Cada informe se puede descargar en Word o imprimir / guardar en PDF. Los informes se guardan en el archivo `informes.json`, sin base de datos.

## Tecnologías

- Node.js
- Express
- HTML, CSS y JavaScript

## Paquetes

| Paquete | Versión | Para qué se usa |
|---|---|---|
| [express](https://expressjs.com) | ^5.3.0 | Servidor web: entrega la página y la API de informes (`/api/informes`) |

Además, el servidor usa los módulos `fs` y `path`, que vienen incluidos en Node.js y no se instalan aparte.

Las fuentes de letra (DM Sans, Fraunces y Poppins) se cargan desde Google Fonts, por lo que se necesita conexión a internet para verlas.

## Requisitos

- [Node.js](https://nodejs.org) instalado
- [Git](https://git-scm.com) instalado

## Cómo ejecutarlo

1. Clonar el repositorio:

   ```bash
   git clone <url-del-repositorio>/SISFIS_VANESA.git
   ```

2. Entrar a la carpeta del proyecto:

   ```bash
   cd SISFIS_VANESA
   ```

3. Instalar las dependencias:

   ```bash
   npm install
   npm install express
   ```

4. Iniciar el servidor:

   ```bash
   npm start
   ```

5. Abrir en el navegador: [http://localhost:3000](http://localhost:3000)

## Autora

Vanesa
