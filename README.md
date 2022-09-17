# downloadPhotosWithExcel
Código básico para descargar fotos de una web/portal con Excel mediante enlace referencial, la foto debe estar en la web (estilo secuencial propio).

Ejecución:

![Imgur](https://i.imgur.com/060vIHr.gif)

* Prerequisitos (solo para windows 7 - 10):

Office 2012-2019 (32/64 bit)


Instrucciones:

* Abrir o crear un Archhivo Excel Habilitado para macros.
* Crear tabla de contenido como se muestra a continuación:

![Imgur1](https://i.imgur.com/gvOJ8Ov.png)


* Importe el codigo como modulo al archivo Excel (.xlsm).

* En las referencias de codigo agragar las siguientes marcadas:

![Imgur2](https://i.imgur.com/YXZphpC.png)

* Seleccione los codigos en la columna trámite y ejecute la macro: "DOWNLOAD_PHOTOS_SELECTION" (antes de ejecutar, asegurese de que el archivo este guardado).

![Imgur3](https://i.imgur.com/060vIHr.gif)

Nota importante: La macro al ser ejecutada crea un arbol de carpetas predefinido de la siguiente manera:
 - En la ubicación actual del archivo crea la carpeta "DOWNLOAD"
 - A continuacion dentro de esta crea otra carpeta con el nombre de la hoja de calculo donde se realizó la seleccion para la descarga de fotos.
 - Finalmente en esta carpeta crea individualmente carpetas con fotos de la columna "Tramite" de la tabla previamente creada.
 - Las imagenes se crean con el nombre del encabezado donde son encontradas en la tabla y se asignan a la carpeta que se creó con el numero de tramite/código.
 
