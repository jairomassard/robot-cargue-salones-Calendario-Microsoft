# robot-cargue-salones-Calendario-Microsoft
Programa para automatizar el cargue de información de Clases y próximas clases en etiquetas informativas ESL tomando info de Microsoft Calendars 360

El programa esta elaborado en Node.JS

Este programa tiene por objetivo automatizar la obtencion de información de las clases que se dictan en un establecimiento educativo o de los eventos que se ofrecen en salones de eventos, alimentando y mostrando en pantallas informativas de papel electronico ESL.  En el calendario de Microsoft 360 se programan diferentes salones o salas de eventos. Cada sala o salón tiene su propio calendario bajo una misma cuenta.   

Se conecta con la API de Microsoft 360 para extraer la información de eventos en el calendario y se conecta con API al software cloud del fabricante de las pantallas informativa ESL, lugar donde se configuran los campos de la Base de datos que se utilizará, se diseña la plantilla de presentación de la información en las pantallas y se administran y gestionan los equipos y se verifica el funcionamiento del HW.

Cada vez que se ejecuta el programa verifica el calendario Microsoft, toma la información que esta vigente y por venir (2 eventos siguientes) sobre eventos en el calendario, y los envia al API del fabricante de las etiquetas ESL para que luego este ultimo se encargue de mandar a cargar la información recibida en las pantallas de los salones a los que se le pidio consultar y tomar la información.
