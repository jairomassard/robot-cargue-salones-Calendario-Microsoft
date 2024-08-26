global.fetch = require('node-fetch');
const { Client } = require('@microsoft/microsoft-graph-client');
require('dotenv').config();
const fs = require('fs');
const md5 = require('md5');
const fetch = require('node-fetch');
const moment = require('moment-timezone');
const { ClientSecretCredential } = require('@azure/identity');
const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

const LOCAL_TIMEZONE = 'America/Bogota';

const configPath = 'C:/Prog_Cargue_Salones_Microsoft/configuraciones/config.json';
const config = JSON.parse(fs.readFileSync(configPath, 'utf8'));

const rutaDirectorioLog = config.logDirectoryPath;
const rutaArchivoLog = `${rutaDirectorioLog}/log.txt`;

async function getAuthenticatedClient() {
    const credential = new ClientSecretCredential(config.tenantId, config.clientId, config.clientSecret);
    const authProvider = new TokenCredentialAuthenticationProvider(credential, { scopes: ['https://graph.microsoft.com/.default'] });
    return Client.initWithMiddleware({ authProvider });
}

function crearLog(contenido) {
    if (!fs.existsSync(rutaDirectorioLog)) {
        fs.mkdirSync(rutaDirectorioLog, { recursive: true });
    }
    // Obtener el timestamp en la zona horaria local
    const localTimestamp = moment().tz(LOCAL_TIMEZONE).format();
    fs.appendFile(rutaArchivoLog, `${contenido};Éxito;${localTimestamp}\n`, 'utf8', (err) => {
        if (err) {
            console.error('Error al escribir en el archivo de log:', err);
        } else {
            console.log('Log guardado!');
        }
    });
}

function enviarYalaInfo(transactionObj) {
    const fecha = moment().add(13, 'hours').tz(LOCAL_TIMEZONE);
    const formatofecha = fecha.format('YYYY-MM-DD');
    const hash1 = md5(config.merchantKey + formatofecha);

    const myHeaders = {
        'veryText': hash1,
        'merchantCode': config.merchantCode,
        'type': '1',
        'Content-Type': 'application/json'
    };

    const raw = JSON.stringify([transactionObj]);
    const requestOptions = {
        method: 'POST',
        headers: myHeaders,
        body: raw,
        redirect: 'follow',
    };

    fetch('https://sg.yalabi.net/open/saveOrGoods', requestOptions)
        .then(response => response.json())
        .then(result => {
            console.log('Registro enviado con éxito:', result);
            console.log(JSON.stringify(transactionObj, null, 2)); // Mejora la visualización del objeto en la consola
            crearLog(`${JSON.stringify(transactionObj)};Éxito;${moment().toISOString()}`);
        })
        .catch(error => {
            console.error('Error al enviar el registro:', error);
            crearLog(`${JSON.stringify(transactionObj)};Error;${moment().toISOString()}`);
        });
}

async function procesarEventos(events, salonName) {
    const now = moment().tz(LOCAL_TIMEZONE);

    const eventoActual = events.find(event => now.isBetween(moment(event.start.dateTime).tz(LOCAL_TIMEZONE), moment(event.end.dateTime).tz(LOCAL_TIMEZONE)));
    const eventosFuturos = events.filter(event => moment(event.start.dateTime).tz(LOCAL_TIMEZONE).isAfter(now));
    const proximoEvento = eventosFuturos[0];
    const segundoProximoEvento = eventosFuturos[1];

    const transactionObj = {
        itemName: salonName,
        itemBarCode: salonName,
        merchantGoodsId: salonName,
        merchantGoodsCategoryId: salonName,
        categoryName: 'Salones de Clase',
        reservedField1: eventoActual ? eventoActual.start.dateTime.split('T')[0] : '',
        reservedField2: eventoActual ? eventoActual.start.dateTime.split('T')[1].slice(0, 5) : '',
        reservedField3: eventoActual ? eventoActual.end.dateTime.split('T')[1].slice(0, 5) : '',
        reservedField4: eventoActual ? eventoActual.subject : '',
        reservedField5: eventoActual ? eventoActual.bodyPreview : '', // Usamos bodyPreview
        reservedField6: eventoActual ? 'OCUPADO' : 'DISPONIBLE',
        reservedField7: proximoEvento ? proximoEvento.start.dateTime.split('T')[0] : '',
        reservedField8: proximoEvento ? proximoEvento.start.dateTime.split('T')[1].slice(0, 5) : '',
        reservedField9: proximoEvento ? proximoEvento.end.dateTime.split('T')[1].slice(0, 5) : '',
        reservedField10: proximoEvento ? proximoEvento.subject : '',
        reservedField11: segundoProximoEvento ? segundoProximoEvento.start.dateTime.split('T')[0] : '',
        reservedField12: segundoProximoEvento ? segundoProximoEvento.start.dateTime.split('T')[1].slice(0, 5) : '',
        reservedField13: segundoProximoEvento ? segundoProximoEvento.end.dateTime.split('T')[1].slice(0, 5) : '',
        reservedField14: segundoProximoEvento ? segundoProximoEvento.subject : '',
    };

    enviarYalaInfo(transactionObj);
}

async function procesarCalendarios() {
    const client = await getAuthenticatedClient();
    for (const [salonName, calendarId] of Object.entries(config.calendarIds)) {
        try {
            const result = await client.api(`/users/${config.userId}/calendars/${calendarId}/events`)
                .header('Prefer', `outlook.timezone="${LOCAL_TIMEZONE}"`)
                .select('subject,start,end,bodyPreview') // Cambiamos 'body' a 'bodyPreview'
                .orderby('start/dateTime')
                .get();
            if (result.value.length > 0) {
                await procesarEventos(result.value, salonName);
            } else {
                console.log(`No se encontraron eventos próximos en el calendario: ${salonName}`);
            }
        } catch (error) {
            console.error(`Error al obtener eventos del calendario ${salonName}:`, error);
        }
    }
}

procesarCalendarios();
