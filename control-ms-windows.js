/***************************************************************************************
 * Script to control Windows PCs
 * -------------------------------------------------------------------------------------
 * Send commands to Windows PCs for shutdown, hibernate, etc.
 * Source: https://forum.iobroker.net/topic/1570/windows-steuerung and https://blog.instalator.ru/archives/47
 * 
 * Aktuelle Version:    https://github.com/Mic-M/iobroker.control-ms-windows
 * Support:             https://forum.iobroker.net/topic/1570/windows-steuerung
 * ---------------------------
 * Change Log:
 *  0.1 Mic  - Initial Version
 * ---------------------------
 * Many thanks to Vladimir Vilisov for GetAdmin. Check out his website at
 * https://blog.instalator.ru/archives/47
 ***************************************************************************************/

/*******************************************************************************
 * Zur Einrichtung von GetAdmin
 ******************************************************************************/
/*
 * 1) Software "GetAdmin" (getestet Version 2.6) auf Zielrechner installieren.
 *    Link: https://blog.instalator.ru/archives/47
 * 2) In GetAdmin, ganz oben links unter "Server": 
 *     - IP: die IP-Adresse der ioBrokers eintragen
 *     - Port: Standard-Port 8585 so lassen
 * 3) In GetAdmin, oben unter "Options" Haken bei Minimize und Startup setzen, 
 *    damit sich GetAdmin bei jedem Rechnerstart startet und das minimiert.
 *    Dann mit "Save" bestätigen.
 * 4) Fertig
 * ------------------------------------------------------------------------------
 * Beispiele für individuelle Einträge in GetAdmin Command list:
 *     a) Ruhezustand: 
 *         - in Spalte 'Command' z.B. "m_hibernate" eintragen
 *         - in Spalte 'PATH or URL' eintragen: shutdown
 *         - in Spalte 'PARAMETERS' eintragen: -h
 *     b) Energie sparen:
 *         - in Spalte 'Command' z.B. "m_sleep" eintragen
 *         - in Spalte 'PATH or URL' eintragen: rundll32.exe
 *         - in Spalte 'PARAMETERS' eintragen: powrprof.dll,SetSuspendState
 */



/*******************************************************************************
 * Konfiguration: Pfade
 ******************************************************************************/
// Pfad, unter dem die States (Datenpunkte) in den Objekten angelegt werden.
// Kann man so bestehen lassen.
const STATE_PATH = 'javascript.'+ instance + '.' + 'Control-PC';

/*******************************************************************************
 * Konfiguration: Geräte
 ******************************************************************************/
// Hier deine Geräte aufnehmen. Du kannst beliebig viele ergänzen.
const CONFIG_DEVICES = [
    {
        name: 'PC-John',   // Für Datenpunkt und Ausgabe
        ip:   '192.168.0.101',
    },
    {
        name: 'Gästezimmer-PC',
        ip:   '10.10.0.102',
    },
];

/*******************************************************************************
 * Konfiguration: Get Admin Commands
 ******************************************************************************/
// Eigene Commands, die in Get Admin in der Command List eingetragen sind, Spalte "Command"
// Bitte ohne Leerzeichen, Sonderzeichen, etc.
// Falls keine eigenen Commands: GETADMIN_COMMANDS_OWN = [];
const GETADMIN_COMMANDS_OWN = ['m_hibernate', 'm_sleep', '', ''];


/*******************************************************************************
 * Konfiguration: Konsolen-Ausgaben
 ******************************************************************************/
// Auf true setzen, wenn ein paar Infos dieses Scripts im Log ausgegeben werden dürfen.
const LOG_INFO = true;

// Auf true setzen, wenn zur Fehlersuche einige Meldungen ausgegeben werden sollen.
// Ansonsten bitte auf false stellen.
const LOG_DEBUG = false;



/*************************************************************************************************************************
 * Ab hier nichts mehr ändern / Stop editing here!
 *************************************************************************************************************************/

/********************************************************************************
 * Durch Get Admin unterstützte Commands
 ********************************************************************************/
const GETADMIN_COMMANDS = ['process', 'shutdown', 'poweroff', 'reboot', 'forceifhung', 'logoff', 'monitor1', 'monitor2'];


/********************************************************************************
 * init - This is executed on every script (re)start.
 ********************************************************************************/
init();
function init() {
    
    // Create our states, if not yet existing.
    createStates();

    // States should have been created, so continue
    setTimeout(function(){    

        // Subscribe to states
        doSubscriptions();

    }, 2000);

}

function doSubscriptions() {

    // Loop through the devices
    for (let lpConfDevice of CONFIG_DEVICES) {

        let name = lpConfDevice['name'];
        let statePath = STATE_PATH + '.' + name;

        /*****************
         * Loop through the commands to subscribe accordingly
         *****************/
        let allCommands = cleanArray([].concat(GETADMIN_COMMANDS, GETADMIN_COMMANDS_OWN)); // merge both into one array
        for (let lpCommand of allCommands) {

            on({id: statePath + '.' + lpCommand, change: 'any', val: true}, function(obj) {

                // First: Get the device + command state portion of obj.id, as variable is not available within "on({id..."
                let stateFull = obj.id // e.g. [javascript.0.Control-PC.PC-Maria.shutdown]
                let stateDeviceAndCommand = stateFull.substring(STATE_PATH.length +1); // e.g. [PC-Maria.shutdown]
                let stateDeviceAndCommandSplit = stateDeviceAndCommand.split('.');
                let stateDevice = stateDeviceAndCommandSplit[0];    // e.g. [PC-Maria]
                let stateCommand = stateDeviceAndCommandSplit[1];   // e.g. [shutdown]

                // Next, get the ip
                let ip = getConfigValuePerKey(CONFIG_DEVICES, 'name', stateDevice, 'ip');

                if( (ip != -1) ) {
                    getAdminSendCommand(name, ip, 'cmd', stateCommand);
                } else {
                    log('No configration found for ' + stateDevice, 'warn');
                }
            });
        }

        /*****************
         * Also subscribe to "sendKey"
         *****************/
        on({id: statePath + '.sendKey', change: 'any'}, function(obj) {

            // First: Get the device + command state portion of obj.id, as variable is not available within "on({id..."
            let stateFull = obj.id // e.g. [javascript.0.Control-PC.PC-Maria.sendKey]
            let stateDeviceAndCommand = stateFull.substring(STATE_PATH.length +1); // e.g. [PC-Maria.sendKey]
            let stateDeviceAndCommandSplit = stateDeviceAndCommand.split('.');
            let stateDevice = stateDeviceAndCommandSplit[0];    // e.g. [PC-Maria]

            // Next, get the ip
            let ip = getConfigValuePerKey(CONFIG_DEVICES, 'name', stateDevice, 'ip');

            getAdminSendCommand(name, ip, 'key', obj.state.val);

        });

    }
}


/* 
 * @param {string}  name     Name des Rechners, nur für Log-Ausgabe
 * @param {string}  host     IP-Adresse des Windows-PCs, z.B. 10.10.0.107
 * @param {string}  action   If command, use 'cmd', if key, use 'key', etc.
 * @param {string}  command  Userspezifischer Command wie z.B. "m_hibernate", oder "poweroff"
 */
function getAdminSendCommand(name, host, action, command){
    
    let request = require('request');
    let options = { url: 'http://' + host + ':' + '8585' + '/?' + action + '=' + command };

    if (LOG_DEBUG) log('Send command to ' + name + ': ' + options.url);
    if (LOG_INFO) log('Send command [' + command + '] to ' + name); 
    request(options, function (error, response, body) {
        if ( (response !== undefined) && !error ) {
            if ( parseInt(response.statusCode) === 200 ) {
                if (LOG_INFO) log(name + ' responds with [OK]'); 
            } else {
                if (LOG_INFO) log(name + ' responds with unexpected status code [' + response.statusCode + ']');
            }
        } else {
            if (LOG_INFO) log('No response from ' + name + ', so it seems to be off.'); 
        }
    });
}


/**
 * Create script states
 */
function createStates() {

    for (let lpConfDevice of CONFIG_DEVICES) {

        let name = lpConfDevice['name'];
        let nameClean = cleanStringForState(name);
        let statePath = STATE_PATH + '.' + nameClean;

        // Create Get Admin Command States
        for (let lpCommand of GETADMIN_COMMANDS) {
            createState(statePath + '.' + lpCommand, {'name':'Command: ' + lpCommand, 'type':'boolean', 'read':false, 'write':true, 'role':'button', 'def':false });
        }

        // Create User Specific Command States
        if (! isLikeEmpty(GETADMIN_COMMANDS_OWN)) {
            for (let lpCommand of GETADMIN_COMMANDS_OWN) {
                createState(statePath + '.' + lpCommand, {'name':'User Command: ' + lpCommand, 'type':'boolean', 'read':false, 'write':true, 'role':'button', 'def':false });
            }
        }

        // Create State for sending a key
        createState(statePath + '.sendKey', {'name':'Send Key', 'type':'string', 'read':true, 'write':true, 'role':'state', 'def':'' });


    }

}



/**
 * Retrieve values from a CONFIG variable, example:
 * const CONF = [{car: 'bmw', color: 'black', hp: '250'}, {car: 'audi', color: 'blue', hp: '190'}]
 * To get the color of the Audi, use: getConfigValuePerKey('car', 'bmw', 'color')
 * To find out which car has 190 hp, use: getConfigValuePerKey('hp', '190', 'car')
 * @param {object}  config     The configuration variable/constant
 * @param {string}  key1       Key to look for.
 * @param {string}  key1Value  The value the key should have
 * @param {string}  key2       The key which value we return
 * @returns {any}    Returns the element's value, or number -1 of nothing found.
 */
function getConfigValuePerKey(config, key1, key1Value, key2) {
    // We need to get all ids of LOG_FILTER into array
    for (let lpConfDevice of config) {
        if ( lpConfDevice[key1] === key1Value ) {
            if (lpConfDevice[key2] === undefined) {
                return -1;
            } else {
                return lpConfDevice[key2];
            }
        }
    }
    return -1;
}



/**
 * Will just keep letters, incl. Umlauts, numbers, "-" and "_" and "."
 * @param  {string}  strInput   Input String
 * @return {string}   the processed string 
 */
function cleanStringForState(strInput) {
    let strResult = strInput.replace(/([^a-zA-ZäöüÄÖÜß0-9\-\._]+)/gi, '');
    return strResult;
}

/**
 * Checks if Array or String is not undefined, null or empty.
 * @param inputVar - Input Array or String, Number, etc.
 * @return true if it is undefined/null/empty, false if it contains value(s)
 * Array or String containing just whitespaces or >'< or >"< is considered empty
 */
function isLikeEmpty(inputVar) {
    if (typeof inputVar !== 'undefined' && inputVar !== null) {
        let strTemp = JSON.stringify(inputVar);
        strTemp = strTemp.replace(/\s+/g, ''); // remove all whitespaces
        strTemp = strTemp.replace(/\"+/g, "");  // remove all >"<
        strTemp = strTemp.replace(/\'+/g, "");  // remove all >'<
        if (strTemp !== '') {
            return false;
        } else {
            return true;
        }
    } else {
        return true;
    }
}

/**
 * Clean Array: Removes all falsy values: undefined, null, 0, false, NaN and "" (empty string)
 * Source: https://stackoverflow.com/questions/281264/remove-empty-elements-from-an-array-in-javascript
 * @param {array} inputArray       Array to process
 * @return {array}  Cleaned array
 */
function cleanArray(inputArray) {
  var newArray = [];
  for (let i = 0; i < inputArray.length; i++) {
    if (inputArray[i]) {
      newArray.push(inputArray[i]);
    }
  }
  return newArray;
}