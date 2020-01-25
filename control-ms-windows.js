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
 *  0.3  Mic  + Support creating states under 0_userdata.0
 *  0.2  Mic  - Fix: disregard empty values in GETADMIN_COMMANDS_OWN
 *  0.1  Mic  - Initial Version
 * ---------------------------
 * Many thanks to Vladimir Vilisov for GetAdmin. Check out his website at
 * https://blog.instalator.ru/archives/47
 ***************************************************************************************/
/*
 * VORAUSSETZUNG:
 * In der Instanz des JavaScript-Adapters die Option [Erlaube das Kommando "setObject"] aktivieren.
 * Das ist notwendig, damit die Datenpunkte unterhalb von 0_userdata.0 angelegt werden können.
 * https://github.com/Mic-M/iobroker.createUserStates
 * Wer das nicht möchte: bitte Script-Version 0.2 verwenden.
 */

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
// Es wird die Anlage sowohl unterhalb '0_userdata.0' als auch 'javascript.x' unterstützt.
const STATE_PATH = '0_userdata.0.Computer.Control-PC';

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


/*************************************************************************************************************************
 * Global variables and constants
 *************************************************************************************************************************/

// Final state path
const FINAL_STATE_LOCATION = validateStatePath(STATE_PATH, false);
const FINAL_STATE_PATH = validateStatePath(STATE_PATH, true);

// Durch Get Admin unterstützte Commands
const GETADMIN_COMMANDS = ['process', 'shutdown', 'poweroff', 'reboot', 'forceifhung', 'logoff', 'monitor1', 'monitor2'];


/********************************************************************************
 * init - This is executed on every script (re)start.
 ********************************************************************************/
init();
function init() {

    // Create all script states
    createUserStates(FINAL_STATE_LOCATION, false, buildNeededStates(), function() {
        // -- All states created, so we continue by using callback

        // Subscribe to states
        setTimeout(doSubscriptions, 2000); // using delay since not all states seem to be ready right away.

    });

}

function doSubscriptions() {

    if (LOG_INFO) log('Start subscribing to state changes...');

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

            if( (ip != -1) ) {
                getAdminSendCommand(name, ip, 'key', obj.state.val);
            } else {
                log('No configration found for ' + stateDevice, 'warn');
            }
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
    
    //let request = require('request'); // Not needed: https://github.com/ioBroker/ioBroker.javascript/issues/471#issuecomment-571759380
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
 * Build an array of states we need to create.
 * @return {array} Array of states to be created. Format: see function createUserStates()
 */
function buildNeededStates() {

    let finalStates = [];

    for (let lpConfDevice of CONFIG_DEVICES) {

        let name = lpConfDevice['name'];
        let nameClean = cleanStringForState(name);
        let statePath = FINAL_STATE_PATH + '.' + nameClean;

        // Create Get Admin Command States
        for (let lpCommand of GETADMIN_COMMANDS) {
            finalStates.push([statePath + '.' + lpCommand, {'name':'Command: ' + lpCommand, 'type':'boolean', 'read':false, 'write':true, 'role':'button', 'def':false }]);
        }

        // Create User Specific Command States
        if (! isLikeEmpty(GETADMIN_COMMANDS_OWN)) {
            for (let lpCommand of cleanArray(GETADMIN_COMMANDS_OWN)) {
                finalStates.push([statePath + '.' + lpCommand, {'name':'User Command: ' + lpCommand, 'type':'boolean', 'read':false, 'write':true, 'role':'button', 'def':false }]);
            }
        }

        // Create State for sending a key
        finalStates.push([statePath + '.sendKey', {'name':'Send Key', 'type':'string', 'read':true, 'write':true, 'role':'state', 'def':'' }]);

    }

    return finalStates;    

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
 * Clean a given string for using in ioBroker as part of a atate
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
 * 08-Sep-2019: added check for [ and ] to also catch arrays with empty strings.
 * @param inputVar - Input Array or String, Number, etc.
 * @return true if it is undefined/null/empty, false if it contains value(s)
 * Array or String containing just whitespaces or >'< or >"< or >[< or >]< is considered empty
 */
function isLikeEmpty(inputVar) {
    if (typeof inputVar !== 'undefined' && inputVar !== null) {
        let strTemp = JSON.stringify(inputVar);
        strTemp = strTemp.replace(/\s+/g, ''); // remove all whitespaces
        strTemp = strTemp.replace(/\"+/g, "");  // remove all >"<
        strTemp = strTemp.replace(/\'+/g, "");  // remove all >'<
        strTemp = strTemp.replace(/\[+/g, "");  // remove all >[<
        strTemp = strTemp.replace(/\]+/g, "");  // remove all >]<
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


/**
 * For a given state path, we extract the location '0_userdata.0' or 'javascript.0' or add '0_userdata.0', if missing.
 * @param {string}  path            Like: 'Computer.Control-PC', 'javascript.0.Computer.Control-PC', '0_userdata.0.Computer.Control-PC'
 * @param {boolean} returnFullPath  If true: full path like '0_userdata.0.Computer.Control-PC', if false: just location like '0_userdata.0' or 'javascript.0'
 * @return {string}                 Path
 */
function validateStatePath(path, returnFullPath) {
    if (path.startsWith('.')) path = path.substr(1);    // Remove first dot
    if (path.endsWith('.'))   path = path.slice(0, -1); // Remove trailing dot
    if (path.length < 1) log('Provided state path is not valid / too short.', 'error')
    let match = path.match(/^((javascript\.([1-9][0-9]|[0-9])\.)|0_userdata\.0\.)/);
    let location = (match == null) ? '0_userdata.0' : match[0].slice(0, -1); // default is '0_userdata.0'.
    if(returnFullPath) {
        return (path.indexOf(location) == 0) ? path : (location + '.' + path);
    } else {
        return location;
    }
}


/**
 * Create states under 0_userdata.0 or javascript.x
 * Current Version:     https://github.com/Mic-M/iobroker.createUserStates
 * Support:             https://forum.iobroker.net/topic/26839/
 * Autor:               Mic (ioBroker) | Mic-M (github)
 * Version:             1.0 (17 January 2020)
 * Example:
 * -----------------------------------------------
    let statesToCreate = [
        ['Test.Test1', {'name':'Test 1', 'type':'string', 'read':true, 'write':true, 'role':'info', 'def':'Hello' }],
        ['Test.Test2', {'name':'Test 2', 'type':'string', 'read':true, 'write':true, 'role':'info', 'def':'Hello' }],
    ];
    createUserStates('0_userdata.0', false, statesToCreate);
 * -----------------------------------------------
 * PLEASE NOTE: Per https://github.com/ioBroker/ioBroker.javascript/issues/474, the used function setObject() 
 *              executes the callback PRIOR to completing the state creation. Therefore, we use a setTimeout and counter. 
 * -----------------------------------------------
 * @param {string} where          Where to create the state: e.g. '0_userdata.0' or 'javascript.x'.
 * @param {boolean} force         Force state creation (overwrite), if state is existing.
 * @param {array} statesToCreate  State(s) to create. single array or array of arrays
 * @param {object} [callback]     Optional: a callback function -- This provided function will be executed after all states are created.
 */
function createUserStates(where, force, statesToCreate, callback = undefined) {
 
    const WARN = false; // Throws warning in log, if state is already existing and force=false. Default is false, so no warning in log, if state exists.
    const LOG_DEBUG = false; // To debug this function, set to true
    // Per issue #474 (https://github.com/ioBroker/ioBroker.javascript/issues/474), the used function setObject() executes the callback 
    // before the state is actual created. Therefore, we use a setTimeout and counter as a workaround.
    // Increase this to 100, if it is not working.
    const DELAY = 50; // Delay in milliseconds (ms)


    // Validate "where"
    if (where.endsWith('.')) where = where.slice(0, -1); // Remove trailing dot
    if ( (where.match(/^javascript.([0-9]|[1-9][0-9])$/) == null) && (where.match(/^0_userdata.0$/) == null) ) {
        log('This script does not support to create states under [' + where + ']', 'error');
        return;
    }

    // Prepare "statesToCreate" since we also allow a single state to create
    if(!Array.isArray(statesToCreate[0])) statesToCreate = [statesToCreate]; // wrap into array, if just one array and not inside an array

    let numStates = statesToCreate.length;
    let counter = -1;
    statesToCreate.forEach(function(param) {
        counter += 1;
        if (LOG_DEBUG) log ('[Debug] Currently processing following state: [' + param[0] + ']');

        // Clean
        let stateId = param[0];
        if (! stateId.startsWith(where)) stateId = where + '.' + stateId; // add where to beginning of string
        stateId = stateId.replace(/\.*\./g, '.'); // replace all multiple dots like '..', '...' with a single '.'
        const FULL_STATE_ID = stateId;

        if( ($(FULL_STATE_ID).length > 0) && (existsState(FULL_STATE_ID)) ) { // Workaround due to https://github.com/ioBroker/ioBroker.javascript/issues/478
            // State is existing.
            if (WARN && !force) log('State [' + FULL_STATE_ID + '] is already existing and will no longer be created.', 'warn');
            if (!WARN && LOG_DEBUG) log('[Debug] State [' + FULL_STATE_ID + '] is already existing. Option force (=overwrite) is set to [' + force + '].');

            if(!force) {
                // State exists and shall not be overwritten since force=false
                // So, we do not proceed.
                numStates--;
                if (numStates === 0) {
                    if (LOG_DEBUG) log('[Debug] All states successfully processed!');
                    if (typeof callback === 'function') { // execute if a function was provided to parameter callback
                        if (LOG_DEBUG) log('[Debug] An optional callback function was provided, which we are going to execute now.');
                        return callback();
                    }
                } else {
                    // We need to go out and continue with next element in loop.
                    return; // https://stackoverflow.com/questions/18452920/continue-in-cursor-foreach
                }
            } // if(!force)
        }

        /************
         * State is not existing or force = true, so we are continuing to create the state through setObject().
         ************/
        let obj = {};
        obj.type = 'state';
        obj.native = {};
        obj.common = param[1];
        setObject(FULL_STATE_ID, obj, function (err) {
            if (err) {
                log('Cannot write object for state [' + FULL_STATE_ID + ']: ' + err);
            } else {
                if (LOG_DEBUG) log('[Debug] Now we are creating new state [' + FULL_STATE_ID + ']')
                let init = null;
                if(param[1].def === undefined) {
                    if(param[1].type === 'number') init = 0;
                    if(param[1].type === 'boolean') init = false;
                    if(param[1].type === 'string') init = '';
                } else {
                    init = param[1].def;
                }
                setTimeout(function() {
                    setState(FULL_STATE_ID, init, true, function() {
                        if (LOG_DEBUG) log('[Debug] setState durchgeführt: ' + FULL_STATE_ID);
                        numStates--;
                        if (numStates === 0) {
                            if (LOG_DEBUG) log('[Debug] All states processed.');
                            if (typeof callback === 'function') { // execute if a function was provided to parameter callback
                                if (LOG_DEBUG) log('[Debug] Function to callback parameter was provided');
                                return callback();
                            }
                        }
                    });
                }, DELAY + (20 * counter) );
            }
        });
    });
}
