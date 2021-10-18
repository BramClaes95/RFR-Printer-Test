//
// Copyright (c) 2016 Ricoh Co., Ltd. All rights reserved.
//
// Abstract:
//    RICOH V4 Printer Driver
//
var bidi = {};

bidi.writePJL = function(printerStream, pjl) {
  var UEL = [0x1B, 0x25, 0x2D, 0x31, 0x32, 0x33, 0x34, 0x35, 0x58];
  var HEADER = [0x40, 0x50, 0x4A, 0x4C, 0x0D, 0x0A];
  var i = 0;
  var writeData = [].concat(UEL).concat(HEADER);

  for (i = 0; i < pjl.length; i++) {
    writeData.push(pjl.charCodeAt(i));
  }
  writeData = writeData.concat(UEL);

  var bytesWritten = printerStream.write(writeData);
  return bytesWritten === writeData.length
}

bidi.readPJL = function(printerStream) {
  var readBuffer = [];
  var bytesRead = 0;
  var readString = "";
  var readEnd = false;

  while (!readEnd) {
    var buffer = printerStream.read(2048);

    readBuffer = readBuffer.concat(buffer);
    bytesRead += buffer.length;

    if (buffer.length === 0) {
      readEnd = true;
    }
    else if (buffer.length < 2048) {
      if (buffer[buffer.length - 1] === 0x0C) {
        readEnd = true;
      }
    }
  }

  for (var i = 0; (bytesRead > 0) && (i < bytesRead); i++) {
    readString += String.fromCharCode(readBuffer.shift());
  }
  return readString;
}

bidi.allInclude = function(source, target) {
  if( !source || !target || target === "\r\n" ) return false;
  var commands = target.split("\r\n");
  for(var i = 0; i < commands.length; i++ ) {
    if( commands[i] && source.indexOf(commands[i]) === -1 ) {
      return false;
    }
  }
  return true;
}

bidi.getPJL = function(printerStream, request) {
  if (!this.writePJL(printerStream, request)) return "";
  var result = "";
  var response = "";
  do {
    response = this.readPJL(printerStream);
    if( response ) {
      result += response;
    }
  } while(response && !this.allInclude(result, request));
  return result;
}

bidi.trimspace = function(value) {
  return value.replace(/^\s+|\s+$/g, '');
}

bidi.queryID = function(tray) {
  var trimmedValue = this.trimspace(tray);
  return trimmedValue === "INTRAYMANUAL" ? "INTRAYM" : trimmedValue;
}

bidi.queryIDController7 = function(tray) {
  var trimmedValue = this.trimspace(tray);
  if (!trimmedValue) return "";
  return trimmedValue === "TRAY4" ? "INTRAYM" : "IN" + trimmedValue;
}

bidi.lockQueryID = function(tray) {
  var trimmedValue = this.trimspace(tray);
  return trimmedValue === "INTRAYMANUAL" ? "INTRAYMULTI" : trimmedValue;
}

bidi.getChildNodeList = function(key, pjl) {
  var re = new RegExp(key + "\\s*\\[\\d+\\s+ENUMERATED\\s*\\]\\s*\\r\\n(?:^\\s+.+\\r\\n)+", "m");
  var list = pjl.match(re);
  if (!list) return null;
  return list[0].split("\r\n").slice(1,-1);
}

bidi.getTrayQueryCommand = function(pjl) {
  var nodelist = this.getChildNodeList("IN TRAYS", pjl);
  if(!nodelist) return "";
  var result = "";
  for (var i = 0; i < nodelist.length; i++) {
    result += "@PJL INQUIRE " + this.queryID(nodelist[i]) + "SIZE\r\n";
    result += "@PJL INQUIRE " + this.queryID(nodelist[i]) + "MEDIA\r\n";
    result += "@PJL INQUIRE " + this.lockQueryID(nodelist[i]) + "\r\n";
  }
  return result;
}

bidi.getTrayQueryCommandController7 = function (pjl) {
  var nodelist = this.getChildNodeList("IN TRAY", pjl);
  if (!nodelist) return "";
  var result = "";
  for (var i = 0; i < nodelist.length; i++) {
    result += "@PJL INQUIRE " + this.queryIDController7(nodelist[i]) + "SIZE\r\n";
  }
  return result;
}

bidi.getCommonInfo = function(printerStream, printerBidiSchemaResponses) {
  var queryCommonInfo = "@PJL INFO CONFIG\r\n";
  var infoResponse = this.getPJL(printerStream, queryCommonInfo);
  if (!infoResponse) return;
  var queryTray = this.getTrayQueryCommand(infoResponse);
  if (!queryTray) return;
  var trayResponse = this.getPJL(printerStream, queryTray);
  if (!trayResponse) return;
  printerBidiSchemaResponses.addString("\\Printer.Configuration.CommonInfo:data", infoResponse + trayResponse);
  printerBidiSchemaResponses.addString("\\Printer.Configuration.CommonInfo:id", Math.random().toString());
}

bidi.getCommonInfoCommandController7 = function (printerStream, printerBidiSchemaResponses) {
  var queryCommonInfo = "@PJL INFO CONFIG\r\n";
  var infoResponse = this.getPJL(printerStream, queryCommonInfo);
  if (!infoResponse) return;
  var queryTray = this.getTrayQueryCommandController7(infoResponse);
  if (!queryTray) return;
  var trayResponse = this.getPJL(printerStream, queryTray);
  if (!trayResponse) return;
  printerBidiSchemaResponses.addString("\\Printer.Configuration.CommonInfoController7:data", infoResponse + trayResponse);
  printerBidiSchemaResponses.addString("\\Printer.Configuration.CommonInfoController7:id", Math.random().toString());
}

bidi.getInfo = function(printerStream, query, responseKey, printerBidiSchemaResponses) {
  var info = this.getPJL(printerStream, query);
  if(!info) return;
  printerBidiSchemaResponses.addString(responseKey + ":data", info);
  printerBidiSchemaResponses.addString(responseKey + ":id", Math.random().toString());
}

bidi.getGenericInfo = function(printerStream, printerBidiSchemaResponses) {
  var queryGenericInfo = "@PJL INQUIRE DATAMODE\r\n@PJL INQUIRE STAPLE\r\n";
  return bidi.getInfo(printerStream, queryGenericInfo, "\\Printer.Configuration.GenericInfo", printerBidiSchemaResponses );
}

bidi.getGenericFunctionInfo = function(printerStream, printerBidiSchemaResponses) {
  var queryGenericFunctionInfo = "@PJL INFO GENERICFUNCKEY\r\n";
  return bidi.getInfo(printerStream, queryGenericFunctionInfo, "\\Printer.Configuration.GenericFunctionInfo", printerBidiSchemaResponses );
}

bidi.getPostScriptInfo = function(printerStream, printerBidiSchemaResponses) {
  var queryPostScriptInfo = "@PJL INFO POSTSCRIPT\r\n";
  return bidi.getInfo(printerStream, queryPostScriptInfo, "\\Printer.Configuration.PostScriptInfo", printerBidiSchemaResponses );
}

function getSchemas(scriptContext, printerStream, schemaRequests, printerBidiSchemaResponses) {
  debugger;
  var index = 0;
  for (index = 0; index < schemaRequests.length; index++) {
    if (schemaRequests[index] === "\\Printer.Cache:data") return 0;
  }
  for (index = 0; index < schemaRequests.length; index++) {
    var key = schemaRequests[index];
    switch (key) {
      case "CommonInfo":
        bidi.getCommonInfo(printerStream, printerBidiSchemaResponses);
        break;
      case "GenericInfo":
        bidi.getGenericInfo(printerStream, printerBidiSchemaResponses);
        break;
      case "GenericFunctionInfo":
        bidi.getGenericFunctionInfo(printerStream, printerBidiSchemaResponses);
        break;
      case "PostScriptInfo":
        bidi.getPostScriptInfo(printerStream, printerBidiSchemaResponses);
        break;
      case "CommonInfoController7":
        bidi.getCommonInfoCommandController7(printerStream, printerBidiSchemaResponses);
        break;
      default:
    }
  }
  return 0;
}
