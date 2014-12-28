// Description
//   A Hubot script that DESCRIPTION
//
// Configuration:
//   None
//
// Commands:
//   hubot XXX [<args>] - DESCRIPTION
//
// Note:
//   $ KEY=$(openssl pkcs12 -in key.p12 -nodes -nocerts)
//   Enter Import Password: notasecret
//   MAC verified OK
//
// Author:
//   bouzuya <m@bouzuya.net>
//
var Promise, config, google, parseConfig, parseString;

parseConfig = require('hubot-config');

google = require('googleapis');

Promise = require('es6-promise').Promise;

parseString = require('xml2js').parseString;

config = parseConfig('google-spreadsheet-items', {
  spreadsheetKey: '0AiNaqqSQT22JdHQ1aE1xcmNqdXMteVpiUDg1MzVBeEE',
  email: null,
  key: null
});

module.exports = function(robot) {
  var authorize, getCellsUrl, getWorksheetUrl, listSpreadsheetsUrl, message, parseXml, projections, request, visibilities;
  if (config.spreadsheetKey == null) {
    message = 'HUBOT_GOOGLE_SPREADSHEET_ITEMS_SPREADSHEET_KEY is not defined';
    robot.logger.error(message);
    return;
  }
  if (config.email == null) {
    message = 'HUBOT_GOOGLE_SPREADSHEET_ITEMS_EMAIL is not defined';
    robot.logger.error(message);
    return;
  }
  if (config.key == null) {
    message = 'HUBOT_GOOGLE_SPREADSHEET_ITEMS_KEY is not defined';
    robot.logger.error(message);
    return;
  }
  visibilities = {
    "private": 'private',
    "public": 'public'
  };
  projections = {
    basic: 'basic',
    full: 'full'
  };
  authorize = function(_arg) {
    var email, key;
    email = _arg.email, key = _arg.key;
    return new Promise(function(resolve, reject) {
      var jwt, scope;
      scope = ['https://spreadsheets.google.com/feeds'];
      jwt = new google.auth.JWT(email, null, key, scope, null);
      return jwt.authorize(function(err) {
        if (err != null) {
          return reject(err);
        } else {
          return resolve(jwt);
        }
      });
    });
  };
  request = function(client, options) {
    return new Promise(function(resolve, reject) {
      return client.request(options, function(err, data) {
        if (err != null) {
          return reject(err);
        } else {
          return resolve(data);
        }
      });
    });
  };
  parseXml = function(xml) {
    return new Promise(function(resolve, reject) {
      return parseString(xml, function(err, parsed) {
        if (err != null) {
          return reject(err);
        } else {
          return resolve(parsed);
        }
      });
    });
  };
  listSpreadsheetsUrl = function(client) {
    var baseUrl;
    baseUrl = 'https://spreadsheets.google.com/feeds';
    visibilities = 'private';
    projections = 'full';
    return "" + baseUrl + "/spreadsheets/" + visibilities + "/" + projections;
  };
  getWorksheetUrl = function(_arg) {
    var baseUrl, key, projections, visibilities;
    key = _arg.key, visibilities = _arg.visibilities, projections = _arg.projections;
    baseUrl = 'https://spreadsheets.google.com/feeds';
    return "" + baseUrl + "/worksheets/" + key + "/" + visibilities + "/" + projections;
  };
  getCellsUrl = function(_arg) {
    var baseUrl, key, projections, visibilities, worksheetId;
    key = _arg.key, worksheetId = _arg.worksheetId, visibilities = _arg.visibilities, projections = _arg.projections;
    baseUrl = 'https://spreadsheets.google.com/feeds';
    return "" + baseUrl + "/cells/" + key + "/" + worksheetId + "/" + visibilities + "/" + projections;
  };
  return robot.respond(/google spreadsheet items/, function(res) {
    var client, worksheetUrl;
    client = null;
    worksheetUrl = getWorksheetUrl({
      key: config.spreadsheetKey,
      visibilities: visibilities["private"],
      projections: projections.basic
    });
    return authorize(config).then(function(c) {
      return client = c;
    }).then(function() {
      robot.logger.info('hubot-google-spreadsheet-items: authorized');
      return request(client, {
        url: worksheetUrl
      }).then(parseXml);
    }).then(function(data) {
      var url, worksheetUrls;
      worksheetUrls = data.feed.entry.map(function(i) {
        return i.id[0];
      });
      url = worksheetUrls[0];
      if (url.indexOf(worksheetUrl) !== 0) {
        throw new Error();
      }
      return url.replace(worksheetUrl + '/', '');
    }).then(function(worksheetId) {
      var cellsUrl;
      cellsUrl = getCellsUrl({
        key: config.spreadsheetKey,
        worksheetId: worksheetId,
        visibilities: visibilities["private"],
        projections: projections.basic
      });
      return request(client, {
        url: cellsUrl
      }).then(parseXml);
    }).then(function(data) {
      message = data.feed.entry.map(function(i) {
        return {
          title: i.title[0]._,
          content: i.content[0]._
        };
      }).filter(function(i) {
        return i.title.match(/^A/);
      }).map(function(i) {
        return i.content;
      }).join('\n');
      return res.send(message);
    })["catch"](function(e) {
      robot.logger.error(e);
      return res.send('hubot-google-spreadsheet-items: error');
    });
  });
};
