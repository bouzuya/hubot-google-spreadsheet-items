# Description
#   A Hubot script that DESCRIPTION
#
# Configuration:
#   HUBOT_GOOGLE_SPREADSHEET_ITEMS_EMAIL
#   HUBOT_GOOGLE_SPREADSHEET_ITEMS_KEY
#   HUBOT_GOOGLE_SPREADSHEET_ITEMS_SPREADSHEET_KEY
#
# Commands:
#   hubot google spreadsheet items - fetch google spreadsheet items
#
# Note:
#   $ KEY=$(openssl pkcs12 -in key.p12 -nodes -nocerts)
#   Enter Import Password: notasecret
#   MAC verified OK
#
# Author:
#   bouzuya <m@bouzuya.net>
#
parseConfig = require 'hubot-config'
google = require 'googleapis'
{Promise} = require 'es6-promise'
{parseString} = require 'xml2js'

config = parseConfig 'google-spreadsheet-items',
  email: null
  key: null
  spreadsheetKey: null

module.exports = (robot) ->
  # validate config
  unless config.spreadsheetKey?
    message = 'HUBOT_GOOGLE_SPREADSHEET_ITEMS_SPREADSHEET_KEY is not defined'
    robot.logger.error(message)
    return

  unless config.email?
    message = 'HUBOT_GOOGLE_SPREADSHEET_ITEMS_EMAIL is not defined'
    robot.logger.error(message)
    return

  unless config.key?
    message = 'HUBOT_GOOGLE_SPREADSHEET_ITEMS_KEY is not defined'
    robot.logger.error(message)
    return

  # constants
  visibilities =
    private: 'private'
    public: 'public'

  projections =
    basic: 'basic'
    full: 'full'

  # functions
  authorize = ({ email, key })->
    new Promise (resolve, reject) ->
      scope = ['https://spreadsheets.google.com/feeds']
      jwt = new google.auth.JWT(email, null, key, scope, null)
      jwt.authorize (err) ->
        if err? then reject(err) else resolve(jwt)

  request = (client, options) ->
    new Promise (resolve, reject) ->
      client.request options, (err, data) ->
        if err? then reject(err) else resolve(data)

  parseXml = (xml) ->
    new Promise (resolve, reject) ->
      parseString xml, (err, parsed) ->
        if err? then reject(err) else resolve(parsed)

  listSpreadsheetsUrl = (client) ->
    baseUrl = 'https://spreadsheets.google.com/feeds'
    visibilities = 'private'
    projections = 'full'
    "#{baseUrl}/spreadsheets/#{visibilities}/#{projections}"

  # visibilities: private / public
  # projections: full / basic
  getWorksheetUrl = ({ key, visibilities, projections }) ->
    baseUrl = 'https://spreadsheets.google.com/feeds'
    "#{baseUrl}/worksheets/#{key}/#{visibilities}/#{projections}"

  # visibilities: private / public
  # projections: full / basic
  getCellsUrl = ({ key, worksheetId, visibilities, projections }) ->
    baseUrl = 'https://spreadsheets.google.com/feeds'
    "#{baseUrl}/cells/#{key}/#{worksheetId}/#{visibilities}/#{projections}"

  robot.respond /google spreadsheet items/, (res) ->

    client = null
    worksheetUrl = getWorksheetUrl
      key: config.spreadsheetKey
      visibilities: visibilities.private
      projections: projections.basic

    authorize(config)
    .then (c) -> client = c
    .then ->
      robot.logger.info 'hubot-google-spreadsheet-items: authorized'
      request(client, { url: worksheetUrl }).then parseXml
    .then (data) ->
      # get worksheet id
      worksheetUrls = data.feed.entry.map (i) -> i.id[0]
      url = worksheetUrls[0]
      throw new Error() if url.indexOf(worksheetUrl) isnt 0
      url.replace(worksheetUrl + '/', '')
    .then (worksheetId) ->
      cellsUrl = getCellsUrl
        key: config.spreadsheetKey
        worksheetId: worksheetId
        visibilities: visibilities.private
        projections: projections.basic
      request(client, { url: cellsUrl }).then parseXml
    .then (data) ->
      message = data.feed.entry.map (i) ->
        { title: i.title[0]._, content: i.content[0]._ }
      .filter (i) ->
        i.title.match(/^A/)
      .map (i) ->
        i.content
      .join '\n'
      res.send message
    .catch (e) ->
      robot.logger.error e
      res.send 'hubot-google-spreadsheet-items: error'
