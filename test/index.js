'use strict';

var should = require('should');
var sinon = require('sinon');

var _ = require('lodash');
var spread2json = require('../');
var api = require('../lib/api');

var opts = require('./opts');
var mock = require('./mock');
var data = require('./data');

var SPREADSHEET_KEY = '1YXVzaaxqkPKsr-excIOXScnTQC7y_DKrUKs0ukzSIgo';
var WORKSHEET_NAMES = [
  'Test1',
  'Test1.list',
  'Test1.#list.list',
  'Test2',
  'Test2.map',
  'Test2.map.$.map'
];

describe('spread2json', function() {
  before(function() {
    spread2json.setup({ api: opts.installed });
    var sandbox = sinon.sandbox.create();
    sandbox.stub(api, 'getWorksheet').yields(null, mock.getWorksheet);
    var stubGetList = sandbox.stub(api, 'getList');
    for (var key in mock.getList) {
      stubGetList.withArgs(undefined, SPREADSHEET_KEY, key).yields(null, mock.getList[key]);
    }
  });

  it('#getWorksheet', function(done) {
    spread2json.getWorksheet(SPREADSHEET_KEY, function(err, result) {
      should.not.exist(err);
      should.exist(result);
      result.should.have.length(6);
      result.forEach(function(d, i) {
        d.should.have.property('title', WORKSHEET_NAMES[i]);
        d.should.have.property('sheetId');
        d.should.have.property('link');
      });

      done();
    });
  });

  it('#getWorksheetDatas', function(done) {
    spread2json.getWorksheetDatas(SPREADSHEET_KEY, WORKSHEET_NAMES, function(err, result) {
      should.not.exist(err);
      should.exist(result);
      result.should.have.length(6);
      result.forEach(function(d) {
        d.should.have.property('name');
        d.should.have.property('opts');
        d.should.have.property('list');
      });

      done();
    });
  });

  it('#toJson', function(done) {
    spread2json.getWorksheetDatas(SPREADSHEET_KEY, WORKSHEET_NAMES, function(err, result) {
      should.not.exist(err);
      should.exist(result);

      spread2json.toJson(result, function(_err, _result) {
        should.not.exist(_err);
        should.exist(_result);
        _.forEach(_result, function(ret, key) {
          _.forEach(ret, function(_ret, _key) {
            _ret.should.eql(data[key][_key]);
          });
        });

        done();
      });
    });
  });
});
