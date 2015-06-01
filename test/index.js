var should = require('should');
var sinon = require('sinon');

var spread2json = require('../');
var api = require('../lib/api');

var opts = require('./opts');
var mock = require('./mock');
var data = require('./data');

var SPREADSHEET_KEY = '1YXVzaaxqkPKsr-excIOXScnTQC7y_DKrUKs0ukzSIgo';
var WORKSHEET_KEYS = [
  'od6',
  'oat5a13',
  'o7qj5m0',
  'ocxr3gc',
  'o139ph2',
  'opanz8'
];

describe('spread2json', function() {
  before(function() {
    spread2json.setup({ api: opts.installed });
    var sandbox = sinon.sandbox.create();
    sandbox.stub(api, 'getWorksheet').yields(null, mock.getWorksheet);
    var stubGetCells = sandbox.stub(api, 'getCells');
    for (var key in mock.getCells) {
      stubGetCells.withArgs(undefined, SPREADSHEET_KEY, key).yields(null, mock.getCells[key]);
    }
  });

  it('#getWorksheet', function(done) {
    spread2json.getWorksheet(SPREADSHEET_KEY, function(err, result) {
      should.not.exist(err);
      should.exist(result);
      result.should.have.length(6);
      result.forEach(function(d, i) {
        d.should.have.property('id', WORKSHEET_KEYS[i]);
        d.should.have.property('updated');
        d.should.have.property('title');
      });

      done();
    });
  });

  it('#getWorksheetDatas', function(done) {
    spread2json.getWorksheetDatas(SPREADSHEET_KEY, WORKSHEET_KEYS, function(err, result) {
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
    spread2json.getWorksheetDatas(SPREADSHEET_KEY, WORKSHEET_KEYS, function(err, result) {
      should.not.exist(err);
      should.exist(result);

      spread2json.toJson(result, function(_err, _result) {
        should.not.exist(_err);
        should.exist(_result);
        _result.should.eql(data);

        done();
      });
    });
  });
});
