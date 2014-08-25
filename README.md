Spread2Json
==========

Can be converted to JSON format any SpreadSheet data.

example SpreadSheet data

|   | A      | B        | C                | D |
|:-:|:-------|:---------|:-----------------|---|
| 1 | {}     |          |                  |   |
| 2 | _id    | obj.code | obj.value:number |   |
| 3 |        |          |                  |   |
| 4 | first  | one      | 1                |   |
| 5 | second | two      | 2                |   |
| 6 |        |          |                  |   |
converted to Object
```
[
    {
        _id: 'first',
        obj: {
            code: 'one',
            value: 1
        }
    }, 
    {
        _id: 'second',
        obj: {
            code: 'two',
            value: 2
        }
    }
]
```

## Installation
```
npm install spread2json
```

## Usage
### Quick start
example

|   | A              | B        | C                | D |
|:-:|:---------------|:---------|:-----------------|---|
| 1 | {name: 'Test'} |          |                  |   |
| 2 | _id            | obj.code | obj.value:number |   |
| 3 |                |          |                  |   |
| 4 | first          | one      | 1                |   |
| 5 | second         | two      | 2                |   |
| 6 |                |          |                  |   |
Sheet1
```
var spread2json = require('spread2json');

var spreadsheetKye = 'spreadsheetkey';
var worksheetIds = ['sheet1'];
spread2json.getWorksheetDatas(spreadsheetKey, worksheetIds, function(err, data) {
    console.log(data);
    // [{
    //    name: 'Sheet1',            // sheet name
    //    opts: { name: 'Test' },    // sheet option (A1)
    //    list: [
    //        { _id: 'first', obj: { code: 'one', value: 1 } }, { _id: 'second', obj: { code: 'two', value: 2 } }
    //    ]
    // }]

    spread2json.toJson(data, function(err, json) {
        console.log(json);
        // {
        //    Test: {
        //        first: {
        //            _id: 'first',
        //            obj: { code: 'one', value: 1 }
        //        },
        //        second: {
        //            _id: 'second',
        //            obj: { code: 'two', value: 2 }
        //        }
        //    }
        // }
    });
});
```

### Setup
Setup options.
```
var spread2json = require('spread2json');

spread2json.setup({
    option_cell: 'A1',    // Cell with a custom sheet option. It is not yet used now. (default: 'A1'
    attr_line: '2',       // Line with a data attribute. (default: '2'
    data_line: '4',       // Line with a data. (default: '4'
    ref_key: '_id',       // ref key. (default: '_id'
    logger: customLogger, // custom Logger
    api: {
        client_id: 'YOUR CLIENT ID HERE',
        client_secret: 'YOUR CLIENT SECRET HERE',
        redirect_url: 'http://localhost',
        token_path: './dist/token.json'
    }
});
```

### Sheet option
sheet option. setting with optionCell (default: 'A1'
* `name`
* `type`
* `key`
* `attr_line`
* `data_line`
* `ref_key`


### Attribute
Specify the key name.

**Special character**
* `#` Use when the array.
* `$` Use when the split array.
* `:number` or `:num` Use when the parameters of type `Number`.
* `:boolean` or `:bool` Use when the parameters of type `Boolean`.
* `:date` Use when the parameters of unix time.
* `:index` Use when the array of array.

### Example
An example of a complex format.


[test.xlsx](https://github.com/yuhei-a/excel2json/raw/master/test/data/test.xlsx) > [test.json](https://github.com/yuhei-a/excel2json/blob/master/test/data/test.json)
```
var spread2json = require('spread2json');
var SPREADSHEET_KEY = '1YXVzaaxqkPKsr-excIOXScnTQC7y_DKrUKs0ukzSIgo';
var WORKSHEET_KEYS = [
    'od6',
    'oat5a13',
    'o7qj5m0',
    'ocxr3gc',
    'o139ph2',
    'opanz8'
];
spread2json.getWorksheetDatas(SPREADSHEET_KEY, WORKSHEET_KEYS, function(err, sheetDatas) {
    spread2json.toJson(sheetDatas, function(err, result) {
        fs.writeFileSync('test.json', JSON.stringify(result, null, 4));
    });
});
```

## Test
Run `npm test` and `npm run-script jshint`
