# json-xlsx-simply

translate json to xlsx simply

# Usage

``` JavaScript

const jsonToXLSX = require('json-xlsx-simply').jsonToXLSX

const filePath = 'demo.xlsx'
const jsonData = [
  {name: 'a', age: 1, description: 'a'},
  {name: 'b', age: 1, description: 'b'},
  {name: 'c', age: 1, description: 'c'},
]
const opts = {
  sheetName: 'person',                    // defalut: sheet
  head: ['person-name', 'person-name'],   // default: all keys name
  fields: ['name', 'age'],                // default: all keys value
}

jsonToXLSX(filePath, jsonData, opts)
  .then(()=> {
    // done and you will get a xlsx file in the file path
  })
```