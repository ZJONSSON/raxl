# Random access xlsx streaming reader

Function that takes in a `directory` object from `unzipper` and returns a list of worksheets that can be streamed.

Example:

```js
const unzipper = require('unzipper');
const etl = require('etl');
const request = require('request');
const raxl = require('./index');

async function main() {
  const directory =  await unzipper.Open.url(request, {
  	url: 'https://www.hud.gov/sites/documents/RM-A_07-31-2014.xlsx'
  });
  const workbook = await raxl(directory);  
  workbook.sheet1().pipe(etl.map(d => console.log(JSON.stringify(d,null,2))));
}

main().then(console.log,console.log)
```