const unzipper = require('unzipper');
const etl = require('etl');
const request = require('request');
const raxl = require('./index');

async function main() {
  const directory =  await unzipper.Open.url(request, {url: 'https://www.fhfa.gov/DataTools/Downloads/Documents/HPI/HPI_AT_3zip.xlsx'});
  const workbook = await raxl(directory, {headerCol:3});  
  workbook.sheet1().pipe(etl.map(d => console.log(JSON.stringify(d,null,2))));
}

main().then(console.log,console.log)