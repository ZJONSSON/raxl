const xmler = require('xmler');
const etl = require('etl');

function lettersToNumber(letters){
  return letters.split('').reduce((r, a) => r * 26 + parseInt(a, 36) - 9, 0)-1;
}

module.exports = async function(directory, options) {
  options = options || {};
  directory = await directory;
  const sharedStringsFile = directory.files.find(d => d.path == 'xl/sharedStrings.xml');
  let sharedStringsPromise = sharedStringsFile.stream().pipe(xmler('t')).pipe(etl.map(d => d.value)).promise();
  let sharedStrings;
  
  return directory.files.reduce( (p,d) => {
    if (d.path.startsWith('xl/worksheets')) {
      const name = d.path.replace(/xl\/worksheets\//,'').replace(/\.xml$/,'');
      const stream = argv => {
        let headers;
        argv = argv || {};
        return d.stream()
          .pipe(xmler('row', {showAttr: true}))
          .pipe(etl.map(async d => {
            sharedStrings = sharedStrings || await sharedStringsPromise;
            let row = [];
            d.value.c.forEach(d => {
              const col = lettersToNumber(/([A-Z]+)/.exec(d.__attr__.r)[1]);
              if (!d.v && !d.is) return;

              let value;
              const type = d.__attr__.t;
              
              if (type === 's') {
                value = sharedStrings[+d.v.text];
              } else if (type == 'inlineStr') {
                
                value = d.is.t.text;
              } else  if (type === 'n') {
                value = +d.v.text;
              } else {
                value = d.v.text;
              }
              
              row[col] = value;
            });
            if (argv.raw) return row;

            if (!headers) {
              if (options.headerCol && !row[options.headerCol]) return;
              headers = row;
            } else {
              return row.reduce( (p,d,i) => {
                p[headers[i] ||  `COL_${i}`] = d;
                return p;
              },{});
            }
          }));
      };
      p[name] = stream;
    }
    return p;
  },{});
};