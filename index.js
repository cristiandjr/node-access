var ADODB = require('node-adodb');
ADODB.debug = true;

//const dir = 'BASEENERGIASMVS-Copys.accdb';
const dir = 'BASE ENERGIAS MVS_be.accdb';
// Connect to the MS Access DB
var connection = ADODB.open('Provider=Microsoft.ACE.OLEDB.12.0;Data Source='+dir+';Persist Security Info=False;');
//var connection = ADODB.open('Provider=Microsoft.ACE.OLEDB.12.0;Data Source='+dir+';Persist Security Info=False;');
//var connection = ADODB.open('Provider=Microsoft.ACE.OLEDB.15.0;Data Source='+dir+';Persist Security Info=False');
connection
  //.execute('SELECT SITIO FROM MVS WHERE Code = "CS009" ')
  .query('SELECT * FROM MVS')
  .then(data => {
    console.log(JSON.stringify(data, null, 2));
    console.log(data.Code)
  })
  .catch(error => {
    console.error(error);
  });
 