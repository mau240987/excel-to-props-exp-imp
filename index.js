'use strict';
const excelToJson = require('convert-excel-to-json');
const fs = require('fs');
const javaProps = require('java-props');

//Parsing *.properties file into JSON (encoding ISO8859-1 <-> latin1)
javaProps.parseFile('storetext_FRG_en_US.properties','latin1').then((props) => {
    //console.log(properties);
    // { a: 'Hello World', b: 'Node.js®', c: 'value', d: 'foobar' }
	
	const result = excelToJson({
		sourceFile: 'Copia di R21-NewLabels_en_US_v2_FE_CL.xlsx',
		header:{
			rows: 21
		}
	});

	//console.log(result);
	
	const newPropsJson = {};

	const sheets = Object.keys(result);

	for(let i=0;i<sheets.length;i++){
		let rows = result[sheets[i]];
		for(let j=0;j<rows.length;j++){
			const row = rows[j];
			if(props[row.B]){
				let key = row.B;
				let value = row.D;
				newPropsJson[key] = value;
				console.log(key  + ' = ' + value);
			}
		}
	}
	
	const propsKeys = Object.keys(props);
	let propsStr = '';
	
	for(let z=0;z<propsKeys.length;z++){
		//Replace new value getting from excel file
		if(newPropsJson[propsKeys[z]]){
			props[propsKeys[z]] = newPropsJson[propsKeys[z]];
		}
		let breakline = '\n';
		if(z===0){
			breakline = '';
		}
		propsStr = propsStr + breakline + propsKeys[z] + ' = ' + props[propsKeys[z]];
	}
	
	//console.log(props);
	
	// specify the path to the file, and create a buffer with characters we want to write
	let path = 'newProperties.properties';
	let buffer = new Buffer(propsStr);

	// open the file in writing mode, adding a callback function where we do the actual writing
	fs.open(path, 'w', function(err, fd) {
		if (err) {
			throw 'could not open file: ' + err;
		}

		// write the contents of the buffer, from position 0 to the end, to the file descriptor returned in opening our file
		fs.write(fd, buffer, 0, buffer.length, null, function(err) {
			if (err) throw 'error writing file: ' + err;
			fs.close(fd, function() {
				console.log('wrote the file successfully');
			});
		});
	});
			
}).catch((err) => {
	console.error(err)
});