const funcBody = `
;;function write_ws_xml_datavalidation(validations) {
	var o = '<dataValidations>';
	for(var i=0; i < validations.length; i++) {
		var validation = validations[i];
		if(validation.type=="list"){
			o += '<dataValidation type="list" allowBlank="1" sqref="' + validation.sqref + '" showInputMessage="1" showErrorMessage="1" errorTitle="sdfsadf" errorMessage="aaaadf">';
			o += '<formula1>&quot;' + validation.values + '&quot;</formula1>';
			o += '</dataValidation>';
		}
		else if(validation.type=="decimal"){
			o += '<dataValidation type="decimal" operator="between" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="' + validation.sqref + '">';
			o += '<formula1>' + validation.min + '</formula1>';
			o += '<formula2>' + validation.max + '</formula2>';
			o += '</dataValidation>';
		}
		else if(validation.type=="date"){
			o += '<dataValidation type="date" operator="between" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="' + validation.sqref + '">';
			o += '<formula1>' + (new Date(validation.start)).getTime()/1000 + '</formula1>';
			o += '<formula2>' + (new Date(validation.end)).getTime()/1000 + '</formula2>';
			o += '</dataValidation>';
		}
		else if(validation.type == "alphabet"){
			o += '<dataValidation type="custom" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="' + validation.sqref + '">';
			o += '<formula1>ISNUMBER(SUMPRODUCT(SEARCH(MID('+validation.sqref.split(':')[0]+',ROW(INDIRECT(&quot;1:&quot;&amp;LEN('+validation.sqref.split(':')[0]+'))),1),&quot;abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ&quot;)))</formula1>';
			o += '</dataValidation>';
		}
		else if(validation.type == "email"){
			o += '<dataValidation type="custom" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="' + validation.sqref + '">';
			o += '<formula1>=ISNUMBER(MATCH(&quot;*@*.???&quot;,'+validation.sqref.split(':')[0]+',0))</formula1>';
			o += '</dataValidation>';
		}
		else if(validation.type == "phone_number"){
			o += '<dataValidation type="custom" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="' + validation.sqref + '">';
			o += '<formula1>=AND(ISNUMBER(SUMPRODUCT(SEARCH(MID('+validation.sqref.split(':')[0]+',ROW(INDIRECT(&quot;1:&quot;&amp;LEN('+validation.sqref.split(':')[0]+'))),1),&quot;0123456789&quot;))),LEN('+validation.sqref.split(':')[0]+')=9,'+validation.sqref.split(':')[0]+'&lt;&gt;&quot;123456789&quot;,'+validation.sqref.split(':')[0]+'&lt;&gt;&quot;987654321&quot;,'+validation.sqref.split(':')[0]+'&lt;&gt;&quot;000000000&quot;)</formula1>';
			o += '</dataValidation>';
		}
	}
	o += '</dataValidations>';
	return o;
};;
`;
module.exports = function(input) {
  const needle = `write_ws_xml_merges(ws['!merges']`;
  const start = input.indexOf(needle);
  const nextLineIndex = input.indexOf('\n', start);
  const funcCallCode = `
    if(ws['!dataValidation']) o[o.length] = write_ws_xml_datavalidation(ws['!dataValidation']);
  `;
  return funcBody + input.substr(0, nextLineIndex) + funcCallCode + input.substr(nextLineIndex + 1);
}

