const funcBody = `
;;var XmlNode = (function () {
  function XmlNode(tagName, attributes, children) {

    if (!(this instanceof XmlNode)) {
      return new XmlNode(tagName, attributes, children);
    }
    this.tagName = tagName;
    this._attributes = attributes || {};
    this._children = children || [];
    this._prefix = '';
    return this;
  }

  XmlNode.prototype.createElement = function () {
    return new XmlNode(arguments)
  }

  XmlNode.prototype.children = function() {
    return this._children;
  }

  XmlNode.prototype.append = function (node) {
    this._children.push(node);
    return this;
  }

  XmlNode.prototype.prefix = function (prefix) {
    if (arguments.length==0) { return this._prefix;}
    this._prefix = prefix;
    return this;
  }

  XmlNode.prototype.attr = function (attr, value) {
    if (value == undefined) {
      delete this._attributes[attr];
      return this;
    }
    if (arguments.length == 0) {
      return this._attributes;
    }
    else if (typeof attr == 'string' && arguments.length == 1) {
      return this._attributes.attr[attr];
    }
    if (typeof attr == 'object' && arguments.length == 1) {
      for (var key in attr) {
        this._attributes[key] = attr[key];
      }
    }
    else if (arguments.length == 2 && typeof attr == 'string') {
      this._attributes[attr] = value;
    }
    return this;
  }

  var APOS = "'"; QUOTE = '"'
  var ESCAPED_QUOTE = {  }
  ESCAPED_QUOTE[QUOTE] = '&quot;'
  ESCAPED_QUOTE[APOS] = '&apos;'

  XmlNode.prototype.escapeAttributeValue = function(att_value) {
    return '"' + att_value.replace(/\"/g,'&quot;') + '"';// TODO Extend with four other codes

  }

  XmlNode.prototype.toXml = function (node) {
    if (!node) node = this;
    var xml = node._prefix;
    xml += '<' + node.tagName;
    if (node._attributes) {
      for (var key in node._attributes) {
        xml += ' ' + key + '=' + this.escapeAttributeValue(''+node._attributes[key]) + ''
      }
    }
    if (node._children && node._children.length > 0) {
      xml += ">";
      for (var i = 0; i < node._children.length; i++) {
        xml += this.toXml(node._children[i]);
      }
      xml += '</' + node.tagName + '>';
    }
    else {
      xml += '/>';
    }
    return xml;
  }
  return XmlNode;
})();
var StyleBuilder = function (options) {

  var customNumFmtId = 164;



  var table_fmt = {
    0:  'General',
    1:  '0',
    2:  '0.00',
    3:  '#,##0',
    4:  '#,##0.00',
    9:  '0%',
    10: '0.00%',
    11: '0.00E+00',
    12: '# ?/?',
    13: '# ??/??',
    14: 'm/d/yy',
    15: 'd-mmm-yy',
    16: 'd-mmm',
    17: 'mmm-yy',
    18: 'h:mm AM/PM',
    19: 'h:mm:ss AM/PM',
    20: 'h:mm',
    21: 'h:mm:ss',
    22: 'm/d/yy h:mm',
    37: '#,##0 ;(#,##0)',
    38: '#,##0 ;[Red](#,##0)',
    39: '#,##0.00;(#,##0.00)',
    40: '#,##0.00;[Red](#,##0.00)',
    45: 'mm:ss',
    46: '[h]:mm:ss',
    47: 'mmss.0',
    48: '##0.0E+0',
    49: '@',
    56: '"上午/下午 "hh"時"mm"分"ss"秒 "'    };
  var fmt_table = {};

  for (var idx in table_fmt) {
    fmt_table[table_fmt[idx]] = idx;
  }


  // cache style specs to avoid excessive duplication
  _hashIndex = {};
  _listIndex = [];

  return {

    initialize: function (options) {

      this.$fonts = XmlNode('fonts').attr('count',0).attr("x14ac:knownFonts","1");
      this.$fills = XmlNode('fills').attr('count',0);
      this.$borders = XmlNode('borders').attr('count',0);
      this.$numFmts = XmlNode('numFmts').attr('count',0);
      this.$cellStyleXfs = XmlNode('cellStyleXfs');
      this.$xf = XmlNode('xf')
        .attr('numFmtId', 0)
        .attr('fontId', 0)
        .attr('fillId', 0)
        .attr('borderId', 0);

      this.$cellXfs = XmlNode('cellXfs').attr('count',0);
      this.$cellStyles = XmlNode('cellStyles')
        .append(XmlNode('cellStyle')
          .attr('name', 'Normal')
          .attr('xfId',0)
          .attr('builtinId',0)
        );
      this.$dxfs = XmlNode('dxfs').attr('count', "0");
      this.$tableStyles = XmlNode('tableStyles')
        .attr('count','0')
        .attr('defaultTableStyle','TableStyleMedium9')
        .attr('defaultPivotStyle','PivotStyleMedium4')


      this.$styles = XmlNode('styleSheet')
        .attr('xmlns:mc','http://schemas.openxmlformats.org/markup-compatibility/2006')
        .attr('xmlns:x14ac','http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac')
        .attr('xmlns','http://schemas.openxmlformats.org/spreadsheetml/2006/main')
        .attr('mc:Ignorable','x14ac')
        .prefix('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        .append(this.$numFmts)
        .append(this.$fonts)
        .append(this.$fills)
        .append(this.$borders)
        .append(this.$cellStyleXfs.append(this.$xf))
        .append(this.$cellXfs)
        .append(this.$cellStyles)
        .append(this.$dxfs)
        .append(this.$tableStyles);


      // need to specify styles at index 0 and 1.
      // the second style MUST be gray125 for some reason

      var defaultStyle = options.defaultCellStyle || {};
      if (!defaultStyle.font) defaultStyle.font = {name: 'Calibri', sz: '12'};
      if (!defaultStyle.font.name) defaultStyle.font.name = 'Calibri';
      if (!defaultStyle.font.sz) defaultStyle.font.sz = 11;
      if (!defaultStyle.fill) defaultStyle.fill = {  patternType: "none", fgColor: {}};
      if (!defaultStyle.border) defaultStyle.border = {};
      if (!defaultStyle.numFmt) defaultStyle.numFmt = 0;

      this.defaultStyle = defaultStyle;

      var gray125Style = JSON.parse(JSON.stringify(defaultStyle));
      gray125Style.fill = {patternType: "gray125", fgColor: { }}

      this.addStyles([defaultStyle, gray125Style]);
      return this;
    },

    // create a style entry and returns an integer index that can be used in the cell .s property
    // these format of this object follows the emerging Common Spreadsheet Format
    addStyle: function (attributes) {

      var hashKey = JSON.stringify(attributes);
      var index = _hashIndex[hashKey];
      if (index == undefined) {

        index = this._addXf(attributes); //_listIndex.push(attributes) -1;
        _hashIndex[hashKey] = index;
      }
      else {
        index = _hashIndex[hashKey];
      }
      return index;
    },

    // create style entries and returns array of integer indexes that can be used in cell .s property
    addStyles: function (styles) {
      var self = this;
      return styles.map(function (style) {
        return self.addStyle(style);
      })
    },

    _duckTypeStyle: function(attributes) {

      if (typeof attributes == 'object' && (attributes.patternFill || attributes.fgColor)) {
        return {fill: attributes }; // this must be read via XLSX.parseFile(...)
      }
      else if (attributes.font || attributes.numFmt || attributes.border || attributes.fill) {
        return attributes;
      }
      else {
        return this._getStyleCSS(attributes)
      }
    },

    _getStyleCSS: function(css) {
      return css; //TODO
    },

    // Create an <xf> record for the style as well as corresponding <font>, <fill>, <border>, <numfmts>
    // Right now this is simple and creates a <font>, <fill>, <border>, <numfmts> for every <xf>
    // We could perhaps get fancier and avoid duplicating  auxiliary entries as Excel presumably intended, but bother.
    _addXf: function (attributes) {


      var fontId = this._addFont(attributes.font);
      var fillId = this._addFill(attributes.fill);
      var borderId = this._addBorder(attributes.border);
      var numFmtId = this._addNumFmt(attributes.numFmt);

      var $xf = XmlNode('xf')
        .attr("numFmtId", numFmtId)
        .attr("fontId", fontId)
        .attr("fillId", fillId)
        .attr("borderId", borderId)
        .attr("xfId", "0");

      if (fontId > 0) {
        $xf.attr('applyFont', "1");
      }
      if (fillId > 0) {
        $xf.attr('applyFill', "1");
      }
      if (borderId > 0) {
        $xf.attr('applyBorder', "1");
      }
      if (numFmtId > 0) {
        $xf.attr('applyNumberFormat', "1");
      }

      if (attributes.alignment) {
        var $alignment = XmlNode('alignment');
        if (attributes.alignment.horizontal) { $alignment.attr('horizontal', attributes.alignment.horizontal);}
        if (attributes.alignment.vertical)  { $alignment.attr('vertical', attributes.alignment.vertical);}
        if (attributes.alignment.indent)  { $alignment.attr('indent', attributes.alignment.indent);}
        if (attributes.alignment.readingOrder)  { $alignment.attr('readingOrder', attributes.alignment.readingOrder);}
        if (attributes.alignment.wrapText)  { $alignment.attr('wrapText', attributes.alignment.wrapText);}
        if (attributes.alignment.textRotation!=undefined)  { $alignment.attr('textRotation', attributes.alignment.textRotation);}

        $xf.append($alignment).attr('applyAlignment',1)

      }
      this.$cellXfs.append($xf);
      var count = +this.$cellXfs.children().length;

      this.$cellXfs.attr('count', count);
      return count - 1;
    },

    _addFont: function (attributes) {

      if (!attributes) {  return 0; }

      var $font = XmlNode('font')
        .append(XmlNode('sz').attr('val', attributes.sz || this.defaultStyle.font.sz))
        .append(XmlNode('name').attr('val', attributes.name || this.defaultStyle.font.name))

      if (attributes.bold) $font.append(XmlNode('b'));
      if (attributes.underline)  $font.append(XmlNode('u'));
      if (attributes.italic)  $font.append(XmlNode('i'));
      if (attributes.strike)  $font.append(XmlNode('strike'));
      if (attributes.outline)  $font.append(XmlNode('outline'));
      if (attributes.shadow)  $font.append(XmlNode('shadow'));

      if (attributes.vertAlign) {
        $font.append(XmlNode('vertAlign').attr('val', attributes.vertAlign))
      }


      if (attributes.color) {
        if (attributes.color.theme) {
          $font.append(XmlNode('color').attr('theme', attributes.color.theme))

          if (attributes.color.tint) { //tint only if theme
            $font.append(XmlNode('tint').attr('theme', attributes.color.tint))
          }

        } else if (attributes.color.rgb) { // not both rgb and theme
          $font.append(XmlNode('color').attr('rgb', attributes.color.rgb))
        }
      }

      this.$fonts.append($font);

      var count = this.$fonts.children().length;
      this.$fonts.attr('count', count);
      return count - 1;
    },

    _addNumFmt: function (numFmt) {
      if (!numFmt) { return 0; }

      if (typeof numFmt == 'string') {
        var numFmtIdx = fmt_table[numFmt];
        if (numFmtIdx >= 0) {
          return numFmtIdx; // we found a match against built in formats
        }
      }

      if (/^[0-9]+$/.exec(numFmt)) {
        return numFmt; // we're matching an integer against some known code
      }
      numFmt = numFmt
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');

      var $numFmt = XmlNode('numFmt')
        .attr('numFmtId', (++customNumFmtId))
        .attr('formatCode', numFmt);

      this.$numFmts.append($numFmt);

      var count = this.$numFmts.children().length;
      this.$numFmts.attr('count', count);
      return customNumFmtId ;
    },

    _addFill: function (attributes) {

      if (!attributes) { return 0; }

      var $patternFill = XmlNode('patternFill')
        .attr('patternType', attributes.patternType || 'solid');

      if (attributes.fgColor) {
        var $fgColor = XmlNode('fgColor');

        //Excel doesn't like it when we set both rgb and theme+tint, but xlsx.parseFile() sets both
        //var $fgColor = createElement('<fgColor/>', null, null, {xmlMode: true}).attr(attributes.fgColor)
        if (attributes.fgColor.rgb) {

          if (attributes.fgColor.rgb.length == 6) {
            attributes.fgColor.rgb = "FF" + attributes.fgColor.rgb /// add alpha to an RGB as Excel expects aRGB
          }

          $fgColor.attr('rgb', attributes.fgColor.rgb);
          $patternFill.append($fgColor);
        }
        else if (attributes.fgColor.theme) {
          $fgColor.attr('theme', attributes.fgColor.theme);
          if (attributes.fgColor.tint) {
            $fgColor.attr('tint', attributes.fgColor.tint);
          }
          $patternFill.append($fgColor);
        }

        if (!attributes.bgColor) {
          attributes.bgColor = { "indexed": "64"}
        }
      }

      if (attributes.bgColor) {
        var $bgColor = XmlNode('bgColor').attr(attributes.bgColor);
        $patternFill.append($bgColor);
      }

      var $fill = XmlNode('fill')
        .append($patternFill);

      this.$fills.append($fill);

      var count = this.$fills.children().length;
      this.$fills.attr('count', count);
      return count - 1;
    },

    _getSubBorder: function(direction, spec) {

      var $direction = XmlNode(direction);
      if (spec){
        if (spec.style) $direction.attr('style', spec.style);
        if (spec.color) {
          var $color = XmlNode('color');
          if (spec.color.auto) {
            $color.attr('auto', spec.color.auto);
          }
          else if (spec.color.rgb) {
            $color.attr('rgb', spec.color.rgb);
          }
          else if (spec.color.theme || spec.color.tint) {
            $color.attr('theme', spec.color.theme || "1");
            $color.attr('tint', spec.color.tint || "0");
          }
          $direction.append($color)
        }
      }
      return $direction;
    },

    _addBorder: function (attributes) {
      if (!attributes) { return 0; }

      var self = this;

      var $border = XmlNode('border')
        .attr("diagonalUp",attributes.diagonalUp)
        .attr("diagonalDown",attributes.diagonalDown);

      var directions = ["left","right","top","bottom","diagonal"];

      directions.forEach(function(direction) {
        $border.append(self._getSubBorder(direction, attributes[direction]))
      });
      this.$borders.append($border);

      var count = this.$borders.children().length;
      this.$borders.attr('count', count);
      return count -1;
    },

    toXml: function () {
      return this.$styles.toXml();
    }
  }.initialize(options||{});
}
;;
;;function write_ws_xml_datavalidation(validations) {
	var o = '<dataValidations>';
	for(var i=0; i < validations.length; i++) {
		var validation = validations[i];
    if(validation.type=="list"){
      o += '<dataValidation type="list" allowBlank="1" sqref="' + validation.sqref + '" showInputMessage="1" showErrorMessage="1" errorTitle="sdfsadf" errorMessage="aaaadf">';
      o += '<formula1>&quot;' + validation.values + '&quot;</formula1>';
      o += '</dataValidation>';
    }
    else if(validation.type=="fixed"){
			o += '<dataValidation type="custom" allowBlank="0" sqref="' + validation.sqref + '" showInputMessage="1" showErrorMessage="1" errorTitle="sdfsadf" errorMessage="aaaadf">';
      o += '<formula1>=ISNUMBER(MATCH(&quot;' + validation.value +'&quot;))</formula1>';
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
		else if(validation.type == "phone"){
			var formula = '<dataValidation type="custom" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="' + validation.sqref + '">';
			var _cell = validation.sqref.split(':')[0];
			formula += '<formula1>OR(IFERROR(IF(LEN(' + _cell +
			')=10,VALUE(' + _cell +
			')*0 + 1,FALSE),FALSE), IFERROR(IF(OR(LEN(' + _cell +
			')=12,LEN(' + _cell + 
			')=13),IF(AND(MID(' + _cell +
			',4,1)=&quot;-&quot;,MID(' + _cell + 
			',8,1)=&quot;-&quot;),VALUE(LEFT(' + _cell + 
			',3) &amp; MID(' + _cell + 
			',5,3) &amp; MID(' + _cell + 
			',9,32767)) * 0 + 1,FALSE),FALSE),FALSE))</formula1>';
			formula += '</dataValidation>';
			console.info(formula);
			o += formula;
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
  input = input.replace(/function write_sty_xml.+\{/, (a) => a + `
    return style_builder.toXml();
  `);
  input = input.replace(/function get_cell_style.+\{/, (a) => a + `
    if (/^\d+$/.exec(cell.s)) { return cell.s}  // if its already an integer index, let it be
    if (cell.s && (cell.s == +cell.s)) { return cell.s}  // if its already an integer index, let it be
    var s = cell.s || {};
    if (cell.z) s.numFmt = cell.z;
    return style_builder.addStyle(s);
  `);
  input = input.replace(/function write_zip_type.+\{/, (a) => a + `
    style_builder  = new StyleBuilder(opts);
  `);
  const needle = `write_ws_xml_merges(ws['!merges']`;
  const start = input.indexOf(needle);
  const nextLineIndex = input.indexOf('\n', start);
  const funcCallCode = `
    if(ws['!dataValidation']) o[o.length] = write_ws_xml_datavalidation(ws['!dataValidation']);
  `;
  return funcBody + input.substr(0, nextLineIndex) + funcCallCode + input.substr(nextLineIndex + 1);
}

