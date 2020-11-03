function XlsxToJsonConverter(workbook) {
    if (!window.XLSX) {
        console.error("xlxs 라이브러리가 존재하지 않습니다.");
        return;
    }

    // this.workbook = workbook;

    this.setWorkbook(workbook);
}

XlsxToJsonConverter.prototype.setWorkbook = function(workbook) {
    if (!workbook) {
        console.error("파라미터가 비어있습니다.");
        return;
    }

    this.workbook = workbook;
}

XlsxToJsonConverter.prototype.convert = function(sheetName) {
    var ctxt = this;

    this.workbook.SheetNames.forEach(function(item, index, array) {

        console.log(workbook.Sheets[item]);
        var table = ctxt.util.tabulateToArray(workbook.Sheets[item]);
        var json = ctxt.util.tableArrayToJson(table, workbook.Sheets[item]["!merges"]);

        console.log(json);
    });
}

XlsxToJsonConverter.prototype.util = (function() {
    var ctxt = this;

    function colToNum(s) {
        var ans = 0;
        for (var i = 0; i < s.length; i++) {
            ans = ans * 26 + s.charCodeAt(i) - 64;
        }
        return ans - 1;
    }

    function numToCol(n) {
        var ans = '';
        for (n++; n-- > 0; n = (n - (n % 26)) / 26) {
            ans = String.fromCharCode((n % 26) + 65) + ans;
        } 
        return ans;
    }

    function parseCell(s) {
        var m = s.match(/([A-Z]+)([0-9]+)/);
        if (m !== null) {
            return [colToNum(m[1]), Number(m[2]) - 1, m[0]];
        }
    }

    function tabulateToArray(sheet){
        var rows = [];
        
        Object.keys(sheet).map(parseCell).filter(function(x) {
            return x !== undefined;
        }).forEach(function(parsedKey){
            var key = parsedKey[2],
            col = parsedKey[0],
            row = parsedKey[1],
            val = sheet[key].v;
    
            if (undefined === rows[row]) {
                rows[row] = [];
            }
    
            rows[row][col] = val;
        });     
        return rows;
    }
    
    // function getMergeCellValue(tableArray)
    
    function tableArrayToJson(tableArray) {
        tableArray.forEach(function(rowArray, row) {
            var obj = {};
            var refs = [];
            var values = [];
            var isValue = false;
    
            for (var column = 0; column < rowArray.length; column++) {
                var value = rowArray[column];
            // rowArray.forEach(function(value, column) {
                if (value === undefined) {
                    continue;
                }
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               
                
                if (!isValue && value.charAt(0) === "$") {
                    var ref = value.replace(/\$(.*?)/g, "$1");
                    var type = value.substring(ref.length);
                    // key = key.substring(1);
    
                    refs.push(ref);
    
                    // switch (type) {
                    //     case "[]": reference[key] = []; break;
                    //     case "{}": reference[key] = {}; break;
                    //     // default: 
                    // }
                    // value.replace("$", )
                } else {
                    isValue = false;
                }
            }
    
            // console.log(refs);
    
            refs.forEach(function(ref) {
                var tempRef;


    
                ref.split(".").reverse().forEach(function(value) {
                    var key = String(/\w+/.exec(ref));
                    var refKeys = (value.match(/(?<=\[).+?(?=\])/g) || []).map(function(value) {
                        return isNaN(value) ? value : parseInt(value);
                    });

                    var type = value.replace(/(\w+|\[(.+?)\])/g, "");

                    for (var i = refKeys.length; i > 0; i++) {
                        
                    }

                    console.log(ref, key, refKeys);
                });
                // ref.split('.').map(function(value) {
                //     var key = String(/\w+/.exec(ref));
                //     var refKeys = (value.match(/(?<=\[).+?(?=\])/g) || []).map(function(value) {
                //         return isNaN(value) ? value : parseInt(value);
                //     });
    
                //     console.log(key, refKeys);
    
                    // if (key !== "null") {
                    //     tempRef
                    // }f
                // })
            });
            // });
        });
    
        return tableArray;
    
    }

    return {
        colToNum: colToNum,
        numToCol: numToCol,
        parseCell: parseCell,

        tabulateToArray: tabulateToArray,
        tableArrayToJson: tableArrayToJson
    }
})();