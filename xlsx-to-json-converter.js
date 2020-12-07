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

    this.json = {};
    this.workbook.SheetNames.forEach(function(item, index, array) {
        var data = ctxt.util.sheetToJson(ctxt.workbook, item);
        ctxt.json[data.startKey] = data.data;
    });
}

XlsxToJsonConverter.prototype.download = function(fileName) {
    console.log(this.workbook);
    var dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(this.json, null, "\t"));
    var downloadAnchorNode = document.createElement('a');
    downloadAnchorNode.setAttribute("href", dataStr);
    downloadAnchorNode.setAttribute("download", (fileName || "xlsx2json") + ".json");
    document.body.appendChild(downloadAnchorNode); // required for firefox
    downloadAnchorNode.click();
    downloadAnchorNode.remove();
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

    function sheetToJson(workbook, sheetName) {
        var parseRefStr = function(refStr) {
            var refs = [];

            var ref = refStr.replace(/\$(.*?)/g, "$1");

            var types = ["[]"];
            var type = types.find(function(type) {
                return (new RegExp(type.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + "$")).test(ref);
            });

            if (!type) {
                type = "str";
            }
            // var type = refStr.replace(/.+?\[.+?\]/g, "");

            // console.log(refStr,ref,type)

            // key = key.substring(1);

            ref.split(".").forEach(function(refOfSplitByDot) {
                var typesRegex = new RegExp("(" + types.map(function(type) {
                    return type.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
                }).join("|") + ")", "g");
                var key = refOfSplitByDot.replace(/\[.+?\]/g, "").replace(typesRegex, "");//String(/\w+/.exec(refOfSplitByDot));
                var refKeys = (refOfSplitByDot.match(/(?<=\[).+?(?=\])/g) || []).map(function(value) {
                    return isNaN(value) ? value : parseInt(value);
                });

                // refKeys.unshift(key);
                // console.log("key", key);
                // var startBracketsIndex = key.indexOf("[");
                if (key.length > 0) {
                    refKeys.unshift(key);
                    // if (startBracketsIndex > 0) {
                    //     if (startBracketsIndex !== refOfSplitByDot.indexOf("[]")) {
                    //         refKeys.unshift(key);
                    //     } else {

                    //     }
                    // } else {
                    //     refKeys.unshift(key);
                    // }
                }

                Array.prototype.push.apply(refs, refKeys);
            });

            return {
                refKeys: refs,
                type: type
            };
        }

        var sheet = workbook.Sheets[sheetName];
        var obj;

        var sheetNameRefs = parseRefStr(sheetName);

        var tableArray = tabulateToArray(sheet);
        tableArray.forEach(function(rowArray, row) {
            var refs = sheetNameRefs.refKeys.slice();
            var values = [];
            var isValue = false;
            var type;

            for (var column = 0; column < rowArray.length; column++) {
                var value = rowArray[column];
                if (value === undefined) {
                    var mergeData = sheet["!merges"].find(function(mergeData) {
                        return (mergeData.s.r <= row && mergeData.s.c <= column)
                            && (mergeData.e.r >= row && mergeData.e.c >= column);
                    });

                    if (!mergeData) {
                        continue;
                    }

                    tableArray[row][column] = tableArray[mergeData.s.r][mergeData.s.c];
                    value = tableArray[mergeData.s.r][mergeData.s.c];
                }
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               
                
                if (!isValue && value.charAt(0) === "$") {
                    var parseRefs = parseRefStr(value);
                    Array.prototype.push.apply(refs, parseRefs.refKeys);
                    
                    type = parseRefs.type;
                } else {
                    if (!isValue) {
                        isValue = true;
                    }
                }
                
                if (isValue) {
                    var parseValue = value;
                    var splitIndex = value.indexOf("$");

                    var replacer = {}
                    replacer["s"] = function(value) {
                        return String(value);
                    }
                    replacer["n"] = function(value) {
                        return Number(value);
                    }
                    replacer["b"] = function(value) {
                        return value === "true";
                    }

                    if (splitIndex > 0) {
                        if (value.charAt(splitIndex - 1) !== "\\") {

                            var replacerName = value.substring(0, splitIndex);

                            if (replacer[replacerName]) {
                                parseValue = replacer[replacerName](value.substring(splitIndex));
                            } else {
                                console.error("존재하지 않는 replacer", replacerName);
                            }
                        }
                    }

                    values.push(parseValue);
                }
            }
    
            var evalRefStr = "obj";
            refs.forEach(function(ref, index, arr) {
                if (index > 0) {
                    var partRefStr = isNaN(ref) ? ("." + ref) : ("[" + ref + "]");

                    var beforeRef = eval(evalRefStr);
                    if (beforeRef === undefined) {
                        eval(evalRefStr + " = " + (isNaN(ref) ? "{}" : "[]"));
                    }

                    evalRefStr += partRefStr;
                }
            });

            if (type === "[]") {
                eval(evalRefStr + " = values");
            } else {
                eval(evalRefStr + " = values.join(\", \")");
            }
        });

        return {
            startKey: sheetNameRefs.refKeys[0],
            data: obj
        };
    }
    

    return {
        colToNum: colToNum,
        numToCol: numToCol,
        parseCell: parseCell,

        tabulateToArray: tabulateToArray,
        sheetToJson: sheetToJson
    }
})();