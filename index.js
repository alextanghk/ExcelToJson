const xlsx = require("xlsx");
const _ = require("lodash");
const fs = require("fs");
const path = require("path");
const { Console } = require("console");
const defaultConfig = {
    debugMode: false
};

function numToLetter(num) {
    var alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    var result = "";
    function printToLetter(number){
        var charIndex = number % alphabet.length
        var quotient = number/alphabet.length
        if(charIndex-1 == -1){
            charIndex = alphabet.length
            quotient--;
        }
        result =  alphabet.charAt(charIndex-1) + result;
        if(quotient>=1){
            printToLetter(parseInt(quotient));
        }
    }
    printToLetter(num);
    return result;
}

class ExcelParser {

    constructor(config) {
        this.parseResult = null;
        this.source = "";
        this.config = config === undefined ? Object.assign({},defaultConfig) : Object.assign({},defaultConfig, config) ;
    }

    ParseJson(source) {

        console.log("Reading file from %s", source);

        // Check if file exist
        if (!fs.existsSync(source)) {
            throw new Error("File not found");
        }

        // Reset variables
        this.parseResult = null;
        this.source = source;

        // Get workbook data
        const workbook = xlsx.readFile(this.source);

        const sheets = workbook.Sheets;
        const sheetsName = Object.keys(sheets);

        console.log("Total number of sheets: %d", sheetsName.length);

        let result = [];
        sheetsName.forEach(n => {
            const sheet = sheets[n];
            const data = xlsx.utils.sheet_to_json(sheet, {defval:""}, {range: sheet["!ref"]});
            result.push({
                name: n,
                total: data.length,
                header: data.length > 0 ? Object.keys(data[0]):[],
                data: data
            })
        })

        this.parseResult = result;
        return this;
    }

    GetJson() {
        return this.parseResult;
    }

    SaveAsJson(target) {
        if (this.parseResult === null) {
            throw new Error("Data not found");
        }

        target = target === undefined || target == "" ? this.source.replace(/\.\w+$/,'.json') : target; // set target path same as the source if target path is empty
        const defaultFileName = path.basename(this.source).replace(/\.\w+$/,'.json'); // Get file name

        // Create target dir if not exist
        if (!fs.existsSync(path.dirname(target))) {
            fs.mkdirSync(path.dirname(target),{recursive: true});
        }

        // Check if target path provide file name or not
        const regex = new RegExp(/\.json$/,'g');
        if (!regex.test(path.basename(target))) {
            target = path.join(target,defaultFileName);
        }

        console.log("Saving file to %s", target);

        // Create the file
        fs.writeFile(target, JSON.stringify(this.parseResult, null, 2), function(err) {
            if(err) {
                throw new Error(err);
            }
            console.log("File has been saved!");
        }); 
    }
}

module.exports = ExcelParser;