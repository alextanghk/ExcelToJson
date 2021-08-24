#!/usr/bin/env node
const ExcelParser = require("./index.js");
const util = require('util');
const path = require("path");
const fs = require("fs");

const yargs = require('yargs/yargs')
const { hideBin } = require('yargs/helpers')

const argv = yargs(hideBin(process.argv))
    .option('source',{
        alias: "s",
        string: true,
        default: "",
        describe: "Full path of the xlsx file."
    })
    .option('target',{
        alias: "t",
        string: true,
        default: "",
        describe: "Target path for output file."
    })
    .help()
    .alias('help', 'h')
    .argv;

const source = argv.source;
let target = argv.target == "" ? source.replace(/\.\w+$/,'.json'): argv.target;

const parser = new ExcelParser();
const result = parser.ParseJson(source,{headerAsIndex:true}).GetJson();
parser.SaveAsJson(target);