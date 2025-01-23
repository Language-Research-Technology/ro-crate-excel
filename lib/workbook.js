/* This is part of RO-Crate-excel a tool for implementing the DataCrate data packaging
spec.  Copyright (C) 2020  University of Technology Sydney

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <http://www.gnu.org/licenses/>.
*/

const Excel = require('exceljs');
const defaults = require('./defaults');
const {ROCrate} = require('ro-crate');
const _ = require("lodash");
const {v4: uuidv4} = require('uuid');

class Workbook {
    constructor(opts) {
        this.propertyWarnings = {};
        if (opts && opts.crate) {
            this.workbook = new Excel.Workbook();
            this.crate = opts.crate;
        }
        this.log = {
            info: [],
            warning: [],
            errors: []
        };
    }

    async crateToWorkbook() {
        this.crate.resolveContext();
        var sheetName = "RootDataset";
        const root = this.crate.rootDataset;
        // Turn @id refs into inferred references
        //this.collapseReferences();
        this.indexByType();

        const worksheet = this.workbook.addWorksheet(sheetName);
        worksheet.views = [
            {state: 'frozen', xSplit: 1, activeCell: 'A2'}
        ];
        worksheet.columns = [
            {header: 'Name', key: 'Name', width: 10},
            {header: 'Value', key: 'Value', width: 100}
        ]
        for (let prop of Object.keys(root)) {
            worksheet.addRow({Name: prop, Value: this.formatMultiple(root[prop])});
        }
        this.addContextSheet();

        /* Add the worksheets for each type to the workbook */
        const types = new Set([...defaults.typeOrder, ...Object.keys(this.crate.types)])
        for (let t of types) {
            if (this.crate.types[t]) {
                const sheet = this.workbook.addWorksheet(`@type=${t.replace(/:/, "")}`);
                const cols = {}
                var columns = []
                for (let item of this.crate.types[t]) {
                    for (let prop of Object.keys(item)) {
                        if (!cols[prop]) {
                            cols[prop] = prop;
                            columns.push({header: prop, key: prop, width: 20});
                        }
                    }
                }
                sheet.columns = columns;
                sheet.addRows(this.crate.types[t].map(item => {
                    return this.formatItem(item)
                }));
            }
        }
    }


    /**
     * Load in an existing spreadshsset
     * @param {string} [filename] - Path to an Excel sheet
     * @param {boolean} [addToCrate] - Optionaly don't treat the spreadsheet as containing a whole crate, just load in worksheets into the existing crate
     */

    async loadExcel(filename, addToCrate) {
        var addToExistingCrate = false;
        if (addToCrate) {
            addToExistingCrate = true;
        }
        this.workbook = new Excel.Workbook();
        await this.workbook.xlsx.readFile(filename);
        await this.workbookToCrate(addToExistingCrate);
    }

    async loadExcelFromBuffer(buffer, addToCrate) {
        var addToExistingCrate = false;
        if (addToCrate) {
            addToExistingCrate = true;
        }
        this.workbook = new Excel.Workbook();
        await this.workbook.xlsx.load(buffer);
        await this.workbookToCrate(addToExistingCrate);
    }

    addContextSheet() {
        const context = this.crate.getJson()["@context"];
        this.contextWorksheet = this.workbook.addWorksheet("@context");
        this.contextWorksheet.columns = [
            {header: "name", key: "name", width: "20"},
            {header: "@id", key: "@id", width: "60"},
        ];
        for (let contextBlock of context) {
            this.addContextTerms(contextBlock);
        }
    }

    addContextTerms(terms) {
        if (typeof terms === 'string') {
            this.contextWorksheet.addRow({"URL": terms});
        } else {
            for (let term of Object.keys(terms)) {
                if (!(term === "@base") && !(term === "@vocab" && terms[term] === "http://schema.org/")) {
                    const item = this.crate.getItem(terms[term]) || {};
                    const row = {};
                    row.name = item["name"] || term;
                    row["@id"] = terms[term];
                    this.contextWorksheet.addRows([row]);
                }
            }
        }
    }

    indexCrateByName() {
        // TODO - warn about duplicate names
        this.crate.itemByName = {};
        for (let item of this.crate.getGraph()) {
            if (item.name) {
                this.crate.itemByName[item.name] = item;
            }
        }

    }

    getItemByName(name) {
        if (this.crate.itemByName[name]) {
            return this.crate.itemByName[name];
        } else {
            return null;
        }
    }

    indexByType() {
        this.crate.types = {};
        for (let item of this.crate.entities()) {
            if (!(item["@id"] && item["@type"] && (item["@id"] === this.crate.rootDataset["@id"] || item["@id"].match(/^ro-crate-metadata.json(ld)?$/)))) {
                // Only need to check first type cos we don't want this thing showing up in two places?
                const t = item["@type"][0];
                if (!defaults.embedTypes.includes(t)) {
                    // const stringifiedItem = this.stringify(item); Legacy feature that was a bad idea
                    if (!this.crate.types[t]) {
                        this.crate.types[t] = [item];
                    } else {
                        this.crate.types[t].push(item);
                    }
                }
            }
        }
    }

    formatItem(item) {
        const formattedItem = {};
        for (let p of Object.keys(item)) {
            formattedItem[p] = this.formatMultiple(item[p])
        }
        return formattedItem;
    }

    formatMultiple(vals) {
        if (Array.isArray(vals)) {
            if (vals.length > 1) {
                return `[${vals.map(v => this.formatSingle(v)).join(", ")}]`;
            } else {
                return this.formatSingle(vals[0]);
            }
        } else {
            return this.formatSingle(vals);
        }
    }

    formatSingle(val) {
        if (val?.["@id"]) {
            return `"${val["@id"]}"`;
        } else {
            return val;
        }
    }

    /*
       For worksheets that have a vertical (Name | Value layout) - turn them into a JSON-LD item
       Returns: {item: object, multipleRefs: array}
       TODO: - deal with wrong headers - no Name / Value (other stuff will be discarded)
     */
    sheetToItem(sheetName) {
        const item = {};
        const sheet = this.workbook.getWorksheet(sheetName);
        const me = this;
        //console.log(sheetName, sheet)
        // TODO - deal with repeated props
        const multipleRefs = [];
        if (sheet) {
            sheet.eachRow(function (row, rowNumber) {
                if (rowNumber > 1) {
                    //TODO - normalise property values to lowercase?
                    const prop = row.values[1];
                    const val = row.getCell(2)

                    // me.validProp(prop);
                    // // TODO check prop is in context and optionaly normalize
                    // const val = row.getCell(2);
                    // item[prop] = me.parseCell(val);

                    if (prop && !prop.startsWith(".")) {
                        me.validProp(prop);
                        const propValue = me.parseCell(val); // TODO -- remove [] and {} if they're there
                        if (prop && prop.startsWith("isRef_") && propValue) {
                            const p = {};
                            p[prop.replace("isRef_", "")] = {"@id": propValue}
                            multipleRefs.push(p);
                        } else if (prop && prop.startsWith("isTerm_") && propValue) {
                            const resolvedTerm = me.crate.resolveTerm(propValue);
                            if (resolvedTerm) {
                                item[prop.replace("isTerm_", "")] = {"@id": resolvedTerm};
                            } else {
                                item[prop.replace("isTerm_", "")] = {"@id": propValue};
                            }
                        } else {
                            if (item[prop]) {
                                item[prop] = [item[prop], propValue];
                            } else {
                                item[prop] = propValue;
                            }
                        }
                    }
                }
            });
        }
        return {item, multipleRefs};
    }

    parseContextSheet() {
        // For worksheets that have a vertical (Name | Value layout) - turn them into a JSON-LD item
        // TODO - deal with wrong headers - no Name / Value (other stuff will be discarded)
        const additionalContext = {};
        const sheet = this.workbook.getWorksheet("@context");
        const me = this;
        if (sheet) {
            sheet.eachRow({includeEmpty: false}, function (row, rowNumber) {
                if (rowNumber > 1) {
                    additionalContext[row.values[1]] = me.parseCell(row.getCell(2));
                }
            });
        }
        return additionalContext;
    }

    parseSheetDefaults() {
        // Look up some config stuff
        // TODO: Use Fields when making NEW sheets??? (do we really need this?)
        this.sheetDefaults = {};
        const sheet = this.workbook.getWorksheet('SheetDefaults');
        const me = this;
        if (sheet) {
            const columns = sheet.getRow(1).values;
            sheet.eachRow({includeEmpty: false}, function (row, rowNumber) {
                if (rowNumber === 2) {
                    var sheetName = "";
                    row.eachCell({includeEmpty: false}, function (cell, cellNumber) {
                        if (cellNumber > 1) {
                            if (columns[cell._column._number]) {
                                const sheetName = columns[cell._column._number]
                                me.sheetDefaults[sheetName] = me.parseCell(cell);
                                me.log.info.push(`Added Sheet: ${me.sheetDefaults[sheetName]?.['@type']}`);
                                console.log(`Added Sheet: ${me.sheetDefaults[sheetName]?.['@type']}`);
                            }
                        }
                    });
                }
            });
        }
    }

    parseCell(val) {
        var cellString = '';
        // TODO: Not sure if these first two are needed
        if (val.text) {
            cellString = val.text;
        } else if (val.result) {
            cellString = val.result;
        } else if (val.value && val.value.richText) {
            for (let t of val.value.richText) {
                cellString += t.text;
            }
        } else {
            cellString = val.value;
        }
        // Look for curly braces - if they're there then it's JSON
        if (!cellString) {
            return "";
        }
        cellString = cellString.toString();
        const curly = cellString.match(/{(.*)}/);
        if (curly) {
            try {
                return JSON.parse(cellString);
            } catch (error) {
                return cellString;
            }
        } else {
            // Look for arrays of strings or references
            const sq = cellString.match(/^\s*\[\s*(.*?)\s*\]\s*$/);
            if (sq) {
                return sq[1].split(/\s*,\s*/)//.map(x => this.parseCell(x));
            }
        }
        if (cellString === `""`) cellString = null;
        return cellString;
    }

    // Turns a worksheet into  set of items, row by row
    sheetToItems(sheetID) {
        // TODO get default type
        const worksheet = this.workbook.getWorksheet(sheetID);
        if (!this.propertyWarnings) {
            this.propertyWarnings = {};
        }
        const columns = worksheet.getRow(1).values;
        const sheetName = worksheet.name;
        var items = [];
        const me = this;
        worksheet.eachRow({includeEmpty: false}, function (row, rowNumber) {
            if (rowNumber > 1) {
                var item = {}
                if (me.sheetDefaults && me.sheetDefaults[sheetName]) {
                    item = _.clone(me.sheetDefaults[sheetName]);
                }
                const additionalTypes = [];
                const multipleRefs = [];
                row.eachCell({includeEmpty: false}, function (cell, cellNumber) {
                    if (columns[cell._column._number]) {
                        let prop = columns[cell._column._number];
                        // Some columns are not plain text strings they can be objects like hyperlinks or formulas
                        if (typeof prop === 'object' && prop !== null) {
                            //text is for hyperlink and result is the result of a formula
                            const value = prop?.text || prop?.result;
                            if (!value) {
                                //Keep prop to print it to console and exit loop
                                me.log.errors.push(`Error with column: ${JSON.stringify(prop)}`);
                                return;
                            } else {
                                prop = value;
                            }
                        }
                        if (prop && !prop.startsWith(".")) {
                            me.validProp(prop);
                            const propValue = me.parseCell(cell); // TODO -- remove [] and {} if they're there
                            if (prop && prop.startsWith("isType_") && propValue) {
                                additionalTypes.push(prop.replace("isType_", ""));
                            } else if (prop && prop.startsWith("isRef_") && propValue) {
                                const p = {};
                                p[prop.replace("isRef_", "")] = {"@id": propValue}
                                multipleRefs.push(p);
                            } else if (prop && prop.startsWith("isTerm_") && propValue) {
                                const resolvedTerm = me.crate.resolveTerm(propValue);
                                if (resolvedTerm) {
                                    item[prop.replace("isTerm_", "")] = {"@id": resolvedTerm};
                                } else {
                                    item[prop.replace("isTerm_", "")] = {"@id": propValue};
                                }
                            } else {
                                if (item[prop]) {
                                    item[prop] = [item[prop], propValue];
                                } else {
                                    item[prop] = propValue;
                                }
                            }
                        }
                    }
                });
                if (!item["@type"]) {
                    item["@type"] = [];
                }
                if (!Array.isArray(item["@type"])) {
                    item["@type"] = [item["@type"]];
                }
                if (multipleRefs.length > 0) {
                    for (let k of multipleRefs) {
                        const obj = [];
                        let p;
                        for (let [key, value] of Object.entries(k)) {
                            p = key;
                            obj.push(value);
                        }
                        if (!item[p]) {
                            item[p] = [].concat(obj);
                        } else {
                            item[p] = item[p].concat(obj);
                        }
                    }
                }
                item["@type"] = item["@type"].concat(additionalTypes);
                if (!item["@type"]) {
                    item["@type"] = ["Thing"];
                }
                items.push(item);
            }

        });
        // deal with embedded stuff
        return items;
        // Add to graph
    }

    async workbookToCrate(addToExistingCrate) {
        console.log('addToExistingCrate', addToExistingCrate)
        if (!addToExistingCrate || !this.crate) {
            this.crate = new ROCrate({"array": true, "link": true});
        }

        await this.crate.resolveContext();
        const parsedContextSheet = this.parseContextSheet();
        if (Object.keys(parsedContextSheet).length) {
            this.crate.addContext(parsedContextSheet);
        }
        this.parseSheetDefaults();
        //Updating the rootId if newRoot; if RootDataset exists
        const newRoot = this.sheetToItem("RootDataset");
        for (let prop of Object.keys(newRoot.item)) {
            if (prop === "@id") {
                this.crate.updateEntityId(this.crate.rootDataset, newRoot.item["@id"]);
            } else {
                this.crate.rootDataset[prop] = newRoot.item[prop];
            }
        }
        this.crate.rootDataset["@type"] = this.crate.rootDataset["@type"].concat(newRoot.additionalTypes);
        if (!this.crate.rootDataset["@type"]) {
            this.crate.rootDataset["@type"] = ["Dataset"];
        }
        if (!Array.isArray(this.crate.rootDataset["@type"])) {
            this.crate.rootDataset["@type"] = [this.crate.rootDataset["@type"]];
        }
        const root = this.crate.rootDataset;
        for (let i of newRoot.multipleRefs) {
            this.crate.addEntity(i, {replace: true, recurse: true});
            const obj = [];
            let p;
            for (let [key, value] of Object.entries(i)) {
                p = key;
                obj.push(value);
            }
            if (!root[p]) {
                root[p] = [].concat(obj);
            } else {
                root[p] = root[p].concat(obj);
            }
        }
        const me = this;
        this.workbook.eachSheet(function (worksheet, worksheetID) {
            if (!["@context", "RootDataset", "config", "SheetDefaults"].includes(me.workbook.getWorksheet(worksheetID).name)) {
                if (!me.workbook.getWorksheet(worksheetID).name.startsWith('.')) {
                    me.log.info.push(`Reading: ${me.workbook.getWorksheet(worksheetID).name}`);
                    const items = me.sheetToItems(worksheetID);
                    for (let item of items) {
                        if (!item["@id"]) {
                            me.log.warning.push(`Item does not have an @id ${JSON.stringify(item)}`);
                            console.log("Warning item does not have an @id", item);
                            item["@id"] = `#${uuidv4()}`;
                        }
                        me.log.info.push(`Added: ${item['@type']} - ${item["@id"]}`);
                        me.crate.addEntity(item, {replace: true, recurse: true});
                    }
                } else {
                    me.log.info.push(`Ignoring Sheet (.): ${me.workbook.getWorksheet(worksheetID).name}`);
                }
            }
        });

        this.resolveLinks();
        this.addBackLinks();
    }

    resolveLinks() {
        // Anything in "" is potentially a reference by ID or Name
        this.crate.index();
        this.indexCrateByName();
        for (let item of this.crate.getGraph()) {
            for (let prop of Object.keys(item)) {
                var vals = [];
                for (let val of this.crate.utils.asArray(item[prop])) {
                    if (val && !val["@id"]) {
                        var linkMatch = val.toString().match(/^"(.*)"$/);
                        if (linkMatch) {
                            const potentialID = linkMatch[1];
                            // log("ID", potentialID)
                            if (this.crate.getItem(potentialID)) {
                                vals.push({"@id": this.crate.getItem(potentialID)["@id"]});
                            } else if (this.crate.resolveTerm(potentialID)) {
                                vals.push({"@id": this.crate.resolveTerm(potentialID)});
                            } else if (this.getItemByName(potentialID)) {
                                vals.push({"@id": this.getItemByName(potentialID)["@id"]});
                            } else if (this.crate.getItem(`#${potentialID}`)) {
                                vals.push({"@id": this.crate.getItem(`#${potentialID}`)["@id"]});
                            } else {
                                vals.push(val);
                            }
                        } else {
                            vals.push(val);
                        }
                    } else {
                        vals.push(val);
                    }
                }
                if (vals.length === 1) {
                    vals = vals[0];
                } else if (vals.length === 0) {
                    vals = "";
                }
                item[prop] = vals;
            }
        }
    }

    backLinkItem(item) {
        for (let p1 in defaults.back_links) {
            const p2 = defaults.back_links[p1];
            if (item[p1]) {
                for (let i of item[p1]) {
                    if (i['@id']) {
                        const target = this.crate.getEntity(i['@id']);
                        if (target) {
                            this.crate.addValues(target, p2, item);
                        }
                    }
                }
            }
            if (item[p2]) {
                for (let i of item[p2]) {
                    if (i['@id']) {
                        const target = this.crate.getEntity(i['@id']);
                        if (target) {
                            this.crate.addValues(target, p1, item);
                        }
                    }
                }
            }
        }
    }

    addBackLinks() {
        // Add @reverse properties if not there
        for (let item of this.crate.graph) {
            this.backLinkItem(item);
        }
    }

    validProp(prop) {
        prop = prop.replace(/^is(.*?)_/, "");
        if (prop && !prop.startsWith("@") && !this.crate.resolveTerm(prop)) {
            if (!this.propertyWarnings[prop]) {
                this.propertyWarnings[prop] = true;
                this.log.warning.push(`Property ${prop} not defined in @context`);
                console.log("Warning: undefined property", prop);
            }
        }
    }
}

module.exports = Workbook;



