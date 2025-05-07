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

/* Test for workbook.js */

const fs = require('fs-extra');
const rocrate = require('ro-crate');
const ROCrate = rocrate.ROCrate;
const Workbook = require("../lib/workbook.js");
const assert = require("assert");
const chai = require('chai');

const expect = chai.expect;

// Fixtures
const metadataPath = "test_data/sample/ro-crate-metadata.json";
const IDRC_metadataPath = "test_data/IDRC/ro-crate-metadata.json";


describe("Create a workbook from a crate", function () {

    it("Should create a workbook with just one sheet", async function () {
        this.timeout(5000);
        const c = new ROCrate({array: true, link: true});
        c.name = "Test"

        const workbook = new Workbook({crate: c});
        await workbook.crateToWorkbook();
        const sheet = workbook.workbook.getWorksheet("RootDataset");
        console.log(sheet.getCell("A1").value, sheet.getCell("A2").value, sheet.getCell("A3").value, sheet.getCell("A4").value)
        assert.equal(
            sheet.getCell("A1").value,
            "Name"
        );
        assert.equal(
            sheet.getCell("B1").value,
            "Value"
        );
        assert.equal(
            sheet.getCell("A2").value,
            "@id"
        );
        assert.equal(
            sheet.getCell("B2").value,
            "./"
        );
    });


    it("Should create a workbook with one sheet and some metadata", async function () {
        this.timeout(5000);
        const c = new ROCrate();
        c.index();
        const root = c.getRootDataset();
        root["name"] = "My dataset";
        root["description"] = "Some old dataset";
        const workbook = new Workbook({crate: c});
        await workbook.crateToWorkbook();
        const rootSheetName = "RootDataset";
        datasetItem = workbook.sheetToItem(rootSheetName);
        assert.equal(Object.keys(datasetItem.item).length, 4)
        assert.equal(datasetItem.item.name, "My dataset");
        assert.equal(datasetItem.item.description, "Some old dataset");
        console.log(workbook.sheetDefaults)

    });


    it("Should create a workbook with two sheets", async function () {
        this.timeout(5000);

        const c = new ROCrate();
        c.index();
        const root = c.getRootDataset();
        root["name"] = "My dataset";
        root["description"] = "Some old dataset";
        c.addItem({
            "@id": "https://ror.org/03f0f6041",
            "name": "Universtiy of Technology Sydney",
            "@type": "Organization"
        })
        const workbook = new Workbook({crate: c});
        await workbook.crateToWorkbook();
        // This is not using the api - may be fragile
        assert.equal(workbook.workbook["_worksheets"].length, 4, "There are only two sheets");


    });

    it("Should handle the sample dataset", async function () {
        this.timeout(5000);

        var c = new ROCrate(JSON.parse(fs.readFileSync(metadataPath)), {array: true, link: true});

        const workbook = new Workbook({crate: c});
        await workbook.crateToWorkbook();
        workbook.workbook.xlsx.writeFile("test.xlsx");

        assert.equal(workbook.workbook["_worksheets"].length, 17, "Right number of tabs")
        const root = workbook.sheetToItem("RootDataset");
        assert.equal(root.item.publisher, `"http://uts.edu.au"`)

        // Name indexing works
        workbook.indexCrateByName();
        const pt = workbook.getItemByName("Peter Sefton")
        assert.equal(pt.name, "Peter Sefton")
        const s = workbook.workbook.getWorksheet("@type=Person");
        console.log("WORKBOOK", s.id)
        const items = workbook.sheetToItems(s.id);
        assert.equal(items.length, 1);
        assert.equal(items[0].name, "Peter Sefton");

    });


    it("Should handle the the IDRC (Cameron Neylon) dataset", async function () {
        this.timeout(5000);
        const excelFilePath = "METADATA_IDRC.xlsx";
        var c = new ROCrate(JSON.parse(fs.readFileSync(IDRC_metadataPath)), {array: true, link: true});

        const workbook = new Workbook({crate: c});
        await workbook.crateToWorkbook();
        //console.log(workbook.excel.Sheets)
        //assert.equal(workbook.workbook["_worksheets"].length, 15, "14 sheets")

        await workbook.workbook.xlsx.writeFile(excelFilePath);

        const workbook2 = new Workbook();
        await workbook2.loadExcel(excelFilePath);
        // Check all our items have survived the round trip
        //fs.writeFileSync("test.json", JSON.stringify(workbook2.crate.getJson(), null, 2));
        //console.log(workbook.crate.getRootDataset())
        for (let item of workbook2.crate.getGraph()) {
            if (item.name) {
                assert.equal(item.name[0], workbook.crate.getItem(item["@id"]).name[0])
            }
        }
        assert.equal(workbook.crate.getGraph().length, workbook2.crate.getGraph().length);


    });


    it("Can add to an existing crate", async function () {
        this.timeout(5000);
        const excelFilePath = "test_data/collections-workbook.xlsx";
        // New empty crate
        var c = new ROCrate({array: true, link: true});

        const workbook2 = new Workbook({crate: c});
        await workbook2.loadExcel(excelFilePath, true); // true here means add to crate not
        //console.log(JSON.stringify(workbook2.crate.toJSON(), null, 2));

        console.log("DEFAULTS", workbook2.sheetDefaults)

        const f = workbook2.crate.getEntity("/object2/1.mp4");
        assert(f);
        console.log(f['@type']);
        assert(f['@type'].includes('PrimaryMaterial'), "Picked up an extra type from isTypePrimaryMaterial column");
    });

    it("Can deal with there being no @context worksheet", async function () {
        this.timeout(5000);
        const excelFilePath = "test_data/collections-workbook-sans-context.xlsx";
        // New empty crate
        var c = new ROCrate({array: true, link: true});

        const workbook2 = new Workbook({crate: c});
        await workbook2.loadExcel(excelFilePath, true); // true here means add to crate
        assert(workbook2.crate.toJSON()["@context"].length === 2)
        //assert.equal(f.linguisticGenre[0]['@id'], "http://purl.archive.org/language-data-commons/terms#Dialogue", "Resolved context term")
        console.log(workbook2.crate.toJSON()["@context"])
    });

    it("Correctly adds the rdfs:label and rdfs:comment from name and description from rdf:Property and rdfs:Class for custom terms", async function () {
        this.timeout(5000);
        const excelFilePath = "test_data/custom_terms.xlsx";
        // New empty crate
        var c = new ROCrate({array: true, link: true});

        const workbook2 = new Workbook({crate: c});
        await workbook2.loadExcel(excelFilePath, true); // true here means add to crate
        const testProperty = c.getEntity('#testProperty')
        assert(testProperty.name[0] === 'Test Property')
        console.log(testProperty)
        assert(testProperty.name[0] === testProperty['rdfs:label'][0])
    });

    it("Can resolve double quoted references", async function () {
        var c = new ROCrate({array: true, list: true});

        c.addEntity({"@id": "#test1", name: "test 1"});
        c.addEntity({"@id": "#test2", name: "test 2"});
        c.addEntity({"@id": "#test3", name: "test 3"});
        c.addEntity({
                "@id": "#test4",
                name: "references",
                author: `"#test1"`, //By ID
                publisher: `"test2"`, // BY ID minus #
                contributor: `"test 3"` // By name
            }
        )
        const workbook = new Workbook({crate: c});
        await workbook.crateToWorkbook();
        workbook.resolveLinks();
        const item4 = workbook.crate.getEntity("#test4")
        //console.log(item4.author)
        assert.equal(item4.author[0]['@id'], "#test1");
        assert.equal(item4.publisher[0]['@id'], "#test2");
        assert.equal(item4.contributor[0]['@id'], "#test3");

    });


    it("Can deal with extra context terms", async function () {
        var c = new ROCrate({array: true, link: true});
        await c.resolveContext();

        c.addEntity({
            "@type": "Property",
            "@id": "http://example.com/mybetterprop",
            "label": "myBetterProp",
            "comment": "My description of my custom property",
        });

        c.addContext({"ldac":"http://w3id.org/ldac/terms#"});
        //Adding twice?
        c.addContext({"ldac":"http://w3id.org/ldac/terms#"});
        const ldacTerm = c.resolveTerm("ldac:linguisticGenre");
        assert(ldacTerm, "http://w3id.org/ldac/terms#linguisticGenre");

        c.addContext({"myBetterProp": "http://example.com/mybetterprop"});

        const term = c.resolveTerm('myBetterProp');
        assert(term, "http://example.com/mybetterprop");

        const workbook = new Workbook({crate: c});
        await workbook.crateToWorkbook();
        await workbook.workbook.xlsx.writeFile("test_context.xlsx");

        const contextSheet = workbook.workbook.getWorksheet("@context")
        assert(contextSheet.getRow(3).values[2], "http://example.com/mybetterprop");

    });


    it("Can export a workbook to a crate", async function () {
        this.timeout(5000);

        var c = new ROCrate(JSON.parse(fs.readFileSync(metadataPath)), {array: true, link: true});
        const graphLength = c.toJSON()["@graph"].length;
        const workbook = new Workbook({crate: c});
        await workbook.workbook.xlsx.writeFile("test-this.xlsx");

        await workbook.crateToWorkbook();

        await workbook.workbookToCrate();
        //console.log(JSON.stringify(workbook.crate.toJSON(), null, 2));
        expect(workbook.crate.toJSON()["@graph"].length).to.eql(graphLength);


    });


    it("Can handle mixed languages and various kinds of cell value", async function () {
        this.timeout(5000);
        const catalogPath = "test_data/mixed_lg/ro-crate-metadata.xlsx";
        const wb = new Workbook();
        await wb.loadExcel(catalogPath);
        // Start with a crate from a spreadsheet
        sourceCrate = wb.crate;
        const item = sourceCrate.getItem("ConcessionHealthCareCard/13655-1706ar.pdf")
        console.log(item.name);
        expect(item.name[0]).to.equal("وبطاقات الرعاية الصحية(بطاقات التخفيض)  Concession");
        const root = sourceCrate.getRootDataset();

        expect(root.datePublished[0]).to.equal("2022-01-10");
        expect(root.testProp[0]).to.equal("وبطاقات الرعاية الصحية(بطاقات التخفيض)  Concession");
        expect(root.SUM[0]).to.equal("5");

        expect(root.REFS[0]).to.equal("5Dataset");

    });


});

describe('Load Buffer', () => {

    it('should be able to load a workbook from buffer', async () => {
        const excelFilePath = "test_data/collections-workbook-sans-context.xlsx";
        const crate = new ROCrate({}, {array: true, link: true});
        const buffer = fs.readFileSync(excelFilePath)
        const wb = new Workbook({crate});
        await wb.loadExcelFromBuffer(buffer, true);
        const object1 = wb.crate.getItem('object1/');
        assert(object1['@type'][0] === 'RepositoryObject')
        assert(wb.log.info.length > 0);
        assert(wb.log.warning.length > 0)
    });
});


describe('Sheets', () => {
    it('should be able to ignore worksheets that start with . (dot)', async () => {
        const excelFilePath = "test_data/collections-workbook-sans-context.xlsx";
        const crate = new ROCrate({}, {array: true, link: true});
        const buffer = fs.readFileSync(excelFilePath);
        const wb = new Workbook({crate});
        await wb.loadExcelFromBuffer(buffer, true);
        const undefinedItem = wb.crate.getItem('/object1/1_sensitive.mpg');
        assert.strictEqual(undefinedItem, undefined, 'because the sheet name starts with dot, this should be undefined');
    });
});


describe('Can handle multi _isRef in a sheet', () => {
    it('should create 2 references to another object', async () => {
        const excelFilePath = "test_data/additional-multi/additional-ro-crate-metadata.xlsx";
        const crate = new ROCrate({}, {array: true, link: false});
        const buffer = fs.readFileSync(excelFilePath);
        const wb = new Workbook({crate});
        await wb.loadExcelFromBuffer(buffer, true);
        const object = wb.crate.getItem('#OBJECT_Juan');
        const speakers = object.speaker;
        assert.strictEqual(Array.isArray(speakers), true);
        assert.strictEqual(speakers.length, 2);
    });
    it('should create 1 reference even if it is referenced twice with the same id', async () => {
        const excelFilePath = "test_data/additional-multi/additional-ro-crate-metadata.xlsx";
        const crate = new ROCrate({}, {array: true, link: false});
        const buffer = fs.readFileSync(excelFilePath);
        const wb = new Workbook({crate});
        await wb.loadExcelFromBuffer(buffer, true);
        const object = wb.crate.getItem('#OBJECT_Emilia');
        const speakers = object.speaker;
        assert.strictEqual(Array.isArray(speakers), true);
        assert.strictEqual(speakers.length, 1);
    });
});

describe('Can send correct warnings', () => {
    it('should list warnings correctly', async () => {
        const excelFilePath = "test_data/additional_underscore_fields/additional-ro-crate-metadata.xlsx";
        const crate = new ROCrate({}, {array: true, link: false});
        const buffer = fs.readFileSync(excelFilePath);
        const wb = new Workbook({crate});
        await wb.loadExcelFromBuffer(buffer, true);
        assert.strictEqual(wb.log.warning.includes('Property something_somethingElse not defined in @context'), true);
    });
});

describe('Can handle excel columns as objects (such as hyperlinks)', () => {
    it('should handle the prop', async () => {
        const excelFilePath = "test_data/additional-with-non-text-cols/additional-ro-crate-metadata.xlsx";
        const crate = new ROCrate({}, {array: true, link: false});
        const buffer = fs.readFileSync(excelFilePath);
        const wb = new Workbook({crate});
        await wb.loadExcelFromBuffer(buffer, true);
        assert.strictEqual(wb.log.errors.length, 0);
    });
});


describe('Can handle root Id other than ./', () => {
    it('should handle the prop', async () => {
        const excelFilePath = "test_data/additional-rootId/additional-ro-crate-metadata.xlsx";
        const crate = new ROCrate({}, {array: true, link: false});
        const buffer = fs.readFileSync(excelFilePath);
        const wb = new Workbook({crate});
        await wb.loadExcelFromBuffer(buffer, true);
        assert.strictEqual(wb.log.errors.length, 0);
        const root = wb.crate.rootDataset;
        assert.strictEqual(root['@id'], 'arcp://name.my_org/My_Dataset!');
    });
});

describe('Can handle _isRef in a RootDataset (vertical)', () => {
    it('should create 2 references to another object', async () => {
        const excelFilePath = "test_data/additional-rootdataset/ro-crate-metadata-RootDataset_isRef.xlsx";
        const crate = new ROCrate({}, {array: true, link: false});
        const buffer = fs.readFileSync(excelFilePath);
        const wb = new Workbook({crate});
        await wb.loadExcelFromBuffer(buffer, true);
        const root = wb.crate.rootDataset;
        assert.strictEqual(root['@id'], 'TEST_ID');
        const author = [{"@id":"#LDaCA"}, {"@id":"#AARNET"}];
        const authorId = root['author'];
        assert.deepStrictEqual(authorId, author);
    });
    it('should handle isTerm_', async () => {
        const excelFilePath = "test_data/additional-rootdataset/ro-crate-metadata-RootDataset_isRef.xlsx";
        const crate = new ROCrate({}, {array: true, link: false});
        const buffer = fs.readFileSync(excelFilePath);
        const wb = new Workbook({crate});
        const term = wb.crate.resolveTerm('ldac:subjectLanguage');
        assert(!term);
        await wb.loadExcelFromBuffer(buffer, true);
        const term2 = wb.crate.resolveTerm('ldac:subjectLanguage');
        assert.strictEqual(term2, 'https://w3id.org/ldac/terms#subjectLanguage');
    });
});

describe('Can handle multiple context', () => {
    it('should handle the @context correctly', async () => {
        const excelFilePath = "test_data/additional-rootdataset/ro-crate-metadata-RootDataset_isRef.xlsx";
        const crate = new ROCrate({}, {array: true, link: false});
        const buffer = fs.readFileSync(excelFilePath);
        const wb = new Workbook({crate});
        await wb.loadExcelFromBuffer(buffer, true);
        const contextLength = wb.crate.context.length;
        await wb.loadExcelFromBuffer(buffer, true);
        console.assert(contextLength === wb.crate.context.length);
    });
});

describe('Can merge terms from different columns', () => {
    it('can handle multiple terms', async () => {
        const excelFilePath = "test_data/additional-multi/additional-ro-crate-metadata.xlsx";
        const crate = new ROCrate({}, {array: true, link: false});
        const buffer = fs.readFileSync(excelFilePath);
        const wb = new Workbook({crate});
        await wb.loadExcelFromBuffer(buffer, true);
        const object = wb.crate.getItem('#OBJECT_Emilia');
        const languages = object.language;
        assert.strictEqual(Array.isArray(languages), true);
        assert.strictEqual(languages.length, 3);
    });
    it('can handle multiple terms in RootDataset', async () => {
        const excelFilePath = "test_data/additional-rootdataset/ro-crate-metadata-RootDataset_isRef.xlsx";
        const crate = new ROCrate({}, {array: true, link: false});
        const buffer = fs.readFileSync(excelFilePath);
        const wb = new Workbook({crate});
        await wb.loadExcelFromBuffer(buffer, true);
        const rootDataset = wb.crate.rootDataset;
        const languages = rootDataset.inLanguage;
        assert.strictEqual(Array.isArray(languages), true);
        assert.strictEqual(languages.length, 3);
    });
});