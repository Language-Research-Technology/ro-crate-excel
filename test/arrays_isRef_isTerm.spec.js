const fs = require('fs-extra');
const rocrate = require('ro-crate');
const ROCrate = rocrate.ROCrate;
const Workbook = require("../lib/workbook.js");
const assert = require("assert");

describe('Can handle arrays in isRef_ and isTerm_ columns in a sheet', function() {


    it('should create 5 references from the isRef array', async () => {
        const excelFilePath = "test_data/arrays_isRef_isTerm.xlsx";
        const crate = new ROCrate({}, {array: true, link: false});
        const buffer = fs.readFileSync(excelFilePath);
        const wb = new Workbook({crate});
        await wb.loadExcelFromBuffer(buffer, true);
        const object = wb.crate.getItem('#TestObject');
        const researchParticipants = object['ldac:researchParticipant'];
        assert.strictEqual(researchParticipants.length, 5);
    });

    it('should create 2 references from the isTerm array', async () => {
        const excelFilePath = "test_data/arrays_isRef_isTerm.xlsx";
        const crate = new ROCrate({}, {array: true, link: false});
        const buffer = fs.readFileSync(excelFilePath);
        const wb = new Workbook({crate});
        await wb.loadExcelFromBuffer(buffer, true);
        const object = wb.crate.getItem('#TestObject');
        const communicationMode = object['ldac:communicationMode'];
        assert.strictEqual(communicationMode.length, 2);
    });
});