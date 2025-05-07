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
const fixtures = require("./fixtures");
const update = require("../lib/update.js");
const assert = require("assert");
const chai = require('chai');
const Workbook = require("../lib/workbook.js");
const expect = chai.expect;   
const path = require("path");
const rocrate = require("ro-crate");
const ROCrate = rocrate.ROCrate;




describe("Create workbooks", function() {
  var testPath;

  before(function () { testPath = fixtures.buildup(); });
  //after(function () { fixtures.teardown(); });

  it("Should create a workbook from the sample data", async function() {
    this.timeout(15000); 

    var msg = await update(testPath, 5);
    assert.equal(msg[0], `No metadata or catalog found, making: ${testPath}/ro-crate-metadata.json`)
    
    // Second go should use the ro-crate.metadata.json file
    console.log("Test path", testPath)
    msg = await update(testPath, 5);
    assert.equal(msg[0], `Using existing spreadsheet: ${testPath}/ro-crate-metadata.xlsx`)

    // Add something to the workbook...
    const wb = new Workbook();
    const catalogPath = path.join(testPath, "ro-crate-metadata.xlsx");

    await wb.loadExcel(catalogPath);

    const sheet = wb.workbook.getWorksheet("RootDataset");
    sheet.getCell("A4").value = "name";
    sheet.getCell("B4").value = "A Dataset";
    await wb.workbook.xlsx.writeFile(catalogPath);
    msg = await update(testPath, 5);
    const finalCrate = new ROCrate(JSON.parse(await fs.readFile(path.join(testPath, "ro-crate-metadata.json"))));
    console.log(finalCrate.toJSON())
    assert.equal("A Dataset", finalCrate.rootDataset.name)

    assert.equal(msg[0], `Using existing spreadsheet: ${testPath}/ro-crate-metadata.xlsx`)

    msg = await update(testPath, 5);



   
  });





});





