# ro-crate-excel

Node library with utilities for converting RO-Crates to Spreadsheet format for data entry and vice versa.

This replaces the Calcyte tool, both the [javascript](https://code.research.uts.edu.au/eresearch/calcytejs) and previous Python versions.

## What is this?

This is a library for building tools to assist in JSON-LD data entry. It has been built for RO-Crate but could be used for more general purpose JSON-LD <-> Spreadsheet conversion.

## Installation

### To use the script

```bash
npm install ro-crate-excel --global
```

### As a docker container (experimental)

Clone this repository, change into the root then make a container:

```bash
docker build -t rocxl .
```

Run the container:

```bash
docker run -v ~/path/to/data/:/data rocxl -d 5 /data
```

### For development

Get this repository (assuming you are working in `~/working`):

`git clone https://github.com/UTS-eResearch/ro-crate-excel.git  ~/working/ro-crate-excel`

Change into the directory:
`cd ~/working/ro-crate-excel`

Install the app:

```bash
npm install
npm link # to install the rocxl script
```

Run the tests:

```bash
npm test
```

or a specific test:

```bash
npm test test/update.crate.spec.js
```

# Usage

To run this code, use the `rocxl` (Excel <-> Research Object) script.

`rocxl` creates RO-Crates with an HTML entry-point in an `ro-crate-preview.html` file.

Usage:

```
./rocxl
Usage: rocxl [options] <directories...>
```

Generates an Excel spreadsheet and RO-Crate metadata from a set of files and
updates the RO-Crate with data filled in the spreadsheet.

To generate an Excel spreadsheet from an `ro-crate-metadata.json` file instead, use the `--JSON` option.

The file system is ALWAYS traversed and file information merged into
existing metadata.

```
Options:
  -V, --version                    output the version number
  -b,  --bag [bag-dir]             Create Bagit Bag(s) under [bag-dir])
  -a,  --add                       Add metadata from additional-ro-crate-metadata.xlsx to an existing ro-crate-metadata.json crate). Does not re-write the Excel input file or create ro-crate-metadata.xlsx.
  -z   --zip                       Zip the bagged ro-crate (only works with --bag
  -j   --JSON                      Use the ro-crate-metafata.json file rather than ro-crate-metadata.xlsx
  -p   --partOf [partOf]           This is part of another RO-Crate, supply the ro-crate-metadata.json path.
  -d,  --depth [depth]             Maximum depth to recurse into directories looking for files
  -r,  --recurse                   Recurse into directories looking for files
  -c,  --cratescript [cratesript]  URL of Crate-script directory
  -m,  --maxfiles [maxfiles]       Maximum number of files to itemise per directory (default is undefined)
  -u, --url [distro]               Distribution URL
  -h, --help                       output usage information
```

To run rocxl on a group of directories, pass it a list of directories.

One directory:

```bash
rocxl test_data/Glop_Pot -r
```

This will:

- Traverse the entire Glop_Pot directory, and generate or update the ro-crate-metadata.xlsx files.
- Create or update the `test_data/Glop_Pot/ro-crate-metadata.json` file
- Create an _[RO-Crate] Website_ with entry-point `test_data/Glop_Pot/ro-crate-metadata.html`

All the sample directories:

```bash
rocxl -r test_data/* -c https://data.research.uts.edu.au/examples/ro-crate/examples/src/crate.js
```

rocxl will generate:

- An `ro-crate-metadata.xlsx` file in each root directory (this is for humans to fill in with metadata about the data)

- An `ro-crate-metadata.json` file containing JSON-LD metadata derived from the spreadsheet some basic file-format information.

- An `ro-crate-preview.html` file generated from `ro-crate-metadata.json`

See the examples in `test_data`.

# To create an Excel file in which to describe a bunch of objects

To describe some things such as large numbers of predictably named files using a spreadsheet but use another tools such as Crate-O to describe the root dataset and top level context:

Assuming data is in `mydir`.

- Create an Excel crate in `mydir`

  `xlro -d 5 mydir`

- Move the resulting .xlsx file out of the way

  `mv mydir/ro-crate-metadata.xlsx mydir/additional-ro-crate-metadata.xlsx`

- Add whatever is needed to additional-ro-crate-metadata.xlsx to describe the files therein, and their relationship to RepositoryObject and RepositoryCollection entities

- Edit the `mydir/ro-crate-metadata.json` file with other tools of your choice

- Re-generate `mydir/ro-crate-metadata.json` with metadata from `mydir/additional-ro-crate-metadata.xlsx` by typing:
  `xlro -a mydir`

# About the spreadsheet format

This library allows transformation between RO-Crate and Excel spreadsheets using multiple worksheets in a workbook which is named `ro-crate-metadata.xlsx` and appears alongside the `ro-crate-metadata.json` file in the root of the dataset.

## SheetDefaults Worksheet

Optionally, a sheet name SheetDefaults can specify a default item template for that worksheet.

This sheet (at the moment) has two rows - the top row lists the names of worksheets that have default values and the second

| SheetName    | File              | RepositoryObject                                         | Person              |
| ------------ | ----------------- | -------------------------------------------------------- | ------------------- |
| itemtemplate | {"@type": "File"} | {"@type": "RepositoryObject", "license" : "LICENSE.txt"} | {"@type": "Person"} |

With the above configuration, every object in the sheet named `RepositoryObject` will have {"@type": "RepositoryObject", "license" : "LICENSE.txt"} as a starting point -- any additional values such as a `@type` column will be _added_ to the item.

## The `Root Dataset Worksheet`

The root dataset item is represented by a worksheet named "RootDataset" referred to as the `Root Dataset Worksheet`; this worksheet has two columns, `Name` and `Value`.
Each value of a property is represented as a row in the spreadsheet.

For example - the worksheet for this `Dataset`:

```json
{
  "@id": "./",
  "@type": "Dataset",
  "datePublished": "2017",
  "name": "Example Dataset",
  "identifier": "https://doi.org/10.4225/59/59672c09f4a4b",
  "description": "Do try to put more info in here than the title. Please.",
  "author": [
    { "@id": "https://orcid.org/0000-0002-3545-944X" },
    { "@id": "https://orcid.org/0000-0001-5152-5307" }
  ],
  "license": { "@id": "https://creativecommons.org/licenses/by-nc-sa/3.0/au/" }
}
```

Is structured as per this table:

| Name        | Value                                                   |
| ----------- | ------------------------------------------------------- |
| @id         | ./                                                      |
| @type       | Dataset                                                 |
| name        | Example Dataset                                         |
| description | Do try to put more info in here than the title. Please. |
| author      | "Peter Sefton"                                          |
| author      | "Michael Lynch"                                         |
| license     | "https://creativecommons.org/licenses/by-nc-sa/3.0/au/" |

### Multiple values use multiple lines

The multi-valued `author` property is represented as two rows as this is convenient to do in this vertically aligned Name/Value format.

### References

The quotes around the `Value`s for license and author indicate that the value is a reference to another item - these references can be to the `name` or `@id` property of the item. If the license and person items are JSON-LD items like these:

```json
{
  "@id": "https://creativecommons.org/licenses/by/4.0/",
  "@type": "CreativeWork",
  "name": "CC BY 4.0",
  "description": "Creative Commons Attribution 4.0 International License"
},
{
  "@id": "http://orcid.org/0000-0002-3545-944X",
  "@type": "Person",
  "name": "Peter Sefton",
  "familyName": "Sefton",
  "givenName": "Peter",
  "affiliation": {
    "@id": "https://ror.org/0384j8v12"
  }
}

{
  "@id": "https://orcid.org/0000-0001-5152-5307",
  "@type": "Person",
  "name": "Michael Lynch",
  "familyName": "Lynch",
  "givenName": "Michael",
  "affiliation": {
    "@id": "https://ror.org/0384j8v12"
  }
}
```

Then the Person and the CreativeWork will be described in two additional worksheets named `@type=Person` and `@type=CreativeWork`, these `Type Worksheets` use a different format to represent one item per line.

The `@type=Person` worksheet is as follows:

| @id                                   | @type  | name          | familyName | givenName | affiliation                 |
| ------------------------------------- | ------ | ------------- | ---------- | --------- | --------------------------- |
| http://orcid.org/0000-0002-3545-944X  | Person | Peter Sefton  | Sefton     | Peter     | "https://ror.org/0384j8v12" |
| https://orcid.org/0000-0001-5152-5307 | Person | Michael Lynch | Lynch      | Michael   | "https://ror.org/0384j8v12" |

And the `@type=CreativeWork` worksheet:

| @id                                          | @type        | name      | description                                            |
| -------------------------------------------- | ------------ | --------- | ------------------------------------------------------ |
| https://creativecommons.org/licenses/by/4.0/ | CreativeWork | CC BY 4.0 | Creative Commons Attribution 4.0 International License |

### Adding additional @types using isType\_\<Type>

If there is a column named `isType_<Type>` such as `istype_Annotation`, then rows representing items will have an additional type (e.g. `Annotation`) if the value of the cell evalutates to True (i.e. it has a non empty, no-zero value).

| @id          | @type        | name          | description      | isType_Annotation |
| ------------ | ------------ | ------------- | ---------------- | ----------------- |
| somefile.txt | CreativeWork | My annotation | A description of |

### Referring to other items

Columns with names that start with isRef\_ are converted as references to an ids references to an @id, e.g.:

| @id          | @type | isRef_hasAnnotation        |
| ------------ | ----- | -------------------------- |
| my_audio.wav | File  | ./my_audio_annotation.json |

Will be converted to:

```json
{
  "@id": "Mmy_audio.wav",
  "@type": "File",
  "hasAnnotation": { "@id": "./my_audio_annotation.json" }
}
```

Columns with names that start with isTerm\_ are treated as references to vocabulary items that are defined in the context:

So assuming the @context sheet contains:

| name | @id                          |
| ---- | ---------------------------- |
| ldac | https://w3id.org/ldac/terms# |

| @id          | @type | isTerm_annotationType |
| ------------ | ----- | --------------------- |
| my_audio.wav | File  | ldac:Dialogue         |

The resulting item will be:

```json
{
  "@id": "my_audio.wav",
  "@type": "File",
  "annotationType": {
    "@id": "http://purl.archive.org/language-data-commons/terms#Dialogue"
  }
}
```

### Representing multiple values

To represent multiple values - for example, if there are multiple affiliations for a person there are two ways to accomplish this:

1.  Repeat a column header with the SAME name as many times as needed (for the maximum number of repeats in the column), OR
2.  Use a comma-separated list enclosed in square brackets

| @id                                  | @type  | name         | FamilyName | givenName | affiliation                                        |
| ------------------------------------ | ------ | ------------ | ---------- | --------- | -------------------------------------------------- |
| http://orcid.org/0000-0002-3545-944X | Person | Peter Sefton | Sefton     | Peter     | ["https://ror.org/0384j8v12", http://ptsefton.com] |

This approach can also be used in the `Root Dataset Worksheet`. The URL is treated as a string value as it is not enclosed in double quotes.

| Name   | Value                                 |
| ------ | ------------------------------------- |
| author | ["Peter Sefton", http://ptsefton.com] |

These values will be interpreted as references, omitting the quotes will cause a value to be interpreted as a string.

### Hiding values (and showing provenance)

To stop a column in the spreadsheet from being copied to the output crate, add a "." to the name. For example, if an orignal data source uses the term `Title`, then in order to show the provenance of the data, create a column called `name` (which is the RO-Crate correct term for the name of a work), and use a formula to copy the data into the `name` column.

| .Title                               | name |
| ------------------------------------ | ---- |
| A Short Introduction to Spreadsheets | =A1  |

### Embedding JSON

To avoid having to create `@type Worksheets` for things such as `GeoCoordinates` or `PropertyVale` items, items or arrays of items may be embedded in a cell using standard JSON-LD.

For example, this `Place` item:

```json
{
  "@id": "http://www.geonames.org/8152662/catalina-park.html",
  "@type": "Place",
  "address": "Katoomba, NSW",
  "description": "Catalina Park is a disused motor racing venue, located at Katoomba, in the Blue Mountains, New South Wales, Australia, and is recognised as an Aboriginal Place due to the long association of the local Gundungarra and Darug clans to the area.",
  "geo": {
    "@id": "#d2c5b5e0-a720-4b21-ad3a-f44ad89488e7"
  },
  "name": "Catalina Park"
}
```

Can be represented in the `@type=Place` worksheet:

| @id                                                | @type | address       | description                                                                                                                                                                                                                                       | geo                                                                                                                                                                      | name          |
| -------------------------------------------------- | ----- | ------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------ | ------------- |
| http://www.geonames.org/8152662/catalina-park.html | Place | Katoomba, NSW | Catalina Park is a disused motor racing venue, located at Katoomba, in the Blue Mountains, New South Wales, Australia, and is recognised as an Aboriginal Place due to the long association of the local Gundungarra and Darug clans to the area. | {"@id":"#d2c5b5e0-a720-4b21-ad3a-f44ad89488e7","@type":"GeoCoordinates","latitude":"-33.7152","longitude":"150.30119","name":"Latitude: -33.7152, Longitude: 150.30119"} | Catalina Park |

NOTE: Any cell that contains at least one `{` and one `}` will be parsed as JSON - if that fails it will be included as an escaped string.

# Adding additional properties to the @context

TODO

## If the URL for a property does not resolve to a useful URL

In this case, define an item of @type `Property` in the `@type=Property` worksheet (or if you're starting with a crate, add an item of `@type` `Property` to the graph.) The `@id` should be the URL of the fully resolved property - to use the example from the spec:

| @id                                       | @type    | name | description | sameAs                                                 |
| ----------------------------------------- | -------- | ---- | ----------- | ------------------------------------------------------ |
| http://purl.org/ontology/bibo/interviewee | Property |      |             | http://neologism.ecs.soton.ac.uk/bibo.html#interviewee |

```json
{
  "@context": [
    "https://w3id.org/ro/crate/1.0/context",
    { "interviewee": "http://purl.org/ontology/bibo/interviewee" }
  ],
  "@graph": [
    {
      "@id": "http://purl.org/ontology/bibo/interviewee",
      "sameAs": "http://neologism.ecs.soton.ac.uk/bibo.html#interviewee",
      "@type": "Property"
    }
  ]
}
```

## If the property is locally defined

To define a local property which is specific to a dataset or because there is no available public ontology that has one, define it in the graph as an item of `@type` `rdf:Property`, as [per the RO-Crate Spec advice on ad hoc terms](https://www.researchobject.org/ro-crate/1.1/appendix/jsonld.html#add-local-definitions-of-ad-hoc-terms).

```json
{
  "@context": [
    "https://w3id.org/ro/crate/1.0/context",
    { "myProp": "https://w3id.org/ro/terms/myNameSpace/#myProp" }
  ],
  "@graph": [
    {
      "@id": "https://w3id.org/ro/terms/myNameSpace/#myProp",
      "@type": "rdf:Property",
      "rdfs:label": "myProp",
      "rdfs:comment": "This is my custom property I want to use in describing things"
    }
  ]
}
```

Which on conversion to Excel would look like:

| @id       | @type  | rdfs:label                                                    | rdfs:comment | sameAs |
| --------- | ------ | ------------------------------------------------------------- | ------------ | ------ |
| \_:myProp | myProp | This is my custom property I want to use in describing things |

TODO: Make @context entries for additional `Property` items automatically show up in the `@context` if not already defined - and force appropriate IDs (they must be either full http(s) URIs or blank node `@id`s and start with a lowercase letter).

# Implementing Workbook to RO-Crate

When converting from a worksheet to a JSON-LD item, the process is to:

- Convert the `Root DataSet Workseet` to an RO-Crate `Root Dataset` - with the necessary `@id` and other

- Convert each `@type sheet` to an item by mapping column names to properties; each row becomes an item in the RO-Crate graph.

- Index the crate by `@id` and by `name`

- For every item in the `@graph` array:
- Normalise the item's `@id`:
- If the `@id` is a URL, or the item is, in RO-Crate terms a `Data Entity` - that has @type `File` or `Dataset` or it starts with `#` then leave it as-is.
- else prepend `#` to the `@id`
- for each value of a property that starts and ends with double quotes:
- If the value matches a known `@id` then add a reference `{"@id": "#someid"}`
- else if the value (without quotes) matches a known name add a reference to the item with that name
- else if the value (without quotes) does not start with `#` prepend `#` and see if it matches a known `@id` - if it does, add it as a reference
