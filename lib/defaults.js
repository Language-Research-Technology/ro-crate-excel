  /* This is part of Calcyte a tool for implementing the RO-Crate data packaging
spec.  Copyright (C) 2018  University of Technology Sydney

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
/* Defaults for Calcyte such as names of key files */

const ro_crate_name = "ro-crate-metadata";
const html_multi_file_dirs = "ro-crate-metadata_files";
const metadata_json_file_name = `${ro_crate_name}.json`;
const metadata_excel_file_name = `${ro_crate_name}.xlsx`;
const additional_metadata_excel_file_name = `additional-${ro_crate_name}.xlsx`;
const html_file_name = "ro-crate-preview.html";
const max_depth = 10; // Number of dirs to recurse into
const maxFilesInDir = 30; // Don't list files in a directory if there are more than this
const ignore_file_regex = new RegExp(
  `(^~)|(^\\.)|^${html_file_name}$|^${metadata_json_file_name}$|^${metadata_excel_file_name}$`
);
const ignore_dir_regex = new RegExp(`(^${html_multi_file_dirs}$)|(^\\.)`);
const BagIt_Profile_Identifier =
  "https://raw.githubusercontent.com/UTS-eResearch/RO-Crate/master/spec/1.0/profile-datacrate-v1.0.json";
const ROCrate_Specification_Identifier =
  "https://github.com/UTS-eResearch/datacrate/blob/master/spec/1.0/data_crate_specification_v1.0.md";
const ROCrate_version = "1.01";
const ROCrate_profile_file = "defaults/profile-datacrate-v" + ROCrate_version + ".json"

const DEFAULTS = './defaults';
const defaults_dir = DEFAULTS;
const context = ["https://researchobject.github.io/ro-crate/1.0/context.jsonld", {"@vocab": "http://schema.org/"}];
const default_context = context;
const metadata_template =  `${DEFAULTS}/metadata_template.html`;
const render_script = "https://code.research.uts.edu.au/eresearch/CalcyteJS/raw/feature/or-crate/lib/render.js";

const display_keys = [
  "name",
  "familyName",
  "givenName",
  "@type",
  "description",
  "funder",
  "memberOf",
  "isPartOf",
  "fileOf",
  "thumbnail",
  "datePublished",
  "creator",
  "path",
  "encodingFormat",
  "contentSize",
  "affiliation",
  "email",

  "@reverse",
];

// RO-Crate specific terms which have inverses
const back_links = {
  hasFile: "fileOf",
  hasPart: "isPartOf",
  hasMember: "memberOf",
  memberOf: "hasMember"
  // hasMember: ["memberOf", "pcdm:memberOf"],
  // "pcdm:hasMember": ["memberOf", "pcdm:memberOf"],
  // memberOf: ["hasMember", "pcdm:hasMember"],
  // "pcdm:memberOf": ["hasMember", "pcdm:hasMember"]
};



const back_back_links = new Set(Object.values(back_links))

module.exports = {
  html_multi_file_dirs: html_multi_file_dirs,
  default_context: default_context,
  metadata_json_file_name: metadata_json_file_name,
  metadata_excel_file_name: metadata_excel_file_name,
  additional_metadata_excel_file_name: additional_metadata_excel_file_name,
  html_file_name: html_file_name,
  ignore_file_regex: ignore_file_regex,
  max_depth: max_depth,
  ignore_dir_regex: ignore_dir_regex,
  maxFilesInDir: maxFilesInDir,
  BagIt_Profile_Identifier: BagIt_Profile_Identifier,
  ROCrate_Specification_Identifier: ROCrate_Specification_Identifier,
  ROCrate_version: ROCrate_version,
  back_links: back_links,
  back_back_links: back_back_links,
  ROCrate_profile_file: ROCrate_profile_file,
  context: context,
  defaults_dir: defaults_dir,
  metadata_template: metadata_template,
  display_keys: display_keys,
  render_script: render_script,
  embedTypes: [], //["PropertyValue", "GeoCoordinates"],
  typeOrder: new Set(["Dataset", "File",  "Person",  "Organization", "Place", "ScholarlyArticle"])

};
