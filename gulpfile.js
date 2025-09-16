const { src, dest } = require('gulp');

function buildIcons() {
  return src('*.svg')
    .pipe(dest('dist/icons'))
    .pipe(dest('dist/nodes/OnlyOffice'));
}

function buildStructure() {
  return src('dist/onlyoffice_credentials.js')
    .pipe(dest('dist/credentials'))
    .pipe(src('dist/onlyoffice_node.js'))
    .pipe(dest('dist/nodes/OnlyOffice'));
}

exports['build:icons'] = buildIcons;
exports['build:structure'] = buildStructure;
