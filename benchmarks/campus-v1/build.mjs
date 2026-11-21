import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { deflateSync } from 'node:zlib';

// Independently authored OPC fixture packages, with fixed timestamps and stored ZIP entries.
const here = path.dirname(fileURLToPath(import.meta.url));
const inputs = path.join(here, 'inputs');
const design = JSON.parse(readFileSync(path.join(here, 'design.json')));
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.';
const cache = Array.from({length: 256}, (_, i) => {
  for (let j=0; j<8; j++) i = i & 1 ? 0xEDB88320 ^ (i >>> 1) : i >>> 1;
  return i >>> 0;
});
function crc32(data) {
  let crc = 0xFFFFFFFF;
  for (const byte of data) crc = cache[(crc ^ byte)&255] ^ (crc >>> 8);
  return (crc ^ 0xFFFFFFFF) >>> 0;
}
function zip(files) {
  const entries=[], directory=[];
  let offset=0;
  for (const [name, raw] of Object.entries(files)) {
    const label=Buffer.from(name), data=Buffer.from(raw), crc=crc32(data);
    const local=Buffer.alloc(30), central=Buffer.alloc(46);
    local.writeUInt32LE(0x04034B50); local.writeUInt16LE(20,4); local.writeUInt16LE(33,12);
    local.writeUInt32LE(crc,14); local.writeUInt32LE(data.length,18); local.writeUInt32LE(data.length,22);
    local.writeUInt16LE(label.length,26);
    central.writeUInt32LE(0x02014B50); central.writeUInt16LE(20,4); central.writeUInt16LE(20,6);
    central.writeUInt16LE(33,14); central.writeUInt32LE(crc,16);
    central.writeUInt32LE(data.length,20); central.writeUInt32LE(data.length,24);
    central.writeUInt16LE(label.length,28); central.writeUInt32LE(offset,42);
    entries.push(local,label,data); directory.push(central,label);
    offset += local.length+label.length+data.length;
  }
  const index=Buffer.concat(directory), end=Buffer.alloc(22);
  end.writeUInt32LE(0x06054B50); end.writeUInt16LE(directory.length/2,8);
  end.writeUInt16LE(directory.length/2,10); end.writeUInt32LE(index.length,12); end.writeUInt32LE(offset,16);
  return Buffer.concat([...entries,index,end]);
}
function swatch(rgb) {
  const chunk = (type, data) => {
    const t=Buffer.from(type), len=Buffer.alloc(4), crc=Buffer.alloc(4);
    len.writeUInt32BE(data.length); crc.writeUInt32BE(crc32(Buffer.concat([t,data])));
    return Buffer.concat([len,t,data,crc]);
  };
  const ihdr=Buffer.alloc(13);
  ihdr.writeUInt32BE(1); ihdr.writeUInt32BE(1,4); ihdr[8]=8; ihdr[9]=2;
  return Buffer.concat([Buffer.from([137,80,78,71,13,10,26,10]),chunk('IHDR',ihdr),
    chunk('IDAT',deflateSync(Buffer.from([0,...rgb]))),chunk('IEND',Buffer.alloc(0))]);
}
const image = '<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="240000" cy="240000"/><wp:docPr id="1" name="Campus mark"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="0" name="mark.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rPic"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="240000" cy="240000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
const text = s => '<w:r><w:t>'+s+'</w:t></w:r>';
const par = s => '<w:p>'+s+'</w:p>';
const field = (code, cache) => '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> '+code+' </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>'+text(cache)+'<w:r><w:fldChar w:fldCharType="end"/></w:r>';
const relationships = spec => '<Relationships xmlns="'+PKG+'">'+spec.map(([id,t,p])=>'<Relationship Id="'+id+'" Type="'+t+'" Target="'+p+'"/>').join('')+'</Relationships>';
const sdef = (front, c) => '<w:sectPr><w:headerReference w:type="default" r:id="rHeader"/><w:footerReference w:type="default" r:id="'+(front?'rFrontFooter':'rBodyFooter')+'"/><w:type w:val="nextPage"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="'+((!front&&c==='A03')?'960':'1740')+'" w:right="1320" w:bottom="1500" w:left="'+((front&&c==='A06')?'1960':'1540')+'" w:header="600" w:footer="640" w:gutter="0"/><w:pgNumType w:fmt="'+(front?'upperRoman':'decimal')+'" w:start="'+((!front&&c==='A06')?'4':'1')+'"/></w:sectPr>';
function make(c, isTemplate=false) {
  const school=c==='template';
  const title=school?text('{{title_')+text('en}}'):text('Archive methods');
  const bodyText='<w:p><w:pPr><w:pStyle w:val="BodyCampus"/></w:pPr><w:r>'+((c==='A05'||c==='A08')?'<w:rPr><w:sz w:val="38"/></w:rPr>':'')+'<w:t>Anchor paragraph.</w:t></w:r></w:p>';
  const head='<w:document xmlns:w="'+W+'" xmlns:r="'+R+'"><w:body>';
  const main = head+par(title)+par(text('Author: ')+text(school?'{{author}}':'Anonymous'))+par(image)+
    par(c==='A04'?text('Contents'):field(c==='A08'?'REF Bibliography':'TOC \\o "1-2"','Contents'))+
    '<w:p><w:pPr>'+sdef(true,c)+'</w:pPr></w:p>'+
    '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>'+text('1  Materials')+'</w:p>'+bodyText+
    par(text(c==='A07'?'Updated ordinary paragraph.':'Unchanged ordinary paragraph.'))+sdef(false,c)+'</w:body></w:document>';
  const zoom = (c==='A01'||c==='A08')?'30':'26';
  const styles = '<w:styles xmlns:w="'+W+'"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="BodyCampus"><w:name w:val="Campus body"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="'+zoom+'"/></w:rPr></w:style></w:styles>';
  const core='<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:creator>'+((c==='A03')?'EXAMPLE-Q17':'Campus examples')+'</dc:creator></cp:coreProperties>';
  const overrides=[['word/document.xml',CT+(isTemplate?'template':'document')+'.main+xml'],['word/styles.xml',CT+'styles+xml'],['word/header1.xml',CT+'header+xml'],['word/footer1.xml',CT+'footer+xml'],['word/footer2.xml',CT+'footer+xml'],['docProps/core.xml','application/vnd.openxmlformats-package.core-properties+xml']];
  return zip({
    '[Content_Types].xml':'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>'+overrides.map(([p,t])=>'<Override PartName="/'+p+'" ContentType="'+t+'"/>').join('')+'</Types>',
    '_rels/.rels':relationships([['rDoc',R+'/officeDocument','word/document.xml'],['rProp',PKG+'/metadata/core-properties','docProps/core.xml']]),
    'docProps/core.xml':core,'word/document.xml':main,'word/styles.xml':styles,
    'word/_rels/document.xml.rels':relationships([['rStyles',R+'/styles','styles.xml'],['rHeader',R+'/header','header1.xml'],['rFrontFooter',R+'/footer','footer1.xml'],['rBodyFooter',R+'/footer','footer2.xml'],['rPic',R+'/image','media/mark.png']]),
    'word/header1.xml':'<w:hdr xmlns:w="'+W+'">'+par(text('Campus archive'))+'</w:hdr>',
    'word/footer1.xml':'<w:ftr xmlns:w="'+W+'">'+par(field('PAGE','I'))+'</w:ftr>',
    'word/footer2.xml':'<w:ftr xmlns:w="'+W+'">'+par(c==='A04'?text('1'):field('PAGE','1'))+(c==='A05'?par(text('EXAMPLE-')+text('Q17')):'')+'</w:ftr>',
    'word/media/mark.png':swatch(c==='A01'||c==='A06'?[222,140,33]:[35,119,180])
  });
}
mkdirSync(inputs,{recursive:true});
const hashes={};
for (const [name, id, isTemplate] of [
  ['school.docx','template',false], ['school.dotx','template',true],
  ...Object.keys(design.cases).map(id=>[id+'.docx',id,false])
]) {
  const data=make(id,isTemplate), file=path.join(inputs,name);
  if (existsSync(file)) assert.deepEqual(readFileSync(file),data,'Refusing to change an existing input: '+name);
  else writeFileSync(file,data);
  hashes[name]=createHash('sha256').update(data).digest('hex');
}
hashes['TASK.md']=createHash('sha256').update(readFileSync(path.join(inputs,'TASK.md'))).digest('hex');
const frozen=JSON.stringify(hashes,null,2)+'\n', manifest=path.join(here,'input-sha256.json');
if (existsSync(manifest)) assert.equal(readFileSync(manifest,'utf8'),frozen,'Frozen input manifest differs');
else writeFileSync(manifest,frozen);
console.log('Built '+Object.keys(hashes).length+' fixed inputs.');
