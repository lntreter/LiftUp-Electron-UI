// const XLSX = require('xlsx');
// const fs = require('fs');

// function convertExcelToJson(inputFilePath, outputFilePath) {
//   try {
//     const workbook = XLSX.readFile(inputFilePath);
//     const sheetName = workbook.SheetNames[0];
//     const worksheet = workbook.Sheets[sheetName];
    
//     // JSON'a çevirme işlemi başlar. İlk satırı atlamak için range: 1 eklenir.
//     let jsonData = XLSX.utils.sheet_to_json(worksheet, { header:1, range: 2 });
    
//     // JSON verilerini istenen yapıya dönüştür
//     jsonData = jsonData.map((row) => {
//       // Yeni bir obje oluştur
//       let newRow = {};
//       // Başlıkları manuel olarak eşleştir
//       newRow[''] = {
//         'Number': row[0],
//       };
//       newRow['Signal'] = {
//         'Name': row[1],
//         'TYPE': row[2],
//         'CATEGORY': row[3],
//         'CURRENT(Max)': row[4]
//       };
//       newRow['CABLE'] = {
//         'TYPE': row[5],
//         'AWG': row[6]
//       };
//       newRow['Source'] = {
//         'ATA CHAPTER': row[7],
//         'PIN NAME': row[8],
//         'LOCATION': row[9],
//         'LRU': row[10],
//         'RD NUMBER': row[11],
//         'Connector': row[12],
//         'Pin No': row[13]
//       };
//       newRow['Destination'] = {
//         'ATA CHAPTER': row[14],
//         'PIN NAME': row[15],
//         'LOCATION': row[16],
//         'LRU': row[17],
//         'RD NUMBER': row[18],
//         'Connector': row[19],
//         'Pin No': row[20]
//       };
        
//       // Eğer daha fazla sütun varsa, onları da burada ekleyebilirsiniz.

//       return newRow;
//     });

//     // JSON verisini bir dosyaya yaz
//     fs.writeFile(outputFilePath, JSON.stringify(jsonData, null, 2), (err) => {
//       if (err) {
//         console.error('JSON dosyası yazılırken bir hata oluştu:', err);
//         return;
//       }
//       console.log(`${outputFilePath} üzerine JSON dosyası başarıyla yazıldı.`);
//     });
//   } catch (error) {
//     console.error('Excel dosyası işlenirken bir hata oluştu:', error);
//   }
// }

// module.exports = { convertExcelToJson };


const XLSX = require('xlsx');
const fs = require('fs');

function convertExcelToXml(inputFilePath, outputFilePath) {
  try {
    const workbook = XLSX.readFile(inputFilePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    
    // İlk satırı atlamak için range: 1 eklenir ve JSON'a çevirme işlemi başlar.
    let jsonData = XLSX.utils.sheet_to_json(worksheet, { header:1, range: 2 });
    
    // Başladığımız XML string'i
    let xmlData = '<?xml version="1.0" encoding="UTF-8"?>\n<root>';
    
    // JSON verilerini XML'e dönüştür
    jsonData.forEach((row) => {
      xmlData += `
  <row>
    <Number>${row[0]}</Number>
    <Signal>
      <Name>${row[1]}</Name>
      <Type>${row[2]}</Type>
      <Category>${row[3]}</Category>
      <CurrentMax>${row[4]}</CurrentMax>
    </Signal>
    <Cable>
      <Type>${row[5]}</Type>
      <AWG>${row[6]}</AWG>
    </Cable>
    <Source>
      <ATAChapter>${row[7]}</ATAChapter>
      <PinName>${row[8]}</PinName>
      <Location>${row[9]}</Location>
      <LRU>${row[10]}</LRU>
      <RDNumber>${row[11]}</RDNumber>
      <Connector>${row[12]}</Connector>
      <PinNo>${row[13]}</PinNo>
    </Source>
    <Destination>
      <ATAChapter>${row[14]}</ATAChapter>
      <PinName>${row[15]}</PinName>
      <Location>${row[16]}</Location>
      <LRU>${row[17]}</LRU>
      <RDNumber>${row[18]}</RDNumber>
      <Connector>${row[19]}</Connector>
      <PinNo>${row[20]}</PinNo>
    </Destination>
  </row>`;
    });

    xmlData += '\n</root>';
    
    // XML verisini bir dosyaya yaz
    fs.writeFile(outputFilePath, xmlData, (err) => {
      if (err) {
        console.error('XML dosyası yazılırken bir hata oluştu:', err);
        return;
      }
      console.log(`${outputFilePath} üzerine XML dosyası başarıyla yazıldı.`);
    });
  } catch (error) {
    console.error('Excel dosyası işlenirken bir hata oluştu:', error);
  }
}

module.exports = { convertExcelToXml };
