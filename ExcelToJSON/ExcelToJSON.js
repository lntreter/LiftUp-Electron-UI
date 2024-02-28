const XLSX = require('xlsx');
const fs = require('fs');

function convertExcelToJson(inputFilePath, outputFilePath) {
  try {
    const workbook = XLSX.readFile(inputFilePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    
    // JSON'a çevirme işlemi başlar. İlk satırı atlamak için range: 1 eklenir.
    let jsonData = XLSX.utils.sheet_to_json(worksheet, { header:1, range: 2 });
    
    // JSON verilerini istenen yapıya dönüştür
    jsonData = jsonData.map((row) => {
      // Yeni bir obje oluştur
      let newRow = {};
      // Başlıkları manuel olarak eşleştir
      newRow[''] = {
        'Number': row[0],
      };
      newRow['Signal'] = {
        'Name': row[1],
        'TYPE': row[2],
        'CATEGORY': row[3],
        'CURRENT(Max)': row[4]
      };
      newRow['CABLE'] = {
        'TYPE': row[5],
        'AWG': row[6]
      };
      newRow['Source'] = {
        'ATA CHAPTER': row[7],
        'PIN NAME': row[8],
        'LOCATION': row[9],
        'LRU': row[10],
        'RD NUMBER': row[11],
        'Connector': row[12],
        'Pin No': row[13]
      };
      newRow['Destination'] = {
        'ATA CHAPTER': row[14],
        'PIN NAME': row[15],
        'LOCATION': row[16],
        'LRU': row[17],
        'RD NUMBER': row[18],
        'Connector': row[19],
        'Pin No': row[20]
      };
        
      // Eğer daha fazla sütun varsa, onları da burada ekleyebilirsiniz.

      return newRow;
    });

    // JSON verisini bir dosyaya yaz
    fs.writeFile(outputFilePath, JSON.stringify(jsonData, null, 2), (err) => {
      if (err) {
        console.error('JSON dosyası yazılırken bir hata oluştu:', err);
        return;
      }
      console.log(`${outputFilePath} üzerine JSON dosyası başarıyla yazıldı.`);
    });
  } catch (error) {
    console.error('Excel dosyası işlenirken bir hata oluştu:', error);
  }
}

module.exports = { convertExcelToJson };
