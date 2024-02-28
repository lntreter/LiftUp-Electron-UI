import { useState, useEffect } from 'react'
import TLogo from './assets/logo.png'
import LULogo from './assets/liftup2.png'
import * as XLSX from 'xlsx'
import './App.css'


function App() {
  const [fileExtension, setFileExtension] = useState('')
  const [fileDirectory, setFileDirectory] = useState('')

  useEffect(() => {
    console.log(fileDirectory);
  }
  , [fileDirectory])
  
  const Convert = async (inputFilePath:string , outputFilePath:string) => {
    try {
      console.log('inputFilePath: ', inputFilePath)
      const fileData = await window.ipcRenderer.invoke('read-excel', inputFilePath);
      const workbook = XLSX.readFile(fileData);
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      
      // JSON'a çevirme işlemi başlar. İlk satırı atlamak için range: 1 eklenir.
      let jsonData = XLSX.utils.sheet_to_json(worksheet, { header:1, range: 2 });

      console.log('jsonData: ', jsonData);
      
      // JSON verilerini istenen yapıya dönüştür
      jsonData = jsonData.map((row:any) => {
        // Yeni bir obje oluştur
        let newRow:any = {};
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
      

      // fs.writeFile(outputFilePath, JSON.stringify(jsonData, null, 2), (err:any) => {
      //   if (err) {
      //     console.error('JSON dosyası yazılırken bir hata oluştu:', err);
      //     return;
      //   }
      //   console.log(`${outputFilePath} üzerine JSON dosyası başarıyla yazıldı.`);
      // });
      
    } catch (error) {
      console.error('Excel dosyası işlenirken bir hata oluştu:', error);
    }
  }

  return (
    <>
      <div>
        <a>
          <img src={LULogo} className="logo" alt="Vite logo" />
        </a>
        <a>
          <img src={TLogo} className="logo react" alt="React logo" />
        </a>
      </div>
      <h1>E3.Series Devre Oluşturma</h1>
      <div>
        <span className='span1'>Excel dosyasını seçin.</span>
        <span></span>
        <div className='dosya'>
          <input className='inputt' type="file" onChange={(e) => {
            const files = e.target.files;
            if (files && files.length > 0) {
              setFileExtension(files[0].name.split('.').pop()!);
              setFileDirectory(files[0].path.replace(/\\/g, '\\\\'));
            }
          }} />
          <>
          <div className='devreB'>
            {fileExtension == "xlsx" ? <button onClick={() => {

              Convert(fileDirectory, './components/Devreler.json')
              console.log(fileDirectory)

            }} className='devreB'> Devreyi çizdir! </button> : 
              <span className='span2'>Lütfen bir Excel dosyası seçin.</span>
            } 
            </div>
          </>
        </div>
      </div>
    </>
  )
}

export default App
