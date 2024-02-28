import { app, BrowserWindow, ipcMain } from 'electron'
import path from 'node:path'
import fs from 'node:fs'
import * as XLSX from 'xlsx'

// The built directory structure
//
// ├─┬─┬ dist
// │ │ └── index.html
// │ │
// │ ├─┬ dist-electron
// │ │ ├── main.js
// │ │ └── preload.js
// │
process.env.DIST = path.join(__dirname, '../dist')
process.env.VITE_PUBLIC = app.isPackaged ? process.env.DIST : path.join(process.env.DIST, '../public')


let win: BrowserWindow | null
// 🚧 Use ['ENV_NAME'] avoid vite:define plugin - Vite@2.x
const VITE_DEV_SERVER_URL = process.env['VITE_DEV_SERVER_URL']

function createWindow() {
  win = new BrowserWindow({
    icon: path.join(process.env.VITE_PUBLIC, 'electron-vite.svg'),
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true, // Bu, varsayılan olarak true'dur
    },
  })

  // Test active push message to Renderer-process.
  win.webContents.on('did-finish-load', () => {
    win?.webContents.send('main-process-message', (new Date).toLocaleString())
  })

  win.webContents.openDevTools();

  if (VITE_DEV_SERVER_URL) {
    win.loadURL(VITE_DEV_SERVER_URL)
  } else {
    // win.loadFile('dist/index.html')
    win.loadFile(path.join(process.env.DIST, 'index.html'))
  }
}

// Quit when all windows are closed, except on macOS. There, it's common
// for applications and their menu bar to stay active until the user quits
// explicitly with Cmd + Q.
app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit()
    win = null
  }
})

app.on('activate', () => {
  // On OS X it's common to re-create a window in the app when the
  // dock icon is clicked and there are no other windows open.
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow()
  }
})

// ipcMain.handle('read-excel', async (event, arg) => {
//   try {
//     if(fs.existsSync(arg)) {
//       console.log('arg: ', arg);
//       const fileBuffer = fs.readFileSync(arg);
//       console.log('fileBuffer: ', fileBuffer);
//       const data = XLSX.read(fileBuffer, {type: 'buffer'});
//       return data;
//     } else {
//       throw new Error(`File does not exist: ${arg}`);
//     }
//   } catch (err) {
//     console.error(err);
//     throw err;
//   }
// });


ipcMain.handle('read-excel', async (_event, arg) => {
  try {
    if(fs.existsSync(arg)) {
      console.log('arg: ', arg);
      const fileBuffer = fs.readFileSync(arg);
      console.log('fileBuffer: ', fileBuffer);
      const data = XLSX.read(fileBuffer, {type: 'buffer'});
      const workbook = data
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

      const outputFilePath = "./src/components/Output.json"


      fs.writeFile(outputFilePath, JSON.stringify(jsonData, null, 2), (err:any) => {
        if (err) {
          console.error('JSON dosyası yazılırken bir hata oluştu:', err);
          return;
        }
        console.log(`${outputFilePath} üzerine JSON dosyası başarıyla yazıldı.`);
      });

      return data;
    } else {
      throw new Error(`File does not exist: ${arg}`);
    }
  } catch (err) {
    console.error(err);
    throw err;
  }
});


app.whenReady().then(createWindow)
