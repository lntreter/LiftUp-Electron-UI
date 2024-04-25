import { useState, useEffect } from 'react'
import TLogo from './assets/logo.png'
import LULogo from './assets/liftup2.png'
import './App.css'
import { ToastContainer, toast } from 'react-toastify';
import 'react-toastify/dist/ReactToastify.css';


function App() {
  const [fileExtension, setFileExtension] = useState('')
  const [fileDirectory, setFileDirectory] = useState('')

  useEffect(() => {
    console.log(fileDirectory);
  }
  , [fileDirectory])

  const Convert = async (inputFilePath:string) => {
    try {
      console.log('inputFilePath: ', inputFilePath)
      await window.ipcRenderer.invoke('read-excel', inputFilePath);

      console.log('XML dosyası oluşturuldu.')
      toast("XML dosyası oluşturuldu.")
      
    } catch (error) {
      console.error('Excel dosyası işlenirken bir hata oluştu:', error);
    }
  }

  return (
    <>
      <ToastContainer />
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

              Convert(fileDirectory)
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