import { useState } from 'react';
import { saveAs } from 'file-saver';
import { processAndGenerateExcel } from './utils/calculator';

export default function App() {
  const [rawBuffer, setRawBuffer] = useState(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [fileName, setFileName] = useState('');
  const [outputFileName, setOutputFileName] = useState('');
  const [isDragging, setIsDragging] = useState(false);
  const [errorMsg, setErrorMsg] = useState(null);

  const handleDownload = async () => {
    if (!rawBuffer) return;
    setIsProcessing(true);
    setErrorMsg(null);

    try {
      // 1. Local Image Processing
      let imageBuffer = null;
      try {
        const response = await fetch(import.meta.env.BASE_URL + 'logoatws.png');
        if (response.ok) {
           imageBuffer = await response.arrayBuffer();
        }
      } catch (err) {
         console.warn("Failed to load local logoatws.png", err);
      }

      const outputBuffer = await processAndGenerateExcel(rawBuffer, imageBuffer);
      
      const blob = new Blob([outputBuffer], { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      });

      let finalName = 'DC_Schedule_Program.xlsx';
      if (outputFileName.trim() !== '') {
          finalName = `${outputFileName.trim()}.xlsx`;
      }

      saveAs(blob, finalName);
    } catch (e) {
      console.error(e);
      setErrorMsg("Error generating download: " + e.message);
    } finally {
      setIsProcessing(false);
    }
  };

  const processFile = (file) => {
    if (!file) return;

    // Validation
    const validExtensions = ['xlsx', 'xls'];
    const fileExt = file.name.split('.').pop().toLowerCase();
    
    if (!validExtensions.includes(fileExt)) {
      setErrorMsg('Pastikan file Excel (.xlsx atau .xls) yang Anda Kirim!');
      return;
    }

    if (file.size > 10 * 1024 * 1024) {
      setErrorMsg('Ukuran file maksimal 10MB.');
      return;
    }

    setFileName(file.name);
    setIsProcessing(true);
    setErrorMsg(null);

    try {
      const reader = new FileReader();

      reader.onload = (e) => {
        setRawBuffer(e.target.result);
        setIsProcessing(false);
      };

      reader.onerror = (error) => {
        console.error("Error reading file:", error);
        setErrorMsg('Gagal membaca file.');
        setIsProcessing(false);
      };

      reader.readAsArrayBuffer(file);
    } catch (error) {
      console.error("Failed to parse Excel file:", error);
      setErrorMsg('Gagal memproses file Excel.');
      setIsProcessing(false);
    }
  };

  const handleDragOver = (e) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = (e) => {
    e.preventDefault();
    setIsDragging(false);
  };

  const handleDrop = (e) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files[0];
    if (file) {
      processFile(file);
    }
  };

  return (
    <div className="bg-gradient-to-tr from-amber-400 to-slate-50 to-50% relative w-full overflow-y-auto min-h-screen flex flex-col items-center justify-center font-sans">
      
      {/* Toast Error Notification */}
      {errorMsg && (
        <div className="fixed top-6 right-auto md:top-6 md:left-5 md:bottom-auto bg-slate-50/10 backdrop-blur-md text-red-600 px-5 py-4 rounded-xl shadow-2xl z-50 flex items-center space-x-3 transition-all duration-300">
          <i className="fa-solid fa-circle-exclamation text-xl"></i>
          <span className="font-medium text-sm md:text-base">{errorMsg}</span>
          <button onClick={() => setErrorMsg(null)} className="text-red-600 hover:text-red-400 ml-4 focus:outline-none transition">
            <i className="fa-solid fa-times text-lg mt-1"></i>
          </button>
        </div>
      )}

      {/* Card Container */}
      <div className="bg-transparent backdrop-blur-none border-transparent shadow-none md:bg-white/10 md:backdrop-blur-xl md:border-white/40 md:shadow-xl p-10 rounded-3xl max-w-lg w-full transform transition duration-500 hover:scale-[1.01]">
        
        {/* Header */}
        <div className="text-center mb-8">
          <div className="w-44 h-auto flex items-center justify-center mx-auto mb-8">
            <img src={`${import.meta.env.BASE_URL}logoatws.png`} alt="ATW Solar"/>
          </div>
          <h1 className="text-3xl font-extrabold text-slate-800 tracking-tight">DC Schedule</h1>
          <p className="text-slate-500 mt-2 font-medium">Upload Excel, Dapatkan Optimasi Kabel</p>
        </div>

        {/* Form */}
        <form className="space-y-6" onSubmit={(e) => e.preventDefault()}>
          
          {/* Upload Area */}
          <div className="relative group">
            <label 
              htmlFor="dropzone-file" 
              className={`flex flex-col items-center justify-center w-full h-52 border-2 border-dashed rounded-2xl cursor-pointer transition-all duration-300 ${
                isDragging ? 'bg-sky-50 border-sky-600 scale-[1.02] shadow-md' : 'bg-slate-50/80 border-slate-300 hover:border-sky-800 hover:bg-slate-100'
              }`}
              onDragOver={handleDragOver}
              onDragLeave={handleDragLeave}
              onDrop={handleDrop}
            >
              <div className="flex flex-col items-center justify-center pt-5 pb-6 pointer-events-none">
                <i className={`fa-solid fa-cloud-arrow-up text-5xl mb-4 transition-colors duration-300 ${isDragging ? 'text-sky-600' : 'text-slate-400 group-hover:text-sky-800'}`}></i>
                <p className={`mb-2 text-sm transition-colors duration-300 ${isDragging ? 'text-sky-800 font-bold' : 'text-slate-600'}`}>
                  {isDragging ? 'Lepaskan file di sini' : <><span className="font-semibold">Klik untuk upload</span> atau drag & drop</>}
                </p>
                <p className="text-xs text-slate-400 font-medium">Format XLSX / XLS (Maks: 10MB)</p>
              </div>
              <input 
                id="dropzone-file" 
                type="file" 
                className="hidden" 
                accept=".xlsx, .xls" 
                onChange={(e) => processFile(e.target.files[0])}
              />
            </label>
          </div>

          {/* File Name Preview (Dynamic) */}
          {fileName && (
            <div className="flex bg-slate-50/40 px-5 py-1 rounded-xl items-center justify-between border border-slate-200 shadow-sm animate-fade-in">
              <div className="flex items-center space-x-4 overflow-hidden">
                <i className="fa-solid fa-file-excel text-green-700 text-sm"></i>
                <span className="text-sm text-slate-700 font-semibold truncate max-w-xs">{fileName}</span>
              </div>
              <button 
                type="button" 
                className="text-slate-400 hover:text-red-600 p-2 rounded-lg transition focus:outline-none"
                onClick={() => {
                   setFileName('');
                   setRawBuffer(null);
                }}
                title="Hapus file"
              >
                <i className="fa-solid fa-trash-can text-sm"></i>
              </button>
            </div>
          )}

          {/* Output Filename (Custom) */}
          <div className="space-y-2">
            <label htmlFor="output_name" className="ps-1 text-sm font-semibold text-slate-600">Nama file output (opsional)</label>
            <div className="relative">
              <div className="absolute inset-y-0 left-0 pl-4 flex items-center pointer-events-none">
                <i className="fa-solid fa-pen text-slate-400 text-sm"></i>
              </div>
              <input 
                id="output_name" 
                type="text" 
                maxLength={100} 
                placeholder="Contoh: DC Schedule - Project Name" 
                className="w-full rounded-xl bg-slate-50/50 border border-slate-200 pl-10 pr-4 py-3 text-sm font-medium text-slate-700 placeholder-slate-400 shadow-sm focus:border-sky-800 focus:outline-none focus:ring-2 focus:ring-sky-600/20 transition-all" 
                value={outputFileName}
                onChange={(e) => setOutputFileName(e.target.value)}
              />
            </div>
          </div>

          {/* Submit / Download Button */}
          {rawBuffer && !isProcessing ? (
             <button 
              type="button" 
              onClick={handleDownload}
              className="w-full text-slate-50 font-bold rounded-xl text-lg px-5 py-4 mt-2 text-center shadow-lg transform transition-all duration-200 bg-gradient-to-r from-sky-900 to-sky-700 hover:from-sky-700 hover:to-sky-700 active:scale-[0.98] focus:outline-none focus:ring-4 focus:ring-green-500/30 flex items-center justify-center"
            >
              <i className="fa-solid fa-download mr-3 text-xl"></i> Dapatkan Hasil Optimasi
            </button>
          ) : (
            <button 
              type="button" 
              disabled={true}
              className={`w-full text-slate-50 font-bold rounded-xl text-lg px-5 py-4 mt-2 text-center shadow-sm transition-all duration-300 flex items-center justify-center ${
                isProcessing ? 'bg-sky-700 animate-pulse' : 'bg-slate-400/60 cursor-not-allowed'
              }`}
            >
              {isProcessing ? (
                <span><i className="fas fa-spinner fa-spin mr-3"></i> Sedang Memproses...</span>
              ) : (
                <span>Upload file terlebih dahulu</span>
              )}
            </button>
          )}

        </form>
      </div>

    </div>
  );
}
