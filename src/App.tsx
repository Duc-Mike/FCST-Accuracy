import { useState, useRef } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { 
  Upload, 
  FileSpreadsheet, 
  CheckCircle2, 
  AlertCircle, 
  Download, 
  ArrowRight,
  Database,
  Layers,
  FileText,
  Loader2
} from 'lucide-react';
import { processExcelFiles } from './utils/excelProcessor';

type Step = 'initial' | 'upload' | 'processing' | 'complete';

export default function App() {
  const [hasBaseData, setHasBaseData] = useState<boolean | null>(null);
  const [step, setStep] = useState<Step>('initial');
  const [files, setFiles] = useState<Record<string, File>>({});
  const [error, setError] = useState<string | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);

  const fileInputRefs = useRef<Record<string, HTMLInputElement | null>>({});

  const handleResponse = (response: boolean) => {
    setHasBaseData(response);
    setStep('upload');
  };

  const handleFileChange = (key: string, file: File | null) => {
    if (file) {
      setFiles(prev => ({ ...prev, [key]: file }));
    }
  };

  const requiredFiles = hasBaseData === true
    ? ['FCST', 'SO', 'SHIP', 'Base Data']
    : ['M FCST', 'M+1 FCST', 'M+2 FCST', 'M+3 FCST', 'M+4 FCST', 'SO', 'SHIP'];

  const isAnyFileUploaded = Object.keys(files).length > 0;

  const handleProcess = async () => {
    setIsProcessing(true);
    setError(null);
    try {
      const buffer = await processExcelFiles(files, hasBaseData!);
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      setDownloadUrl(url);
      setStep('complete');
    } catch (err: any) {
      setError(err.message || 'An error occurred during processing');
    } finally {
      setIsProcessing(false);
    }
  };

  const reset = () => {
    setHasBaseData(null);
    setStep('initial');
    setFiles({});
    setError(null);
    setDownloadUrl(null);
  };

  return (
    <div className="min-h-screen bg-[#F8FAFC] text-[#1E293B] font-sans selection:bg-indigo-100">
      {/* Decorative Background Elements */}
      <div className="fixed inset-0 overflow-hidden pointer-events-none">
        <div className="absolute -top-[10%] -left-[10%] w-[40%] h-[40%] bg-indigo-50 rounded-full blur-3xl opacity-50" />
        <div className="absolute top-[60%] -right-[10%] w-[30%] h-[30%] bg-emerald-50 rounded-full blur-3xl opacity-50" />
      </div>

      <main className="relative z-10 max-w-4xl mx-auto px-6 py-12 md:py-20">
        <header className="text-center mb-12">
          <motion.div 
            initial={{ opacity: 0, y: -20 }}
            animate={{ opacity: 1, y: 0 }}
            className="inline-flex items-center justify-center p-3 bg-white rounded-2xl shadow-sm border border-slate-100 mb-6"
          >
            <Layers className="w-8 h-8 text-indigo-600" />
          </motion.div>
          <motion.h1 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            transition={{ delay: 0.1 }}
            className="text-4xl md:text-5xl font-bold tracking-tight text-slate-900 mb-4"
          >
            FCST SO Accuracy
          </motion.h1>
          <motion.p 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            transition={{ delay: 0.2 }}
            className="text-lg text-slate-500 max-w-xl mx-auto"
          >
            Merge and transform your Excel data with precision and ease.
          </motion.p>
        </header>

        <div className="bg-white rounded-3xl shadow-xl shadow-slate-200/50 border border-slate-100 overflow-hidden">
          <AnimatePresence mode="wait">
            {step === 'initial' && (
              <motion.div
                key="initial"
                initial={{ opacity: 0, scale: 0.95 }}
                animate={{ opacity: 1, scale: 1 }}
                exit={{ opacity: 0, scale: 1.05 }}
                className="p-12 text-center"
              >
                <Database className="w-16 h-16 text-indigo-500 mx-auto mb-6 opacity-20" />
                <h2 className="text-2xl font-semibold mb-8">You have base data?</h2>
                <div className="flex flex-col sm:flex-row gap-4 justify-center">
                  <button
                    onClick={() => handleResponse(true)}
                    className="group relative px-8 py-4 bg-indigo-600 text-white rounded-2xl font-semibold transition-all hover:bg-indigo-700 hover:shadow-lg hover:shadow-indigo-200 active:scale-95"
                  >
                    <span className="flex items-center justify-center gap-2">
                      Yes, I have
                      <ArrowRight className="w-4 h-4 transition-transform group-hover:translate-x-1" />
                    </span>
                  </button>
                  <button
                    onClick={() => handleResponse(false)}
                    className="px-8 py-4 bg-slate-100 text-slate-700 rounded-2xl font-semibold transition-all hover:bg-slate-200 active:scale-95"
                  >
                    No, I don't
                  </button>
                </div>
              </motion.div>
            )}

            {step === 'upload' && (
              <motion.div
                key="upload"
                initial={{ opacity: 0, x: 20 }}
                animate={{ opacity: 1, x: 0 }}
                exit={{ opacity: 0, x: -20 }}
                className="p-8 md:p-12"
              >
                <div className="flex items-center justify-between mb-8">
                  <h2 className="text-xl font-bold flex items-center gap-2">
                    <Upload className="w-5 h-5 text-indigo-500" />
                    Upload Files
                  </h2>
                  <button 
                    onClick={reset}
                    className="text-sm text-slate-400 hover:text-slate-600 transition-colors"
                  >
                    Back to start
                  </button>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-10">
                  {requiredFiles.map((key) => (
                    <div 
                      key={key}
                      onClick={() => fileInputRefs.current[key]?.click()}
                      className={`relative group cursor-pointer p-5 rounded-2xl border-2 border-dashed transition-all ${
                        files[key] 
                          ? 'border-emerald-200 bg-emerald-50/30' 
                          : 'border-slate-200 hover:border-indigo-300 hover:bg-indigo-50/30'
                      }`}
                    >
                      <input
                        type="file"
                        accept=".xlsx, .xls"
                        className="hidden"
                        ref={el => fileInputRefs.current[key] = el}
                        onChange={(e) => handleFileChange(key, e.target.files?.[0] || null)}
                      />
                      <div className="flex items-center gap-4">
                        <div className={`p-3 rounded-xl ${files[key] ? 'bg-emerald-100 text-emerald-600' : 'bg-slate-100 text-slate-400 group-hover:bg-indigo-100 group-hover:text-indigo-600'}`}>
                          {files[key] ? <CheckCircle2 className="w-5 h-5" /> : <FileSpreadsheet className="w-5 h-5" />}
                        </div>
                        <div className="flex-1 min-w-0">
                          <p className="text-sm font-bold truncate">{key}</p>
                          <p className="text-xs text-slate-400 truncate">
                            {files[key] ? files[key].name : 'Click to select file'}
                          </p>
                        </div>
                      </div>
                    </div>
                  ))}
                </div>

                {error && (
                  <motion.div 
                    initial={{ opacity: 0, y: 10 }}
                    animate={{ opacity: 1, y: 0 }}
                    className="mb-8 p-4 bg-red-50 border border-red-100 rounded-2xl flex items-center gap-3 text-red-600 text-sm"
                  >
                    <AlertCircle className="w-5 h-5 shrink-0" />
                    {error}
                  </motion.div>
                )}

                <button
                  disabled={!isAnyFileUploaded || isProcessing}
                  onClick={handleProcess}
                  className={`w-full py-5 rounded-2xl font-bold text-lg transition-all flex items-center justify-center gap-3 ${
                    isAnyFileUploaded && !isProcessing
                      ? 'bg-indigo-600 text-white hover:bg-indigo-700 shadow-lg shadow-indigo-200 active:scale-[0.98]'
                      : 'bg-slate-100 text-slate-400 cursor-not-allowed'
                  }`}
                >
                  {isProcessing ? (
                    <>
                      <Loader2 className="w-6 h-6 animate-spin" />
                      Processing Data...
                    </>
                  ) : (
                    <>
                      <Layers className="w-6 h-6" />
                      Generate FCST SO Accuracy
                    </>
                  )}
                </button>
              </motion.div>
            )}

            {step === 'complete' && (
              <motion.div
                key="complete"
                initial={{ opacity: 0, scale: 0.95 }}
                animate={{ opacity: 1, scale: 1 }}
                className="p-16 text-center"
              >
                <div className="w-24 h-24 bg-emerald-100 text-emerald-600 rounded-full flex items-center justify-center mx-auto mb-8">
                  <CheckCircle2 className="w-12 h-12" />
                </div>
                <h2 className="text-3xl font-bold mb-4">Merge Successful!</h2>
                <p className="text-slate-500 mb-10">Your FCST SO Accuracy file is ready for download.</p>
                
                <div className="flex flex-col sm:flex-row gap-4 justify-center">
                  <a
                    href={downloadUrl!}
                    download="FCST SO Accuracy.xlsx"
                    className="flex items-center justify-center gap-2 px-10 py-5 bg-emerald-600 text-white rounded-2xl font-bold text-lg hover:bg-emerald-700 transition-all shadow-lg shadow-emerald-200 active:scale-95"
                  >
                    <Download className="w-6 h-6" />
                    Download Excel
                  </a>
                  <button
                    onClick={reset}
                    className="px-10 py-5 bg-slate-100 text-slate-700 rounded-2xl font-bold text-lg hover:bg-slate-200 transition-all active:scale-95"
                  >
                    Start New Merge
                  </button>
                </div>
              </motion.div>
            )}
          </AnimatePresence>
        </div>

        <footer className="mt-12 text-center text-slate-400 text-sm flex items-center justify-center gap-4">
          <div className="flex items-center gap-1.5">
            <FileText className="w-4 h-4" />
            <span>Output: FCST SO Accuracy.xlsx</span>
          </div>
          <div className="w-1 h-1 bg-slate-200 rounded-full" />
          <span>v1.0.0</span>
        </footer>
      </main>
    </div>
  );
}
