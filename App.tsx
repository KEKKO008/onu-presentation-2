import React, { useState, useEffect } from 'react';
import { slides } from './data';
import SlidePreview from './components/SlidePreview';
import PythonCodeViewer from './components/PythonCodeViewer';
import { generatePythonScript } from './utils/generatePythonScript';
import { generateClientPPTX } from './utils/generateClientPPTX';
import { Layout, Code, ChevronLeft, ChevronRight, Presentation } from 'lucide-react';

const App: React.FC = () => {
  const [currentSlideIndex, setCurrentSlideIndex] = useState(0);
  const [viewMode, setViewMode] = useState<'preview' | 'code'>('preview');
  const [pythonCode, setPythonCode] = useState('');
  const [isGenerating, setIsGenerating] = useState(false);

  useEffect(() => {
    setPythonCode(generatePythonScript(slides));
  }, []);

  const nextSlide = () => {
    setCurrentSlideIndex((prev) => Math.min(prev + 1, slides.length - 1));
  };

  const prevSlide = () => {
    setCurrentSlideIndex((prev) => Math.max(prev - 1, 0));
  };

  const handleDownloadPPTX = async () => {
    try {
      setIsGenerating(true);
      await generateClientPPTX(slides);
    } catch (error) {
      console.error("Failed to generate PPTX:", error);
      alert("Failed to generate PPTX. See console for details.");
    } finally {
      setIsGenerating(false);
    }
  };

  const currentSlide = slides[currentSlideIndex];

  return (
    <div className="flex h-screen bg-gray-100 overflow-hidden font-sans">
      {/* Sidebar */}
      <aside className="w-64 bg-white border-r border-gray-200 flex flex-col z-10">
        <div className="p-6 border-b border-gray-100">
          <div className="flex items-center gap-2 mb-1">
            <div className="w-8 h-8 rounded bg-[#009EDB] flex items-center justify-center text-white font-bold text-sm">UN</div>
            <h1 className="font-bold text-gray-800 text-lg leading-tight">Presentation<br/>Builder</h1>
          </div>
          <p className="text-xs text-gray-500 mt-2">Generate Python-PPTX scripts effortlessly.</p>
        </div>
        
        <nav className="flex-1 overflow-y-auto p-4 space-y-2">
          <div className="text-xs font-bold text-gray-400 uppercase tracking-wider mb-3">Slides</div>
          {slides.map((slide, index) => (
            <button
              key={slide.id}
              onClick={() => {
                setCurrentSlideIndex(index);
                setViewMode('preview');
              }}
              className={`w-full text-left px-3 py-2 rounded-md text-sm transition-colors flex items-center gap-2 ${
                index === currentSlideIndex && viewMode === 'preview'
                  ? 'bg-blue-50 text-blue-600 font-medium'
                  : 'text-gray-600 hover:bg-gray-50'
              }`}
            >
              <span className="text-xs opacity-50 w-4">{index + 1}.</span>
              <span className="truncate">{slide.title}</span>
            </button>
          ))}
        </nav>

        <div className="p-4 border-t border-gray-200">
          <div className="text-xs font-bold text-gray-400 uppercase tracking-wider mb-3">Actions</div>
          <button
            onClick={() => setViewMode('code')}
            className={`w-full flex items-center gap-2 px-3 py-2 rounded-md text-sm transition-colors mb-2 ${
               viewMode === 'code' ? 'bg-gray-800 text-white' : 'bg-gray-100 text-gray-700 hover:bg-gray-200'
            }`}
          >
            <Code size={16} />
            <span>View Python Code</span>
          </button>
        </div>
      </aside>

      {/* Main Content */}
      <main className="flex-1 flex flex-col h-full relative">
        {/* Header Toolbar */}
        <header className="h-16 bg-white border-b border-gray-200 flex items-center justify-between px-6 shadow-sm z-10">
          <div className="flex items-center gap-4">
             <div className="flex bg-gray-100 rounded-lg p-1">
                <button 
                  onClick={() => setViewMode('preview')}
                  className={`px-4 py-1.5 rounded-md text-sm font-medium transition-all flex items-center gap-2 ${viewMode === 'preview' ? 'bg-white shadow-sm text-gray-900' : 'text-gray-500 hover:text-gray-700'}`}
                >
                  <Layout size={16} /> Preview
                </button>
                <button 
                  onClick={() => setViewMode('code')}
                  className={`px-4 py-1.5 rounded-md text-sm font-medium transition-all flex items-center gap-2 ${viewMode === 'code' ? 'bg-white shadow-sm text-purple-600' : 'text-gray-500 hover:text-gray-700'}`}
                >
                  <Code size={16} /> Script
                </button>
             </div>
          </div>

          <div className="flex items-center gap-3">
             <button
                onClick={handleDownloadPPTX}
                disabled={isGenerating}
                className="flex items-center gap-2 bg-[#D4AF37] hover:bg-[#C09F30] text-white px-4 py-2 rounded-md transition-colors text-sm font-bold shadow-md active:scale-95"
             >
                {isGenerating ? (
                  <span className="animate-pulse">Generating...</span>
                ) : (
                  <>
                    <Presentation size={16} />
                    Download PPTX
                  </>
                )}
             </button>

             {viewMode === 'preview' && (
               <>
                 <div className="h-6 w-px bg-gray-200 mx-2"></div>
                 <button 
                    onClick={prevSlide}
                    disabled={currentSlideIndex === 0}
                    className="p-2 rounded-full hover:bg-gray-100 disabled:opacity-30 transition-colors"
                 >
                   <ChevronLeft size={20} />
                 </button>
                 <button 
                    onClick={nextSlide}
                    disabled={currentSlideIndex === slides.length - 1}
                    className="p-2 rounded-full hover:bg-gray-100 disabled:opacity-30 transition-colors"
                 >
                   <ChevronRight size={20} />
                 </button>
               </>
             )}
          </div>
        </header>

        {/* Workspace */}
        <div className="flex-1 overflow-hidden relative">
          {viewMode === 'preview' ? (
             <SlidePreview slide={currentSlide} />
          ) : (
             <div className="h-full w-full p-6 bg-gray-900">
                <PythonCodeViewer code={pythonCode} />
             </div>
          )}
        </div>
      </main>
    </div>
  );
};

export default App;
