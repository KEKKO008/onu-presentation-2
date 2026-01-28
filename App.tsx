import React, { useState, useEffect } from 'react';
import { motion, AnimatePresence } from 'framer-motion';
import { 
  Layout, Code, ChevronLeft, ChevronRight, 
  Check, Copy, Terminal, Play
} from 'lucide-react';
import { slides } from './data';

// ----------------------------------------------------------------------------
// TYPE DEFINITIONS
// ----------------------------------------------------------------------------

enum SlideType {
  TITLE = 'title',
  CONTENT = 'content'
}

type LayoutType = 'standard' | 'title' | 'split' | 'grid' | 'image-focus' | 'quadrant' | 'quote';

interface SlideData {
  id: number;
  type: SlideType;
  layout: LayoutType;
  title: string;
  subtitle?: string;
  points?: string[];
  footer?: string;
  imageUrl?: string;
  quoteAuthor?: string;
}

interface Colors {
  primary: string;
  secondary: string;
  text: string;
  accent: string;
}

// ----------------------------------------------------------------------------
// CONSTANTS
// ----------------------------------------------------------------------------

const THEME: Colors = {
  primary: '#009EDB',    // Official UN Blue
  secondary: '#D4AF37',  // Metallic Gold
  text: '#F0F4F8',       // Ice White
  accent: '#00B4D8',     // Lighter Cyan
};

const FOOTER_TEXT = "UN Strategic Vision 2024";

// ----------------------------------------------------------------------------
// UTILITY: Python Script Generator
// ----------------------------------------------------------------------------

const generatePythonScript = (slidesList: SlideData[]): string => {
  // Convert HEX to RGB tuple for Python
  const hexToRgb = (hex: string) => {
    const r = parseInt(hex.slice(1, 3), 16);
    const g = parseInt(hex.slice(3, 5), 16);
    const b = parseInt(hex.slice(5, 7), 16);
    return `${r}, ${g}, ${b}`;
  };

  const primaryRgb = hexToRgb(THEME.primary);
  const secondaryRgb = hexToRgb(THEME.secondary);
  const accentRgb = hexToRgb(THEME.accent);
  
  let script = `import requests
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# --- 1. SETUP & CONFIGURATION ---
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# Brand Colors (Diplomatic Luxury Theme)
C_PRI = RGBColor(${primaryRgb})    # UN Blue
C_SEC = RGBColor(${secondaryRgb})  # Metallic Gold
C_ACCENT = RGBColor(${accentRgb})  # Cyan Glow
C_WHITE = RGBColor(240, 244, 248)  # Ice White
C_DARK = RGBColor(10, 25, 47)      # Deep Navy
C_TRANS_BLACK = RGBColor(0, 0, 0)

# --- 2. HELPER FUNCTIONS ---

def add_remote_image(slide, url, left, top, width, height):
    """Downloads an image from a URL and adds it to the slide."""
    try:
        response = requests.get(url, timeout=10)
        img_stream = BytesIO(response.content)
        pic = slide.shapes.add_picture(img_stream, left, top, width, height)
        return pic
    except Exception as e:
        print(f"Could not download image: {e}")
        # Fallback: Deep Navy Box
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
        shape.fill.solid()
        shape.fill.fore_color.rgb = C_DARK
        return shape

def add_overlay(slide, left, top, width, height, transparency=0.4):
    """Adds a dark overlay to make text pop over images."""
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = C_TRANS_BLACK
    shape.fill.transparency = transparency
    shape.line.fill.background()

def add_text(slide, text, left, top, width, height, font_size, bold=False, color=C_WHITE, align=PP_ALIGN.LEFT, font_name='Arial'):
    tb = slide.shapes.add_textbox(left, top, width, height)
    p = tb.text_frame.paragraphs[0]
    p.text = text
    p.font.name = font_name
    p.font.size = Pt(font_size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.alignment = align
    return tb

# --- 3. LAYOUT GENERATORS ---
# (Full layout logic implementation...)
`;

  // We are keeping the script concise for the UI, looping through slides to generate specific calls
  slidesList.forEach((slide, index) => {
    script += `\n# Slide ${index + 1}: ${slide.title}
slide = prs.slides.add_slide(prs.slide_layouts[6])
data = {
    'title': "${slide.title}",
    'subtitle': "${slide.subtitle || ''}",
    'imageUrl': "${slide.imageUrl || ''}",
    'quoteAuthor': "${slide.quoteAuthor || ''}",
    'points': [${slide.points ? slide.points.map((p: string) => `"${p}"`).join(', ') : ''}]
}
`;
    if (slide.layout === 'title') script += `create_cinematic_title(slide, data)\n`;
    else if (slide.layout === 'split') script += `create_split_layout(slide, data)\n`;
    else if (slide.layout === 'quadrant') script += `create_quadrant_layout(slide, data)\n`;
    else if (slide.layout === 'quote') script += `create_quote_layout(slide, data)\n`;
    else if (slide.layout === 'grid') script += `create_grid_layout(slide, data)\n`;
    else script += `create_image_focus(slide, data)\n`;
  });

  script += `\n# Save Presentation
prs.save('UN_Diplomatic_Presentation.pptx')
print("Generated UN_Diplomatic_Presentation.pptx")`;

  return script;
};

// ----------------------------------------------------------------------------
// COMPONENT: SlidePreview
// ----------------------------------------------------------------------------

interface SlidePreviewProps {
  slide: SlideData;
}

const SlidePreview: React.FC<SlidePreviewProps> = ({ slide }) => {
  const isTitle = slide.layout === 'title';
  const isSplit = slide.layout === 'split';
  const isGrid = slide.layout === 'grid';
  const isQuadrant = slide.layout === 'quadrant';
  const isQuote = slide.layout === 'quote';
  const isFocus = slide.layout === 'image-focus' || slide.layout === 'standard';

  return (
    <div className="w-full h-full flex items-center justify-center p-8 bg-[#0a192f] overflow-hidden">
      <AnimatePresence mode="wait">
        <motion.div 
          key={slide.id}
          initial={{ opacity: 0, scale: 0.98, filter: 'blur(4px)' }}
          animate={{ opacity: 1, scale: 1, filter: 'blur(0px)' }}
          exit={{ opacity: 0, scale: 1.02, filter: 'blur(2px)' }}
          transition={{ duration: 0.6, ease: [0.22, 1, 0.36, 1] }}
          className="relative shadow-2xl w-full max-w-6xl aspect-[16/9] flex flex-col overflow-hidden text-white rounded-lg border border-white/5"
          style={{ 
            fontFamily: "'Inter', sans-serif",
            backgroundImage: `url(${slide.imageUrl})`,
            backgroundSize: 'cover',
            backgroundPosition: 'center',
          }}
        >
          {/* Cinematic Overlay */}
          <div 
            className="absolute inset-0 z-0 pointer-events-none" 
            style={{ background: 'linear-gradient(135deg, rgba(10,25,47,0.85) 0%, rgba(0,30,60,0.5) 100%)' }}
          />

          {/* --- TITLE LAYOUT --- */}
          {isTitle && (
            <div className="relative z-10 flex flex-col items-center justify-center h-full text-center space-y-8 p-12">
              <motion.div 
                initial={{ scale: 0.8, opacity: 0 }}
                animate={{ scale: 1, opacity: 1 }}
                transition={{ delay: 0.2 }}
                className="px-8 py-2 border-2 border-[#D4AF37] text-[#D4AF37] uppercase tracking-[0.3em] text-sm font-bold bg-black/30 backdrop-blur-sm"
              >
                United Nations
              </motion.div>
              
              <motion.h1 
                initial={{ y: 20, opacity: 0 }}
                animate={{ y: 0, opacity: 1 }}
                transition={{ delay: 0.4, duration: 0.8 }}
                className="text-6xl md:text-8xl font-black tracking-tighter uppercase drop-shadow-2xl text-white"
              >
                {slide.title}
              </motion.h1>
              
              <motion.div 
                 initial={{ width: 0 }}
                 animate={{ width: "200px" }}
                 transition={{ delay: 0.6, duration: 0.8 }}
                 className="h-1 bg-[#009EDB] shadow-[0_0_15px_#009EDB]"
              />

              <motion.p 
                initial={{ y: 20, opacity: 0 }}
                animate={{ y: 0, opacity: 1 }}
                transition={{ delay: 0.8 }}
                className="text-xl md:text-2xl font-light tracking-wide text-gray-300 max-w-2xl"
              >
                {slide.subtitle}
              </motion.p>
            </div>
          )}

          {/* --- SPLIT LAYOUT --- */}
          {isSplit && (
            <div className="relative z-10 grid grid-cols-2 h-full">
              <div className="col-span-1 relative group">
                  <div className="absolute inset-0 bg-black/20 group-hover:bg-transparent transition-colors duration-500"></div>
              </div>
              <div className="col-span-1 bg-[#0a192f]/95 backdrop-blur-md h-full p-16 flex flex-col justify-center border-l border-[#D4AF37]/30 relative">
                <h2 className="text-5xl font-bold mb-10 text-white relative z-10 leading-none">
                  {slide.title}
                  <span className="block h-1 w-24 bg-[#D4AF37] mt-4"></span>
                </h2>
                <div className="space-y-8 relative z-10">
                  {slide.points?.map((point: string, idx: number) => (
                    <motion.div 
                      key={idx}
                      initial={{ x: 20, opacity: 0 }}
                      animate={{ x: 0, opacity: 1 }}
                      transition={{ delay: 0.3 + (0.1 * idx) }}
                      className="flex items-start gap-4"
                    >
                      <div className="w-1.5 h-1.5 mt-2.5 bg-[#009EDB] shadow-[0_0_8px_#009EDB] shrink-0 rotate-45" />
                      <p className="text-xl text-gray-300 font-light">{point}</p>
                    </motion.div>
                  ))}
                </div>
              </div>
            </div>
          )}

          {/* --- QUADRANT LAYOUT --- */}
          {isQuadrant && (
             <div className="relative z-10 w-full h-full p-12 flex flex-col">
                <div className="text-center mb-10">
                    <h2 className="text-4xl font-bold uppercase tracking-[0.2em] text-white drop-shadow-lg">{slide.title}</h2>
                </div>
                <div className="grid grid-cols-2 grid-rows-2 flex-1 gap-6">
                    {slide.points?.map((point: string, idx: number) => (
                        <motion.div 
                          key={idx}
                          initial={{ opacity: 0, y: 20 }}
                          animate={{ opacity: 1, y: 0 }}
                          transition={{ delay: 0.2 + (idx * 0.1) }}
                          className="bg-white/5 backdrop-blur-md border border-white/10 flex items-center justify-center relative overflow-hidden group hover:bg-[#009EDB]/20 transition-all duration-500 cursor-default"
                        >
                            <div className="absolute top-4 left-4 text-[#D4AF37] font-serif text-4xl opacity-50 group-hover:opacity-100 transition-opacity">
                              0{idx+1}
                            </div>
                            <h3 className="text-2xl font-bold tracking-widest uppercase z-10 text-center px-4">{point}</h3>
                            <div className="absolute bottom-0 right-0 w-8 h-8 border-b-2 border-r-2 border-[#D4AF37] opacity-0 group-hover:opacity-100 transition-opacity duration-300"></div>
                        </motion.div>
                    ))}
                </div>
             </div>
          )}

          {/* --- QUOTE LAYOUT --- */}
          {isQuote && (
              <div className="relative z-10 w-full h-full flex flex-col items-center justify-center p-24 text-center">
                  <div className="absolute inset-0 bg-black/70 backdrop-blur-sm z-0"></div>
                  <div className="relative z-10">
                      <div className="text-[#D4AF37] text-7xl font-serif mb-4 opacity-80">“</div>
                      <motion.h2 
                          initial={{ opacity: 0, scale: 0.95 }}
                          animate={{ opacity: 1, scale: 1 }}
                          transition={{ duration: 0.8 }}
                          className="text-5xl font-serif leading-tight text-white mb-10 italic"
                      >
                          {slide.points?.[0]}
                      </motion.h2>
                      <div className="flex flex-col items-center gap-2">
                          <div className="w-12 h-[2px] bg-[#009EDB]"></div>
                          <p className="text-lg uppercase tracking-widest text-[#D4AF37] font-bold mt-4">
                              {slide.quoteAuthor}
                          </p>
                      </div>
                  </div>
              </div>
          )}

          {/* --- GRID LAYOUT --- */}
          {isGrid && (
            <div className="relative z-10 flex flex-col h-full p-12">
               <div className="absolute inset-0 bg-gradient-to-r from-[#0a192f] via-[#0a192f]/90 to-transparent z-0"></div>
               
               <div className="relative z-10 mb-10 pl-6 border-l-4 border-[#D4AF37]">
                   <h2 className="text-5xl font-bold text-white">{slide.title}</h2>
               </div>
              
              <div className="relative z-10 grid grid-cols-2 md:grid-cols-4 gap-6 flex-1 items-center">
                {slide.points?.map((point: string, idx: number) => (
                  <motion.div 
                    key={idx}
                    initial={{ opacity: 0, y: 20 }}
                    animate={{ opacity: 1, y: 0 }}
                    transition={{ delay: 0.2 + (idx * 0.1) }}
                    whileHover={{ y: -5 }}
                    className="aspect-[4/5] bg-gradient-to-br from-white/10 to-transparent backdrop-blur-md border border-white/10 p-6 flex flex-col items-center justify-center text-center hover:border-[#009EDB] transition-colors duration-300"
                  >
                    <div className="text-4xl font-black text-white/10 mb-4">0{idx + 1}</div>
                    <p className="text-lg font-bold leading-tight text-white">{point}</p>
                  </motion.div>
                ))}
              </div>
            </div>
          )}

          {/* --- IMAGE FOCUS / STANDARD --- */}
          {isFocus && (
            <div className="relative z-10 p-16 h-full flex flex-col justify-center items-start">
               <div className="absolute inset-0 bg-gradient-to-r from-[#0a192f] via-[#0a192f]/60 to-transparent z-[-1]" />
               
               <div className="max-w-2xl bg-[#0a192f]/80 backdrop-blur-lg p-10 border border-white/5 shadow-2xl">
                   <div className="flex items-center gap-3 mb-6">
                       <span className="text-[#009EDB] uppercase tracking-[0.2em] text-xs font-bold">Key Objectives</span>
                   </div>
                  <h2 className="text-5xl font-bold mb-8 text-white leading-none">
                    {slide.title}
                  </h2>
                  <div className="space-y-6">
                    {slide.points?.map((point: string, idx: number) => (
                      <div key={idx} className="flex gap-4 items-center group">
                        <div className="w-8 h-[1px] bg-white/30 group-hover:bg-[#D4AF37] transition-colors"></div>
                        <p className="text-xl text-gray-200 font-light">{point}</p>
                      </div>
                    ))}
                  </div>
               </div>
            </div>
          )}

          {/* FOOTER */}
          {!isTitle && (
              <div className="absolute bottom-0 w-full px-12 py-6 flex justify-between items-center z-20 pointer-events-none">
                  <div className="flex items-center gap-3">
                       <div className="w-3 h-3 border border-[#D4AF37] rounded-full"></div>
                       <span className="text-[10px] text-white/40 tracking-[0.2em] uppercase">{FOOTER_TEXT}</span>
                  </div>
                  <span className="text-xs font-mono text-[#009EDB] opacity-70">
                      SLIDE {slide.id.toString().padStart(2, '0')}
                  </span>
              </div>
          )}
        </motion.div>
      </AnimatePresence>
    </div>
  );
};

// ----------------------------------------------------------------------------
// COMPONENT: PythonCodeViewer
// ----------------------------------------------------------------------------

interface PythonCodeViewerProps {
  code: string;
}

const PythonCodeViewer: React.FC<PythonCodeViewerProps> = ({ code }) => {
  const [copied, setCopied] = React.useState(false);

  const handleCopy = () => {
    navigator.clipboard.writeText(code);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const syntaxHighlight = (code: string): string => {
    return code
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/(from|import|def|return|if|else|for|in|class)/g, '<span class="text-purple-400">$1</span>')
      .replace(/('.*?')|(".*?")/g, '<span class="text-green-400">$1</span>')
      .replace(/\b(\d+)\b/g, '<span class="text-orange-400">$1</span>')
      .replace(/#.*/g, '<span class="text-gray-500">$1</span>')
      .replace(/\b(Presentation|Inches|Pt|RGBColor)\b/g, '<span class="text-yellow-400">$1</span>');
  };

  return (
    <div className="flex flex-col h-full bg-[#1e1e1e] text-gray-300 rounded-lg overflow-hidden shadow-2xl border border-gray-800">
      <div className="flex items-center justify-between px-4 py-3 bg-[#252526] border-b border-black">
        <div className="flex items-center gap-2 text-sm font-medium text-blue-400">
          <Terminal size={16} />
          <span>generate_presentation.py</span>
        </div>
        <button
          onClick={handleCopy}
          className="flex items-center gap-2 px-3 py-1.5 text-xs font-medium text-white bg-blue-600 hover:bg-blue-700 rounded transition-colors"
        >
          {copied ? <Check size={14} /> : <Copy size={14} />}
          {copied ? 'Copied!' : 'Copy'}
        </button>
      </div>
      <div className="flex-1 overflow-auto p-4 font-mono text-sm leading-6">
        <pre className="whitespace-pre">
          <code dangerouslySetInnerHTML={{ __html: syntaxHighlight(code) }} />
        </pre>
      </div>
    </div>
  );
};

// ----------------------------------------------------------------------------
// MAIN APP COMPONENT
// ----------------------------------------------------------------------------

const App: React.FC = () => {
  const [currentSlideIndex, setCurrentSlideIndex] = useState(0);
  const [viewMode, setViewMode] = useState<'preview' | 'code'>('preview');
  const [pythonCode, setPythonCode] = useState('');

  useEffect(() => {
    setPythonCode(generatePythonScript(slides));
  }, []);

  const nextSlide = () => {
    setCurrentSlideIndex((prev) => Math.min(prev + 1, slides.length - 1));
  };

  const prevSlide = () => {
    setCurrentSlideIndex((prev) => Math.max(prev - 1, 0));
  };

  const currentSlide = slides[currentSlideIndex];

  return (
    <div className="flex h-screen bg-gray-100 overflow-hidden font-sans">
      {/* Sidebar */}
      <aside className="w-64 bg-white border-r border-gray-200 flex flex-col z-20 shadow-xl">
        <div className="p-6 border-b border-gray-100">
          <div className="flex items-center gap-2 mb-1">
            <div className="w-8 h-8 rounded bg-[#009EDB] flex items-center justify-center text-white font-bold text-sm">UN</div>
            <h1 className="font-bold text-gray-800 text-lg leading-tight">Presentation<br/>Builder</h1>
          </div>
          <p className="text-xs text-gray-500 mt-2">Interactive Deck & Python Generator</p>
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
          <div className="text-xs font-bold text-gray-400 uppercase tracking-wider mb-3">Developer</div>
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
        <header className="h-16 bg-white border-b border-gray-200 flex items-center justify-between px-6 shadow-sm z-20">
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
             {viewMode === 'preview' && (
               <>
                 <span className="text-sm font-medium text-gray-500 mr-2">
                    Slide {currentSlideIndex + 1} of {slides.length}
                 </span>
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
        <div className="flex-1 overflow-hidden relative bg-gray-100">
          {viewMode === 'preview' ? (
             <SlidePreview slide={currentSlide} />
          ) : (
             <div className="h-full w-full p-6 bg-[#0a192f] flex items-center justify-center">
                <div className="w-full max-w-4xl h-full">
                    <PythonCodeViewer code={pythonCode} />
                </div>
             </div>
          )}
        </div>
      </main>
    </div>
  );
};

export default App;
