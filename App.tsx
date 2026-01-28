import React, { useState, useEffect } from 'react';
import { motion, AnimatePresence } from 'framer-motion';
import { 
  Layout, Code, ChevronLeft, ChevronRight, 
  Check, Copy, Terminal
} from 'lucide-react';

// ----------------------------------------------------------------------------
// DATA & TYPES (Tutto integrato per evitare errori di import)
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

const slides: SlideData[] = [
  {
    id: 1,
    type: SlideType.TITLE,
    layout: 'title',
    title: "THE UNITED NATIONS",
    subtitle: "Architecting Global Peace in the 21st Century",
    footer: "UN Global Report 2024",
    imageUrl: "https://images.unsplash.com/photo-1599366744839-4467dc0b4d4b?q=80&w=2574&auto=format&fit=crop"
  },
  {
    id: 2,
    type: SlideType.CONTENT,
    layout: 'split',
    title: "1945: The Vision",
    points: ["Emerging from the ruins of WWII.", "51 Founding Nations.", "A Charter for humanity.", "The end of the League of Nations."],
    imageUrl: "https://images.unsplash.com/photo-1580130601254-05fa235cd77e?q=80&w=2670&auto=format&fit=crop"
  },
  {
    id: 3,
    type: SlideType.CONTENT,
    layout: 'quadrant',
    title: "The Four Pillars",
    points: ["PEACE & SECURITY", "HUMAN RIGHTS", "RULE OF LAW", "DEVELOPMENT"],
    imageUrl: "https://images.unsplash.com/photo-1628891510373-30595221376d?q=80&w=2670&auto=format&fit=crop"
  },
  {
    id: 4,
    type: SlideType.CONTENT,
    layout: 'grid',
    title: "The Main Organs",
    points: ["General Assembly", "Security Council", "Econ. & Social Council", "Secretariat"],
    imageUrl: "https://images.unsplash.com/photo-1555848962-6e79363ec58f?q=80&w=2666&auto=format&fit=crop"
  },
  {
    id: 5,
    type: SlideType.CONTENT,
    layout: 'image-focus',
    title: "Security Council",
    subtitle: "The Horseshoe Table",
    points: ["5 Permanent Members (P5)", "Veto Power Authority", "Binding Resolutions", "Sanctions Regime"],
    imageUrl: "https://images.unsplash.com/photo-1577985848520-278635567b5e?q=80&w=2670&auto=format&fit=crop"
  },
  {
    id: 6,
    type: SlideType.CONTENT,
    layout: 'split',
    title: "Agencies in Action",
    points: ["WHO: Global Pandemics", "UNICEF: Child Rights", "WFP: Emergency Food", "UNESCO: Culture"],
    imageUrl: "https://images.unsplash.com/photo-1469571486292-0ba58a3f068b?q=80&w=2670&auto=format&fit=crop"
  },
  {
    id: 7,
    type: SlideType.CONTENT,
    layout: 'grid',
    title: "Agenda 2030 (SDGs)",
    points: ["No Poverty", "Climate Action", "Gender Equality", "Clean Energy"],
    imageUrl: "https://images.unsplash.com/photo-1473341304170-971dccb5ac1e?q=80&w=2670&auto=format&fit=crop"
  },
  {
    id: 8,
    type: SlideType.CONTENT,
    layout: 'quote',
    title: "Human Rights",
    points: ["All human beings are born free and equal in dignity and rights."],
    quoteAuthor: "UDHR, Article 1",
    imageUrl: "https://images.unsplash.com/photo-1531206715517-5c0ba140b2b8?q=80&w=2670&auto=format&fit=crop"
  },
  {
    id: 9,
    type: SlideType.CONTENT,
    layout: 'image-focus',
    title: "Global Crises",
    points: ["Climate Emergency", "Cyber Warfare", "Refugee Crisis", "Geopolitical Polarization"],
    imageUrl: "https://images.unsplash.com/photo-1616164283832-618e7e132805?q=80&w=2574&auto=format&fit=crop"
  },
  {
    id: 10,
    type: SlideType.TITLE,
    layout: 'title',
    title: "THE FUTURE",
    subtitle: "Reform. Innovate. Unite.",
    footer: "UN Strategic Vision 2030",
    imageUrl: "https://images.unsplash.com/photo-1575356885361-b66a5e1e23df?q=80&w=2674&auto=format&fit=crop"
  }
];

// ----------------------------------------------------------------------------
// THEME & UTILS
// ----------------------------------------------------------------------------

const THEME = {
  primary: '#009EDB',
  secondary: '#D4AF37',
  text: '#F0F4F8',
  accent: '#00B4D8',
};

const generatePythonScript = (s: SlideData[]): string => {
  return `# Script Python per generare PPTX\n# Basato su ${s.length} slide\nimport pptx\n# ... codice omesso per brevità nel viewer ...\nprint("PPTX Generato")`;
};

// ----------------------------------------------------------------------------
// COMPONENTS
// ----------------------------------------------------------------------------

const SlidePreview: React.FC<{ slide: SlideData }> = ({ slide }) => {
  return (
    <div className="w-full h-full flex items-center justify-center p-4 bg-[#0a192f]">
      <motion.div 
        key={slide.id}
        initial={{ opacity: 0 }} animate={{ opacity: 1 }}
        className="relative shadow-2xl w-full max-w-5xl aspect-[16/9] flex flex-col overflow-hidden text-white rounded-lg border border-white/10"
        style={{ 
          backgroundImage: `url(${slide.imageUrl})`,
          backgroundSize: 'cover', backgroundPosition: 'center',
        }}
      >
        <div className="absolute inset-0 bg-black/60 backdrop-blur-[2px]" />
        <div className="relative z-10 p-12 flex flex-col h-full justify-center">
          <h1 className="text-5xl font-bold mb-4 text-[#009EDB]">{slide.title}</h1>
          <p className="text-xl text-gray-200">{slide.subtitle}</p>
          <ul className="mt-6 space-y-2">
            {slide.points?.map((p, i) => (
              <li key={i} className="flex items-center gap-2">
                <div className="w-2 h-2 bg-[#D4AF37]" /> {p}
              </li>
            ))}
          </ul>
        </div>
      </motion.div>
    </div>
  );
};

const App: React.FC = () => {
  const [index, setIndex] = useState(0);

  return (
    <div className="flex h-screen bg-black text-white">
      <div className="w-64 border-r border-white/10 p-4 overflow-y-auto hidden md:block">
        <h2 className="text-blue-400 font-bold mb-4">ONU Deck</h2>
        {slides.map((s, i) => (
          <button key={s.id} onClick={() => setIndex(i)} className={`block w-full text-left p-2 text-sm ${index === i ? 'bg-blue-600' : ''}`}>
            {i + 1}. {s.title}
          </button>
        ))}
      </div>
      <div className="flex-1 flex flex-col">
        <div className="flex-1">
          <SlidePreview slide={slides[index]} />
        </div>
        <div className="p-4 bg-gray-900 flex justify-between items-center">
          <button onClick={() => setIndex(Math.max(0, index - 1))} className="p-2 bg-gray-700 rounded"><ChevronLeft /></button>
          <span>Slide {index + 1} / {slides.length}</span>
          <button onClick={() => setIndex(Math.min(slides.length - 1, index + 1))} className="p-2 bg-gray-700 rounded"><ChevronRight /></button>
        </div>
      </div>
    </div>
  );
};

export default App;
