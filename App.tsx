import React, { useState, useEffect } from 'react';
import { motion } from 'framer-motion';
import { 
  Layout, Code, ChevronLeft, ChevronRight, Presentation, 
  Check, Copy, Terminal 
} from 'lucide-react';
import PptxGenJS from 'pptxgenjs';
import { slides } from './data';

// ----------------------------------------------------------------------------
// TYPE DEFINITIONS (Inlined)
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
// CONSTANTS (Inlined)
// ----------------------------------------------------------------------------

const THEME: Colors = {
  primary: '#009EDB',    // Official UN Blue
  secondary: '#D4AF37',  // Metallic Gold
  text: '#F0F4F8',       // Ice White
  accent: '#00B4D8',     // Lighter Cyan
};

const FOOTER_TEXT = "UN Strategic Vision 2024";

// ----------------------------------------------------------------------------
// UTILITY: Python Script Generator (Inlined)
// ----------------------------------------------------------------------------

const generatePythonScript = (slidesList: any[]): string => {
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

def create_cinematic_title(slide, data):
    if data.get('imageUrl'):
        add_remote_image(slide, data['imageUrl'], 0, 0, Inches(13.333), Inches(7.5))
    
    add_overlay(slide, 0, 0, Inches(13.333), Inches(7.5), 0.3)
    
    # UN Badge Top
    badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.66), Inches(1), Inches(2), Inches(0.4))
    badge.fill.solid()
    badge.fill.fore_color.rgb = C_TRANS_BLACK
    badge.fill.transparency = 0.3
    badge.line.color.rgb = C_SEC
    badge.text_frame.text = "UNITED NATIONS"
    badge.text_frame.paragraphs[0].font.size = Pt(10)
    badge.text_frame.paragraphs[0].font.color.rgb = C_SEC
    
    # Big Title
    add_text(slide, data['title'].upper(), Inches(0.5), Inches(2.5), Inches(12.333), Inches(2), 72, True, C_WHITE, PP_ALIGN.CENTER, 'Arial Black')
    
    # Accent Bar
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(5.66), Inches(4.8), Inches(2), Inches(0.05))
    line.fill.solid()
    line.fill.fore_color.rgb = C_PRI
    line.line.fill.background()

    # Subtitle
    add_text(slide, data.get('subtitle', ''), Inches(1), Inches(5), Inches(11.333), Inches(1), 24, False, C_WHITE, PP_ALIGN.CENTER)

def create_split_layout(slide, data):
    # Left Half Image
    if data.get('imageUrl'):
        add_remote_image(slide, data['imageUrl'], 0, 0, Inches(6.6), Inches(7.5))
        add_overlay(slide, 0, 0, Inches(6.6), Inches(7.5), 0.1)
    
    # Right Half Dark
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.6), 0, Inches(6.8), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = C_DARK
    bg.line.fill.background()
    
    # Gold Line Vertical
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.6), 0, Inches(0.05), Inches(7.5))
    line.fill.solid()
    line.fill.fore_color.rgb = C_SEC
    
    add_text(slide, data['title'], Inches(7.1), Inches(1), Inches(5.5), Inches(1.5), 44, True, C_WHITE)
    
    # Gold Underline Title
    uline = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7.1), Inches(2.2), Inches(1.5), Inches(0.05))
    uline.fill.solid()
    uline.fill.fore_color.rgb = C_SEC
    
    y = 3.0
    for point in data.get('points', []):
        # Blue Diamond Bullet
        dot = slide.shapes.add_shape(MSO_SHAPE.DIAMOND, Inches(7.1), Inches(y + 0.1), Inches(0.15), Inches(0.15))
        dot.fill.solid()
        dot.fill.fore_color.rgb = C_PRI
        dot.line.fill.background()
        
        add_text(slide, point, Inches(7.4), Inches(y), Inches(5), Inches(0.8), 20, False, C_WHITE)
        y += 0.8

def create_quadrant_layout(slide, data):
    # Background
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = C_DARK
    
    add_text(slide, data['title'].upper(), Inches(0), Inches(0.5), Inches(13.333), Inches(1), 40, True, C_WHITE, PP_ALIGN.CENTER)
    
    points = data.get('points', [])
    positions = [
        (Inches(1), Inches(2)), (Inches(7), Inches(2)),
        (Inches(1), Inches(4.8)), (Inches(7), Inches(4.8))
    ]
    
    for i, point in enumerate(points):
        if i >= 4: break
        left, top = positions[i]
        
        # Box
        box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, Inches(5.3), Inches(2.2))
        box.fill.solid()
        box.fill.fore_color.rgb = C_TRANS_BLACK
        box.fill.transparency = 0.5 # Darker
        box.line.color.rgb = C_WHITE
        box.line.width = Pt(0.5)
        
        # Number
        add_text(slide, f"0{i+1}", left + Inches(0.2), top + Inches(0.2), Inches(1), Inches(0.5), 24, True, C_SEC, PP_ALIGN.LEFT, 'Times New Roman')
        
        # Text
        add_text(slide, point, left + Inches(0.2), top + Inches(0.8), Inches(4.9), Inches(1), 24, True, C_WHITE, PP_ALIGN.CENTER)

def create_quote_layout(slide, data):
    if data.get('imageUrl'):
        add_remote_image(slide, data['imageUrl'], 0, 0, Inches(13.333), Inches(7.5))
    add_overlay(slide, 0, 0, Inches(13.333), Inches(7.5), 0.7)
    
    text = data.get('points', [''])[0]
    author = data.get('quoteAuthor', '')
    
    # Big Quote Mark
    add_text(slide, '"', Inches(1), Inches(1.5), Inches(1), Inches(1), 100, True, C_SEC, PP_ALIGN.CENTER, 'Times New Roman')
    
    # Quote Text
    add_text(slide, text, Inches(2), Inches(2.5), Inches(9.333), Inches(3), 44, True, C_WHITE, PP_ALIGN.CENTER, 'Times New Roman')
    
    # Author
    add_text(slide, author.upper(), Inches(2), Inches(5.5), Inches(9.333), Inches(1), 16, True, C_SEC, PP_ALIGN.CENTER)

def create_grid_layout(slide, data):
    # Left Dark Panel
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(4), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = C_DARK
    
    # Right Image Panel
    if data.get('imageUrl'):
        add_remote_image(slide, data['imageUrl'], Inches(4), 0, Inches(9.333), Inches(7.5))
        add_overlay(slide, Inches(4), 0, Inches(9.333), Inches(7.5), 0.2)
    
    # Title on Left
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(3), Inches(3), Inches(3))
    p = tb.text_frame.paragraphs[0]
    p.text = data['title']
    p.font.bold = True
    p.font.size = Pt(40)
    p.font.color.rgb = C_WHITE
    p.alignment = PP_ALIGN.LEFT
    
    # Grid Content Overlay on Right
    points = data.get('points', [])
    positions = [
        (Inches(4.5), Inches(1)), (Inches(8.5), Inches(1)),
        (Inches(4.5), Inches(4)), (Inches(8.5), Inches(4))
    ]
    
    for i, point in enumerate(points):
        if i >= 4: break
        left, top = positions[i]
        
        # Glass Card
        card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, Inches(3.5), Inches(2.5))
        card.fill.solid()
        card.fill.fore_color.rgb = C_DARK
        card.fill.transparency = 0.2
        card.line.color.rgb = C_WHITE
        
        add_text(slide, f"0{i+1}", left + Inches(0.2), top + Inches(0.2), Inches(0.5), Inches(0.5), 24, True, C_WHITE)
        add_text(slide, point, left + Inches(0.2), top + Inches(1), Inches(3.1), Inches(1), 20, True, C_WHITE, PP_ALIGN.CENTER)

def create_image_focus(slide, data):
    if data.get('imageUrl'):
        add_remote_image(slide, data['imageUrl'], 0, 0, Inches(13.333), Inches(7.5))
    add_overlay(slide, 0, 0, Inches(13.333), Inches(7.5), 0.5)
    
    # Solid Card
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(1.5), Inches(5.5), Inches(5))
    box.fill.solid()
    box.fill.fore_color.rgb = C_DARK
    box.fill.transparency = 0.1
    box.line.color.rgb = C_SEC
    
    add_text(slide, "KEY OBJECTIVES", Inches(1.5), Inches(1.8), Inches(4), Inches(0.5), 10, True, C_PRI)
    add_text(slide, data['title'], Inches(1.5), Inches(2.2), Inches(4.5), Inches(1.2), 40, True, C_WHITE)
    
    y = 3.6
    for point in data.get('points', []):
        add_text(slide, point, Inches(1.5), y, Inches(4.5), Inches(0.6), 20, False, C_WHITE)
        y += 0.7

# --- 4. MAIN LOOP ---
def create_slides():
    blank_layout = prs.slide_layouts[6]
`;

  slidesList.forEach((slide, index) => {
    script += `
    # Slide ${index + 1}
    slide = prs.slides.add_slide(blank_layout)
    data = {
        'title': "${slide.title}",
        'subtitle': "${slide.subtitle || ''}",
        'imageUrl': "${slide.imageUrl || ''}",
        'quoteAuthor': "${slide.quoteAuthor || ''}",
        'points': [${slide.points ? slide.points.map((p: string) => `"${p}"`).join(', ') : ''}]
    }
    `;

    if (slide.layout === 'title') {
      script += `create_cinematic_title(slide, data)\n`;
    } else if (slide.layout === 'split') {
      script += `create_split_layout(slide, data)\n`;
    } else if (slide.layout === 'quadrant') {
      script += `create_quadrant_layout(slide, data)\n`;
    } else if (slide.layout === 'quote') {
      script += `create_quote_layout(slide, data)\n`;
    } else if (slide.layout === 'grid') {
      script += `create_grid_layout(slide, data)\n`;
    } else {
      script += `create_image_focus(slide, data)\n`;
    }
  });

  script += `
    # Save
    prs.save('UN_Diplomatic_Presentation.pptx')
    print("Generated UN_Diplomatic_Presentation.pptx")

if __name__ == "__main__":
    create_slides()
`;

  return script;
};

// ----------------------------------------------------------------------------
// UTILITY: Client-Side PPTX Generator (Inlined)
// ----------------------------------------------------------------------------

const generateClientPPTX = async (slidesList: any[]) => {
  const pres = new PptxGenJS();
  
  // Configure Presentation
  pres.layout = 'LAYOUT_16x9';
  pres.title = 'UN Diplomatic Presentation';
  pres.company = 'United Nations';

  const C_PRI = THEME.primary;
  const C_SEC = THEME.secondary;
  const C_WHITE = THEME.text;
  const C_DARK = "0A192F"; // Deep Navy
  
  // Helper to add footer
  const addFooter = (slide: PptxGenJS.Slide, index: number) => {
    slide.addText(
      [
        { text: `${index}`, options: { color: C_WHITE, fontSize: 10 } },
        { text: `  |  ${FOOTER_TEXT}`, options: { color: C_PRI, fontSize: 10 } }
      ],
      { x: 0.5, y: '95%', w: '90%', align: 'left' }
    );
  };

  for (let i = 0; i < slidesList.length; i++) {
    const data = slidesList[i];
    const slide = pres.addSlide();
    
    // Background Dark
    slide.background = { color: C_DARK };

    // --- LAYOUT LOGIC ---

    if (data.layout === 'title') {
      // Background Image
      if (data.imageUrl) {
        slide.addImage({ path: data.imageUrl, x: 0, y: 0, w: '100%', h: '100%' });
      }
      // Overlay
      slide.addShape('rect', { x: 0, y: 0, w: '100%', h: '100%', fill: { color: '000000', transparency: 30 } });

      // UN Badge
      slide.addShape('roundRect', { x: '42%', y: 1, w: '16%', h: 0.5, fill: { color: '000000', transparency: 30 }, line: { color: C_SEC, width: 1 } });
      slide.addText("UNITED NATIONS", { x: '42%', y: 1.1, w: '16%', align: 'center', fontSize: 12, color: C_SEC, bold: true });

      // Title
      slide.addText(data.title.toUpperCase(), { x: 0.5, y: 2.5, w: '90%', align: 'center', fontSize: 60, color: C_WHITE, bold: true, fontFace: 'Arial Black' });
      
      // Accent Line
      slide.addShape('rect', { x: '42%', y: 4.8, w: '16%', h: 0.05, fill: { color: C_PRI } });

      // Subtitle
      if (data.subtitle) {
        slide.addText(data.subtitle, { x: 1, y: 5.2, w: '80%', align: 'center', fontSize: 24, color: C_WHITE });
      }

    } else if (data.layout === 'split') {
      // Left Image
      if (data.imageUrl) {
        slide.addImage({ path: data.imageUrl, x: 0, y: 0, w: '50%', h: '100%' });
        slide.addShape('rect', { x: 0, y: 0, w: '50%', h: '100%', fill: { color: '000000', transparency: 10 } });
      }
      
      // Vertical Gold Line
      slide.addShape('rect', { x: '50%', y: 0, w: 0.05, h: '100%', fill: { color: C_SEC } });

      // Right Content
      slide.addText(data.title, { x: '53%', y: 0.8, w: '45%', fontSize: 40, color: C_WHITE, bold: true });
      slide.addShape('rect', { x: '53%', y: 1.8, w: '15%', h: 0.05, fill: { color: C_SEC } });

      if (data.points) {
        data.points.forEach((point: string, idx: number) => {
          slide.addText(point, { x: '53%', y: 2.5 + (idx * 0.8), w: '45%', fontSize: 18, color: C_WHITE, bullet: { type: 'number', numberType: 'arabicPeriod' } });
        });
      }
      addFooter(slide, i + 1);

    } else if (data.layout === 'quadrant') {
      if (data.imageUrl) {
        slide.addImage({ path: data.imageUrl, x: 0, y: 0, w: '100%', h: '100%' });
      }
      slide.addShape('rect', { x: 0, y: 0, w: '100%', h: '100%', fill: { color: C_DARK, transparency: 10 } }); // Darker tint

      // Title
      slide.addText(data.title.toUpperCase(), { x: 0, y: 0.4, w: '100%', align: 'center', fontSize: 36, color: C_WHITE, bold: true });

      // Grid
      const positions = [
        { x: 1, y: 1.8 }, { x: 6.5, y: 1.8 },
        { x: 1, y: 4.5 }, { x: 6.5, y: 4.5 }
      ];

      if (data.points) {
        data.points.forEach((point: string, idx: number) => {
          if (idx < 4) {
            const pos = positions[idx];
            // Box
            slide.addShape('rect', { x: pos.x, y: pos.y, w: 4.5, h: 2.2, fill: { color: '000000', transparency: 40 }, line: { color: C_WHITE, width: 0.5 } });
            // Number
            slide.addText(`0${idx + 1}`, { x: pos.x + 0.2, y: pos.y + 0.2, fontSize: 24, color: C_SEC, bold: true });
            // Text
            slide.addText(point, { x: pos.x + 0.2, y: pos.y + 0.8, w: 4.1, align: 'center', fontSize: 20, color: C_WHITE, bold: true });
          }
        });
      }
      addFooter(slide, i + 1);

    } else if (data.layout === 'quote') {
      if (data.imageUrl) {
        slide.addImage({ path: data.imageUrl, x: 0, y: 0, w: '100%', h: '100%' });
      }
      slide.addShape('rect', { x: 0, y: 0, w: '100%', h: '100%', fill: { color: '000000', transparency: 60 } });

      // Quote Mark
      slide.addText('"', { x: 1, y: 1.5, fontSize: 80, color: C_SEC, align: 'center' });

      // Text
      if (data.points && data.points[0]) {
        slide.addText(data.points[0], { x: 2, y: 2.5, w: '60%', align: 'center', fontSize: 36, color: C_WHITE, italic: true, fontFace: 'Times New Roman' });
      }

      // Author
      if (data.quoteAuthor) {
        slide.addText(data.quoteAuthor.toUpperCase(), { x: 2, y: 5.5, w: '60%', align: 'center', fontSize: 16, color: C_SEC, bold: true });
      }
      addFooter(slide, i + 1);

    } else if (data.layout === 'grid') {
       // Left Panel
       slide.addShape('rect', { x: 0, y: 0, w: '35%', h: '100%', fill: { color: C_DARK } });
       
       // Image Right
       if (data.imageUrl) {
         slide.addImage({ path: data.imageUrl, x: '35%', y: 0, w: '65%', h: '100%' });
         slide.addShape('rect', { x: '35%', y: 0, w: '65%', h: '100%', fill: { color: '000000', transparency: 20 } });
       }

       // Title Left
       slide.addText(data.title, { x: 0.5, y: 2.5, w: '30%', fontSize: 36, color: C_WHITE, bold: true });

       // Cards Right
       const positions = [
         { x: 4.2, y: 1 }, { x: 7.8, y: 1 },
         { x: 4.2, y: 4 }, { x: 7.8, y: 4 }
       ];
       
       if (data.points) {
         data.points.forEach((point: string, idx: number) => {
            if (idx < 4) {
              const pos = positions[idx];
              slide.addShape('rect', { x: pos.x, y: pos.y, w: 3.2, h: 2.5, fill: { color: C_DARK, transparency: 20 }, line: { color: C_WHITE } });
              slide.addText(`0${idx + 1}`, { x: pos.x + 0.1, y: pos.y + 0.1, fontSize: 14, color: C_WHITE });
              slide.addText(point, { x: pos.x + 0.2, y: pos.y + 1, w: 2.8, align: 'center', fontSize: 18, color: C_WHITE, bold: true });
            }
         });
       }
       addFooter(slide, i + 1);

    } else {
      // Image Focus / Standard
      if (data.imageUrl) {
        slide.addImage({ path: data.imageUrl, x: 0, y: 0, w: '100%', h: '100%' });
      }
      slide.addShape('rect', { x: 0, y: 0, w: '100%', h: '100%', fill: { color: '000000', transparency: 40 } });

      // Content Box
      slide.addShape('rect', { x: 1, y: 1.5, w: 5.5, h: 5, fill: { color: C_DARK, transparency: 10 }, line: { color: C_SEC } });
      
      slide.addText("KEY OBJECTIVES", { x: 1.5, y: 1.8, fontSize: 10, color: C_PRI, bold: true });
      slide.addText(data.title, { x: 1.5, y: 2.2, w: 4.5, fontSize: 36, color: C_WHITE, bold: true });

      if (data.points) {
        data.points.forEach((point: string, idx: number) => {
          slide.addText(point, { x: 1.5, y: 3.5 + (idx * 0.7), w: 4.5, fontSize: 18, color: C_WHITE, bullet: { type: 'number' } });
        });
      }
      addFooter(slide, i + 1);
    }
  }

  // Generate File
  await pres.writeFile({ fileName: 'UN_Diplomatic_Presentation.pptx' });
};

// ----------------------------------------------------------------------------
// COMPONENT: SlidePreview (Inlined)
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
      <motion.div 
        key={slide.id}
        initial={{ opacity: 0, scale: 0.98 }}
        animate={{ opacity: 1, scale: 1 }}
        transition={{ duration: 0.5, ease: "easeOut" }}
        className="relative shadow-2xl w-full max-w-5xl aspect-[16/9] flex flex-col overflow-hidden text-white rounded-lg border border-white/5"
        style={{ 
          fontFamily: "'Inter', sans-serif",
          backgroundImage: `url(${slide.imageUrl})`,
          backgroundSize: 'cover',
          backgroundPosition: 'center',
        }}
      >
        {/* Cinematic Blue/Gold Tint Overlay */}
        <div 
          className="absolute inset-0 z-0 pointer-events-none" 
          style={{ background: 'linear-gradient(135deg, rgba(10,25,47,0.85) 0%, rgba(0,30,60,0.4) 100%)' }}
        />

        {/* --- TITLE LAYOUT --- */}
        {isTitle && (
          <div className="relative z-10 flex flex-col items-center justify-center h-full text-center space-y-8">
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
              transition={{ delay: 0.4 }}
              className="text-7xl md:text-8xl font-black tracking-tighter uppercase drop-shadow-2xl text-white"
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
              className="text-xl font-light tracking-wide text-gray-300 max-w-2xl"
            >
              {slide.subtitle}
            </motion.p>
          </div>
        )}

        {/* --- SPLIT LAYOUT --- */}
        {isSplit && (
          <div className="relative z-10 grid grid-cols-2 h-full">
            <div className="col-span-1 relative group">
                {/* Image side - subtly revealing image on hover */}
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
                    transition={{ delay: 0.1 * idx }}
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
                        initial={{ opacity: 0 }}
                        animate={{ opacity: 1 }}
                        transition={{ delay: idx * 0.1 }}
                        className="bg-white/5 backdrop-blur-md border border-white/10 flex items-center justify-center relative overflow-hidden group hover:bg-[#009EDB]/20 transition-all duration-500 cursor-default"
                      >
                          <div className="absolute top-4 left-4 text-[#D4AF37] font-serif text-4xl opacity-50 group-hover:opacity-100 transition-opacity">
                            0{idx+1}
                          </div>
                          <h3 className="text-2xl font-bold tracking-widest uppercase z-10 text-center px-4">{point}</h3>
                          
                          {/* Corner Accents */}
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
    </div>
  );
};

// ----------------------------------------------------------------------------
// COMPONENT: PythonCodeViewer (Inlined)
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

  // Simple syntax highlighting helper
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
    <div className="flex flex-col h-full bg-[#1e1e1e] text-gray-300 rounded-lg overflow-hidden shadow-2xl">
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
          {copied ? 'Copied!' : 'Copy Script'}
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
