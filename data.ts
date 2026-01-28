import { SlideData, SlideType } from './types';

export const slides: SlideData[] = [
  {
    id: 1,
    type: SlideType.TITLE,
    layout: 'title',
    title: "THE UNITED NATIONS",
    subtitle: "Architecting Global Peace in the 21st Century",
    footer: "UN Global Report 2024",
    // UN Flags flying outside Headquarters
    imageUrl: "https://images.unsplash.com/photo-1599366744839-4467dc0b4d4b?q=80&w=2574&auto=format&fit=crop"
  },
  {
    id: 2,
    type: SlideType.CONTENT,
    layout: 'split',
    title: "1945: The Vision",
    points: [
      "Emerging from the ruins of WWII.",
      "51 Founding Nations.",
      "A Charter for humanity.",
      "The end of the League of Nations."
    ],
    // Historical/Document vibe
    imageUrl: "https://images.unsplash.com/photo-1580130601254-05fa235cd77e?q=80&w=2670&auto=format&fit=crop"
  },
  {
    id: 3,
    type: SlideType.CONTENT,
    layout: 'quadrant',
    title: "The Four Pillars",
    points: [
      "PEACE & SECURITY",
      "HUMAN RIGHTS",
      "RULE OF LAW",
      "DEVELOPMENT"
    ],
    // Peacekeeper / Blue Helmet vibe abstract
    imageUrl: "https://images.unsplash.com/photo-1628891510373-30595221376d?q=80&w=2670&auto=format&fit=crop"
  },
  {
    id: 4,
    type: SlideType.CONTENT,
    layout: 'grid',
    title: "The Main Organs",
    points: [
      "General Assembly",
      "Security Council",
      "Econ. & Social Council",
      "Secretariat"
    ],
    // The General Assembly Hall
    imageUrl: "https://images.unsplash.com/photo-1555848962-6e79363ec58f?q=80&w=2666&auto=format&fit=crop"
  },
  {
    id: 5,
    type: SlideType.CONTENT,
    layout: 'image-focus',
    title: "Security Council",
    subtitle: "The Horseshoe Table",
    points: [
      "5 Permanent Members (P5)",
      "Veto Power Authority",
      "Binding Resolutions",
      "Sanctions Regime"
    ],
    // Security Council Chamber / Round Table feel
    imageUrl: "https://images.unsplash.com/photo-1577985848520-278635567b5e?q=80&w=2670&auto=format&fit=crop"
  },
  {
    id: 6,
    type: SlideType.CONTENT,
    layout: 'split',
    title: "Agencies in Action",
    points: [
      "WHO: Global Pandemics",
      "UNICEF: Child Rights",
      "WFP: Emergency Food",
      "UNESCO: Culture"
    ],
    // Humanitarian Aid / Field Work
    imageUrl: "https://images.unsplash.com/photo-1469571486292-0ba58a3f068b?q=80&w=2670&auto=format&fit=crop"
  },
  {
    id: 7,
    type: SlideType.CONTENT,
    layout: 'grid',
    title: "Agenda 2030 (SDGs)",
    points: [
      "No Poverty",
      "Climate Action",
      "Gender Equality",
      "Clean Energy"
    ],
    // Nature / Green Energy / Sustainability
    imageUrl: "https://images.unsplash.com/photo-1473341304170-971dccb5ac1e?q=80&w=2670&auto=format&fit=crop"
  },
  {
    id: 8,
    type: SlideType.CONTENT,
    layout: 'quote',
    title: "Human Rights",
    points: [
      "All human beings are born free and equal in dignity and rights."
    ],
    quoteAuthor: "UDHR, Article 1",
    // Diverse hands / Unity
    imageUrl: "https://images.unsplash.com/photo-1531206715517-5c0ba140b2b8?q=80&w=2670&auto=format&fit=crop"
  },
  {
    id: 9,
    type: SlideType.CONTENT,
    layout: 'image-focus',
    title: "Global Crises",
    points: [
      "Climate Emergency",
      "Cyber Warfare",
      "Refugee Crisis",
      "Geopolitical Polarization"
    ],
    // Storm / Conflict / Dramatic
    imageUrl: "https://images.unsplash.com/photo-1616164283832-618e7e132805?q=80&w=2574&auto=format&fit=crop"
  },
  {
    id: 10,
    type: SlideType.CONTENT,
    layout: 'title',
    title: "THE FUTURE",
    subtitle: "Reform. Innovate. Unite.",
    footer: "UN Strategic Vision 2030",
    // Geneva / Palace of Nations exterior
    imageUrl: "https://images.unsplash.com/photo-1575356885361-b66a5e1e23df?q=80&w=2674&auto=format&fit=crop"
  }
];
