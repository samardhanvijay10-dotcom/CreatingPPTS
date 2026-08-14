import React, { useState, useEffect } from 'react';
import pptxgen from 'pptxgenjs';
import confetti from 'canvas-confetti';
import {
  Sparkles,
  Plus,
  Trash2,
  Copy,
  ChevronUp,
  ChevronDown,
  Download,
  Play,
  Layers,
  Palette,
  Layout,
  Grid,
  X,
  ChevronLeft,
  ChevronRight,
  Quote as QuoteIcon,
  CheckCircle2,
  Sliders,
  Wand2,
  Clock,
  Users,
  Check,
  Search,
  FileText
} from 'lucide-react';

// Expanded Theme Preset Definitions
const THEMES = {
  docuCraft: {
    id: 'docuCraft',
    name: 'DocuCraft Signature',
    category: 'Signature',
    bg: '#030712',
    cardBg: '#0B132B',
    accent: '#3B82F6',
    accentText: '#60A5FA',
    textMain: '#F8FAFC',
    textMuted: '#94A3B8',
    border: '#1E293B',
    font: 'font-sans',
    isDark: true
  },
  sleekTech: {
    id: 'sleekTech',
    name: 'Sleek Tech',
    category: 'Modern',
    bg: '#0F172A',
    cardBg: '#1E293B',
    accent: '#6366F1',
    accentText: '#818CF8',
    textMain: '#F8FAFC',
    textMuted: '#94A3B8',
    border: '#334155',
    font: 'font-sans',
    isDark: true
  },
  sunsetGlow: {
    id: 'sunsetGlow',
    name: 'Sunset Glow',
    category: 'Vibrant',
    bg: '#180D1C',
    cardBg: '#2D1236',
    accent: '#F43F5E',
    accentText: '#FB7185',
    textMain: '#FFF1F2',
    textMuted: '#FDA4AF',
    border: '#4C1D54',
    font: 'font-sans',
    isDark: true
  },
  cyberpunk: {
    id: 'cyberpunk',
    name: 'Cyberpunk Neon',
    category: 'Vibrant',
    bg: '#080811',
    cardBg: '#121225',
    accent: '#EC4899',
    accentText: '#F472B6',
    textMain: '#F9FAFB',
    textMuted: '#A1A1AA',
    border: '#27272A',
    font: 'font-sans',
    isDark: true
  },
  warmEditorial: {
    id: 'warmEditorial',
    name: 'Warm Editorial',
    category: 'Classic',
    bg: '#FAF7F2',
    cardBg: '#F2ECE1',
    accent: '#991B1B',
    accentText: '#B91C1C',
    textMain: '#1C1917',
    textMuted: '#78716C',
    border: '#E7E5E4',
    font: 'font-serif',
    isDark: false
  },
  pastelNordic: {
    id: 'pastelNordic',
    name: 'Pastel Nordic',
    category: 'Minimal',
    bg: '#F3F4F6',
    cardBg: '#FFFFFF',
    accent: '#0D9488',
    accentText: '#0F766E',
    textMain: '#111827',
    textMuted: '#6B7280',
    border: '#E5E7EB',
    font: 'font-sans',
    isDark: false
  },
  midnightAurora: {
    id: 'midnightAurora',
    name: 'Midnight Aurora',
    category: 'Modern',
    bg: '#030712',
    cardBg: '#111827',
    accent: '#10B981',
    accentText: '#34D399',
    textMain: '#F9FAFB',
    textMuted: '#9CA3AF',
    border: '#1F2937',
    font: 'font-sans',
    isDark: true
  },
  corporateSlate: {
    id: 'corporateSlate',
    name: 'Executive Slate',
    category: 'Executive',
    bg: '#0A192F',
    cardBg: '#112240',
    accent: '#38BDF8',
    accentText: '#7DD3FC',
    textMain: '#F8FAFC',
    textMuted: '#8892B0',
    border: '#233554',
    font: 'font-sans',
    isDark: true
  },
  swissMinimalist: {
    id: 'swissMinimalist',
    name: 'Swiss Brutalist',
    category: 'Brutalist',
    bg: '#000000',
    cardBg: '#141414',
    accent: '#FACC15',
    accentText: '#FDE047',
    textMain: '#FFFFFF',
    textMuted: '#A3A3A3',
    border: '#262626',
    font: 'font-mono',
    isDark: true
  }
};

// Preset Decks
const PRESET_DECKS = {
  startup: {
    name: '🚀 Startup Pitch Deck',
    slides: [
      {
        id: 's1',
        layout: 'title',
        theme: 'docuCraft',
        tag: 'SERIES A PITCH',
        title: 'DocuCraft AI Studio',
        subtitle: 'The autonomous document & presentation intelligence platform for modern enterprises.',
        footerText: 'DocuCraft Studio • Confidential & Proprietary 2025'
      },
      {
        id: 's2',
        layout: 'split',
        theme: 'docuCraft',
        tag: 'PROBLEM & SOLUTION',
        title: 'Unifying Segmented Workflows',
        subtitle: 'Modern enterprise knowledge teams waste 18+ hours per week formatting slides & report assets.',
        col1Title: 'Fragmented Tools',
        col1Text: 'Disconnected document editors, manual chart formatting, and inconsistent brand guidelines slow down delivery.',
        col2Title: 'DocuCraft Engine',
        col2Text: 'AI-native layout generation, instant brand compliance, and multi-format publishing in one cohesive studio.',
        bullets: ['Real-time collaborative editing', 'One-click PowerPoint & PDF export', 'Enterprise RBAC security']
      },
      {
        id: 's3',
        layout: 'timeline',
        theme: 'docuCraft',
        tag: 'PRODUCT ROADMAP',
        title: 'Execution Milestones',
        subtitle: 'Key strategic objectives planned over the next 12 months.',
        timelineSteps: [
          { step: '01', title: 'Core Launch', desc: 'Presentation engine & PowerPoint export release.' },
          { step: '02', title: 'AI Copilot', desc: 'Autonomous content generation & real-time formatting.' },
          { step: '03', title: 'Enterprise Hub', desc: 'Custom template builder & team permission management.' }
        ]
      },
      {
        id: 's4',
        layout: 'stat',
        theme: 'docuCraft',
        tag: 'MARKET TRACTION',
        title: 'Hyper-Growth Scale',
        subtitle: 'Rapid enterprise adoption within 12 months of soft launch.',
        statValue: '$18.4M',
        statLabel: 'Annual Recurring Revenue (ARR)',
        statSubtext: '+340% YoY expansion with zero churn in Fortune 500 tier.',
        bullets: ['650+ Active Enterprise Accounts', '99.4% User Retention Rate', 'Net Promoter Score (NPS) 82']
      },
      {
        id: 's5',
        layout: 'grid',
        theme: 'docuCraft',
        tag: 'KEY CAPABILITIES',
        title: 'Built for Performance',
        subtitle: 'Comprehensive toolkit engineered for modern design & documentation teams.',
        gridItems: [
          { title: 'Instant Theme Engine', desc: 'Switch visual identities across full decks in 1 click.' },
          { title: 'Vector PPTX Export', desc: 'Fully editable PowerPoint shapes & formatted text cards.' },
          { title: 'Interactive Present Mode', desc: 'Smooth fullscreen presenter mode with timer and controls.' },
          { title: 'Brand Asset Sync', desc: 'Maintain strict typography, color standards, and badges.' }
        ]
      },
      {
        id: 's6',
        layout: 'quote',
        theme: 'docuCraft',
        tag: 'CUSTOMER SUCCESS',
        title: 'Executive Endorsement',
        quote: 'DocuCraft Studio completely elevated how our team presents to enterprise buyers. We close deals 40% faster.',
        quoteAuthor: 'Sarah Jenkins',
        quoteRole: 'Chief Commercial Officer, Veloce Global'
      }
    ]
  },
  portfolio: {
    name: '🎨 Creative Portfolio',
    slides: [
      {
        id: 'p1',
        layout: 'title',
        theme: 'warmEditorial',
        tag: 'PORTFOLIO 2025',
        title: 'Aethel Design Studio',
        subtitle: 'Crafting timeless digital products, editorial systems, and immersive brand identities.',
        footerText: 'DocuCraft Studio • Selected Works'
      },
      {
        id: 'p2',
        layout: 'content',
        theme: 'warmEditorial',
        tag: 'DESIGN PHILOSOPHY',
        title: 'Simplicity & Precision',
        subtitle: 'Good design is invisible—it empowers users while evoking emotion and clarity.',
        bullets: [
          'Human-centered interaction design and tactile typography',
          'Rigorous grid systems paired with warm, expressive palettes',
          'Sustainable digital architecture engineered for lightning performance'
        ]
      },
      {
        id: 'p3',
        layout: 'team',
        theme: 'warmEditorial',
        tag: 'CREATIVE LEADERSHIP',
        title: 'Meet the Artisans',
        subtitle: 'Multi-disciplinary designers and creative technologists.',
        teamMembers: [
          { name: 'Elena Rostova', role: 'Design Director', bio: 'Former Lead at Studio Mono. 10+ years in typography & brand architecture.' },
          { name: 'Marcus Vance', role: 'Head of Interaction', bio: 'Pioneer in spatial UX and generative design workflows.' },
          { name: 'Aria Takahashi', role: 'Principal Strategist', bio: 'Specializes in luxury branding and digital storytelling.' }
        ]
      },
      {
        id: 'p4',
        layout: 'quote',
        theme: 'warmEditorial',
        tag: 'CLIENT IMPRESSION',
        title: 'Testimonial',
        quote: 'Working with Aethel transformed how our company communicates value. Their design gave us immediate authority.',
        quoteAuthor: 'Julian Croft',
        quoteRole: 'Founder, Luminary Press'
      }
    ]
  },
  techRoadmap: {
    name: '⚡ Technical Architecture',
    slides: [
      {
        id: 'r1',
        layout: 'title',
        theme: 'cyberpunk',
        tag: 'SYSTEM ARCHITECTURE v3.0',
        title: 'NextGen Edge Cloud Grid',
        subtitle: 'Distributed, zero-trust edge compute infrastructure built for real-time document AI.',
        footerText: 'DocuCraft Studio • Engineering Sync'
      },
      {
        id: 'r2',
        layout: 'stat',
        theme: 'cyberpunk',
        tag: 'PERFORMANCE BENCHMARKS',
        title: 'Sub-Millisecond Execution',
        subtitle: 'Benchmarking global edge node processing under peak workloads.',
        statValue: '< 8ms',
        statLabel: 'P99 Global Latency',
        statSubtext: 'Optimized via Rust WebAssembly isolates deployed across 200+ edge locations.',
        bullets: ['99.999% SLA Uptime Guarantee', 'Automated DDoS & Anomaly Shield', 'Zero cold-start execution']
      },
      {
        id: 'r3',
        layout: 'comparison',
        theme: 'cyberpunk',
        tag: 'INFRASTRUCTURE STACK',
        title: 'Legacy vs NextGen Engine',
        subtitle: 'Comparing traditional document rendering with our edge WASM engine.',
        compTitle1: 'Legacy Server Render',
        compText1: 'Heavy Node.js server instances, high memory consumption, 450ms average render delay.',
        compTitle2: 'DocuCraft WASM Edge',
        compText2: 'Sub-10ms edge rendering, zero memory leaks, local hardware acceleration.',
        bullets: ['85% reduction in cloud compute costs', 'Instant client offline preview', 'End-to-end client encryption']
      }
    ]
  },
  executiveQBR: {
    name: '📈 Executive QBR Review',
    slides: [
      {
        id: 'q1',
        layout: 'title',
        theme: 'corporateSlate',
        tag: 'Q4 BUSINESS REVIEW',
        title: 'Global Growth Strategy',
        subtitle: 'Performance report, key achievements, and strategic objectives for Q1 2026.',
        footerText: 'DocuCraft Studio • Executive Board Presentation'
      },
      {
        id: 'q2',
        layout: 'stat',
        theme: 'corporateSlate',
        tag: 'FINANCIAL PERFORMANCE',
        title: 'Record Quarterly Revenue',
        subtitle: 'Exceeded all Q4 targets across North America and European enterprise markets.',
        statValue: '$34.5M',
        statLabel: 'Q4 Gross Revenue',
        statSubtext: '148% of quarterly goal driven by enterprise expansion.',
        bullets: ['ARR surpassed $120M landmark', 'Gross Margins expanded to 84%', 'EBITDA positive for 3rd straight quarter']
      },
      {
        id: 'q3',
        layout: 'timeline',
        theme: 'corporateSlate',
        tag: 'STRATEGIC HORIZON',
        title: '2026 Growth Pillars',
        subtitle: 'Three key initiatives driving market dominance.',
        timelineSteps: [
          { step: '01', title: 'APAC Expansion', desc: 'Opening Tokyo & Singapore offices.' },
          { step: '02', title: 'AI Automation', desc: 'Deploying autonomous document workflow agents.' },
          { step: '03', title: 'M&A Integration', desc: 'Integrating recent analytics studio acquisition.' }
        ]
      }
    ]
  }
};

export default function App() {
  const [slides, setSlides] = useState(PRESET_DECKS.startup.slides);
  const [activeSlideIndex, setActiveSlideIndex] = useState(0);
  const [showGridPattern, setShowGridPattern] = useState(true);
  const [showGlowEffect, setShowGlowEffect] = useState(true);
  const [showSlideNumber, setShowSlideNumber] = useState(true);
  const [showGlassCard, setShowGlassCard] = useState(true);
  const [aspectRatio, setAspectRatio] = useState('16:9'); // '16:9' or '4:3'
  const [isPresenting, setIsPresenting] = useState(false);
  const [isExporting, setIsExporting] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');
  const [toastMessage, setToastMessage] = useState(null);
  const [showQuickModal, setShowQuickModal] = useState(false);

  // Presenter Timer
  const [presenterTime, setPresenterTime] = useState(0);

  const activeSlide = slides[activeSlideIndex] || slides[0];
  const activeTheme = THEMES[activeSlide?.theme] || THEMES.docuCraft;

  // Show Toast helper
  const triggerToast = (msg) => {
    setToastMessage(msg);
    setTimeout(() => {
      setToastMessage(null);
    }, 3200);
  };

  // Keyboard Navigation in Presenter Mode
  useEffect(() => {
    let interval = null;
    if (isPresenting) {
      interval = setInterval(() => {
        setPresenterTime((prev) => prev + 1);
      }, 1000);
    } else {
      setPresenterTime(0);
    }
    return () => clearInterval(interval);
  }, [isPresenting]);

  useEffect(() => {
    const handleKeyDown = (e) => {
      if (!isPresenting) return;
      if (e.key === 'ArrowRight' || e.key === 'Space') {
        setActiveSlideIndex((prev) => Math.min(prev + 1, slides.length - 1));
      } else if (e.key === 'ArrowLeft') {
        setActiveSlideIndex((prev) => Math.max(prev - 1, 0));
      } else if (e.key === 'Escape') {
        setIsPresenting(false);
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [isPresenting, slides.length]);

  // Format Timer
  const formatTime = (secs) => {
    const m = Math.floor(secs / 60);
    const s = secs % 60;
    return `${m.toString().padStart(2, '0')}:${s.toString().padStart(2, '0')}`;
  };

  // Slide CRUD Operations
  const updateActiveSlide = (fields) => {
    setSlides((prev) =>
      prev.map((slide, idx) => (idx === activeSlideIndex ? { ...slide, ...fields } : slide))
    );
  };

  const addSlide = (layoutType = 'content') => {
    const defaultSlide = {
      id: 's_' + Date.now(),
      layout: layoutType,
      theme: activeSlide?.theme || 'docuCraft',
      tag: 'NEW SECTION',
      title: 'Slide Headline Here',
      subtitle: 'Add a concise narrative or key takeaway here.',
      bullets: ['Key milestone objective', 'Analytical insight or metric', 'Actionable decision point']
    };

    if (layoutType === 'timeline') {
      defaultSlide.timelineSteps = [
        { step: '01', title: 'Phase One', desc: 'Initial discovery and scoping.' },
        { step: '02', title: 'Phase Two', desc: 'Execution and iteration.' },
        { step: '03', title: 'Phase Three', desc: 'Global rollout and review.' }
      ];
    } else if (layoutType === 'comparison') {
      defaultSlide.compTitle1 = 'Option A';
      defaultSlide.compText1 = 'Description of the primary strategy or existing system.';
      defaultSlide.compTitle2 = 'Option B';
      defaultSlide.compText2 = 'Description of the proposed solution or future system.';
    } else if (layoutType === 'grid') {
      defaultSlide.gridItems = [
        { title: 'Feature One', desc: 'Detailed capability summary.' },
        { title: 'Feature Two', desc: 'Detailed capability summary.' },
        { title: 'Feature Three', desc: 'Detailed capability summary.' },
        { title: 'Feature Four', desc: 'Detailed capability summary.' }
      ];
    } else if (layoutType === 'team') {
      defaultSlide.teamMembers = [
        { name: 'Alex Rivera', role: 'Project Director', bio: 'Specialist in enterprise design workflows.' },
        { name: 'Sam Chen', role: 'Lead Architect', bio: 'Focuses on edge compute and real-time systems.' },
        { name: 'Morgan Vance', role: 'Strategy Director', bio: 'Expert in commercial positioning.' }
      ];
    }

    const newSlides = [...slides];
    newSlides.splice(activeSlideIndex + 1, 0, defaultSlide);
    setSlides(newSlides);
    setActiveSlideIndex(activeSlideIndex + 1);
    triggerToast('Added new slide layout');
  };

  const duplicateSlide = () => {
    const dup = {
      ...JSON.parse(JSON.stringify(activeSlide)),
      id: 's_' + Date.now()
    };
    const newSlides = [...slides];
    newSlides.splice(activeSlideIndex + 1, 0, dup);
    setSlides(newSlides);
    setActiveSlideIndex(activeSlideIndex + 1);
    triggerToast('Duplicated slide');
  };

  const deleteSlide = () => {
    if (slides.length <= 1) return;
    const newSlides = slides.filter((_, idx) => idx !== activeSlideIndex);
    setSlides(newSlides);
    setActiveSlideIndex(Math.max(0, activeSlideIndex - 1));
    triggerToast('Deleted slide');
  };

  const moveSlide = (direction) => {
    const targetIdx = activeSlideIndex + direction;
    if (targetIdx < 0 || targetIdx >= slides.length) return;
    const newSlides = [...slides];
    const temp = newSlides[activeSlideIndex];
    newSlides[activeSlideIndex] = newSlides[targetIdx];
    newSlides[targetIdx] = temp;
    setSlides(newSlides);
    setActiveSlideIndex(targetIdx);
  };

  const loadPresetDeck = (key) => {
    if (PRESET_DECKS[key]) {
      setSlides(PRESET_DECKS[key].slides);
      setActiveSlideIndex(0);
      triggerToast(`Loaded ${PRESET_DECKS[key].name}`);
    }
  };

  // Bullet Point Operations
  const updateBullet = (idx, value) => {
    const newBullets = [...(activeSlide.bullets || [])];
    newBullets[idx] = value;
    updateActiveSlide({ bullets: newBullets });
  };

  const addBullet = () => {
    const newBullets = [...(activeSlide.bullets || []), 'New key takeaway point'];
    updateActiveSlide({ bullets: newBullets });
  };

  const removeBullet = (idx) => {
    const newBullets = (activeSlide.bullets || []).filter((_, i) => i !== idx);
    updateActiveSlide({ bullets: newBullets });
  };

  // Timeline Step Operations
  const updateTimelineStep = (idx, key, val) => {
    const steps = [...(activeSlide.timelineSteps || [])];
    if (!steps[idx]) return;
    steps[idx][key] = val;
    updateActiveSlide({ timelineSteps: steps });
  };

  // Grid Item Operations
  const updateGridItem = (idx, key, val) => {
    const items = [...(activeSlide.gridItems || [])];
    if (!items[idx]) return;
    items[idx][key] = val;
    updateActiveSlide({ gridItems: items });
  };

  // Team Member Operations
  const updateTeamMember = (idx, key, val) => {
    const members = [...(activeSlide.teamMembers || [])];
    if (!members[idx]) return;
    members[idx][key] = val;
    updateActiveSlide({ teamMembers: members });
  };

  // PowerPoint PPTX Native Exporter
  const exportToPowerPoint = async () => {
    try {
      setIsExporting(true);
      const pptx = new pptxgen();

      // Configure Layout Widescreen or Standard
      pptx.layout = aspectRatio === '4:3' ? 'LAYOUT_4x3' : 'LAYOUT_16x9';
      pptx.title = 'DocuCraft Studio Presentation';

      slides.forEach((slide) => {
        const theme = THEMES[slide.theme] || THEMES.docuCraft;
        const pptxSlide = pptx.addSlide();

        const customBg = slide.customBgColor || theme.bg;
        const customAccent = slide.customAccentColor || theme.accent;

        // Background Color
        pptxSlide.background = { color: customBg.replace('#', '') };

        const accentHex = customAccent.replace('#', '');
        const textMainHex = theme.textMain.replace('#', '');
        const textMutedHex = theme.textMuted.replace('#', '');

        // Category Badge
        if (slide.tag) {
          pptxSlide.addText(slide.tag.toUpperCase(), {
            x: 0.8,
            y: 0.5,
            w: 8.4,
            h: 0.4,
            fontSize: 12,
            bold: true,
            color: accentHex,
            fontFace: 'Arial'
          });
        }

        // Main Title
        if (slide.title) {
          pptxSlide.addText(slide.title, {
            x: 0.8,
            y: slide.layout === 'title' ? 1.8 : 0.9,
            w: 8.4,
            h: 0.9,
            fontSize: slide.layout === 'title' ? 36 : 26,
            bold: true,
            color: textMainHex,
            fontFace: 'Arial'
          });
        }

        // Subtitle
        if (slide.subtitle) {
          pptxSlide.addText(slide.subtitle, {
            x: 0.8,
            y: slide.layout === 'title' ? 2.9 : 1.8,
            w: 8.4,
            h: 0.7,
            fontSize: 15,
            color: textMutedHex,
            fontFace: 'Arial'
          });
        }

        // Layout Specific Handlers
        if (slide.layout === 'content' && slide.bullets) {
          const bulletItems = slide.bullets.map((b) => ({
            text: b,
            options: { fontSize: 15, color: textMainHex, breakLine: true, bullet: true }
          }));
          pptxSlide.addText(bulletItems, {
            x: 0.8,
            y: 2.7,
            w: 8.4,
            h: 2.5,
            fontFace: 'Arial'
          });
        } else if (slide.layout === 'split') {
          if (slide.col1Title) {
            pptxSlide.addText(slide.col1Title, {
              x: 0.8,
              y: 2.7,
              w: 4.0,
              h: 0.4,
              fontSize: 17,
              bold: true,
              color: accentHex,
              fontFace: 'Arial'
            });
            pptxSlide.addText(slide.col1Text || '', {
              x: 0.8,
              y: 3.2,
              w: 4.0,
              h: 1.8,
              fontSize: 13,
              color: textMutedHex,
              fontFace: 'Arial'
            });
          }
          if (slide.col2Title) {
            pptxSlide.addText(slide.col2Title, {
              x: 5.2,
              y: 2.7,
              w: 4.0,
              h: 0.4,
              fontSize: 17,
              bold: true,
              color: accentHex,
              fontFace: 'Arial'
            });
            pptxSlide.addText(slide.col2Text || '', {
              x: 5.2,
              y: 3.2,
              w: 4.0,
              h: 1.8,
              fontSize: 13,
              color: textMutedHex,
              fontFace: 'Arial'
            });
          }
        } else if (slide.layout === 'stat') {
          if (slide.statValue) {
            pptxSlide.addText(slide.statValue, {
              x: 0.8,
              y: 2.6,
              w: 4.0,
              h: 1.2,
              fontSize: 46,
              bold: true,
              color: accentHex,
              fontFace: 'Arial'
            });
          }
          if (slide.statLabel) {
            pptxSlide.addText(slide.statLabel, {
              x: 0.8,
              y: 3.8,
              w: 4.0,
              h: 0.4,
              fontSize: 15,
              bold: true,
              color: textMainHex,
              fontFace: 'Arial'
            });
          }
          if (slide.statSubtext) {
            pptxSlide.addText(slide.statSubtext, {
              x: 0.8,
              y: 4.2,
              w: 4.0,
              h: 0.8,
              fontSize: 12,
              color: textMutedHex,
              fontFace: 'Arial'
            });
          }
          if (slide.bullets) {
            const bulletItems = slide.bullets.map((b) => ({
              text: b,
              options: { fontSize: 13, color: textMainHex, breakLine: true, bullet: true }
            }));
            pptxSlide.addText(bulletItems, {
              x: 5.2,
              y: 2.7,
              w: 4.0,
              h: 2.2,
              fontFace: 'Arial'
            });
          }
        } else if (slide.layout === 'quote') {
          if (slide.quote) {
            pptxSlide.addText(`“${slide.quote}”`, {
              x: 1.0,
              y: 2.6,
              w: 8.0,
              h: 1.6,
              fontSize: 22,
              italic: true,
              color: textMainHex,
              fontFace: 'Arial'
            });
          }
          if (slide.quoteAuthor) {
            pptxSlide.addText(`— ${slide.quoteAuthor} (${slide.quoteRole || ''})`, {
              x: 1.0,
              y: 4.3,
              w: 8.0,
              h: 0.5,
              fontSize: 14,
              bold: true,
              color: accentHex,
              fontFace: 'Arial'
            });
          }
        } else if (slide.layout === 'timeline' && slide.timelineSteps) {
          slide.timelineSteps.forEach((st, idx) => {
            const posX = 0.8 + idx * 2.8;
            pptxSlide.addText(st.step || `0${idx + 1}`, {
              x: posX,
              y: 2.6,
              w: 2.5,
              h: 0.4,
              fontSize: 18,
              bold: true,
              color: accentHex,
              fontFace: 'Arial'
            });
            pptxSlide.addText(st.title || '', {
              x: posX,
              y: 3.0,
              w: 2.5,
              h: 0.4,
              fontSize: 14,
              bold: true,
              color: textMainHex,
              fontFace: 'Arial'
            });
            pptxSlide.addText(st.desc || '', {
              x: posX,
              y: 3.4,
              w: 2.5,
              h: 1.4,
              fontSize: 12,
              color: textMutedHex,
              fontFace: 'Arial'
            });
          });
        } else if (slide.layout === 'comparison') {
          if (slide.compTitle1) {
            pptxSlide.addText(slide.compTitle1, {
              x: 0.8,
              y: 2.6,
              w: 4.0,
              h: 0.4,
              fontSize: 16,
              bold: true,
              color: accentHex,
              fontFace: 'Arial'
            });
            pptxSlide.addText(slide.compText1 || '', {
              x: 0.8,
              y: 3.1,
              w: 4.0,
              h: 1.8,
              fontSize: 13,
              color: textMutedHex,
              fontFace: 'Arial'
            });
          }
          if (slide.compTitle2) {
            pptxSlide.addText(slide.compTitle2, {
              x: 5.2,
              y: 2.6,
              w: 4.0,
              h: 0.4,
              fontSize: 16,
              bold: true,
              color: accentHex,
              fontFace: 'Arial'
            });
            pptxSlide.addText(slide.compText2 || '', {
              x: 5.2,
              y: 3.1,
              w: 4.0,
              h: 1.8,
              fontSize: 13,
              color: textMutedHex,
              fontFace: 'Arial'
            });
          }
        } else if (slide.layout === 'grid' && slide.gridItems) {
          slide.gridItems.forEach((item, idx) => {
            const col = idx % 2;
            const row = Math.floor(idx / 2);
            const posX = 0.8 + col * 4.4;
            const posY = 2.6 + row * 1.3;

            pptxSlide.addText(item.title || '', {
              x: posX,
              y: posY,
              w: 4.0,
              h: 0.3,
              fontSize: 14,
              bold: true,
              color: accentHex,
              fontFace: 'Arial'
            });
            pptxSlide.addText(item.desc || '', {
              x: posX,
              y: posY + 0.3,
              w: 4.0,
              h: 0.8,
              fontSize: 12,
              color: textMutedHex,
              fontFace: 'Arial'
            });
          });
        } else if (slide.layout === 'team' && slide.teamMembers) {
          slide.teamMembers.forEach((member, idx) => {
            const posX = 0.8 + idx * 2.8;
            pptxSlide.addText(member.name || '', {
              x: posX,
              y: 2.7,
              w: 2.5,
              h: 0.4,
              fontSize: 15,
              bold: true,
              color: textMainHex,
              fontFace: 'Arial'
            });
            pptxSlide.addText(member.role || '', {
              x: posX,
              y: 3.1,
              w: 2.5,
              h: 0.3,
              fontSize: 12,
              bold: true,
              color: accentHex,
              fontFace: 'Arial'
            });
            pptxSlide.addText(member.bio || '', {
              x: posX,
              y: 3.4,
              w: 2.5,
              h: 1.4,
              fontSize: 11,
              color: textMutedHex,
              fontFace: 'Arial'
            });
          });
        }

        // Footer Text
        if (slide.footerText) {
          pptxSlide.addText(slide.footerText, {
            x: 0.8,
            y: 5.2,
            w: 8.4,
            h: 0.3,
            fontSize: 10,
            color: textMutedHex,
            fontFace: 'Arial'
          });
        }
      });

      await pptx.writeFile({ fileName: 'DocuCraft_Studio_Presentation.pptx' });
      confetti({
        particleCount: 140,
        spread: 90,
        origin: { y: 0.6 }
      });
      triggerToast('PowerPoint deck exported successfully!');
    } catch (err) {
      console.error('PPT Export Error:', err);
      alert('Error exporting PowerPoint deck. Check console details.');
    } finally {
      setIsExporting(false);
    }
  };

  // Filter slides by search query
  const filteredSlides = slides.filter((s) => {
    if (!searchQuery) return true;
    const q = searchQuery.toLowerCase();
    return (
      (s.title && s.title.toLowerCase().includes(q)) ||
      (s.subtitle && s.subtitle.toLowerCase().includes(q)) ||
      (s.tag && s.tag.toLowerCase().includes(q)) ||
      (s.layout && s.layout.toLowerCase().includes(q))
    );
  });

  return (
    <div className="min-h-screen bg-[#06080E] text-slate-100 flex flex-col font-sans selection:bg-blue-600 selection:text-white">
      {/* TOAST NOTIFICATION BADGE */}
      {toastMessage && (
        <div className="fixed bottom-6 right-6 z-50 bg-indigo-600 text-white px-4 py-2.5 rounded-xl shadow-2xl flex items-center gap-2 text-xs font-semibold animate-bounce border border-indigo-400/40">
          <Sparkles className="w-4 h-4 text-amber-300 animate-spin" />
          <span>{toastMessage}</span>
        </div>
      )}

      {/* TOP BRAND HEADER NAVBAR */}
      <header className="h-16 border-b border-slate-800/80 bg-[#0A0E17]/90 backdrop-blur-md px-6 flex items-center justify-between z-20 sticky top-0">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-xl bg-gradient-to-tr from-blue-600 via-indigo-600 to-purple-600 flex items-center justify-center shadow-lg shadow-blue-500/25 border border-blue-400/30">
            <Sparkles className="w-5 h-5 text-white" />
          </div>
          <div>
            <div className="flex items-center gap-2">
              <h1 className="text-base font-extrabold tracking-wide text-white">
                DOCUCRAFT <span className="text-blue-400">STUDIO</span>
              </h1>
              <span className="text-[10px] px-2 py-0.5 rounded-full bg-blue-500/15 text-blue-400 font-mono font-semibold border border-blue-500/30">
                PPT MAKER v2.0
              </span>
            </div>
            <p className="text-[11px] text-slate-400">Aesthetic Autonomous Presentation Generator</p>
          </div>
        </div>

        {/* TOP BAR ACTIONS & PRESETS */}
        <div className="flex items-center gap-3">
          <div className="hidden xl:flex items-center gap-1.5 bg-slate-900/90 border border-slate-800 rounded-lg p-1 text-xs">
            <span className="text-slate-400 px-2 font-medium flex items-center gap-1">
              <Wand2 className="w-3.5 h-3.5 text-amber-400" /> Presets:
            </span>
            <button
              onClick={() => loadPresetDeck('startup')}
              className="px-2.5 py-1 rounded-md hover:bg-slate-800 text-slate-300 transition hover:text-white"
            >
              Startup Pitch
            </button>
            <button
              onClick={() => loadPresetDeck('portfolio')}
              className="px-2.5 py-1 rounded-md hover:bg-slate-800 text-slate-300 transition hover:text-white"
            >
              Portfolio
            </button>
            <button
              onClick={() => loadPresetDeck('techRoadmap')}
              className="px-2.5 py-1 rounded-md hover:bg-slate-800 text-slate-300 transition hover:text-white"
            >
              Tech Architecture
            </button>
            <button
              onClick={() => loadPresetDeck('executiveQBR')}
              className="px-2.5 py-1 rounded-md hover:bg-slate-800 text-slate-300 transition hover:text-white"
            >
              Executive QBR
            </button>
          </div>

          <button
            onClick={() => setShowQuickModal(true)}
            className="flex items-center gap-1.5 px-3 py-2 rounded-lg bg-slate-900 hover:bg-slate-800 text-slate-300 text-xs font-medium border border-slate-800 transition"
            title="Quick Slide Generator"
          >
            <Wand2 className="w-3.5 h-3.5 text-amber-400" />
            <span className="hidden sm:inline">Add Layout</span>
          </button>

          <button
            onClick={() => setIsPresenting(true)}
            className="flex items-center gap-2 px-3.5 py-2 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs font-medium transition border border-slate-700 shadow-sm"
          >
            <Play className="w-3.5 h-3.5 text-emerald-400 fill-emerald-400" />
            <span>Present</span>
          </button>

          <button
            onClick={exportToPowerPoint}
            disabled={isExporting}
            className="flex items-center gap-2 px-4 py-2 rounded-lg bg-gradient-to-r from-blue-600 via-indigo-600 to-purple-600 hover:from-blue-500 hover:to-purple-500 text-white text-xs font-semibold shadow-lg shadow-indigo-600/30 transition transform active:scale-95 disabled:opacity-50"
          >
            <Download className="w-3.5 h-3.5" />
            <span>{isExporting ? 'Exporting...' : 'Export PPTX'}</span>
          </button>
        </div>
      </header>

      {/* MAIN WORKSPACE LAYOUT */}
      <div className="flex-1 flex overflow-hidden">
        {/* LEFT SIDEBAR: SLIDE THUMBNAILS & SEARCH */}
        <aside className="w-72 border-r border-slate-800/80 bg-[#080B12] flex flex-col z-10">
          <div className="p-3.5 border-b border-slate-800/80 space-y-2">
            <div className="flex items-center justify-between">
              <span className="text-xs font-bold uppercase tracking-wider text-slate-300 flex items-center gap-1.5">
                <Layers className="w-4 h-4 text-blue-400" /> Slide Deck ({slides.length})
              </span>
              <div className="flex items-center gap-1">
                <button
                  onClick={() => addSlide('content')}
                  title="Add Standard Slide"
                  className="p-1.5 rounded-lg bg-blue-600/20 hover:bg-blue-600 text-blue-400 hover:text-white transition border border-blue-500/30"
                >
                  <Plus className="w-4 h-4" />
                </button>
              </div>
            </div>

            {/* SEARCH SLIDES INPUT */}
            <div className="relative">
              <Search className="w-3.5 h-3.5 absolute left-2.5 top-2.5 text-slate-500" />
              <input
                type="text"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                placeholder="Search slides..."
                className="w-full bg-slate-900/90 border border-slate-800 rounded-lg pl-8 pr-3 py-1.5 text-xs text-slate-200 placeholder-slate-500 focus:outline-none focus:border-blue-500"
              />
              {searchQuery && (
                <button
                  onClick={() => setSearchQuery('')}
                  className="absolute right-2 top-2 text-slate-500 hover:text-slate-300"
                >
                  <X className="w-3 h-3" />
                </button>
              )}
            </div>
          </div>

          {/* SLIDES LIST */}
          <div className="flex-1 overflow-y-auto p-3 space-y-3 custom-scrollbar">
            {filteredSlides.map((slide) => {
              const originalIdx = slides.findIndex((s) => s.id === slide.id);
              const th = THEMES[slide.theme] || THEMES.docuCraft;
              const isActive = originalIdx === activeSlideIndex;
              const customBg = slide.customBgColor || th.bg;

              return (
                <div
                  key={slide.id || originalIdx}
                  onClick={() => setActiveSlideIndex(originalIdx)}
                  className={`group relative rounded-xl p-2.5 transition cursor-pointer border ${
                    isActive
                      ? 'border-blue-500 bg-blue-950/30 shadow-lg shadow-blue-500/10'
                      : 'border-slate-800/80 bg-slate-900/40 hover:border-slate-700 hover:bg-slate-900/80'
                  }`}
                >
                  <div className="flex items-center justify-between mb-2">
                    <span className="text-[10px] font-mono font-bold px-1.5 py-0.5 rounded bg-slate-800 text-slate-400">
                      {(originalIdx + 1).toString().padStart(2, '0')}
                    </span>
                    <span className="text-[10px] uppercase font-semibold text-slate-500 group-hover:text-slate-400">
                      {slide.layout}
                    </span>
                  </div>

                  {/* THUMBNAIL PREVIEW CARD */}
                  <div
                    className="w-full aspect-[16/9] rounded-lg p-2.5 flex flex-col justify-between overflow-hidden relative border shadow-inner transition group-hover:scale-[1.02]"
                    style={{
                      backgroundColor: customBg,
                      borderColor: th.border,
                      color: th.textMain
                    }}
                  >
                    <div className="space-y-1">
                      {slide.tag && (
                        <div
                          className="w-7 h-1 rounded-full opacity-80"
                          style={{ backgroundColor: slide.customAccentColor || th.accent }}
                        ></div>
                      )}
                      <div
                        className="text-[9px] font-bold truncate leading-tight"
                        style={{ color: th.textMain }}
                      >
                        {slide.title || 'Untitled Slide'}
                      </div>
                      <div
                        className="text-[7px] truncate opacity-75"
                        style={{ color: th.textMuted }}
                      >
                        {slide.subtitle || ''}
                      </div>
                    </div>
                    <div
                      className="text-[6px] font-mono opacity-50 flex justify-between"
                      style={{ color: th.textMuted }}
                    >
                      <span>{th.name}</span>
                      <span>16:9</span>
                    </div>
                  </div>
                </div>
              );
            })}
          </div>

          {/* SLIDE ACTIONS */}
          <div className="p-3 border-t border-slate-800/80 grid grid-cols-4 gap-1.5 bg-slate-950/50">
            <button
              onClick={() => moveSlide(-1)}
              disabled={activeSlideIndex === 0}
              title="Move Slide Up"
              className="p-2 rounded-lg bg-slate-900 hover:bg-slate-800 text-slate-300 disabled:opacity-30 border border-slate-800 flex items-center justify-center"
            >
              <ChevronUp className="w-4 h-4" />
            </button>
            <button
              onClick={() => moveSlide(1)}
              disabled={activeSlideIndex === slides.length - 1}
              title="Move Slide Down"
              className="p-2 rounded-lg bg-slate-900 hover:bg-slate-800 text-slate-300 disabled:opacity-30 border border-slate-800 flex items-center justify-center"
            >
              <ChevronDown className="w-4 h-4" />
            </button>
            <button
              onClick={duplicateSlide}
              title="Duplicate Active Slide"
              className="p-2 rounded-lg bg-slate-900 hover:bg-slate-800 text-slate-300 border border-slate-800 flex items-center justify-center"
            >
              <Copy className="w-4 h-4" />
            </button>
            <button
              onClick={deleteSlide}
              disabled={slides.length <= 1}
              title="Delete Active Slide"
              className="p-2 rounded-lg bg-rose-950/40 hover:bg-rose-900/60 text-rose-400 disabled:opacity-30 border border-rose-900/50 flex items-center justify-center"
            >
              <Trash2 className="w-4 h-4" />
            </button>
          </div>
        </aside>

        {/* MAIN CANVAS WORKSPACE PREVIEW */}
        <main className="flex-1 bg-[#05070D] p-6 lg:p-8 flex flex-col items-center justify-center relative overflow-y-auto">
          {/* CANVAS DECOR & CONTROLS TOOLBAR */}
          <div className="mb-4 flex flex-wrap items-center gap-3 bg-slate-900/90 border border-slate-800 rounded-full px-4 py-1.5 shadow-2xl text-xs backdrop-blur z-10">
            <span className="text-slate-400 font-medium">Controls:</span>
            <label className="flex items-center gap-1.5 cursor-pointer text-slate-300 hover:text-white">
              <input
                type="checkbox"
                checked={showGridPattern}
                onChange={(e) => setShowGridPattern(e.target.checked)}
                className="rounded bg-slate-800 border-slate-700 text-blue-600 focus:ring-0"
              />
              <Grid className="w-3.5 h-3.5 text-blue-400" /> Grid
            </label>
            <div className="w-px h-3 bg-slate-800"></div>
            <label className="flex items-center gap-1.5 cursor-pointer text-slate-300 hover:text-white">
              <input
                type="checkbox"
                checked={showGlowEffect}
                onChange={(e) => setShowGlowEffect(e.target.checked)}
                className="rounded bg-slate-800 border-slate-700 text-blue-600 focus:ring-0"
              />
              <Sparkles className="w-3.5 h-3.5 text-amber-400" /> Glow
            </label>
            <div className="w-px h-3 bg-slate-800"></div>
            <label className="flex items-center gap-1.5 cursor-pointer text-slate-300 hover:text-white">
              <input
                type="checkbox"
                checked={showSlideNumber}
                onChange={(e) => setShowSlideNumber(e.target.checked)}
                className="rounded bg-slate-800 border-slate-700 text-blue-600 focus:ring-0"
              />
              <Sliders className="w-3.5 h-3.5 text-emerald-400" /> Badge
            </label>
            <div className="w-px h-3 bg-slate-800"></div>
            <label className="flex items-center gap-1.5 cursor-pointer text-slate-300 hover:text-white">
              <input
                type="checkbox"
                checked={showGlassCard}
                onChange={(e) => setShowGlassCard(e.target.checked)}
                className="rounded bg-slate-800 border-slate-700 text-blue-600 focus:ring-0"
              />
              <FileText className="w-3.5 h-3.5 text-purple-400" /> Glass Cards
            </label>

            <div className="w-px h-3 bg-slate-800"></div>
            <div className="flex items-center gap-1 bg-slate-950/60 p-0.5 rounded-full border border-slate-800">
              <button
                onClick={() => setAspectRatio('16:9')}
                className={`px-2 py-0.5 rounded-full text-[10px] font-mono transition ${
                  aspectRatio === '16:9' ? 'bg-blue-600 text-white font-bold' : 'text-slate-400 hover:text-slate-200'
                }`}
              >
                16:9
              </button>
              <button
                onClick={() => setAspectRatio('4:3')}
                className={`px-2 py-0.5 rounded-full text-[10px] font-mono transition ${
                  aspectRatio === '4:3' ? 'bg-blue-600 text-white font-bold' : 'text-slate-400 hover:text-slate-200'
                }`}
              >
                4:3
              </button>
            </div>
          </div>

          {/* SLIDE CANVAS CONTAINER */}
          <div
            className={`w-full max-w-4xl relative shadow-2xl rounded-2xl overflow-hidden transition-all duration-300 border border-slate-700/50 flex flex-col justify-between p-8 lg:p-12 ${
              aspectRatio === '4:3' ? 'aspect-[4/3]' : 'aspect-[16/9]'
            }`}
            style={{
              backgroundColor: activeSlide.customBgColor || activeTheme.bg,
              color: activeTheme.textMain,
              fontFamily:
                activeSlide.fontFamily === 'serif' || activeTheme.font === 'font-serif'
                  ? 'Georgia, serif'
                  : activeSlide.fontFamily === 'mono' || activeTheme.font === 'font-mono'
                  ? 'Courier New, monospace'
                  : 'system-ui, -apple-system, sans-serif'
            }}
          >
            {/* GRID PATTERN DECOR */}
            {showGridPattern && (
              <div
                className="absolute inset-0 pointer-events-none opacity-10"
                style={{
                  backgroundImage: `radial-gradient(${activeSlide.customAccentColor || activeTheme.accent} 1px, transparent 1px)`,
                  backgroundSize: '24px 24px'
                }}
              ></div>
            )}

            {/* GLOW GLOW DECOR */}
            {showGlowEffect && (
              <div
                className="absolute -top-24 -right-24 w-72 h-72 rounded-full blur-3xl pointer-events-none opacity-20"
                style={{ backgroundColor: activeSlide.customAccentColor || activeTheme.accent }}
              ></div>
            )}

            {/* SLIDE TOP SECTION: BADGE & SLIDE BADGE */}
            <div className="relative z-10 flex items-center justify-between">
              {activeSlide.tag ? (
                <div
                  className="inline-flex items-center gap-2 px-3 py-1 rounded-full text-xs font-mono font-bold tracking-wider uppercase border shadow-sm"
                  style={{
                    backgroundColor: `${activeSlide.customAccentColor || activeTheme.accent}15`,
                    color: activeSlide.customAccentColor || activeTheme.accentText,
                    borderColor: `${activeSlide.customAccentColor || activeTheme.accent}40`
                  }}
                >
                  <span
                    className="w-1.5 h-1.5 rounded-full"
                    style={{ backgroundColor: activeSlide.customAccentColor || activeTheme.accent }}
                  ></span>
                  {activeSlide.tag}
                </div>
              ) : (
                <div></div>
              )}

              {showSlideNumber && (
                <span
                  className="text-xs font-mono opacity-60 font-medium tracking-widest"
                  style={{ color: activeTheme.textMuted }}
                >
                  DOCUCRAFT // 0{activeSlideIndex + 1}
                </span>
              )}
            </div>

            {/* SLIDE MAIN CONTENT RENDERER */}
            <div className="relative z-10 my-auto py-4">
              {/* LAYOUT: TITLE */}
              {activeSlide.layout === 'title' && (
                <div className="space-y-4 max-w-3xl">
                  <h2
                    className="text-3xl lg:text-5xl font-black tracking-tight leading-tight"
                    style={{ color: activeTheme.textMain }}
                  >
                    {activeSlide.title || 'Slide Title Here'}
                  </h2>
                  <p
                    className="text-base lg:text-xl font-normal leading-relaxed opacity-90"
                    style={{ color: activeTheme.textMuted }}
                  >
                    {activeSlide.subtitle || 'Add narrative description or core message.'}
                  </p>
                </div>
              )}

              {/* LAYOUT: BULLET CONTENT */}
              {activeSlide.layout === 'content' && (
                <div className="space-y-6">
                  <div>
                    <h2
                      className="text-2xl lg:text-4xl font-bold tracking-tight"
                      style={{ color: activeTheme.textMain }}
                    >
                      {activeSlide.title}
                    </h2>
                    {activeSlide.subtitle && (
                      <p className="text-sm lg:text-base mt-1" style={{ color: activeTheme.textMuted }}>
                        {activeSlide.subtitle}
                      </p>
                    )}
                  </div>

                  {activeSlide.bullets && (
                    <ul className="space-y-3 max-w-2xl">
                      {activeSlide.bullets.map((bullet, i) => (
                        <li key={i} className="flex items-start gap-3 text-sm lg:text-base font-medium">
                          <CheckCircle2
                            className="w-5 h-5 flex-shrink-0 mt-0.5"
                            style={{ color: activeSlide.customAccentColor || activeTheme.accent }}
                          />
                          <span style={{ color: activeTheme.textMain }}>{bullet}</span>
                        </li>
                      ))}
                    </ul>
                  )}
                </div>
              )}

              {/* LAYOUT: SPLIT COLUMNS */}
              {activeSlide.layout === 'split' && (
                <div className="space-y-6">
                  <div>
                    <h2 className="text-2xl lg:text-4xl font-bold tracking-tight" style={{ color: activeTheme.textMain }}>
                      {activeSlide.title}
                    </h2>
                    {activeSlide.subtitle && (
                      <p className="text-sm lg:text-base mt-1" style={{ color: activeTheme.textMuted }}>
                        {activeSlide.subtitle}
                      </p>
                    )}
                  </div>

                  <div className="grid grid-cols-2 gap-6">
                    <div
                      className={`p-4 rounded-xl border ${
                        showGlassCard ? 'backdrop-blur-md bg-opacity-70 shadow-lg' : ''
                      }`}
                      style={{
                        backgroundColor: activeTheme.cardBg,
                        borderColor: activeTheme.border
                      }}
                    >
                      <h3 className="text-base font-bold mb-1" style={{ color: activeSlide.customAccentColor || activeTheme.accentText }}>
                        {activeSlide.col1Title || 'Column 1'}
                      </h3>
                      <p className="text-xs lg:text-sm leading-relaxed" style={{ color: activeTheme.textMuted }}>
                        {activeSlide.col1Text || 'Details for left topic.'}
                      </p>
                    </div>

                    <div
                      className={`p-4 rounded-xl border ${
                        showGlassCard ? 'backdrop-blur-md bg-opacity-70 shadow-lg' : ''
                      }`}
                      style={{
                        backgroundColor: activeTheme.cardBg,
                        borderColor: activeTheme.border
                      }}
                    >
                      <h3 className="text-base font-bold mb-1" style={{ color: activeSlide.customAccentColor || activeTheme.accentText }}>
                        {activeSlide.col2Title || 'Column 2'}
                      </h3>
                      <p className="text-xs lg:text-sm leading-relaxed" style={{ color: activeTheme.textMuted }}>
                        {activeSlide.col2Text || 'Details for right topic.'}
                      </p>
                    </div>
                  </div>
                </div>
              )}

              {/* LAYOUT: BIG STAT METRIC */}
              {activeSlide.layout === 'stat' && (
                <div className="grid grid-cols-1 lg:grid-cols-2 gap-8 items-center">
                  <div className="space-y-2">
                    <div
                      className="text-5xl lg:text-7xl font-black tracking-tight"
                      style={{ color: activeSlide.customAccentColor || activeTheme.accent }}
                    >
                      {activeSlide.statValue || '$0'}
                    </div>
                    <div className="text-base font-bold" style={{ color: activeTheme.textMain }}>
                      {activeSlide.statLabel || 'Metric Label'}
                    </div>
                    <div className="text-xs lg:text-sm" style={{ color: activeTheme.textMuted }}>
                      {activeSlide.statSubtext || 'Metric detailed subtext explaining growth or traction.'}
                    </div>
                  </div>

                  {activeSlide.bullets && (
                    <div
                      className={`p-5 rounded-2xl border space-y-3 ${
                        showGlassCard ? 'backdrop-blur-md bg-opacity-70 shadow-lg' : ''
                      }`}
                      style={{
                        backgroundColor: activeTheme.cardBg,
                        borderColor: activeTheme.border
                      }}
                    >
                      <h4
                        className="text-xs font-bold uppercase tracking-wider"
                        style={{ color: activeSlide.customAccentColor || activeTheme.accentText }}
                      >
                        Highlights
                      </h4>
                      <ul className="space-y-2 text-xs lg:text-sm">
                        {activeSlide.bullets.map((b, i) => (
                          <li key={i} className="flex items-center gap-2" style={{ color: activeTheme.textMain }}>
                            <div
                              className="w-1.5 h-1.5 rounded-full"
                              style={{ backgroundColor: activeSlide.customAccentColor || activeTheme.accent }}
                            ></div>
                            {b}
                          </li>
                        ))}
                      </ul>
                    </div>
                  )}
                </div>
              )}

              {/* LAYOUT: QUOTE */}
              {activeSlide.layout === 'quote' && (
                <div className="space-y-6 max-w-2xl mx-auto text-center">
                  <QuoteIcon
                    className="w-10 h-10 mx-auto opacity-30"
                    style={{ color: activeSlide.customAccentColor || activeTheme.accent }}
                  />
                  <blockquote
                    className="text-xl lg:text-3xl font-medium italic leading-relaxed"
                    style={{ color: activeTheme.textMain }}
                  >
                    “{activeSlide.quote || 'Insert quote or testimonial statement.'}”
                  </blockquote>
                  <div>
                    <div
                      className="text-sm font-bold"
                      style={{ color: activeSlide.customAccentColor || activeTheme.accentText }}
                    >
                      {activeSlide.quoteAuthor || 'Author Name'}
                    </div>
                    <div className="text-xs opacity-75" style={{ color: activeTheme.textMuted }}>
                      {activeSlide.quoteRole || 'Title, Organization'}
                    </div>
                  </div>
                </div>
              )}

              {/* LAYOUT: TIMELINE STEPS */}
              {activeSlide.layout === 'timeline' && (
                <div className="space-y-6">
                  <div>
                    <h2 className="text-2xl lg:text-4xl font-bold tracking-tight" style={{ color: activeTheme.textMain }}>
                      {activeSlide.title}
                    </h2>
                    {activeSlide.subtitle && (
                      <p className="text-sm lg:text-base mt-1" style={{ color: activeTheme.textMuted }}>
                        {activeSlide.subtitle}
                      </p>
                    )}
                  </div>

                  <div className="grid grid-cols-3 gap-4">
                    {(activeSlide.timelineSteps || []).map((st, i) => (
                      <div
                        key={i}
                        className={`p-4 rounded-xl border relative space-y-2 ${
                          showGlassCard ? 'backdrop-blur-md bg-opacity-70 shadow-lg' : ''
                        }`}
                        style={{
                          backgroundColor: activeTheme.cardBg,
                          borderColor: activeTheme.border
                        }}
                      >
                        <div
                          className="text-lg font-mono font-bold"
                          style={{ color: activeSlide.customAccentColor || activeTheme.accent }}
                        >
                          {st.step || `0${i + 1}`}
                        </div>
                        <h4 className="text-sm font-bold" style={{ color: activeTheme.textMain }}>
                          {st.title || 'Step Title'}
                        </h4>
                        <p className="text-xs leading-relaxed" style={{ color: activeTheme.textMuted }}>
                          {st.desc || 'Step description text.'}
                        </p>
                      </div>
                    ))}
                  </div>
                </div>
              )}

              {/* LAYOUT: COMPARISON MATRIX */}
              {activeSlide.layout === 'comparison' && (
                <div className="space-y-6">
                  <div>
                    <h2 className="text-2xl lg:text-4xl font-bold tracking-tight" style={{ color: activeTheme.textMain }}>
                      {activeSlide.title}
                    </h2>
                    {activeSlide.subtitle && (
                      <p className="text-sm lg:text-base mt-1" style={{ color: activeTheme.textMuted }}>
                        {activeSlide.subtitle}
                      </p>
                    )}
                  </div>

                  <div className="grid grid-cols-2 gap-6">
                    <div
                      className="p-5 rounded-xl border space-y-2"
                      style={{
                        backgroundColor: activeTheme.cardBg,
                        borderColor: activeTheme.border
                      }}
                    >
                      <div className="flex items-center gap-2">
                        <X className="w-4 h-4 text-rose-400" />
                        <h3 className="text-sm font-bold text-rose-400">
                          {activeSlide.compTitle1 || 'Option A'}
                        </h3>
                      </div>
                      <p className="text-xs lg:text-sm leading-relaxed" style={{ color: activeTheme.textMuted }}>
                        {activeSlide.compText1 || 'Description of option A.'}
                      </p>
                    </div>

                    <div
                      className="p-5 rounded-xl border space-y-2"
                      style={{
                        backgroundColor: activeTheme.cardBg,
                        borderColor: activeSlide.customAccentColor || activeTheme.accent
                      }}
                    >
                      <div className="flex items-center gap-2">
                        <Check className="w-4 h-4" style={{ color: activeSlide.customAccentColor || activeTheme.accent }} />
                        <h3
                          className="text-sm font-bold"
                          style={{ color: activeSlide.customAccentColor || activeTheme.accentText }}
                        >
                          {activeSlide.compTitle2 || 'Option B'}
                        </h3>
                      </div>
                      <p className="text-xs lg:text-sm leading-relaxed" style={{ color: activeTheme.textMuted }}>
                        {activeSlide.compText2 || 'Description of option B.'}
                      </p>
                    </div>
                  </div>
                </div>
              )}

              {/* LAYOUT: 2X2 FEATURE GRID */}
              {activeSlide.layout === 'grid' && (
                <div className="space-y-6">
                  <div>
                    <h2 className="text-2xl lg:text-4xl font-bold tracking-tight" style={{ color: activeTheme.textMain }}>
                      {activeSlide.title}
                    </h2>
                    {activeSlide.subtitle && (
                      <p className="text-sm lg:text-base mt-1" style={{ color: activeTheme.textMuted }}>
                        {activeSlide.subtitle}
                      </p>
                    )}
                  </div>

                  <div className="grid grid-cols-2 gap-4">
                    {(activeSlide.gridItems || []).map((item, idx) => (
                      <div
                        key={idx}
                        className={`p-4 rounded-xl border space-y-1.5 ${
                          showGlassCard ? 'backdrop-blur-md bg-opacity-70 shadow-lg' : ''
                        }`}
                        style={{
                          backgroundColor: activeTheme.cardBg,
                          borderColor: activeTheme.border
                        }}
                      >
                        <div
                          className="text-xs font-bold uppercase tracking-wider flex items-center gap-1.5"
                          style={{ color: activeSlide.customAccentColor || activeTheme.accentText }}
                        >
                          <Sparkles className="w-3.5 h-3.5" />
                          {item.title || `Capability 0${idx + 1}`}
                        </div>
                        <p className="text-xs leading-relaxed" style={{ color: activeTheme.textMuted }}>
                          {item.desc || 'Feature or capability details explanation.'}
                        </p>
                      </div>
                    ))}
                  </div>
                </div>
              )}

              {/* LAYOUT: TEAM CARDS */}
              {activeSlide.layout === 'team' && (
                <div className="space-y-6">
                  <div>
                    <h2 className="text-2xl lg:text-4xl font-bold tracking-tight" style={{ color: activeTheme.textMain }}>
                      {activeSlide.title}
                    </h2>
                    {activeSlide.subtitle && (
                      <p className="text-sm lg:text-base mt-1" style={{ color: activeTheme.textMuted }}>
                        {activeSlide.subtitle}
                      </p>
                    )}
                  </div>

                  <div className="grid grid-cols-3 gap-4">
                    {(activeSlide.teamMembers || []).map((m, idx) => (
                      <div
                        key={idx}
                        className={`p-4 rounded-xl border space-y-2 text-center ${
                          showGlassCard ? 'backdrop-blur-md bg-opacity-70 shadow-lg' : ''
                        }`}
                        style={{
                          backgroundColor: activeTheme.cardBg,
                          borderColor: activeTheme.border
                        }}
                      >
                        <div
                          className="w-10 h-10 mx-auto rounded-full flex items-center justify-center font-bold text-sm shadow-md"
                          style={{
                            backgroundColor: `${activeSlide.customAccentColor || activeTheme.accent}25`,
                            color: activeSlide.customAccentColor || activeTheme.accentText,
                            border: `1px solid ${activeSlide.customAccentColor || activeTheme.accent}50`
                          }}
                        >
                          <Users className="w-5 h-5" />
                        </div>
                        <div>
                          <h4 className="text-sm font-bold" style={{ color: activeTheme.textMain }}>
                            {m.name || 'Member Name'}
                          </h4>
                          <span
                            className="text-xs font-medium"
                            style={{ color: activeSlide.customAccentColor || activeTheme.accentText }}
                          >
                            {m.role || 'Title / Role'}
                          </span>
                        </div>
                        <p className="text-[11px] leading-snug" style={{ color: activeTheme.textMuted }}>
                          {m.bio || 'Short background summary.'}
                        </p>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>

            {/* SLIDE FOOTER */}
            <div
              className="relative z-10 pt-4 border-t flex items-center justify-between text-[11px] font-mono"
              style={{
                borderColor: `${activeTheme.border}80`,
                color: activeTheme.textMuted
              }}
            >
              <span>{activeSlide.footerText || 'DocuCraft Studio Presentation'}</span>
              <span className="flex items-center gap-1.5">
                <span
                  className="w-2 h-2 rounded-full animate-ping"
                  style={{ backgroundColor: activeSlide.customAccentColor || activeTheme.accent }}
                ></span>
                {activeTheme.name}
              </span>
            </div>
          </div>
        </main>

        {/* RIGHT SIDEBAR: REAL-TIME SLIDE INSPECTOR & EDITOR */}
        <aside className="w-80 border-l border-slate-800/80 bg-[#080B12] flex flex-col z-10 overflow-y-auto custom-scrollbar">
          <div className="p-4 border-b border-slate-800/80 flex items-center justify-between bg-slate-950/40">
            <span className="text-xs font-bold uppercase tracking-wider text-slate-300 flex items-center gap-1.5">
              <Sliders className="w-4 h-4 text-purple-400" /> Slide Customizer
            </span>
            <span className="text-[10px] bg-slate-800 px-2 py-0.5 rounded text-slate-400 font-mono">
              #{activeSlideIndex + 1}
            </span>
          </div>

          <div className="p-4 space-y-6">
            {/* THEME PRESET SELECTOR */}
            <div className="space-y-2">
              <label className="text-xs font-bold text-slate-300 flex items-center justify-between">
                <span className="flex items-center gap-1.5">
                  <Palette className="w-3.5 h-3.5 text-blue-400" /> Color Aesthetics
                </span>
                <span className="text-[10px] text-slate-500 font-mono">9 Themes</span>
              </label>

              <div className="grid grid-cols-1 gap-1.5 max-h-48 overflow-y-auto custom-scrollbar pr-1">
                {Object.values(THEMES).map((th) => (
                  <button
                    key={th.id}
                    onClick={() => updateActiveSlide({ theme: th.id })}
                    className={`flex items-center justify-between p-2 rounded-xl border text-left text-xs transition ${
                      activeSlide.theme === th.id
                        ? 'border-blue-500 bg-blue-950/40 text-white font-semibold shadow-sm'
                        : 'border-slate-800/80 bg-slate-900/40 text-slate-400 hover:border-slate-700 hover:text-slate-200'
                    }`}
                  >
                    <div className="flex items-center gap-2">
                      <div
                        className="w-3.5 h-3.5 rounded-full border border-white/20 flex-shrink-0"
                        style={{ backgroundColor: th.accent }}
                      ></div>
                      <span className="truncate">{th.name}</span>
                    </div>
                    <span className="text-[9px] font-mono opacity-60 uppercase">{th.category}</span>
                  </button>
                ))}
              </div>
            </div>

            {/* LAYOUT SELECTOR */}
            <div className="space-y-2">
              <label className="text-xs font-bold text-slate-300 flex items-center gap-1.5">
                <Layout className="w-3.5 h-3.5 text-purple-400" /> Layout Style
              </label>
              <div className="grid grid-cols-3 gap-1.5 text-xs">
                {[
                  { id: 'title', label: 'Title' },
                  { id: 'content', label: 'Bullets' },
                  { id: 'split', label: 'Split' },
                  { id: 'stat', label: 'Metric' },
                  { id: 'quote', label: 'Quote' },
                  { id: 'timeline', label: 'Roadmap' },
                  { id: 'comparison', label: 'Matrix' },
                  { id: 'grid', label: 'Grid' },
                  { id: 'team', label: 'Team' }
                ].map((ly) => (
                  <button
                    key={ly.id}
                    onClick={() => updateActiveSlide({ layout: ly.id })}
                    className={`py-1.5 px-1 rounded-lg border text-center transition text-[11px] font-medium ${
                      activeSlide.layout === ly.id
                        ? 'border-purple-500 bg-purple-950/40 text-purple-300'
                        : 'border-slate-800 bg-slate-900/40 text-slate-400 hover:border-slate-700 hover:text-slate-200'
                    }`}
                  >
                    {ly.label}
                  </button>
                ))}
              </div>
            </div>

            {/* FONT FAMILY & COLOR OVERRIDES */}
            <div className="space-y-3 pt-1 border-t border-slate-800/80">
              <label className="text-xs font-bold text-slate-300">Typography & Color Overrides</label>
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="text-[10px] text-slate-400 block mb-1">Font Family</label>
                  <select
                    value={activeSlide.fontFamily || 'default'}
                    onChange={(e) => updateActiveSlide({ fontFamily: e.target.value })}
                    className="w-full bg-slate-900 border border-slate-800 rounded-lg p-1.5 text-xs text-slate-200"
                  >
                    <option value="default">Theme Default</option>
                    <option value="sans">Modern Sans</option>
                    <option value="serif">Warm Serif</option>
                    <option value="mono">Brutalist Mono</option>
                  </select>
                </div>

                <div>
                  <label className="text-[10px] text-slate-400 block mb-1">Accent Override</label>
                  <input
                    type="color"
                    value={activeSlide.customAccentColor || activeTheme.accent}
                    onChange={(e) => updateActiveSlide({ customAccentColor: e.target.value })}
                    className="w-full h-7 bg-slate-900 border border-slate-800 rounded-lg cursor-pointer p-0.5"
                  />
                </div>
              </div>
            </div>

            <hr className="border-slate-800/80" />

            {/* FIELD EDITORS */}
            <div className="space-y-3">
              <div className="space-y-1">
                <label className="text-xs font-medium text-slate-400">Category Tag / Badge</label>
                <input
                  type="text"
                  value={activeSlide.tag || ''}
                  onChange={(e) => updateActiveSlide({ tag: e.target.value })}
                  placeholder="e.g. STRATEGIC ROADMAP"
                  className="w-full bg-slate-900 border border-slate-800 rounded-lg px-3 py-1.5 text-xs text-slate-100 focus:outline-none focus:border-blue-500"
                />
              </div>

              <div className="space-y-1">
                <label className="text-xs font-medium text-slate-400">Main Title</label>
                <input
                  type="text"
                  value={activeSlide.title || ''}
                  onChange={(e) => updateActiveSlide({ title: e.target.value })}
                  placeholder="Headline..."
                  className="w-full bg-slate-900 border border-slate-800 rounded-lg px-3 py-1.5 text-xs text-slate-100 focus:outline-none focus:border-blue-500"
                />
              </div>

              <div className="space-y-1">
                <label className="text-xs font-medium text-slate-400">Subtitle / Narrative</label>
                <textarea
                  rows={2}
                  value={activeSlide.subtitle || ''}
                  onChange={(e) => updateActiveSlide({ subtitle: e.target.value })}
                  placeholder="Subtitle text..."
                  className="w-full bg-slate-900 border border-slate-800 rounded-lg px-3 py-1.5 text-xs text-slate-100 focus:outline-none focus:border-blue-500 resize-none"
                />
              </div>

              {/* LAYOUT EDITORS: BULLETS */}
              {(activeSlide.layout === 'content' || activeSlide.layout === 'stat') && (
                <div className="space-y-2 pt-1">
                  <div className="flex items-center justify-between">
                    <label className="text-xs font-medium text-slate-400">Bullet Highlights</label>
                    <button
                      onClick={addBullet}
                      className="text-[11px] text-blue-400 hover:text-blue-300 flex items-center gap-1 font-medium"
                    >
                      <Plus className="w-3 h-3" /> Add Point
                    </button>
                  </div>
                  <div className="space-y-1.5">
                    {(activeSlide.bullets || []).map((bullet, idx) => (
                      <div key={idx} className="flex items-center gap-1.5">
                        <input
                          type="text"
                          value={bullet}
                          onChange={(e) => updateBullet(idx, e.target.value)}
                          className="flex-1 bg-slate-900 border border-slate-800 rounded-lg px-2.5 py-1 text-xs text-slate-100 focus:outline-none focus:border-blue-500"
                        />
                        <button
                          onClick={() => removeBullet(idx)}
                          className="p-1 text-slate-500 hover:text-rose-400"
                        >
                          <X className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    ))}
                  </div>
                </div>
              )}

              {/* LAYOUT EDITORS: TIMELINE STEPS */}
              {activeSlide.layout === 'timeline' && (
                <div className="space-y-3 pt-1">
                  <label className="text-xs font-medium text-slate-400">Roadmap Milestones</label>
                  {(activeSlide.timelineSteps || []).map((st, idx) => (
                    <div key={idx} className="p-2.5 bg-slate-900 border border-slate-800 rounded-lg space-y-1.5">
                      <div className="flex gap-2">
                        <input
                          type="text"
                          value={st.step || ''}
                          onChange={(e) => updateTimelineStep(idx, 'step', e.target.value)}
                          placeholder="01"
                          className="w-12 bg-slate-950 border border-slate-800 rounded px-2 py-1 text-xs text-blue-400 font-mono"
                        />
                        <input
                          type="text"
                          value={st.title || ''}
                          onChange={(e) => updateTimelineStep(idx, 'title', e.target.value)}
                          placeholder="Title"
                          className="flex-1 bg-slate-950 border border-slate-800 rounded px-2 py-1 text-xs text-slate-100"
                        />
                      </div>
                      <input
                        type="text"
                        value={st.desc || ''}
                        onChange={(e) => updateTimelineStep(idx, 'desc', e.target.value)}
                        placeholder="Description..."
                        className="w-full bg-slate-950 border border-slate-800 rounded px-2 py-1 text-xs text-slate-300"
                      />
                    </div>
                  ))}
                </div>
              )}

              {/* LAYOUT EDITORS: COMPARISON MATRIX */}
              {activeSlide.layout === 'comparison' && (
                <div className="space-y-3 pt-1">
                  <div className="space-y-1">
                    <label className="text-xs font-medium text-slate-400">Option 1 Header & Text</label>
                    <input
                      type="text"
                      value={activeSlide.compTitle1 || ''}
                      onChange={(e) => updateActiveSlide({ compTitle1: e.target.value })}
                      className="w-full bg-slate-900 border border-slate-800 rounded-lg px-2.5 py-1 text-xs text-slate-100"
                    />
                    <textarea
                      rows={2}
                      value={activeSlide.compText1 || ''}
                      onChange={(e) => updateActiveSlide({ compText1: e.target.value })}
                      className="w-full bg-slate-900 border border-slate-800 rounded-lg px-2.5 py-1 text-xs text-slate-300 resize-none"
                    />
                  </div>

                  <div className="space-y-1">
                    <label className="text-xs font-medium text-slate-400">Option 2 Header & Text</label>
                    <input
                      type="text"
                      value={activeSlide.compTitle2 || ''}
                      onChange={(e) => updateActiveSlide({ compTitle2: e.target.value })}
                      className="w-full bg-slate-900 border border-slate-800 rounded-lg px-2.5 py-1 text-xs text-slate-100"
                    />
                    <textarea
                      rows={2}
                      value={activeSlide.compText2 || ''}
                      onChange={(e) => updateActiveSlide({ compText2: e.target.value })}
                      className="w-full bg-slate-900 border border-slate-800 rounded-lg px-2.5 py-1 text-xs text-slate-300 resize-none"
                    />
                  </div>
                </div>
              )}

              {/* LAYOUT EDITORS: 2X2 GRID */}
              {activeSlide.layout === 'grid' && (
                <div className="space-y-2 pt-1">
                  <label className="text-xs font-medium text-slate-400">Grid Cards</label>
                  {(activeSlide.gridItems || []).map((item, idx) => (
                    <div key={idx} className="p-2 bg-slate-900 border border-slate-800 rounded-lg space-y-1">
                      <input
                        type="text"
                        value={item.title || ''}
                        onChange={(e) => updateGridItem(idx, 'title', e.target.value)}
                        placeholder="Card Title"
                        className="w-full bg-slate-950 border border-slate-800 rounded px-2 py-1 text-xs text-slate-100 font-bold"
                      />
                      <input
                        type="text"
                        value={item.desc || ''}
                        onChange={(e) => updateGridItem(idx, 'desc', e.target.value)}
                        placeholder="Card description..."
                        className="w-full bg-slate-950 border border-slate-800 rounded px-2 py-1 text-xs text-slate-300"
                      />
                    </div>
                  ))}
                </div>
              )}

              {/* LAYOUT EDITORS: TEAM MEMBERS */}
              {activeSlide.layout === 'team' && (
                <div className="space-y-2 pt-1">
                  <label className="text-xs font-medium text-slate-400">Team Profiles</label>
                  {(activeSlide.teamMembers || []).map((m, idx) => (
                    <div key={idx} className="p-2 bg-slate-900 border border-slate-800 rounded-lg space-y-1">
                      <div className="flex gap-1.5">
                        <input
                          type="text"
                          value={m.name || ''}
                          onChange={(e) => updateTeamMember(idx, 'name', e.target.value)}
                          placeholder="Name"
                          className="flex-1 bg-slate-950 border border-slate-800 rounded px-2 py-1 text-xs text-slate-100 font-bold"
                        />
                        <input
                          type="text"
                          value={m.role || ''}
                          onChange={(e) => updateTeamMember(idx, 'role', e.target.value)}
                          placeholder="Role"
                          className="flex-1 bg-slate-950 border border-slate-800 rounded px-2 py-1 text-xs text-blue-400"
                        />
                      </div>
                      <input
                        type="text"
                        value={m.bio || ''}
                        onChange={(e) => updateTeamMember(idx, 'bio', e.target.value)}
                        placeholder="Bio summary..."
                        className="w-full bg-slate-950 border border-slate-800 rounded px-2 py-1 text-xs text-slate-300"
                      />
                    </div>
                  ))}
                </div>
              )}

              {/* FOOTER NOTE */}
              <div className="space-y-1 pt-2">
                <label className="text-xs font-medium text-slate-400">Footer Text</label>
                <input
                  type="text"
                  value={activeSlide.footerText || ''}
                  onChange={(e) => updateActiveSlide({ footerText: e.target.value })}
                  placeholder="DocuCraft Studio disclaimer..."
                  className="w-full bg-slate-900 border border-slate-800 rounded-lg px-3 py-1.5 text-xs text-slate-100"
                />
              </div>
            </div>
          </div>
        </aside>
      </div>

      {/* QUICK SLIDE TEMPLATE SELECTION MODAL */}
      {showQuickModal && (
        <div className="fixed inset-0 z-50 bg-black/80 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-[#0B0F19] border border-slate-800 rounded-2xl max-w-lg w-full p-6 space-y-5 shadow-2xl relative">
            <div className="flex items-center justify-between border-b border-slate-800/80 pb-3">
              <div className="flex items-center gap-2">
                <Wand2 className="w-5 h-5 text-amber-400" />
                <h3 className="text-base font-bold text-white">Add New Slide Layout</h3>
              </div>
              <button
                onClick={() => setShowQuickModal(false)}
                className="p-1 rounded-lg text-slate-400 hover:text-white hover:bg-slate-800"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <p className="text-xs text-slate-400">
              Select a pre-designed layout component to instantly insert into your deck:
            </p>

            <div className="grid grid-cols-2 gap-2.5">
              {[
                { type: 'content', title: 'Bullets & Points', desc: 'Standard key takeaways list' },
                { type: 'split', title: 'Split Columns', desc: 'Side-by-side comparison cards' },
                { type: 'timeline', title: 'Roadmap Timeline', desc: '3-stage milestone sequence' },
                { type: 'stat', title: 'Big Metric Callout', desc: 'Highlighted stat & key facts' },
                { type: 'comparison', title: 'Matrix Comparison', desc: 'Pros vs Cons or Option A vs B' },
                { type: 'grid', title: '2x2 Capability Grid', desc: 'Four feature cards grid' },
                { type: 'team', title: 'Team Showcase', desc: 'Speaker & leadership cards' },
                { type: 'quote', title: 'Quote & Endorsement', desc: 'High-impact customer quote' }
              ].map((item) => (
                <button
                  key={item.type}
                  onClick={() => {
                    addSlide(item.type);
                    setShowQuickModal(false);
                  }}
                  className="p-3 bg-slate-900/80 hover:bg-slate-800/90 border border-slate-800 hover:border-blue-500 rounded-xl text-left transition space-y-1 group"
                >
                  <div className="text-xs font-bold text-slate-200 group-hover:text-blue-400 flex items-center justify-between">
                    <span>{item.title}</span>
                    <Plus className="w-3.5 h-3.5 opacity-0 group-hover:opacity-100 transition" />
                  </div>
                  <div className="text-[10px] text-slate-400 leading-tight">{item.desc}</div>
                </button>
              ))}
            </div>
          </div>
        </div>
      )}

      {/* FULLSCREEN PRESENTATION MODE OVERLAY */}
      {isPresenting && (
        <div className="fixed inset-0 z-50 bg-black flex flex-col justify-between p-6 lg:p-12 animate-fade-in">
          {/* PRESENTATION TOP HEADER BAR */}
          <div className="flex items-center justify-between z-20 text-slate-300">
            <div className="flex items-center gap-3">
              <span className="text-xs font-mono bg-white/10 px-3 py-1 rounded-full flex items-center gap-2">
                <Clock className="w-3.5 h-3.5 text-emerald-400" />
                <span>{formatTime(presenterTime)}</span>
              </span>
              <span className="text-xs font-mono text-slate-400">
                SLIDE {activeSlideIndex + 1} OF {slides.length}
              </span>
            </div>

            <button
              onClick={() => setIsPresenting(false)}
              className="p-2 rounded-full bg-white/10 hover:bg-white/20 text-white transition flex items-center gap-1.5 text-xs px-3"
            >
              <X className="w-4 h-4" />
              <span>Exit (ESC)</span>
            </button>
          </div>

          {/* PRESENTATION SLIDE CENTER STAGE */}
          <div
            className={`w-full max-w-6xl mx-auto relative rounded-3xl p-10 lg:p-16 flex flex-col justify-between shadow-2xl overflow-hidden border border-white/10 ${
              aspectRatio === '4:3' ? 'aspect-[4/3]' : 'aspect-[16/9]'
            }`}
            style={{
              backgroundColor: activeSlide.customBgColor || activeTheme.bg,
              color: activeTheme.textMain,
              fontFamily:
                activeSlide.fontFamily === 'serif' || activeTheme.font === 'font-serif'
                  ? 'Georgia, serif'
                  : activeSlide.fontFamily === 'mono' || activeTheme.font === 'font-mono'
                  ? 'Courier New, monospace'
                  : 'system-ui, -apple-system, sans-serif'
            }}
          >
            {showGridPattern && (
              <div
                className="absolute inset-0 pointer-events-none opacity-10"
                style={{
                  backgroundImage: `radial-gradient(${activeSlide.customAccentColor || activeTheme.accent} 1px, transparent 1px)`,
                  backgroundSize: '32px 32px'
                }}
              ></div>
            )}

            {/* Slide Header Tag */}
            <div className="flex items-center justify-between">
              {activeSlide.tag && (
                <div
                  className="inline-flex items-center gap-2 px-4 py-1.5 rounded-full text-sm font-mono font-bold tracking-wider uppercase border shadow-sm"
                  style={{
                    backgroundColor: `${activeSlide.customAccentColor || activeTheme.accent}20`,
                    color: activeSlide.customAccentColor || activeTheme.accentText,
                    borderColor: `${activeSlide.customAccentColor || activeTheme.accent}50`
                  }}
                >
                  <span
                    className="w-2 h-2 rounded-full"
                    style={{ backgroundColor: activeSlide.customAccentColor || activeTheme.accent }}
                  ></span>
                  {activeSlide.tag}
                </div>
              )}
            </div>

            {/* Slide Content Body */}
            <div className="my-auto py-6">
              {activeSlide.layout === 'title' && (
                <div className="space-y-6 max-w-4xl">
                  <h1 className="text-4xl lg:text-6xl font-black tracking-tight leading-tight">
                    {activeSlide.title}
                  </h1>
                  <p className="text-xl lg:text-2xl font-normal leading-relaxed opacity-90" style={{ color: activeTheme.textMuted }}>
                    {activeSlide.subtitle}
                  </p>
                </div>
              )}

              {activeSlide.layout === 'content' && (
                <div className="space-y-8">
                  <div>
                    <h2 className="text-3xl lg:text-5xl font-bold">{activeSlide.title}</h2>
                    {activeSlide.subtitle && (
                      <p className="text-lg lg:text-xl mt-2" style={{ color: activeTheme.textMuted }}>
                        {activeSlide.subtitle}
                      </p>
                    )}
                  </div>
                  {activeSlide.bullets && (
                    <ul className="space-y-4 max-w-3xl">
                      {activeSlide.bullets.map((b, i) => (
                        <li key={i} className="flex items-center gap-4 text-lg lg:text-2xl font-medium">
                          <CheckCircle2
                            className="w-7 h-7 flex-shrink-0"
                            style={{ color: activeSlide.customAccentColor || activeTheme.accent }}
                          />
                          <span>{b}</span>
                        </li>
                      ))}
                    </ul>
                  )}
                </div>
              )}

              {activeSlide.layout === 'split' && (
                <div className="space-y-8">
                  <h2 className="text-3xl lg:text-5xl font-bold">{activeSlide.title}</h2>
                  <div className="grid grid-cols-2 gap-8">
                    <div
                      className="p-8 rounded-2xl border"
                      style={{ backgroundColor: activeTheme.cardBg, borderColor: activeTheme.border }}
                    >
                      <h3
                        className="text-2xl font-bold mb-3"
                        style={{ color: activeSlide.customAccentColor || activeTheme.accentText }}
                      >
                        {activeSlide.col1Title}
                      </h3>
                      <p className="text-base lg:text-lg" style={{ color: activeTheme.textMuted }}>
                        {activeSlide.col1Text}
                      </p>
                    </div>
                    <div
                      className="p-8 rounded-2xl border"
                      style={{ backgroundColor: activeTheme.cardBg, borderColor: activeTheme.border }}
                    >
                      <h3
                        className="text-2xl font-bold mb-3"
                        style={{ color: activeSlide.customAccentColor || activeTheme.accentText }}
                      >
                        {activeSlide.col2Title}
                      </h3>
                      <p className="text-base lg:text-lg" style={{ color: activeTheme.textMuted }}>
                        {activeSlide.col2Text}
                      </p>
                    </div>
                  </div>
                </div>
              )}

              {activeSlide.layout === 'stat' && (
                <div className="grid grid-cols-2 gap-12 items-center">
                  <div className="space-y-4">
                    <div
                      className="text-6xl lg:text-8xl font-black"
                      style={{ color: activeSlide.customAccentColor || activeTheme.accent }}
                    >
                      {activeSlide.statValue}
                    </div>
                    <div className="text-2xl font-bold">{activeSlide.statLabel}</div>
                    <div className="text-base" style={{ color: activeTheme.textMuted }}>
                      {activeSlide.statSubtext}
                    </div>
                  </div>
                  {activeSlide.bullets && (
                    <div
                      className="p-8 rounded-2xl border space-y-4"
                      style={{ backgroundColor: activeTheme.cardBg, borderColor: activeTheme.border }}
                    >
                      <h4
                        className="text-sm font-bold uppercase tracking-wider"
                        style={{ color: activeSlide.customAccentColor || activeTheme.accentText }}
                      >
                        Highlights
                      </h4>
                      <ul className="space-y-3 text-lg">
                        {activeSlide.bullets.map((b, i) => (
                          <li key={i} className="flex items-center gap-3">
                            <div
                              className="w-2.5 h-2.5 rounded-full"
                              style={{ backgroundColor: activeSlide.customAccentColor || activeTheme.accent }}
                            ></div>
                            {b}
                          </li>
                        ))}
                      </ul>
                    </div>
                  )}
                </div>
              )}

              {activeSlide.layout === 'quote' && (
                <div className="space-y-8 max-w-3xl mx-auto text-center">
                  <QuoteIcon
                    className="w-16 h-16 mx-auto opacity-30"
                    style={{ color: activeSlide.customAccentColor || activeTheme.accent }}
                  />
                  <blockquote className="text-2xl lg:text-4xl font-medium italic leading-relaxed">
                    “{activeSlide.quote}”
                  </blockquote>
                  <div>
                    <div
                      className="text-xl font-bold"
                      style={{ color: activeSlide.customAccentColor || activeTheme.accentText }}
                    >
                      {activeSlide.quoteAuthor}
                    </div>
                    <div className="text-base opacity-75" style={{ color: activeTheme.textMuted }}>
                      {activeSlide.quoteRole}
                    </div>
                  </div>
                </div>
              )}

              {activeSlide.layout === 'timeline' && (
                <div className="space-y-8">
                  <h2 className="text-3xl lg:text-5xl font-bold">{activeSlide.title}</h2>
                  <div className="grid grid-cols-3 gap-6">
                    {(activeSlide.timelineSteps || []).map((st, i) => (
                      <div
                        key={i}
                        className="p-6 rounded-2xl border space-y-3"
                        style={{ backgroundColor: activeTheme.cardBg, borderColor: activeTheme.border }}
                      >
                        <div
                          className="text-2xl font-mono font-bold"
                          style={{ color: activeSlide.customAccentColor || activeTheme.accent }}
                        >
                          {st.step}
                        </div>
                        <h4 className="text-lg font-bold">{st.title}</h4>
                        <p className="text-sm opacity-80" style={{ color: activeTheme.textMuted }}>
                          {st.desc}
                        </p>
                      </div>
                    ))}
                  </div>
                </div>
              )}

              {activeSlide.layout === 'comparison' && (
                <div className="space-y-8">
                  <h2 className="text-3xl lg:text-5xl font-bold">{activeSlide.title}</h2>
                  <div className="grid grid-cols-2 gap-8">
                    <div
                      className="p-8 rounded-2xl border space-y-3"
                      style={{ backgroundColor: activeTheme.cardBg, borderColor: activeTheme.border }}
                    >
                      <h3 className="text-xl font-bold text-rose-400">{activeSlide.compTitle1}</h3>
                      <p className="text-base" style={{ color: activeTheme.textMuted }}>
                        {activeSlide.compText1}
                      </p>
                    </div>
                    <div
                      className="p-8 rounded-2xl border space-y-3"
                      style={{
                        backgroundColor: activeTheme.cardBg,
                        borderColor: activeSlide.customAccentColor || activeTheme.accent
                      }}
                    >
                      <h3
                        className="text-xl font-bold"
                        style={{ color: activeSlide.customAccentColor || activeTheme.accentText }}
                      >
                        {activeSlide.compTitle2}
                      </h3>
                      <p className="text-base" style={{ color: activeTheme.textMuted }}>
                        {activeSlide.compText2}
                      </p>
                    </div>
                  </div>
                </div>
              )}

              {activeSlide.layout === 'grid' && (
                <div className="space-y-8">
                  <h2 className="text-3xl lg:text-5xl font-bold">{activeSlide.title}</h2>
                  <div className="grid grid-cols-2 gap-6">
                    {(activeSlide.gridItems || []).map((item, i) => (
                      <div
                        key={i}
                        className="p-6 rounded-2xl border space-y-2"
                        style={{ backgroundColor: activeTheme.cardBg, borderColor: activeTheme.border }}
                      >
                        <h4
                          className="text-base font-bold"
                          style={{ color: activeSlide.customAccentColor || activeTheme.accentText }}
                        >
                          {item.title}
                        </h4>
                        <p className="text-sm" style={{ color: activeTheme.textMuted }}>
                          {item.desc}
                        </p>
                      </div>
                    ))}
                  </div>
                </div>
              )}

              {activeSlide.layout === 'team' && (
                <div className="space-y-8">
                  <h2 className="text-3xl lg:text-5xl font-bold">{activeSlide.title}</h2>
                  <div className="grid grid-cols-3 gap-6">
                    {(activeSlide.teamMembers || []).map((m, i) => (
                      <div
                        key={i}
                        className="p-6 rounded-2xl border text-center space-y-3"
                        style={{ backgroundColor: activeTheme.cardBg, borderColor: activeTheme.border }}
                      >
                        <div
                          className="w-14 h-14 mx-auto rounded-full flex items-center justify-center font-bold text-lg"
                          style={{
                            backgroundColor: `${activeSlide.customAccentColor || activeTheme.accent}30`,
                            color: activeSlide.customAccentColor || activeTheme.accentText
                          }}
                        >
                          <Users className="w-7 h-7" />
                        </div>
                        <div>
                          <h4 className="text-lg font-bold">{m.name}</h4>
                          <p
                            className="text-sm font-semibold"
                            style={{ color: activeSlide.customAccentColor || activeTheme.accentText }}
                          >
                            {m.role}
                          </p>
                        </div>
                        <p className="text-xs" style={{ color: activeTheme.textMuted }}>
                          {m.bio}
                        </p>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>

            {/* Slide Footer */}
            <div className="flex items-center justify-between text-sm opacity-60 font-mono">
              <span>{activeSlide.footerText}</span>
              <span>DocuCraft Studio</span>
            </div>
          </div>

          {/* BOTTOM CONTROLS */}
          <div className="flex items-center justify-center gap-6 z-20">
            <button
              onClick={() => setActiveSlideIndex((prev) => Math.max(prev - 1, 0))}
              disabled={activeSlideIndex === 0}
              className="p-3 rounded-full bg-white/10 hover:bg-white/20 text-white disabled:opacity-30 transition"
            >
              <ChevronLeft className="w-6 h-6" />
            </button>

            <div className="flex items-center gap-2 max-w-md overflow-x-auto px-2">
              {slides.map((_, i) => (
                <button
                  key={i}
                  onClick={() => setActiveSlideIndex(i)}
                  className={`w-3 h-3 rounded-full transition flex-shrink-0 ${
                    i === activeSlideIndex ? 'bg-blue-500 scale-125' : 'bg-white/20 hover:bg-white/40'
                  }`}
                />
              ))}
            </div>

            <button
              onClick={() => setActiveSlideIndex((prev) => Math.min(prev + 1, slides.length - 1))}
              disabled={activeSlideIndex === slides.length - 1}
              className="p-3 rounded-full bg-white/10 hover:bg-white/20 text-white disabled:opacity-30 transition"
            >
              <ChevronRight className="w-6 h-6" />
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
