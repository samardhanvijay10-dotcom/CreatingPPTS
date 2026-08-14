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
  Wand2
} from 'lucide-react';

// Preset Themes Definitions
const THEMES = {
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
  swissMinimalist: {
    id: 'swissMinimalist',
    name: 'Swiss Brutalist',
    category: 'Minimal',
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
        theme: 'sleekTech',
        tag: 'SERIES A PITCH',
        title: 'Nova AI Platform',
        subtitle: 'The intelligent workflow engine designed for modern enterprise teams.',
        footerText: 'Confidential & Proprietary • 2025'
      },
      {
        id: 's2',
        layout: 'split',
        theme: 'sleekTech',
        tag: 'THE PROBLEM & SOLUTION',
        title: 'Bridging Siloed Operations',
        subtitle: 'Teams lose over 15 hours per week switching between fragmented productivity tools.',
        col1Title: 'Traditional Process',
        col1Text: 'Manual data entry, disconnected chat apps, and missing context lead to costly delays and project friction.',
        col2Title: 'Nova AI Solution',
        col2Text: 'Autonomous agents unify project management, documentation, and real-time analytical reporting.',
        bullets: ['Automated project orchestration', 'Sub-second response times', 'Enterprise-grade end-to-end security']
      },
      {
        id: 's3',
        layout: 'stat',
        theme: 'sleekTech',
        tag: 'MARKET TRACTION',
        title: 'Rapid Enterprise Adoption',
        subtitle: 'Unprecedented growth in our first 12 months post-launch.',
        statValue: '$14.2M',
        statLabel: 'Annual Recurrent Revenue (ARR)',
        statSubtext: '+320% Year-over-Year revenue expansion with 98% net retention rate.',
        bullets: ['500+ Active enterprise clients', 'Zero churn in Top 50 tier', 'Net Promoter Score of 78']
      },
      {
        id: 's4',
        layout: 'quote',
        theme: 'sleekTech',
        tag: 'CLIENT TESTIMONIAL',
        title: 'Industry Endorsement',
        quote: 'Nova AI transformed our product development velocity. We shipped 3x faster with half the operational overhead.',
        quoteAuthor: 'Elena Rostova',
        quoteRole: 'VP of Engineering, Apex Global'
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
        tag: 'DESIGN PORTFOLIO 2025',
        title: 'Aethel Design Studio',
        subtitle: 'Crafting timeless digital products and immersive brand experiences.',
        footerText: 'Selected Works • Volume IV'
      },
      {
        id: 'p2',
        layout: 'content',
        theme: 'warmEditorial',
        tag: 'CORE PHILOSOPHIES',
        title: 'Simplicity & Precision',
        subtitle: 'Good design is invisible—it empowers users while evoking emotion and purpose.',
        bullets: [
          'Human-centered interaction design',
          'Tactile, bespoke typography and grid layouts',
          'Sustainable digital architecture and performance'
        ]
      },
      {
        id: 'p3',
        layout: 'quote',
        theme: 'warmEditorial',
        tag: 'WHAT CLIENTS SAY',
        title: 'Client Impression',
        quote: 'Working with Aethel changed how we articulate our values. The brand design gave us instant authority.',
        quoteAuthor: 'Marcus Vance',
        quoteRole: 'Founder, Luminary Press'
      }
    ]
  },
  techRoadmap: {
    name: '⚡ Technical Roadmap',
    slides: [
      {
        id: 'r1',
        layout: 'title',
        theme: 'cyberpunk',
        tag: 'SYSTEM ARCHITECTURE v3.0',
        title: 'NextGen Cloud Grid',
        subtitle: 'Distributed, zero-trust edge infrastructure for real-time compute.',
        footerText: 'Internal Engineering Sync'
      },
      {
        id: 'r2',
        layout: 'stat',
        theme: 'cyberpunk',
        tag: 'PERFORMANCE METRICS',
        title: 'Sub-Millisecond Speed',
        subtitle: 'Benchmarking global edge node distribution under peak load.',
        statValue: '< 12ms',
        statLabel: 'Global P99 Response Time',
        statSubtext: 'Optimized via WebAssembly isolates deployed across 180+ POPs globally.',
        bullets: ['99.999% SLA Uptime', 'Automatic DDoS mitigation', 'Zero cold-start delay']
      },
      {
        id: 'r3',
        layout: 'split',
        theme: 'cyberpunk',
        tag: 'Q3 & Q4 TARGETS',
        title: 'Execution Roadmap',
        subtitle: 'Key technical milestones scheduled for the upcoming quarters.',
        col1Title: 'Q3: Foundation',
        col1Text: 'Migration to rust-native parser, state replication tuning, and multi-region failover.',
        col2Title: 'Q4: Intelligence',
        col2Text: 'On-device LLM inference integration and automated anomaly detection.',
        bullets: ['Global VPC peering', 'SOC2 Type II compliance', 'Developer SDK 2.0 launch']
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
  const [isPresenting, setIsPresenting] = useState(false);
  const [isExporting, setIsExporting] = useState(false);

  const activeSlide = slides[activeSlideIndex] || slides[0];
  const activeTheme = THEMES[activeSlide?.theme] || THEMES.sleekTech;

  // Handle keyboard shortcut in presentation mode
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

  // Slide CRUD Operations
  const updateActiveSlide = (fields) => {
    setSlides((prev) =>
      prev.map((slide, idx) => (idx === activeSlideIndex ? { ...slide, ...fields } : slide))
    );
  };

  const addSlide = () => {
    const newSlide = {
      id: 's_' + Date.now(),
      layout: 'content',
      theme: activeSlide.theme || 'sleekTech',
      tag: 'NEW SECTION',
      title: 'Slide Title Here',
      subtitle: 'Add a concise summary or key takeaways here.',
      bullets: ['First key point highlight', 'Second important detail or stat', 'Actionable takeaway']
    };
    const newSlides = [...slides];
    newSlides.splice(activeSlideIndex + 1, 0, newSlide);
    setSlides(newSlides);
    setActiveSlideIndex(activeSlideIndex + 1);
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
  };

  const deleteSlide = () => {
    if (slides.length <= 1) return;
    const newSlides = slides.filter((_, idx) => idx !== activeSlideIndex);
    setSlides(newSlides);
    setActiveSlideIndex(Math.max(0, activeSlideIndex - 1));
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
    }
  };

  // Helper to add bullet point
  const updateBullet = (idx, value) => {
    const newBullets = [...(activeSlide.bullets || [])];
    newBullets[idx] = value;
    updateActiveSlide({ bullets: newBullets });
  };

  const addBullet = () => {
    const newBullets = [...(activeSlide.bullets || []), 'New key point'];
    updateActiveSlide({ bullets: newBullets });
  };

  const removeBullet = (idx) => {
    const newBullets = (activeSlide.bullets || []).filter((_, i) => i !== idx);
    updateActiveSlide({ bullets: newBullets });
  };

  // PowerPoint PPTX Exporting
  const exportToPowerPoint = async () => {
    try {
      setIsExporting(true);
      const pptx = new pptxgen();
      pptx.layout = 'LAYOUT_16x9';
      pptx.title = 'Aesthetic Presentation';

      slides.forEach((slide) => {
        const theme = THEMES[slide.theme] || THEMES.sleekTech;
        const pptxSlide = pptx.addSlide();

        // Background Color
        pptxSlide.background = { color: theme.bg.replace('#', '') };

        // Card Container background if desired
        const accentHex = theme.accent.replace('#', '');
        const textMainHex = theme.textMain.replace('#', '');
        const textMutedHex = theme.textMuted.replace('#', '');

        // Category Tag
        if (slide.tag) {
          pptxSlide.addText(slide.tag.toUpperCase(), {
            x: 0.8,
            y: 0.6,
            w: 8.4,
            h: 0.4,
            fontSize: 12,
            bold: true,
            color: accentHex,
            fontFace: 'Arial'
          });
        }

        // Title
        if (slide.title) {
          pptxSlide.addText(slide.title, {
            x: 0.8,
            y: slide.layout === 'title' ? 1.8 : 1.0,
            w: 8.4,
            h: 1.0,
            fontSize: slide.layout === 'title' ? 36 : 28,
            bold: true,
            color: textMainHex,
            fontFace: 'Arial'
          });
        }

        // Subtitle
        if (slide.subtitle) {
          pptxSlide.addText(slide.subtitle, {
            x: 0.8,
            y: slide.layout === 'title' ? 3.0 : 2.0,
            w: 8.4,
            h: 0.8,
            fontSize: 16,
            color: textMutedHex,
            fontFace: 'Arial'
          });
        }

        // Layout specific PPT content
        if (slide.layout === 'content' && slide.bullets) {
          const bulletItems = slide.bullets.map((b) => ({
            text: b,
            options: { fontSize: 16, color: textMainHex, breakLine: true, bullet: true }
          }));
          pptxSlide.addText(bulletItems, {
            x: 0.8,
            y: 3.0,
            w: 8.4,
            h: 2.2,
            fontFace: 'Arial'
          });
        } else if (slide.layout === 'split') {
          if (slide.col1Title) {
            pptxSlide.addText(slide.col1Title, {
              x: 0.8,
              y: 2.9,
              w: 4.0,
              h: 0.4,
              fontSize: 18,
              bold: true,
              color: accentHex,
              fontFace: 'Arial'
            });
            pptxSlide.addText(slide.col1Text || '', {
              x: 0.8,
              y: 3.4,
              w: 4.0,
              h: 1.6,
              fontSize: 14,
              color: textMutedHex,
              fontFace: 'Arial'
            });
          }
          if (slide.col2Title) {
            pptxSlide.addText(slide.col2Title, {
              x: 5.2,
              y: 2.9,
              w: 4.0,
              h: 0.4,
              fontSize: 18,
              bold: true,
              color: accentHex,
              fontFace: 'Arial'
            });
            pptxSlide.addText(slide.col2Text || '', {
              x: 5.2,
              y: 3.4,
              w: 4.0,
              h: 1.6,
              fontSize: 14,
              color: textMutedHex,
              fontFace: 'Arial'
            });
          }
        } else if (slide.layout === 'stat') {
          if (slide.statValue) {
            pptxSlide.addText(slide.statValue, {
              x: 0.8,
              y: 2.8,
              w: 4.0,
              h: 1.2,
              fontSize: 48,
              bold: true,
              color: accentHex,
              fontFace: 'Arial'
            });
          }
          if (slide.statLabel) {
            pptxSlide.addText(slide.statLabel, {
              x: 0.8,
              y: 4.0,
              w: 4.0,
              h: 0.4,
              fontSize: 16,
              bold: true,
              color: textMainHex,
              fontFace: 'Arial'
            });
          }
          if (slide.statSubtext) {
            pptxSlide.addText(slide.statSubtext, {
              x: 0.8,
              y: 4.4,
              w: 4.0,
              h: 0.8,
              fontSize: 13,
              color: textMutedHex,
              fontFace: 'Arial'
            });
          }
          if (slide.bullets) {
            const bulletItems = slide.bullets.map((b) => ({
              text: b,
              options: { fontSize: 14, color: textMainHex, breakLine: true, bullet: true }
            }));
            pptxSlide.addText(bulletItems, {
              x: 5.2,
              y: 3.0,
              w: 4.0,
              h: 2.0,
              fontFace: 'Arial'
            });
          }
        } else if (slide.layout === 'quote') {
          if (slide.quote) {
            pptxSlide.addText(`“${slide.quote}”`, {
              x: 1.0,
              y: 2.8,
              w: 8.0,
              h: 1.5,
              fontSize: 22,
              italic: true,
              color: textMainHex,
              fontFace: 'Arial'
            });
          }
          if (slide.quoteAuthor) {
            pptxSlide.addText(`— ${slide.quoteAuthor} (${slide.quoteRole || ''})`, {
              x: 1.0,
              y: 4.4,
              w: 8.0,
              h: 0.5,
              fontSize: 15,
              bold: true,
              color: accentHex,
              fontFace: 'Arial'
            });
          }
        }

        // Footer
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

      await pptx.writeFile({ fileName: 'Aesthetic_Presentation.pptx' });
      confetti({
        particleCount: 120,
        spread: 80,
        origin: { y: 0.6 }
      });
    } catch (err) {
      console.error('PPT Export Error:', err);
      alert('Error exporting PowerPoint deck. Please check console.');
    } finally {
      setIsExporting(false);
    }
  };

  return (
    <div className="min-h-screen bg-[#07090E] text-slate-100 flex flex-col font-sans selection:bg-indigo-500 selection:text-white">
      {/* HEADER NAVBAR */}
      <header className="h-16 border-b border-slate-800/80 bg-[#0B0F17]/90 backdrop-blur px-6 flex items-center justify-between z-20">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-xl bg-gradient-to-tr from-indigo-600 via-purple-600 to-pink-500 flex items-center justify-center shadow-lg shadow-indigo-500/20">
            <Sparkles className="w-5 h-5 text-white animate-pulse" />
          </div>
          <div>
            <h1 className="text-lg font-bold tracking-wide text-white flex items-center gap-2">
              DOCUCRAFT <span className="text-xs px-2 py-0.5 rounded-full bg-indigo-500/20 text-indigo-400 font-mono border border-indigo-500/30">PPT STUDIO</span>
            </h1>
            <p className="text-xs text-slate-400">Next-Generation Aesthetic Presentation Creator</p>
          </div>
        </div>

        {/* PRESET DECKS & ACTION BUTTONS */}
        <div className="flex items-center gap-3">
          <div className="hidden lg:flex items-center gap-1.5 bg-slate-900/80 border border-slate-800 rounded-lg p-1 text-xs">
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
              Tech Roadmap
            </button>
          </div>

          <button
            onClick={() => setIsPresenting(true)}
            className="flex items-center gap-2 px-4 py-2 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-200 text-sm font-medium transition border border-slate-700 shadow-sm"
          >
            <Play className="w-4 h-4 text-emerald-400 fill-emerald-400" />
            <span>Present</span>
          </button>

          <button
            onClick={exportToPowerPoint}
            disabled={isExporting}
            className="flex items-center gap-2 px-4 py-2 rounded-lg bg-gradient-to-r from-indigo-600 to-purple-600 hover:from-indigo-500 hover:to-purple-500 text-white text-sm font-semibold shadow-lg shadow-indigo-600/30 transition transform active:scale-95 disabled:opacity-50"
          >
            <Download className="w-4 h-4" />
            <span>{isExporting ? 'Exporting...' : 'Export PPTX'}</span>
          </button>
        </div>
      </header>

      {/* MAIN WORKSPACE LAYOUT */}
      <div className="flex-1 flex overflow-hidden">
        {/* LEFT SIDEBAR: SLIDE THUMBNAILS & OUTLINE */}
        <aside className="w-64 border-r border-slate-800/80 bg-[#0A0D14] flex flex-col z-10">
          <div className="p-4 border-b border-slate-800/80 flex items-center justify-between">
            <span className="text-xs font-semibold uppercase tracking-wider text-slate-400 flex items-center gap-1.5">
              <Layers className="w-4 h-4 text-indigo-400" /> Slides ({slides.length})
            </span>
            <button
              onClick={addSlide}
              title="Add Slide"
              className="p-1.5 rounded-lg bg-indigo-600/20 hover:bg-indigo-600 text-indigo-400 hover:text-white transition border border-indigo-500/30"
            >
              <Plus className="w-4 h-4" />
            </button>
          </div>

          <div className="flex-1 overflow-y-auto p-3 space-y-3 custom-scrollbar">
            {slides.map((slide, idx) => {
              const th = THEMES[slide.theme] || THEMES.sleekTech;
              const isActive = idx === activeSlideIndex;
              return (
                <div
                  key={slide.id || idx}
                  onClick={() => setActiveSlideIndex(idx)}
                  className={`group relative rounded-xl p-2.5 transition cursor-pointer border ${
                    isActive
                      ? 'border-indigo-500 bg-indigo-950/30 shadow-lg shadow-indigo-500/10'
                      : 'border-slate-800/80 bg-slate-900/40 hover:border-slate-700 hover:bg-slate-900/80'
                  }`}
                >
                  <div className="flex items-center justify-between mb-2">
                    <span className="text-[10px] font-mono font-bold px-1.5 py-0.5 rounded bg-slate-800 text-slate-400">
                      0{idx + 1}
                    </span>
                    <span className="text-[10px] uppercase font-semibold text-slate-500 group-hover:text-slate-400">
                      {slide.layout}
                    </span>
                  </div>

                  {/* THUMBNAIL MINI PREVIEW */}
                  <div
                    className="w-full aspect-[16/9] rounded-lg p-2 flex flex-col justify-between overflow-hidden relative border shadow-inner transition group-hover:scale-[1.02]"
                    style={{
                      backgroundColor: th.bg,
                      borderColor: th.border,
                      color: th.textMain
                    }}
                  >
                    <div className="space-y-1">
                      {slide.tag && (
                        <div
                          className="w-8 h-1 rounded-full opacity-80"
                          style={{ backgroundColor: th.accent }}
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
                      <span>{slide.theme}</span>
                      <span>16:9</span>
                    </div>
                  </div>
                </div>
              );
            })}
          </div>

          {/* SLIDE CONTROL BUTTONS */}
          <div className="p-3 border-t border-slate-800/80 grid grid-cols-4 gap-1.5 bg-slate-950/50">
            <button
              onClick={() => moveSlide(-1)}
              disabled={activeSlideIndex === 0}
              title="Move Up"
              className="p-2 rounded-lg bg-slate-900 hover:bg-slate-800 text-slate-300 disabled:opacity-30 border border-slate-800 flex items-center justify-center"
            >
              <ChevronUp className="w-4 h-4" />
            </button>
            <button
              onClick={() => moveSlide(1)}
              disabled={activeSlideIndex === slides.length - 1}
              title="Move Down"
              className="p-2 rounded-lg bg-slate-900 hover:bg-slate-800 text-slate-300 disabled:opacity-30 border border-slate-800 flex items-center justify-center"
            >
              <ChevronDown className="w-4 h-4" />
            </button>
            <button
              onClick={duplicateSlide}
              title="Duplicate Slide"
              className="p-2 rounded-lg bg-slate-900 hover:bg-slate-800 text-slate-300 border border-slate-800 flex items-center justify-center"
            >
              <Copy className="w-4 h-4" />
            </button>
            <button
              onClick={deleteSlide}
              disabled={slides.length <= 1}
              title="Delete Slide"
              className="p-2 rounded-lg bg-rose-950/40 hover:bg-rose-900/60 text-rose-400 disabled:opacity-30 border border-rose-900/50 flex items-center justify-center"
            >
              <Trash2 className="w-4 h-4" />
            </button>
          </div>
        </aside>

        {/* CENTER CANVAS PREVIEW WORKSPACE */}
        <main className="flex-1 bg-[#06080F] p-6 lg:p-10 flex flex-col items-center justify-center relative overflow-y-auto">
          {/* VISUAL PREVIEW CONTROL TABS */}
          <div className="mb-4 flex items-center gap-4 bg-slate-900/90 border border-slate-800 rounded-full px-4 py-1.5 shadow-xl text-xs backdrop-blur z-10">
            <span className="text-slate-400 font-medium">Canvas Decor:</span>
            <label className="flex items-center gap-1.5 cursor-pointer text-slate-300 hover:text-white">
              <input
                type="checkbox"
                checked={showGridPattern}
                onChange={(e) => setShowGridPattern(e.target.checked)}
                className="rounded bg-slate-800 border-slate-700 text-indigo-600 focus:ring-0"
              />
              <Grid className="w-3.5 h-3.5 text-indigo-400" /> Grid Lines
            </label>
            <div className="w-px h-3 bg-slate-800"></div>
            <label className="flex items-center gap-1.5 cursor-pointer text-slate-300 hover:text-white">
              <input
                type="checkbox"
                checked={showGlowEffect}
                onChange={(e) => setShowGlowEffect(e.target.checked)}
                className="rounded bg-slate-800 border-slate-700 text-indigo-600 focus:ring-0"
              />
              <Sparkles className="w-3.5 h-3.5 text-amber-400" /> Glow Effect
            </label>
            <div className="w-px h-3 bg-slate-800"></div>
            <label className="flex items-center gap-1.5 cursor-pointer text-slate-300 hover:text-white">
              <input
                type="checkbox"
                checked={showSlideNumber}
                onChange={(e) => setShowSlideNumber(e.target.checked)}
                className="rounded bg-slate-800 border-slate-700 text-indigo-600 focus:ring-0"
              />
              <Sliders className="w-3.5 h-3.5 text-emerald-400" /> Slide Badge
            </label>
          </div>

          {/* 16:9 SLIDE CONTAINER */}
          <div className="w-full max-w-4xl aspect-[16/9] relative shadow-2xl rounded-2xl overflow-hidden transition-all duration-300 border border-slate-700/50 flex flex-col justify-between p-8 lg:p-12"
               style={{
                 backgroundColor: activeTheme.bg,
                 color: activeTheme.textMain,
                 fontFamily: activeTheme.font === 'font-serif' ? 'Georgia, serif' : activeTheme.font === 'font-mono' ? 'Courier New, monospace' : 'system-ui, sans-serif'
               }}>

            {/* OPTIONAL BACKGROUND DECORATIONS */}
            {showGridPattern && (
              <div
                className="absolute inset-0 pointer-events-none opacity-10"
                style={{
                  backgroundImage: `radial-gradient(${activeTheme.accent} 1px, transparent 1px)`,
                  backgroundSize: '24px 24px'
                }}
              ></div>
            )}

            {showGlowEffect && (
              <div
                className="absolute -top-24 -right-24 w-72 h-72 rounded-full blur-3xl pointer-events-none opacity-25"
                style={{ backgroundColor: activeTheme.accent }}
              ></div>
            )}

            {/* SLIDE TOP SECTION: TAG / CATEGORY */}
            <div className="relative z-10 flex items-center justify-between">
              {activeSlide.tag ? (
                <div
                  className="inline-flex items-center gap-2 px-3 py-1 rounded-full text-xs font-mono font-bold tracking-wider uppercase border shadow-sm"
                  style={{
                    backgroundColor: `${activeTheme.accent}15`,
                    color: activeTheme.accentText,
                    borderColor: `${activeTheme.accent}40`
                  }}
                >
                  <span className="w-1.5 h-1.5 rounded-full" style={{ backgroundColor: activeTheme.accent }}></span>
                  {activeSlide.tag}
                </div>
              ) : <div></div>}

              {showSlideNumber && (
                <span
                  className="text-xs font-mono opacity-60 font-medium tracking-widest"
                  style={{ color: activeTheme.textMuted }}
                >
                  SLIDE // 0{activeSlideIndex + 1}
                </span>
              )}
            </div>

            {/* SLIDE BODY CONTENT BASED ON LAYOUT */}
            <div className="relative z-10 my-auto py-4">
              {/* LAYOUT: TITLE SLIDE */}
              {activeSlide.layout === 'title' && (
                <div className="space-y-4 max-w-2xl">
                  <h2
                    className="text-3xl lg:text-5xl font-black tracking-tight leading-tight"
                    style={{ color: activeTheme.textMain }}
                  >
                    {activeSlide.title || 'Click to edit title'}
                  </h2>
                  <p
                    className="text-base lg:text-xl font-normal leading-relaxed opacity-90"
                    style={{ color: activeTheme.textMuted }}
                  >
                    {activeSlide.subtitle || 'Click to edit subtitle'}
                  </p>
                </div>
              )}

              {/* LAYOUT: CONTENT / BULLETS */}
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
                      <p
                        className="text-sm lg:text-base mt-1"
                        style={{ color: activeTheme.textMuted }}
                      >
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
                            style={{ color: activeTheme.accent }}
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

                  <div className="grid grid-cols-2 gap-6">
                    <div
                      className="p-4 rounded-xl border"
                      style={{
                        backgroundColor: activeTheme.cardBg,
                        borderColor: activeTheme.border
                      }}
                    >
                      <h3 className="text-base font-bold mb-1" style={{ color: activeTheme.accentText }}>
                        {activeSlide.col1Title || 'Column 1'}
                      </h3>
                      <p className="text-xs lg:text-sm leading-relaxed" style={{ color: activeTheme.textMuted }}>
                        {activeSlide.col1Text || 'Content description for left column.'}
                      </p>
                    </div>

                    <div
                      className="p-4 rounded-xl border"
                      style={{
                        backgroundColor: activeTheme.cardBg,
                        borderColor: activeTheme.border
                      }}
                    >
                      <h3 className="text-base font-bold mb-1" style={{ color: activeTheme.accentText }}>
                        {activeSlide.col2Title || 'Column 2'}
                      </h3>
                      <p className="text-xs lg:text-sm leading-relaxed" style={{ color: activeTheme.textMuted }}>
                        {activeSlide.col2Text || 'Content description for right column.'}
                      </p>
                    </div>
                  </div>
                </div>
              )}

              {/* LAYOUT: BIG STAT CALLOUT */}
              {activeSlide.layout === 'stat' && (
                <div className="grid grid-cols-1 lg:grid-cols-2 gap-8 items-center">
                  <div className="space-y-2">
                    <div
                      className="text-5xl lg:text-7xl font-black tracking-tight"
                      style={{ color: activeTheme.accent }}
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
                      className="p-5 rounded-2xl border space-y-3"
                      style={{
                        backgroundColor: activeTheme.cardBg,
                        borderColor: activeTheme.border
                      }}
                    >
                      <h4 className="text-xs font-bold uppercase tracking-wider" style={{ color: activeTheme.accentText }}>
                        Key Highlights
                      </h4>
                      <ul className="space-y-2 text-xs lg:text-sm">
                        {activeSlide.bullets.map((b, i) => (
                          <li key={i} className="flex items-center gap-2" style={{ color: activeTheme.textMain }}>
                            <div className="w-1.5 h-1.5 rounded-full" style={{ backgroundColor: activeTheme.accent }}></div>
                            {b}
                          </li>
                        ))}
                      </ul>
                    </div>
                  )}
                </div>
              )}

              {/* LAYOUT: QUOTE / TESTIMONIAL */}
              {activeSlide.layout === 'quote' && (
                <div className="space-y-6 max-w-2xl mx-auto text-center">
                  <QuoteIcon
                    className="w-10 h-10 mx-auto opacity-30"
                    style={{ color: activeTheme.accent }}
                  />
                  <blockquote
                    className="text-xl lg:text-3xl font-medium italic leading-relaxed"
                    style={{ color: activeTheme.textMain }}
                  >
                    “{activeSlide.quote || 'Insert an impactful quote or testimonial here.'}”
                  </blockquote>
                  <div>
                    <div className="text-sm font-bold" style={{ color: activeTheme.accentText }}>
                      {activeSlide.quoteAuthor || 'Author Name'}
                    </div>
                    <div className="text-xs opacity-75" style={{ color: activeTheme.textMuted }}>
                      {activeSlide.quoteRole || 'Title, Organization'}
                    </div>
                  </div>
                </div>
              )}
            </div>

            {/* SLIDE FOOTER SECTION */}
            <div
              className="relative z-10 pt-4 border-t flex items-center justify-between text-[11px] font-mono"
              style={{
                borderColor: `${activeTheme.border}80`,
                color: activeTheme.textMuted
              }}
            >
              <span>{activeSlide.footerText || 'DocuCraft Presentation System'}</span>
              <span className="flex items-center gap-1.5">
                <span className="w-2 h-2 rounded-full animate-ping" style={{ backgroundColor: activeTheme.accent }}></span>
                {activeTheme.name}
              </span>
            </div>
          </div>
        </main>

        {/* RIGHT SIDEBAR: REAL-TIME INSPECTOR & EDITOR */}
        <aside className="w-80 border-l border-slate-800/80 bg-[#0A0D14] flex flex-col z-10 overflow-y-auto custom-scrollbar">
          <div className="p-4 border-b border-slate-800/80 flex items-center justify-between bg-slate-950/40">
            <span className="text-xs font-semibold uppercase tracking-wider text-slate-300 flex items-center gap-1.5">
              <Sliders className="w-4 h-4 text-purple-400" /> Slide Editor
            </span>
            <span className="text-[10px] bg-slate-800 px-2 py-0.5 rounded text-slate-400 font-mono">
              #{activeSlideIndex + 1}
            </span>
          </div>

          <div className="p-4 space-y-6">
            {/* THEME SELECTOR */}
            <div className="space-y-2">
              <label className="text-xs font-semibold text-slate-400 flex items-center gap-1.5">
                <Palette className="w-3.5 h-3.5 text-indigo-400" /> Slide Aesthetics & Theme
              </label>
              <div className="grid grid-cols-1 gap-2">
                {Object.values(THEMES).map((th) => (
                  <button
                    key={th.id}
                    onClick={() => updateActiveSlide({ theme: th.id })}
                    className={`flex items-center justify-between p-2.5 rounded-xl border text-left text-xs transition ${
                      activeSlide.theme === th.id
                        ? 'border-indigo-500 bg-indigo-950/40 text-white font-medium shadow-sm'
                        : 'border-slate-800 bg-slate-900/50 text-slate-400 hover:border-slate-700 hover:text-slate-200'
                    }`}
                  >
                    <div className="flex items-center gap-2.5">
                      <div
                        className="w-4 h-4 rounded-full border border-white/20 flex-shrink-0"
                        style={{ backgroundColor: th.accent }}
                      ></div>
                      <span>{th.name}</span>
                    </div>
                    <span className="text-[10px] font-mono opacity-60 uppercase">{th.category}</span>
                  </button>
                ))}
              </div>
            </div>

            {/* LAYOUT SELECTOR */}
            <div className="space-y-2">
              <label className="text-xs font-semibold text-slate-400 flex items-center gap-1.5">
                <Layout className="w-3.5 h-3.5 text-purple-400" /> Layout Template
              </label>
              <div className="grid grid-cols-3 gap-1.5 text-xs">
                {[
                  { id: 'title', label: 'Title' },
                  { id: 'content', label: 'Bullets' },
                  { id: 'split', label: 'Split' },
                  { id: 'stat', label: 'Metric' },
                  { id: 'quote', label: 'Quote' }
                ].map((ly) => (
                  <button
                    key={ly.id}
                    onClick={() => updateActiveSlide({ layout: ly.id })}
                    className={`py-2 px-1 rounded-lg border text-center transition font-medium ${
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

            <hr className="border-slate-800/80" />

            {/* CONTENT EDITING FIELDS */}
            <div className="space-y-4">
              <div className="space-y-1">
                <label className="text-xs font-medium text-slate-400">Category Badge / Tag</label>
                <input
                  type="text"
                  value={activeSlide.tag || ''}
                  onChange={(e) => updateActiveSlide({ tag: e.target.value })}
                  placeholder="e.g. KEY METRICS"
                  className="w-full bg-slate-900 border border-slate-800 rounded-lg px-3 py-2 text-xs text-slate-100 focus:outline-none focus:border-indigo-500"
                />
              </div>

              <div className="space-y-1">
                <label className="text-xs font-medium text-slate-400">Slide Main Title</label>
                <input
                  type="text"
                  value={activeSlide.title || ''}
                  onChange={(e) => updateActiveSlide({ title: e.target.value })}
                  placeholder="Slide Headline"
                  className="w-full bg-slate-900 border border-slate-800 rounded-lg px-3 py-2 text-xs text-slate-100 focus:outline-none focus:border-indigo-500"
                />
              </div>

              <div className="space-y-1">
                <label className="text-xs font-medium text-slate-400">Subtitle / Description</label>
                <textarea
                  rows={2}
                  value={activeSlide.subtitle || ''}
                  onChange={(e) => updateActiveSlide({ subtitle: e.target.value })}
                  placeholder="Brief summary sentence..."
                  className="w-full bg-slate-900 border border-slate-800 rounded-lg px-3 py-2 text-xs text-slate-100 focus:outline-none focus:border-indigo-500 resize-none"
                />
              </div>

              {/* SPECIFIC FIELDS: BULLETS */}
              {(activeSlide.layout === 'content' || activeSlide.layout === 'stat') && (
                <div className="space-y-2">
                  <div className="flex items-center justify-between">
                    <label className="text-xs font-medium text-slate-400">Bullet Points</label>
                    <button
                      onClick={addBullet}
                      className="text-[11px] text-indigo-400 hover:text-indigo-300 flex items-center gap-1 font-medium"
                    >
                      <Plus className="w-3 h-3" /> Add Point
                    </button>
                  </div>
                  <div className="space-y-2">
                    {(activeSlide.bullets || []).map((bullet, idx) => (
                      <div key={idx} className="flex items-center gap-2">
                        <input
                          type="text"
                          value={bullet}
                          onChange={(e) => updateBullet(idx, e.target.value)}
                          className="flex-1 bg-slate-900 border border-slate-800 rounded-lg px-3 py-1.5 text-xs text-slate-100 focus:outline-none focus:border-indigo-500"
                        />
                        <button
                          onClick={() => removeBullet(idx)}
                          className="p-1.5 text-slate-500 hover:text-rose-400"
                        >
                          <X className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    ))}
                  </div>
                </div>
              )}

              {/* SPECIFIC FIELDS: SPLIT COLUMNS */}
              {activeSlide.layout === 'split' && (
                <div className="space-y-3 pt-2">
                  <div className="space-y-1">
                    <label className="text-xs font-medium text-slate-400">Column 1 Header</label>
                    <input
                      type="text"
                      value={activeSlide.col1Title || ''}
                      onChange={(e) => updateActiveSlide({ col1Title: e.target.value })}
                      className="w-full bg-slate-900 border border-slate-800 rounded-lg px-3 py-1.5 text-xs text-slate-100"
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-medium text-slate-400">Column 1 Text</label>
                    <textarea
                      rows={2}
                      value={activeSlide.col1Text || ''}
                      onChange={(e) => updateActiveSlide({ col1Text: e.target.value })}
                      className="w-full bg-slate-900 border border-slate-800 rounded-lg px-3 py-1.5 text-xs text-slate-100 resize-none"
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-medium text-slate-400">Column 2 Header</label>
                    <input
                      type="text"
                      value={activeSlide.col2Title || ''}
                      onChange={(e) => updateActiveSlide({ col2Title: e.target.value })}
                      className="w-full bg-slate-900 border border-slate-800 rounded-lg px-3 py-1.5 text-xs text-slate-100"
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-medium text-slate-400">Column 2 Text</label>
                    <textarea
                      rows={2}
                      value={activeSlide.col2Text || ''}
                      onChange={(e) => updateActiveSlide({ col2Text: e.target.value })}
                      className="w-full bg-slate-900 border border-slate-800 rounded-lg px-3 py-1.5 text-xs text-slate-100 resize-none"
                    />
                  </div>
                </div>
              )}

              {/* SPECIFIC FIELDS: STAT */}
              {activeSlide.layout === 'stat' && (
                <div className="space-y-3 pt-2">
                  <div className="space-y-1">
                    <label className="text-xs font-medium text-slate-400">Big Stat Value</label>
                    <input
                      type="text"
                      value={activeSlide.statValue || ''}
                      onChange={(e) => updateActiveSlide({ statValue: e.target.value })}
                      className="w-full bg-slate-900 border border-slate-800 rounded-lg px-3 py-1.5 text-xs text-slate-100"
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-medium text-slate-400">Metric Label</label>
                    <input
                      type="text"
                      value={activeSlide.statLabel || ''}
                      onChange={(e) => updateActiveSlide({ statLabel: e.target.value })}
                      className="w-full bg-slate-900 border border-slate-800 rounded-lg px-3 py-1.5 text-xs text-slate-100"
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-medium text-slate-400">Metric Subtext</label>
                    <input
                      type="text"
                      value={activeSlide.statSubtext || ''}
                      onChange={(e) => updateActiveSlide({ statSubtext: e.target.value })}
                      className="w-full bg-slate-900 border border-slate-800 rounded-lg px-3 py-1.5 text-xs text-slate-100"
                    />
                  </div>
                </div>
              )}

              {/* SPECIFIC FIELDS: QUOTE */}
              {activeSlide.layout === 'quote' && (
                <div className="space-y-3 pt-2">
                  <div className="space-y-1">
                    <label className="text-xs font-medium text-slate-400">Quote Text</label>
                    <textarea
                      rows={3}
                      value={activeSlide.quote || ''}
                      onChange={(e) => updateActiveSlide({ quote: e.target.value })}
                      className="w-full bg-slate-900 border border-slate-800 rounded-lg px-3 py-1.5 text-xs text-slate-100 resize-none"
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-medium text-slate-400">Author Name</label>
                    <input
                      type="text"
                      value={activeSlide.quoteAuthor || ''}
                      onChange={(e) => updateActiveSlide({ quoteAuthor: e.target.value })}
                      className="w-full bg-slate-900 border border-slate-800 rounded-lg px-3 py-1.5 text-xs text-slate-100"
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-medium text-slate-400">Author Role / Company</label>
                    <input
                      type="text"
                      value={activeSlide.quoteRole || ''}
                      onChange={(e) => updateActiveSlide({ quoteRole: e.target.value })}
                      className="w-full bg-slate-900 border border-slate-800 rounded-lg px-3 py-1.5 text-xs text-slate-100"
                    />
                  </div>
                </div>
              )}

              <div className="space-y-1 pt-2">
                <label className="text-xs font-medium text-slate-400">Footer Note</label>
                <input
                  type="text"
                  value={activeSlide.footerText || ''}
                  onChange={(e) => updateActiveSlide({ footerText: e.target.value })}
                  placeholder="Footer disclaimer..."
                  className="w-full bg-slate-900 border border-slate-800 rounded-lg px-3 py-1.5 text-xs text-slate-100"
                />
              </div>
            </div>
          </div>
        </aside>
      </div>

      {/* FULLSCREEN PRESENTATION MODE OVERLAY */}
      {isPresenting && (
        <div className="fixed inset-0 z-50 bg-black flex flex-col justify-between p-8 lg:p-16 animate-fade-in">
          {/* TOP CONTROLS */}
          <div className="flex items-center justify-between z-20 text-white/70">
            <div className="text-xs font-mono">
              PRESENTING ({activeSlideIndex + 1}/{slides.length})
            </div>
            <button
              onClick={() => setIsPresenting(false)}
              className="p-2 rounded-full bg-white/10 hover:bg-white/20 text-white transition"
            >
              <X className="w-6 h-6" />
            </button>
          </div>

          {/* PRESENTATION SLIDE CENTER */}
          <div className="w-full max-w-6xl aspect-[16/9] mx-auto relative rounded-3xl p-12 lg:p-20 flex flex-col justify-between shadow-2xl overflow-hidden border border-white/10"
               style={{
                 backgroundColor: activeTheme.bg,
                 color: activeTheme.textMain,
                 fontFamily: activeTheme.font === 'font-serif' ? 'Georgia, serif' : activeTheme.font === 'font-mono' ? 'Courier New, monospace' : 'system-ui, sans-serif'
               }}>

            {showGridPattern && (
              <div
                className="absolute inset-0 pointer-events-none opacity-10"
                style={{
                  backgroundImage: `radial-gradient(${activeTheme.accent} 1px, transparent 1px)`,
                  backgroundSize: '32px 32px'
                }}
              ></div>
            )}

            {/* Slide Header */}
            <div className="flex items-center justify-between">
              {activeSlide.tag && (
                <div
                  className="inline-flex items-center gap-2 px-4 py-1.5 rounded-full text-sm font-mono font-bold tracking-wider uppercase border shadow-sm"
                  style={{
                    backgroundColor: `${activeTheme.accent}20`,
                    color: activeTheme.accentText,
                    borderColor: `${activeTheme.accent}50`
                  }}
                >
                  <span className="w-2 h-2 rounded-full" style={{ backgroundColor: activeTheme.accent }}></span>
                  {activeSlide.tag}
                </div>
              )}
            </div>

            {/* Slide Content */}
            <div className="my-auto py-6">
              {activeSlide.layout === 'title' && (
                <div className="space-y-6 max-w-3xl">
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
                          <CheckCircle2 className="w-7 h-7 flex-shrink-0" style={{ color: activeTheme.accent }} />
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
                    <div className="p-8 rounded-2xl border" style={{ backgroundColor: activeTheme.cardBg, borderColor: activeTheme.border }}>
                      <h3 className="text-2xl font-bold mb-3" style={{ color: activeTheme.accentText }}>
                        {activeSlide.col1Title}
                      </h3>
                      <p className="text-base lg:text-lg" style={{ color: activeTheme.textMuted }}>
                        {activeSlide.col1Text}
                      </p>
                    </div>
                    <div className="p-8 rounded-2xl border" style={{ backgroundColor: activeTheme.cardBg, borderColor: activeTheme.border }}>
                      <h3 className="text-2xl font-bold mb-3" style={{ color: activeTheme.accentText }}>
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
                    <div className="text-6xl lg:text-8xl font-black" style={{ color: activeTheme.accent }}>
                      {activeSlide.statValue}
                    </div>
                    <div className="text-2xl font-bold">{activeSlide.statLabel}</div>
                    <div className="text-base" style={{ color: activeTheme.textMuted }}>
                      {activeSlide.statSubtext}
                    </div>
                  </div>
                  {activeSlide.bullets && (
                    <div className="p-8 rounded-2xl border space-y-4" style={{ backgroundColor: activeTheme.cardBg, borderColor: activeTheme.border }}>
                      <h4 className="text-sm font-bold uppercase tracking-wider" style={{ color: activeTheme.accentText }}>
                        Key Highlights
                      </h4>
                      <ul className="space-y-3 text-lg">
                        {activeSlide.bullets.map((b, i) => (
                          <li key={i} className="flex items-center gap-3">
                            <div className="w-2.5 h-2.5 rounded-full" style={{ backgroundColor: activeTheme.accent }}></div>
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
                  <QuoteIcon className="w-16 h-16 mx-auto opacity-30" style={{ color: activeTheme.accent }} />
                  <blockquote className="text-2xl lg:text-4xl font-medium italic leading-relaxed">
                    “{activeSlide.quote}”
                  </blockquote>
                  <div>
                    <div className="text-xl font-bold" style={{ color: activeTheme.accentText }}>
                      {activeSlide.quoteAuthor}
                    </div>
                    <div className="text-base opacity-75" style={{ color: activeTheme.textMuted }}>
                      {activeSlide.quoteRole}
                    </div>
                  </div>
                </div>
              )}
            </div>

            {/* Slide Footer */}
            <div className="flex items-center justify-between text-sm opacity-60">
              <span>{activeSlide.footerText}</span>
              <span>DocuCraft Deck</span>
            </div>
          </div>

          {/* BOTTOM NAVIGATION CONTROLS */}
          <div className="flex items-center justify-center gap-6 z-20">
            <button
              onClick={() => setActiveSlideIndex((prev) => Math.max(prev - 1, 0))}
              disabled={activeSlideIndex === 0}
              className="p-3 rounded-full bg-white/10 hover:bg-white/20 text-white disabled:opacity-30 transition"
            >
              <ChevronLeft className="w-6 h-6" />
            </button>

            <div className="flex items-center gap-2">
              {slides.map((_, i) => (
                <button
                  key={i}
                  onClick={() => setActiveSlideIndex(i)}
                  className={`w-3 h-3 rounded-full transition ${
                    i === activeSlideIndex ? 'bg-indigo-500 scale-125' : 'bg-white/20 hover:bg-white/40'
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
